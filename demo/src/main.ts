import { existsSync } from "node:fs";
import { mkdir, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  MxcError,
  getPlatformSupport,
  probe,
  run,
  type ContainerRequest,
  type ExecutionResult,
} from "@microsoft/mxc-sdk/v1";

interface DemoPolicy {
  filesystem: {
    readonlyPaths: string[];
    readwritePaths: string[];
    deniedPaths: string[];
  };
  network: NonNullable<ContainerRequest["network"]>;
  ui: NonNullable<ContainerRequest["ui"]>;
}

const demoRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const policyPath = path.join(demoRoot, "config", "policy.json");

function resolvePolicyPath(relativePath: string): string {
  const resolved = path.resolve(demoRoot, relativePath);
  const relative = path.relative(demoRoot, resolved);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Policy path escapes the demo directory: ${relativePath}`);
  }
  return resolved;
}

async function loadPolicy(): Promise<DemoPolicy> {
  return JSON.parse(await readFile(policyPath, "utf8")) as DemoPolicy;
}

function resolveFilesystemPolicy(
  policy: DemoPolicy,
): NonNullable<ContainerRequest["filesystem"]> {
  const { filesystem } = policy;
  return {
    readonlyPaths: filesystem.readonlyPaths.map(resolvePolicyPath),
    readwritePaths: filesystem.readwritePaths.map(resolvePolicyPath),
    deniedPaths: filesystem.deniedPaths.map(resolvePolicyPath),
  };
}

function allowedCommand(): string {
  return os.platform() === "win32"
    ? 'cmd.exe /d /s /c type "%MXC_INPUT%" > "%MXC_OUTPUT%" && '
      + 'echo Result: access-controlled write >> "%MXC_OUTPUT%" && '
      + "echo ALLOW: input read and output write succeeded"
    : "sh -c 'cat \"$MXC_INPUT\" > \"$MXC_OUTPUT\" && "
      + "printf \"Result: access-controlled write\\n\" >> \"$MXC_OUTPUT\" && "
      + "echo \"ALLOW: input read and output write succeeded\"'";
}

function deniedCommand(): string {
  return os.platform() === "win32"
    ? 'cmd.exe /d /s /c type "%MXC_DENIED_FILE%" >nul 2>&1 & '
      + "if errorlevel 1 (echo DENY: protected file remained inaccessible & exit /b 0) "
      + "else (echo SECURITY ERROR: protected file was readable 1>&2 & exit /b 1)"
    : "sh -c 'if cat \"$MXC_DENIED_FILE\" >/dev/null 2>&1; then "
      + "echo \"SECURITY ERROR: protected file was readable\" >&2; exit 1; "
      + "else echo \"DENY: protected file remained inaccessible\"; fi'";
}

function baseRequest(
  policy: DemoPolicy,
  command: string,
  environment: Record<string, string>,
): ContainerRequest {
  return {
    command,
    containment: { type: "process" },
    filesystem: resolveFilesystemPolicy(policy),
    network: policy.network,
    ui: policy.ui,
    workingDirectory: resolvePolicyPath(policy.filesystem.readwritePaths[0]),
    environment,
    inheritDefaultEnvironment: true,
    timeoutMs: 30_000,
  };
}

function assertSupported(request: ContainerRequest): void {
  const support = getPlatformSupport();
  if (!support.isSupported) {
    const reasons = Object.entries(support.unavailableReasons ?? {})
      .map(([backend, reason]) => `${backend}: ${reason}`)
      .join("\n");
    throw new Error(
      `MXC is not available on ${os.platform()}/${os.arch()}.`
      + `${support.reason ? ` ${support.reason}` : ""}`
      + `${reasons ? `\n${reasons}` : ""}`,
    );
  }

  console.log(`MXC host backends: ${support.availableMethods.join(", ")}`);
  if (os.platform() === "win32") {
    const requestProbe = probe(request);
    for (const warning of requestProbe.warnings) {
      console.warn(`MXC probe warning: ${warning}`);
    }
    if (requestProbe.error !== undefined) {
      throw new Error(`MXC request probe failed: ${requestProbe.error}`);
    }
    console.log(`MXC ProcessContainer tier: ${requestProbe.tier}`);
  }
}

function assertCompleted(label: string, result: ExecutionResult): void {
  for (const warning of result.warnings) {
    console.warn(`${label} warning: ${warning}`);
  }
  if (result.timedOut) {
    throw new Error(`${label} timed out`);
  }
  if (result.exitCode !== 0) {
    throw new Error(
      `${label} failed with exit code ${result.exitCode}: ${result.stderr.trim()}`,
    );
  }
  process.stdout.write(result.stdout);
  process.stderr.write(result.stderr);
}

async function main(): Promise<void> {
  const policy = await loadPolicy();
  const inputPath = resolvePolicyPath("input/request.txt");
  const outputDir = resolvePolicyPath(policy.filesystem.readwritePaths[0]);
  const outputPath = path.join(outputDir, "report.txt");
  const deniedPath = resolvePolicyPath(policy.filesystem.deniedPaths[0]);

  if (!existsSync(inputPath) || !existsSync(deniedPath)) {
    throw new Error("Demo input and protected files must exist before execution.");
  }
  await mkdir(outputDir, { recursive: true });
  await rm(outputPath, { force: true });

  const allowedRequest = baseRequest(policy, allowedCommand(), {
    MXC_INPUT: inputPath,
    MXC_OUTPUT: outputPath,
  });
  assertSupported(allowedRequest);
  assertCompleted("allowed file operation", await run(allowedRequest));

  const report = await readFile(outputPath, "utf8");
  if (!report.includes("Summarize the approved input")
      || !report.includes("access-controlled write")) {
    throw new Error("The allowed workload did not produce the expected report.");
  }

  const deniedRequest = baseRequest(policy, deniedCommand(), {
    MXC_DENIED_FILE: deniedPath,
  });
  assertCompleted("protected file denial", await run(deniedRequest));
  console.log(`MXC demo report: ${outputPath}`);
}

try {
  await main();
} catch (error) {
  if (error instanceof MxcError) {
    console.error(`MXC error [${error.code}]: ${error.message}`);
  } else if (error instanceof Error) {
    console.error(error.message);
  } else {
    console.error(error);
  }
  process.exitCode = 1;
}
