import { existsSync } from "node:fs";
import { createServer } from "node:http";
import { once } from "node:events";
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
  network: NonNullable<ContainerRequest["network"]> = policy.network,
): ContainerRequest {
  return {
    command,
    containment: { type: "process" },
    filesystem: resolveFilesystemPolicy(policy),
    network,
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

async function fileAllowedStep(policy: DemoPolicy): Promise<void> {
  const inputPath = resolvePolicyPath("input/request.txt");
  const outputDir = resolvePolicyPath(policy.filesystem.readwritePaths[0]);
  const outputPath = path.join(outputDir, "report.txt");
  if (!existsSync(inputPath)) {
    throw new Error("Demo input file must exist before execution.");
  }
  await mkdir(outputDir, { recursive: true });
  await rm(outputPath, { force: true });
  const request = baseRequest(policy, allowedCommand(), {
    MXC_INPUT: inputPath,
    MXC_OUTPUT: outputPath,
  });
  assertSupported(request);
  assertCompleted("allowed file operation", await run(request));
  const report = await readFile(outputPath, "utf8");
  if (!report.includes("Summarize the approved input")
      || !report.includes("access-controlled write")) {
    throw new Error("The allowed workload did not produce the expected report.");
  }
  console.log(`STEP 2 COMPLETE: report written to ${outputPath}`);
}

async function fileDeniedStep(policy: DemoPolicy): Promise<void> {
  const deniedPath = resolvePolicyPath(policy.filesystem.deniedPaths[0]);
  if (!existsSync(deniedPath)) {
    throw new Error("Protected demo file must exist before execution.");
  }
  const request = baseRequest(policy, deniedCommand(), {
    MXC_DENIED_FILE: deniedPath,
  });
  assertSupported(request);
  assertCompleted("protected file denial", await run(request));
  console.log("STEP 3 COMPLETE: protected host file remained inaccessible");
}

async function networkStep(
  policy: DemoPolicy,
  hostLoopback: "allow" | "deny",
): Promise<void> {
  const server = createServer((_request, response) => {
    response.writeHead(200, { "content-type": "text/plain" });
    response.end("MXC_LOOPBACK_OK");
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  if (address === null || typeof address === "string") {
    server.close();
    throw new Error("Could not determine the local endpoint port.");
  }

  try {
    const endpoint = `http://127.0.0.1:${address.port}`;
    const expected = hostLoopback === "allow";
    const command = expected
      ? `cmd.exe /d /s /c curl.exe -sS --noproxy "*" --max-time 5 "${endpoint}" `
        + '| findstr /C:"MXC_LOOPBACK_OK" >nul && '
        + '(echo ALLOW: approved loopback endpoint reached & exit /b 0) '
        + "|| (echo ERROR: loopback endpoint did not return expected data 1>&2 & exit /b 1)"
      : `cmd.exe /d /s /c curl.exe -sS --noproxy "*" --max-time 2 "${endpoint}" >nul 2>&1 && `
        + "(echo SECURITY ERROR: denied loopback endpoint was reachable 1>&2 & exit /b 1) "
        + "|| (echo DENY: host loopback remained blocked & exit /b 0)";
    const request = baseRequest(
      policy,
      command,
      { MXC_ENDPOINT: endpoint },
      {
        egress: { default: "deny" },
        ingress: {
          default: "deny",
          hostLoopback,
        },
      },
    );
    request.containment = {
      type: "processcontainer",
      config: {
        capabilities: ["internetClient"],
      },
    };
    assertSupported(request);
    assertCompleted(
      `network ${hostLoopback} operation`,
      await run(request),
    );
    console.log(
      `STEP ${expected ? 4 : 5} COMPLETE: host loopback ${hostLoopback}`,
    );
  } finally {
    await new Promise<void>((resolve, reject) => {
      server.close((error) => error ? reject(error) : resolve());
    });
  }
}

export async function runProcessStep(
  step: "discover" | "file-allow" | "file-deny" | "network-allow" | "network-deny",
): Promise<void> {
  const policy = await loadPolicy();
  if (step === "discover") {
    const request = baseRequest(policy, allowedCommand(), {
      MXC_INPUT: resolvePolicyPath("input/request.txt"),
      MXC_OUTPUT: resolvePolicyPath("output/report.txt"),
    });
    assertSupported(request);
    console.log("STEP 1 COMPLETE: platform and ProcessContainer policy discovered");
    return;
  }
  if (step === "file-allow") return fileAllowedStep(policy);
  if (step === "file-deny") return fileDeniedStep(policy);
  if (step === "network-allow") return networkStep(policy, "allow");
  return networkStep(policy, "deny");
}

try {
  const argument = process.argv.find((value) => value.startsWith("--step="));
  const step = argument?.slice("--step=".length);
  const knownSteps = [
    "discover",
    "file-allow",
    "file-deny",
    "network-allow",
    "network-deny",
  ] as const;
  if (step === undefined) {
    for (const current of knownSteps) {
      await runProcessStep(current);
    }
  } else {
    if (!knownSteps.includes(step as typeof knownSteps[number])) {
      throw new Error(`Unknown process demo step: ${step}`);
    }
    await runProcessStep(step as typeof knownSteps[number]);
  }
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
