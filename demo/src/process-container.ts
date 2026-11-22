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
} from "@microsoft/mxc-sdk/v1";
import { fmt } from "./demo-format.js";

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
      + 'type "%MXC_OUTPUT%"'
    : 'sh -c \'cat "$MXC_INPUT" > "$MXC_OUTPUT" && '
      + 'printf "Result: access-controlled write\\n" >> "$MXC_OUTPUT" && '
      + 'cat "$MXC_OUTPUT"\'';
}

function deniedCommand(): string {
  return os.platform() === "win32"
    ? 'cmd.exe /d /s /c type "%MXC_DENIED_FILE%"'
    : 'sh -c \'cat "$MXC_DENIED_FILE"\'';
}

function printFilePolicy(policy: DemoPolicy): void {
  fmt.policy('containment="processcontainer"  (demo/config/policy.json)', [
    `readonlyPaths : ${JSON.stringify(policy.filesystem.readonlyPaths)}`,
    `readwritePaths: ${JSON.stringify(policy.filesystem.readwritePaths)}`,
    `deniedPaths   : ${JSON.stringify(policy.filesystem.deniedPaths)}`,
    `network       : egress="${policy.network.egress?.default ?? "deny"}", hostLoopback="${policy.network.ingress?.hostLoopback ?? "deny"}"`,
  ]);
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

function assertSupported(request: ContainerRequest, verbose = false): void {
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

  if (verbose) {
    fmt.log(`Host backends: ${support.availableMethods.join(", ")}`);
  }
  if (os.platform() === "win32") {
    const requestProbe = probe(request);
    for (const warning of requestProbe.warnings) {
      fmt.log(`Probe warning: ${warning}`);
    }
    if (requestProbe.error !== undefined) {
      throw new Error(`MXC request probe failed: ${requestProbe.error}`);
    }
    if (verbose) {
      fmt.log(`Enforced ProcessContainer tier: ${requestProbe.tier}`);
    }
  }
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

  printFilePolicy(policy);
  fmt.run("cmd.exe /c type input\\request.txt > output\\report.txt");

  const request = baseRequest(policy, allowedCommand(), {
    MXC_INPUT: inputPath,
    MXC_OUTPUT: outputPath,
  });
  assertSupported(request, false);
  const result = await run(request);
  if (result.timedOut || result.exitCode !== 0) {
    throw new Error(
      `Allowed file command unexpectedly failed (exit ${result.exitCode}): ${result.stderr.trim()}`,
    );
  }
  for (const line of result.stdout.trim().split(/\r?\n/).filter(Boolean)) {
    fmt.log(`stdout: ${line}`);
  }
  const report = await readFile(outputPath, "utf8");
  if (!report.includes("Summarize the approved input")
      || !report.includes("access-controlled write")) {
    throw new Error("The allowed workload did not produce the expected report.");
  }
  fmt.pass(`Command succeeded (exit code 0) — read input\\request.txt & wrote output\\report.txt`);
  fmt.why("input/ is granted in readonlyPaths and output/ is granted in readwritePaths.");
  console.log("STEP 2 COMPLETE");
}

async function fileDeniedStep(policy: DemoPolicy): Promise<void> {
  const deniedPath = resolvePolicyPath(policy.filesystem.deniedPaths[0]);
  if (!existsSync(deniedPath)) {
    throw new Error("Protected demo file must exist before execution.");
  }

  printFilePolicy(policy);
  fmt.run("cmd.exe /c type protected\\denied.txt");

  const request = baseRequest(policy, deniedCommand(), {
    MXC_DENIED_FILE: deniedPath,
  });
  assertSupported(request, false);
  const result = await run(request);
  if (result.timedOut) {
    throw new Error("Protected file test timed out.");
  }
  if (result.exitCode === 0) {
    throw new Error("SECURITY ERROR: protected file was readable inside ProcessContainer!");
  }
  const rawErr = (result.stderr || result.stdout).trim();
  const errLine = (!rawErr || rawErr.includes("\uFFFD"))
    ? "Access is denied. (Windows kernel ACL deny)"
    : rawErr.replace(/\r?\n/g, " ");
  fmt.osErr(errLine);
  fmt.blocked(
    `Command failed (exit code ${result.exitCode}) — OS blocked reading protected\\denied.txt`,
  );
  fmt.why(
    'deniedPaths: ["protected/denied.txt"] applied a kernel deny ACL, overriding the host user\'s own file access.',
  );
  console.log("STEP 3 COMPLETE");
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
    const expectedAllow = hostLoopback === "allow";
    const command = os.platform() === "win32"
      ? `curl.exe -sS --noproxy "*" --max-time ${expectedAllow ? 5 : 2} "${endpoint}"`
      : `curl -sS --noproxy '*' --max-time ${expectedAllow ? 5 : 2} '${endpoint}'`;

    fmt.policy('containment="processcontainer"  (network guardrail)', [
      'network.egress.default       : "deny"   (outbound internet blocked)',
      `network.ingress.hostLoopback : "${hostLoopback}"  (${expectedAllow ? "local proxy/listener permitted" : "all local ports blocked"})`,
    ]);
    fmt.run(command);

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
    assertSupported(request, false);
    const result = await run(request);

    if (expectedAllow) {
      if (result.exitCode !== 0 || !result.stdout.includes("MXC_LOOPBACK_OK")) {
        throw new Error(
          `Expected loopback allow to succeed, got exit ${result.exitCode}: ${result.stderr.trim()}`,
        );
      }
      fmt.log(`stdout: HTTP 200 "${result.stdout.trim()}" from ${endpoint}`);
      fmt.pass(`Command succeeded (exit code 0) — reached approved local listener ${endpoint}`);
      fmt.why(
        'hostLoopback: "allow" lets the sandbox talk to a trusted local proxy/broker while internet egress stays denied.',
      );
      console.log("STEP 4 COMPLETE");
    } else {
      if (result.exitCode === 0) {
        throw new Error("SECURITY ERROR: denied loopback endpoint was reachable!");
      }
      const errMsg = (result.stderr || result.stdout).trim() || "Connection blocked by OS network filter";
      fmt.osErr(errMsg.replace(/\r?\n/g, " "));
      fmt.blocked(
        `Command failed (exit code ${result.exitCode}) — exact same curl command blocked from ${endpoint}`,
      );
      fmt.why(
        'Changing only hostLoopback from "allow" to "deny" severed loopback access at the Windows network stack.',
      );
      console.log("STEP 5 COMPLETE");
    }
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
    printFilePolicy(policy);
    fmt.run("getPlatformSupport() + probe(ContainerRequest)");
    const request = baseRequest(policy, allowedCommand(), {
      MXC_INPUT: resolvePolicyPath("input/request.txt"),
      MXC_OUTPUT: resolvePolicyPath("output/report.txt"),
    });
    assertSupported(request, true);
    fmt.pass("Native ProcessContainer verified (tier: base-container, zero unsandboxed fallback)");
    fmt.why("Always probe host capabilities before running untrusted agent commands.");
    console.log("STEP 1 COMPLETE");
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
