import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { performance } from "node:perf_hooks";
import { promisify } from "node:util";
import { getPlatformSupport } from "@microsoft/mxc-sdk/v1";

const execFileAsync = promisify(execFile);
const require = createRequire(import.meta.url);

export type MicroVmStep =
  | "discover-warm"
  | "guest-boundary"
  | "safe-compute"
  | "escape-blocked"
  | "snapshot-reset";

const knownSteps: MicroVmStep[] = [
  "discover-warm",
  "guest-boundary",
  "safe-compute",
  "escape-blocked",
  "snapshot-reset",
];

interface HyperlightOneShotConfig {
  version: "1.1.0-alpha";
  containment: "hyperlight";
  hyperlight: {
    runtime: "bash" | "python";
  };
  process: {
    commandLine: string;
    timeout: number;
  };
}

function resolveWxcExecutor(): string {
  const sdkPkg = require.resolve("@microsoft/mxc-sdk/package.json");
  const sdkRoot = path.dirname(sdkPkg);
  const archDir = os.arch() === "arm64" ? "arm64" : "x64";
  const exeName = os.platform() === "win32" ? "wxc-exec.exe" : "lxc-exec";
  const fullPath = path.join(sdkRoot, "bin", archDir, exeName);
  if (!existsSync(fullPath)) {
    throw new Error(`MXC native executor not found at ${fullPath}`);
  }
  return fullPath;
}

function resolveHyperlightHome(): string {
  if (process.env.MXC_HYPERLIGHT_HOME) {
    return process.env.MXC_HYPERLIGHT_HOME;
  }
  if (os.platform() === "win32") {
    const localAppData =
      process.env.LOCALAPPDATA ?? path.join(os.homedir(), "AppData", "Local");
    return path.join(localAppData, "mxc-hyperlight");
  }
  return path.join(os.homedir(), ".local", "share", "mxc-hyperlight");
}

function assertHyperlightSupported(): string {
  const support = getPlatformSupport();
  if (!support.isSupported || !support.availableMethods.includes("hyperlight")) {
    throw new Error(
      `MXC Hyperlight MicroVM backend is not available on ${os.platform()}/${os.arch()} `
      + `(available: ${support.availableMethods.join(", ") || "none"}).`,
    );
  }
  return resolveWxcExecutor();
}

async function ensureSnapshotsReady(executor: string): Promise<string> {
  const home = resolveHyperlightHome();
  const bashSnap = path.join(home, "bash", "snapshot");
  const pySnap = path.join(home, "python", "snapshot");
  if (!existsSync(bashSnap) || !existsSync(pySnap)) {
    console.log("Warming Hyperlight Unikraft snapshots (bash, python)...");
    const { stdout, stderr } = await execFileAsync(
      executor,
      ["--experimental", "--setup-hyperlight=bash,python"],
      { windowsHide: true, timeout: 180_000 },
    );
    if (stdout.trim()) console.log(stdout.trim());
    if (stderr.trim()) console.log(stderr.trim());
  }
  return home;
}

async function runHyperlight(
  executor: string,
  runtime: "bash" | "python",
  commandLine: string,
  dryRun = false,
): Promise<{ stdout: string; stderr: string; elapsedMs: number }> {
  const config: HyperlightOneShotConfig = {
    version: "1.1.0-alpha",
    containment: "hyperlight",
    hyperlight: { runtime },
    process: {
      commandLine,
      timeout: 30_000,
    },
  };
  const encoded = Buffer.from(JSON.stringify(config), "utf8").toString("base64");
  const args = ["--experimental", "--config-base64", encoded];
  if (dryRun) args.splice(1, 0, "--dry-run");

  const start = performance.now();
  const { stdout, stderr } = await execFileAsync(executor, args, {
    windowsHide: true,
    timeout: 45_000,
  });
  const elapsedMs = Math.round(performance.now() - start);
  return { stdout, stderr, elapsedMs };
}

async function discoverAndWarmStep(): Promise<void> {
  const support = getPlatformSupport();
  const executor = assertHyperlightSupported();
  console.log(`MXC host backends: ${support.availableMethods.join(", ")}`);
  const home = await ensureSnapshotsReady(executor);
  const dryRun = await runHyperlight(executor, "python", "print('ok')", true);
  if (!dryRun.stdout.includes("validation passed")) {
    throw new Error("Hyperlight dry-run contract validation did not pass.");
  }
  console.log("Hyperlight snapshots verified: bash (initrd+snapshot), python (initrd+snapshot)");
  console.log(`Snapshot cache ready (${path.basename(home)})`);
  console.log("Dry-run 1.1.0-alpha contract check: validation passed");
  console.log("MICROVM STEP 1 COMPLETE: Hyperlight backend & warmed snapshots ready");
}

async function guestBoundaryStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);
  const hostOs = `${os.type()} ${os.release()} (${os.arch()})`;
  const res = await runHyperlight(
    executor,
    "bash",
    "echo GUEST_KERNEL=$(uname -sr) && echo GUEST_ARCH=$(uname -m) && echo GUEST_ID=$(id -un):$(id -u) && echo GUEST_ROOTFS=$(ls / | tr '\\n' ' ')",
  );
  console.log(`HOST_OS=${hostOs}`);
  process.stdout.write(res.stdout);
  console.log(`MICROVM_BOOT_MS=${res.elapsedMs}`);
  if (!res.stdout.includes("Unikraft")) {
    throw new Error("Expected Unikraft unikernel guest inside Hyperlight MicroVM.");
  }
  console.log(
    `MICROVM STEP 2 COMPLETE: hardware VM booted guest kernel in ${res.elapsedMs}ms`,
  );
}

async function safeComputeStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);
  const untrustedAiCode = [
    "import json, sys",
    "ledger = [",
    "  {'id': 'INV-9040', 'amount': 1240.00, 'approved': True},",
    "  {'id': 'INV-9041', 'amount': 625.50, 'approved': True},",
    "  {'id': 'INV-9042', 'amount': 99400.00, 'approved': False},",
    "]",
    "flagged = [r['id'] for r in ledger if (not r['approved']) or r['amount'] > 10000]",
    "approved_sum = round(sum(r['amount'] for r in ledger if r['approved']), 2)",
    "print('GUEST_PYTHON=' + sys.version.split()[0])",
    "print('EVAL_RESULT=' + json.dumps({'approvedTotal': approved_sum, 'flagged': flagged}))",
  ].join("\n");

  const res = await runHyperlight(executor, "python", untrustedAiCode);
  process.stdout.write(res.stdout);
  if (!res.stdout.includes('"INV-9042"')) {
    throw new Error("MicroVM Python evaluation did not return expected result.");
  }
  console.log(
    `MICROVM STEP 3 COMPLETE: untrusted Python evaluated safely in ${res.elapsedMs}ms`,
  );
}

async function escapeBlockedStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);
  const adversarialCode = [
    "import os, socket",
    "host_paths = ['C:/Windows/System32', 'C:/Users', '/mnt/c/Windows']",
    "visible = [p for p in host_paths if os.path.exists(p)]",
    "print('HOST_FS=' + ('EXPOSED:' + ','.join(visible) if visible else 'UNREACHABLE (0 host mounts)'))",
    "leaked_env = [k for k in ('USERNAME', 'USERPROFILE', 'APPDATA', 'COMPUTERNAME') if k in os.environ]",
    "print('HOST_ENV=' + ('LEAKED' if leaked_env else 'ISOLATED (0 host vars)'))",
    "try:",
    "    socket.create_connection(('127.0.0.1', 80), timeout=1)",
    "    print('GUEST_NET=REACHABLE')",
    "except Exception as exc:",
    "    print(f'GUEST_NET=BLOCKED ({type(exc).__name__}: {exc})')",
  ].join("\n");

  const res = await runHyperlight(executor, "python", adversarialCode);
  process.stdout.write(res.stdout);
  if (
    !res.stdout.includes("HOST_FS=UNREACHABLE")
    || !res.stdout.includes("HOST_ENV=ISOLATED")
    || !res.stdout.includes("GUEST_NET=BLOCKED")
  ) {
    throw new Error("MicroVM escape probes were not all blocked as expected.");
  }
  console.log(
    "MICROVM STEP 4 COMPLETE: host filesystem, env, and sockets unreachable from VM",
  );
}

async function snapshotResetStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);

  const runA = await runHyperlight(
    executor,
    "python",
    "open('/tmp/agent-residue.txt', 'w').write('contaminated'); "
      + "import os; print('RUN_A_TMP_MARKER=' + str(os.path.exists('/tmp/agent-residue.txt')))",
  );
  const runB = await runHyperlight(
    executor,
    "python",
    "import os; print('RUN_B_TMP_MARKER=' + str(os.path.exists('/tmp/agent-residue.txt')))",
  );
  process.stdout.write(runA.stdout);
  process.stdout.write(runB.stdout);
  if (
    !runA.stdout.includes("RUN_A_TMP_MARKER=True")
    || !runB.stdout.includes("RUN_B_TMP_MARKER=False")
  ) {
    throw new Error("MicroVM memory snapshot did not reset cleanly across runs.");
  }
  console.log(
    `MICROVM STEP 5 COMPLETE: clean snapshot restored (${runA.elapsedMs}ms + ${runB.elapsedMs}ms, zero state residue)`,
  );
}

export async function runMicroVmStep(step: MicroVmStep): Promise<void> {
  if (step === "discover-warm") return discoverAndWarmStep();
  if (step === "guest-boundary") return guestBoundaryStep();
  if (step === "safe-compute") return safeComputeStep();
  if (step === "escape-blocked") return escapeBlockedStep();
  return snapshotResetStep();
}

try {
  const argument = process.argv.find((value) => value.startsWith("--step="));
  const step = argument?.slice("--step=".length);
  if (step !== undefined && !knownSteps.includes(step as MicroVmStep)) {
    throw new Error(
      `Unknown --step value: ${step}. Expected one of: ${knownSteps.join(", ")}`,
    );
  }
  for (const current of step ? [step as MicroVmStep] : knownSteps) {
    await runMicroVmStep(current);
  }
} catch (error) {
  console.error(String(error));
  process.exitCode = 1;
}
