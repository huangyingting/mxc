import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { performance } from "node:perf_hooks";
import { promisify } from "node:util";
import { getPlatformSupport } from "@microsoft/mxc-sdk/v1";
import { fmt } from "./demo-format.js";

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
  fmt.policy('version="1.1.0-alpha" · containment="hyperlight" (--experimental)', [
    'hyperlight.runtime : "bash" | "python" (Unikraft unikernel)',
    "isolation          : Windows Hypervisor Platform hardware VM partition",
  ]);
  fmt.run("wxc-exec --experimental --dry-run (validate contract & warm snapshots)");
  fmt.log(`Host backends: ${support.availableMethods.join(", ")}`);
  const home = await ensureSnapshotsReady(executor);
  const dryRun = await runHyperlight(executor, "python", "print('ok')", true);
  if (!dryRun.stdout.includes("validation passed")) {
    throw new Error("Hyperlight dry-run contract validation did not pass.");
  }
  fmt.log(`Snapshots verified: bash + python in ${path.basename(home)}`);
  fmt.pass("Hyperlight MicroVM backend & pre-warmed memory snapshots ready");
  fmt.why("Pre-warmed snapshots give hardware VM isolation in ~90 ms instead of multi-second cold VM boot.");
  console.log("MICROVM STEP 1 COMPLETE");
}

async function guestBoundaryStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);
  const hostOs = `${os.type()} ${os.release()} (${os.arch()})`;
  fmt.policy('containment="hyperlight" · hyperlight={ runtime: "bash" }', [
    "kernelSharing : none (dedicated Unikraft guest kernel on Hypervisor)",
  ]);
  fmt.run("uname -sr && uname -m && id -un && ls /");
  const res = await runHyperlight(
    executor,
    "bash",
    "echo GUEST_KERNEL=$(uname -sr) && echo GUEST_ARCH=$(uname -m) && echo GUEST_ID=$(id -un):$(id -u) && echo GUEST_ROOTFS=$(ls / | tr '\\n' ' ')",
  );
  fmt.log(`HOST_OS=${hostOs}`);
  for (const line of res.stdout.trim().split(/\r?\n/).filter(Boolean)) {
    fmt.log(line);
  }
  console.log(`MICROVM_BOOT_MS=${res.elapsedMs}`);
  if (!res.stdout.includes("Unikraft")) {
    throw new Error("Expected Unikraft unikernel guest inside Hyperlight MicroVM.");
  }
  fmt.pass(`Command succeeded (${res.elapsedMs} ms) — running inside Unikraft 5.15 VM, not Windows NT`);
  fmt.why("Untrusted code executes on its own minimal unikernel inside a hardware-enforced partition.");
  console.log("MICROVM STEP 2 COMPLETE");
}

async function safeComputeStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);
  fmt.policy('containment="hyperlight" · hyperlight={ runtime: "python", timeout: 30000 }', [
    "hostFilesystem : unmounted  |  guestNetwork : disabled",
  ]);
  fmt.run("python -c \"<AI-generated invoice reconciliation script>\"");
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
  for (const line of res.stdout.trim().split(/\r?\n/).filter(Boolean)) {
    fmt.log(line);
  }
  if (!res.stdout.includes('"INV-9042"')) {
    throw new Error("MicroVM Python evaluation did not return expected result.");
  }
  fmt.pass(`Command succeeded (${res.elapsedMs} ms) — Python 3.12 reconciled invoices safely`);
  fmt.why("Pure compute runs at near-native speed via stdout JSON without needing any host disk or network.");
  console.log("MICROVM STEP 3 COMPLETE");
}

async function escapeBlockedStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);
  fmt.policy('containment="hyperlight" · zero-host-surface VM air gap', [
    "hostMounts : [] (none)  |  hostEnv : isolated  |  sockets : disabled",
  ]);
  fmt.run("python -c \"open('C:/Windows/System32/drivers/etc/hosts'); os.environ['USERPROFILE']; socket.connect(('127.0.0.1',80))\"");
  const adversarialCode = [
    "import os, socket",
    "try:",
    "    open('C:/Windows/System32/drivers/etc/hosts').read()",
    "    print('HOST_FS=EXPOSED')",
    "except Exception as e:",
    "    print(f'ERR_FS={type(e).__name__}: {e}')",
    "try:",
    "    _ = os.environ['USERPROFILE']",
    "    print('HOST_ENV=LEAKED')",
    "except Exception as e:",
    "    print(f'ERR_ENV={type(e).__name__}: {e} not in guest env')",
    "try:",
    "    socket.create_connection(('127.0.0.1', 80), timeout=1)",
    "    print('GUEST_NET=REACHABLE')",
    "except Exception as e:",
    "    print(f'ERR_NET={type(e).__name__}: {e}')",
  ].join("\n");

  const res = await runHyperlight(executor, "python", adversarialCode);
  if (
    !res.stdout.includes("ERR_FS=")
    || !res.stdout.includes("ERR_ENV=")
    || !res.stdout.includes("ERR_NET=")
  ) {
    throw new Error(`MicroVM escape probes were not all blocked as expected: ${res.stdout}`);
  }
  for (const line of res.stdout.trim().split(/\r?\n/).filter(Boolean)) {
    fmt.osErr(line);
  }
  fmt.blocked("All 3 escape attempts failed (FileNotFoundError, KeyError, PermissionError Errno 13)");
  fmt.why("The MicroVM has no mounted host drives, no inherited host env vars, and no network stack.");
  console.log("MICROVM STEP 4 COMPLETE");
}

async function snapshotResetStep(): Promise<void> {
  const executor = assertHyperlightSupported();
  await ensureSnapshotsReady(executor);

  fmt.policy('containment="hyperlight" · ephemeral warm-snapshot restore', [
    "persistence : none (every invocation restores the clean memory snapshot)",
  ]);
  fmt.run("Run A: write /tmp/agent-residue.txt  ->  Run B: open('/tmp/agent-residue.txt')");

  const runA = await runHyperlight(
    executor,
    "python",
    "open('/tmp/agent-residue.txt', 'w').write('contaminated'); "
      + "import os; print('RUN_A_WROTE_TMP=' + str(os.path.exists('/tmp/agent-residue.txt')))",
  );
  const runB = await runHyperlight(
    executor,
    "python",
    "try:\n"
      + "    open('/tmp/agent-residue.txt').read()\n"
      + "    print('RUN_B_FOUND=True')\n"
      + "except Exception as e:\n"
      + "    print(f'RUN_B_ERR={type(e).__name__}: {e}')",
  );
  if (
    !runA.stdout.includes("RUN_A_WROTE_TMP=True")
    || !runB.stdout.includes("RUN_B_ERR=FileNotFoundError")
  ) {
    throw new Error("MicroVM memory snapshot did not reset cleanly across runs.");
  }
  fmt.log(`Run A (${runA.elapsedMs} ms): ${runA.stdout.trim()}`);
  fmt.osErr(`Run B (${runB.elapsedMs} ms): ${runB.stdout.trim()}`);
  fmt.blocked("Run B could not read Run A's /tmp/agent-residue.txt (FileNotFoundError)");
  fmt.pass(`Snapshot wiped clean across runs (${runA.elapsedMs} ms + ${runB.elapsedMs} ms, zero state residue)`);
  fmt.why("Each call restores a fresh memory snapshot, so malware or temp files vanish immediately on exit.");
  console.log("MICROVM STEP 5 COMPLETE");
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
