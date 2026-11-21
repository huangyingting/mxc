import { execFileSync, spawn } from "node:child_process";
import { copyFile, readFile } from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  deprovisionContainer,
  provisionContainer,
  spawnInContainerWithPty,
  startContainer,
  stopContainer,
} from "@microsoft/mxc-sdk/v1";

const demoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const repoRoot = path.resolve(demoRoot, "..");
const htmlPath = path.join(repoRoot, "dist", "html", "index.html");
const port = Number.parseInt(process.env.MXC_DEMO_PORT ?? "8787", 10);
const processScript = path.join(demoRoot, "dist", "process-container.js");
const microvmScript = path.join(demoRoot, "dist", "microvm.js");
const clients = new Set();
const processSteps = [
  "discover",
  "file-allow",
  "file-deny",
  "network-allow",
  "network-deny",
];
const microvmSteps = [
  "discover-warm",
  "guest-boundary",
  "safe-compute",
  "escape-blocked",
  "snapshot-reset",
];
let active = null;
let busy = false;
let ptyState = null;
let currentDemo = null;
let processStepIndex = 0;
let ptyStepIndex = 0;

const hostMeta = {
  username: os.userInfo().username,
  sessionId: execFileSync(
    process.platform === "win32" ? "powershell.exe" : "sh",
    process.platform === "win32"
      ? [
          "-NoProfile",
          "-NonInteractive",
          "-Command",
          "(Get-Process -Id $PID).SessionId",
        ]
      : ["-c", "echo unavailable"],
    { encoding: "utf8" },
  ).trim(),
};

function send(value) {
  const payload = `data: ${JSON.stringify(value)}\n\n`;
  for (const client of clients) client.write(payload);
}

function beginStep(demo, step) {
  busy = true;
  currentDemo = demo;
  send({ kind: "step-start", demo, step });
}

function finishStep(demo, step, code = 0) {
  send({ kind: "step-done", demo, step, code });
  busy = false;
}

function waitForMarker(marker, timeoutMs = 30_000) {
  return new Promise((resolve, reject) => {
    let output = "";
    const timer = setTimeout(() => {
      ptyState?.terminal.output.off("data", listener);
      reject(new Error(`PTY step timed out waiting for ${marker}`));
    }, timeoutMs);
    const listener = (chunk) => {
      output += String(chunk);
      const match = output.match(marker);
      if (match !== null) {
        clearTimeout(timer);
        ptyState?.terminal.output.off("data", listener);
        resolve(match[0]);
      }
    };
    ptyState?.terminal.output.on("data", listener);
  });
}

function writePtyCommand(command) {
  ptyState?.terminal.input.write(`${command}\r`);
}

async function cleanupPty() {
  if (ptyState === null) return;
  const state = ptyState;
  ptyState = null;
  if (state.started) {
    try {
      await stopContainer(state.containerId);
    } catch {
      // Deprovision still removes the isolated agent user.
    }
  }
  await deprovisionContainer(state.containerId);
  state.terminal.dispose();
  send({ kind: "output", text: "\nPTY session stopped\nPTY session deprovisioned\n" });
}

async function provisionPty() {
  if (ptyState !== null) throw new Error("The PTY Session is already provisioned.");
  const provisioned = await provisionContainer({
    containment: "isolation_session",
    appId: "MxcLearningKit.WebPtyDemo",
    network: {
      egress: { default: "allow" },
      ingress: { default: "allow", hostLoopback: "allow" },
    },
  });
  const workspace = provisioned.metadata?.ephemeralWorkspacePath;
  if (workspace === undefined) {
    await deprovisionContainer(provisioned.containerId);
    throw new Error("MXC did not return an IsolationSession workspace.");
  }

  await Promise.all([
    copyFile(
      path.join(demoRoot, "input", "incident-request.json"),
      path.join(workspace, "incident-request.json"),
    ),
    copyFile(
      path.join(demoRoot, "input", "incident-approval.json"),
      path.join(workspace, "incident-approval.json"),
    ),
  ]);
  await startContainer(provisioned.containerId);
  const terminal = await spawnInContainerWithPty(
    provisioned.containerId,
    {
      command:
        'powershell.exe -NoLogo -NoProfile -NoExit -Command "Remove-Module PSReadLine -ErrorAction SilentlyContinue"',
      workingDirectory: workspace,
      inheritDefaultEnvironment: true,
      timeoutMs: 0,
    },
    { size: { rows: 30, columns: 120 } },
  );
  ptyState = {
    containerId: provisioned.containerId,
    workspace,
    terminal,
    started: true,
    ended: false,
  };
  const promptReady = waitForMarker(/PS [^\r\n>]+>/, 15_000).catch(() => {});
  terminal.output.on("data", (chunk) => {
    send({ kind: "output", text: String(chunk) });
  });
  send({
    kind: "phase",
    name: "provisioned",
    user: provisioned.metadata?.agentUserName ?? "",
    workspace,
    text: "PTY incident staged: INC-1042, read-only diagnostics, operator approval attached\nPTY connected to isolated session: powershell.exe -NoLogo -NoProfile; terminal=120x30\n",
  });
  await promptReady;
  ptyStepIndex = 1;
  finishStep("pty", 1);
}

async function runPtyStep(step) {
  if (step === 1) {
    await provisionPty();
    return;
  }
  if (ptyState === null) throw new Error("Run PTY step 1 first.");
  const state = ptyState;

  if (step === 2) {
    const marker = /__WEB_IDENTITY_VAL__=([^|\r\n]+)\|(\d+)/;
    const pending = waitForMarker(marker);
    writePtyCommand(
      'Write-Output ("__WEB_IDENTITY_" + "VAL__=$env:USERNAME|$((Get-Process -Id $PID).SessionId)")',
    );
    const result = await pending;
    const match = result.match(marker);
    if (match) {
      send({ kind: "identity", user: match[1].trim(), sessionId: match[2] });
    }
  } else if (step === 3) {
    const marker = /__WEB_APPROVAL_VAL__=([^|\r\n]+)\|(read-only-diagnostics)/;
    const pending = waitForMarker(marker);
    writePtyCommand(
      '$request=Get-Content .\\incident-request.json -Raw | ConvertFrom-Json; '
      + '$approval=Get-Content .\\incident-approval.json -Raw | ConvertFrom-Json; '
      + 'Write-Output ("__WEB_APPROVAL_" + "VAL__=$($request.ticket)|$($approval.scope)")',
    );
    const result = await pending;
    const match = result.match(marker);
    if (match) {
      send({ kind: "approval", ticket: match[1].trim(), scope: match[2] });
    }
  } else if (step === 4) {
    const marker = /__WEB_SERVICES_VAL__=([^|\r\n]+)\|([^\r\n]+)/;
    const pending = waitForMarker(marker);
    writePtyCommand(
      '$spooler=Get-Service Spooler;$time=Get-Service W32Time; '
      + 'Write-Output ("__WEB_SERVICES_" + "VAL__=$($spooler.Status)|$($time.Status)")',
    );
    const result = await pending;
    const match = result.match(marker);
    if (match) {
      send({ kind: "services", spooler: match[1].trim(), w32time: match[2].trim() });
    }
  } else if (step === 5) {
    const pending = waitForMarker(/__WEB_EVIDENCE_VAL__=written/);
    writePtyCommand(
      '$request=Get-Content .\\incident-request.json -Raw | ConvertFrom-Json; '
      + '$approval=Get-Content .\\incident-approval.json -Raw | ConvertFrom-Json; '
      + '$spooler=Get-Service Spooler;$time=Get-Service W32Time; '
      + '$evidence=@("ticket=$($request.ticket)","scope=$($approval.scope)",'
      + '"operator=$env:USERNAME","session_id=$([Diagnostics.Process]::GetCurrentProcess().SessionId)",'
      + '"service=Spooler,status=$($spooler.Status)","service=W32Time,status=$($time.Status)"); '
      + "$evidence | Set-Content .\\triage-evidence.txt -Encoding utf8; "
      + 'Write-Output ("__WEB_EVIDENCE_" + "VAL__=written")',
    );
    await pending;
    const evidence = await readFile(
      path.join(state.workspace, "triage-evidence.txt"),
      "utf8",
    );
    if (!evidence.includes("ticket=INC-1042")) {
      throw new Error("Host could not validate the PTY evidence.");
    }
    send({ kind: "evidence", verified: true, content: evidence.trim() });
  } else if (step === 6) {
    const resultPromise = state.terminal.wait();
    state.terminal.input.end("exit\r");
    const result = await resultPromise;
    state.ended = true;
    await stopContainer(state.containerId);
    await deprovisionContainer(state.containerId);
    state.terminal.dispose();
    ptyState = null;
    send({ kind: "output", text: "\nPTY session stopped\nPTY session deprovisioned\n" });
    if (result.exitCode !== 0 && !result.timedOut) {
      throw new Error(`PTY exited with code ${result.exitCode}`);
    }
  } else {
    throw new Error(`Unknown PTY step: ${step}`);
  }
  ptyStepIndex = step;
  finishStep("pty", step);
}

async function runProcessStep(step) {
  if (step < 1 || step > processSteps.length) {
    throw new Error(`Unknown ProcessContainer step: ${step}`);
  }
  await new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [processScript, `--step=${processSteps[step - 1]}`],
      { cwd: demoRoot, env: { ...process.env }, windowsHide: true },
    );
    active = child;
    child.stdout.on("data", (chunk) => send({ kind: "output", text: chunk.toString() }));
    child.stderr.on("data", (chunk) => send({ kind: "output", text: chunk.toString() }));
    child.on("error", reject);
    child.on("close", (code) => {
      active = null;
      if (code === 0) resolve();
      else reject(new Error(`ProcessContainer step exited with code ${code}`));
    });
  });
  processStepIndex = step;
  finishStep("process", step);
}

async function runMicroVmStep(step) {
  if (step < 1 || step > microvmSteps.length) {
    throw new Error(`Unknown MicroVM step: ${step}`);
  }
  await new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [microvmScript, `--step=${microvmSteps[step - 1]}`],
      { cwd: demoRoot, env: { ...process.env }, windowsHide: true },
    );
    active = child;
    child.stdout.on("data", (chunk) => send({ kind: "output", text: chunk.toString() }));
    child.stderr.on("data", (chunk) => send({ kind: "output", text: chunk.toString() }));
    child.on("error", reject);
    child.on("close", (code) => {
      active = null;
      if (code === 0) resolve();
      else reject(new Error(`MicroVM step exited with code ${code}`));
    });
  });
  finishStep("microvm", step);
}

async function runStep(demo, step) {
  if (busy) throw new Error("A demo step is already running.");
  if (demo !== "pty" && ptyState !== null) {
    throw new Error("Finish or reset the active PTY session before switching.");
  }
  if (!["process", "pty", "microvm"].includes(demo)) {
    throw new Error(`Unknown demo: ${demo}`);
  }
  beginStep(demo, step);
  try {
    if (demo === "process") {
      await runProcessStep(step);
      currentDemo = null;
      processStepIndex = step === processSteps.length ? 0 : step;
    } else if (demo === "microvm") {
      await runMicroVmStep(step);
      currentDemo = null;
    } else {
      if (step === 1 && ptyState !== null) {
        throw new Error("PTY session is already active. Run Steps 2–6 or click Reset.");
      }
      if (step !== 1 && ptyState === null) {
        throw new Error("Run PTY Step 1 first to provision the isolated session.");
      }
      await runPtyStep(step);
      if (step === 6) {
        currentDemo = null;
        ptyStepIndex = 0;
      }
    }
  } catch (error) {
    busy = false;
    send({ kind: "step-error", demo, step, error: String(error) });
    throw error;
  }
}

async function cancelDemo() {
  if (active !== null) {
    active.kill?.();
    active = null;
  }
  await cleanupPty();
  busy = false;
  currentDemo = null;
  processStepIndex = 0;
  ptyStepIndex = 0;
  send({ kind: "cancelled" });
}

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", `http://127.0.0.1:${port}`);

  if (url.pathname === "/api/events") {
    response.writeHead(200, {
      "content-type": "text/event-stream",
      "cache-control": "no-cache",
      connection: "keep-alive",
    });
    response.write("retry: 500\n\n");
    clients.add(response);
    request.on("close", () => clients.delete(response));
    return;
  }

  if (url.pathname === "/api/meta") {
    response.writeHead(200, {
      "content-type": "application/json",
      "cache-control": "no-store",
    });
    response.end(JSON.stringify(hostMeta));
    return;
  }

  if (
    (url.pathname === "/api/run-step" || url.pathname === "/api/cancel" || url.pathname === "/api/pty-input")
    && request.method === "POST"
  ) {
    if (request.headers["x-mxc-demo"] !== "1") {
      response.writeHead(403, { "content-type": "application/json" });
      response.end(JSON.stringify({ error: "Request header missing." }));
      return;
    }
    try {
      if (url.pathname === "/api/cancel") {
        await cancelDemo();
      } else if (url.pathname === "/api/pty-input") {
        if (ptyState === null || !ptyState.started || ptyState.ended) {
          throw new Error("Provision the PTY session (Step 1) before sending live commands.");
        }
        const chunks = [];
        for await (const chunk of request) chunks.push(chunk);
        const cmd = Buffer.concat(chunks).toString("utf8").trim();
        if (cmd.length > 0) {
          writePtyCommand(cmd);
        }
      } else {
        const demo = url.searchParams.get("demo");
        const step = Number.parseInt(url.searchParams.get("step") ?? "", 10);
        await runStep(demo, step);
      }
      response.writeHead(202, { "content-type": "application/json" });
      response.end(JSON.stringify({ status: "completed" }));
    } catch (error) {
      response.writeHead(409, { "content-type": "application/json" });
      response.end(JSON.stringify({ error: String(error) }));
    }
    return;
  }

  if (url.pathname === "/" || url.pathname === "/index.html") {
    try {
      const html = await readFile(htmlPath);
      response.writeHead(200, {
        "content-type": "text/html; charset=utf-8",
        "cache-control": "no-store",
      });
      response.end(html);
    } catch (error) {
      response.writeHead(503, { "content-type": "text/plain; charset=utf-8" });
      response.end(
        `Presentation HTML is unavailable. Run "python build.py --target html".\n${error}`,
      );
    }
    return;
  }

  response.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
  response.end("Not found");
});

server.listen(port, "127.0.0.1", () => {
  console.log(`MXC web demo: http://127.0.0.1:${port}/`);
  console.log("Open slides 25–27 (or press D) to launch Demo 1, Demo 2, or Demo 3.");
});

process.on("SIGINT", async () => {
  await cancelDemo();
  server.close();
});
