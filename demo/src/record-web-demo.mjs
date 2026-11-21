import { spawn } from "node:child_process";
import { mkdir, writeFile } from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const demoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const repoRoot = path.resolve(demoRoot, "..");
const outPath = path.join(repoRoot, "presentation", "assets", "demo-recordings.json");
const port = Number.parseInt(process.env.MXC_RECORD_PORT ?? "8793", 10);
const baseUrl = `http://127.0.0.1:${port}`;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function createSanitizer() {
  const hostUser = os.userInfo().username;
  const hostName = os.hostname();
  const replacements = [
    [new RegExp(escapeRegExp(repoRoot), "gi"), "C:\\mxc"],
    [new RegExp(escapeRegExp(repoRoot.replace(/\\/g, "/")), "gi"), "C:/mxc"],
  ];
  if (hostName) {
    replacements.push([new RegExp(escapeRegExp(hostName), "gi"), "WIN11-HOST"]);
  }
  if (hostUser) {
    replacements.push([new RegExp(escapeRegExp(hostUser), "gi"), "employee-desktop"]);
  }
  return function sanitize(input) {
    if (typeof input === "string") {
      let out = input;
      for (const [pattern, target] of replacements) {
        out = out.replace(pattern, target);
      }
      return out;
    }
    if (Array.isArray(input)) return input.map(sanitize);
    if (input && typeof input === "object") {
      const next = {};
      for (const [key, val] of Object.entries(input)) {
        next[key] = sanitize(val);
      }
      return next;
    }
    return input;
  };
}

function cleanConPtyText(value) {
  return value
    .replace(/\x1b\[2J[\s\S]*?\x1b\[H/g, "")
    .replace(/\x1b\[\d+;1H/g, "\n")
    .replace(/\x1b\[(\d*)C/g, (_, n) => " ".repeat(Math.min(8, Number(n || 1))))
    .replace(/\x1b\][^\x07]*(?:\x07|\x1b\\)/g, "")
    .replace(/\x1b\[[0-?]*[ -/]*[@-~]/g, "")
    .replace(/\r/g, "")
    .replace(/\n{3,}/g, "\n");
}

function compactEvents(rawEvents) {
  const merged = [];
  for (const item of rawEvents) {
    const ev = structuredClone(item.event);
    if (typeof ev.text === "string") {
      ev.text = cleanConPtyText(ev.text);
      if (!ev.text && ev.kind === "output") continue;
    }
    const prev = merged[merged.length - 1];
    if (
      prev
      && prev.event.kind === "output"
      && ev.kind === "output"
      && item.delayMs <= 35
    ) {
      prev.event.text += ev.text;
      continue;
    }
    merged.push({
      delayMs: Math.max(0, Math.min(140, item.delayMs)),
      event: ev,
    });
  }
  return merged;
}

async function waitForServer(maxAttempts = 40) {
  for (let i = 0; i < maxAttempts; i += 1) {
    try {
      const res = await fetch(`${baseUrl}/api/meta`);
      if (res.ok) return await res.json();
    } catch {
      // Server still booting.
    }
    await sleep(150);
  }
  throw new Error(`Timed out waiting for web-demo-server on ${baseUrl}`);
}

function openEventCollector() {
  const listeners = new Set();
  let req = null;
  const ready = new Promise((resolve, reject) => {
    req = http.get(`${baseUrl}/api/events`, (res) => {
      res.setEncoding("utf8");
      resolve();
      let buffer = "";
      res.on("data", (chunk) => {
        buffer += chunk;
        let boundary = buffer.indexOf("\n\n");
        while (boundary !== -1) {
          const frame = buffer.slice(0, boundary);
          buffer = buffer.slice(boundary + 2);
          for (const line of frame.split("\n")) {
            if (line.startsWith("data: ")) {
              const parsed = JSON.parse(line.slice(6));
              for (const fn of listeners) fn(parsed);
            }
          }
          boundary = buffer.indexOf("\n\n");
        }
      });
    });
    req.on("error", (err) => {
      if (err.code !== "ECONNRESET") reject(err);
    });
  });
  return {
    ready,
    subscribe(fn) {
      listeners.add(fn);
      return () => listeners.delete(fn);
    },
    close() {
      req?.destroy();
    },
  };
}

async function main() {
  const sanitize = createSanitizer();
  const server = spawn(process.execPath, ["src/web-demo-server.mjs"], {
    cwd: demoRoot,
    env: { ...process.env, MXC_DEMO_PORT: String(port) },
    stdio: ["ignore", "pipe", "pipe"],
    windowsHide: true,
  });

  let serverLogs = "";
  server.stdout.on("data", (d) => { serverLogs += String(d); });
  server.stderr.on("data", (d) => { serverLogs += String(d); });

  let collector = null;
  try {
    const rawMeta = await waitForServer();
    collector = openEventCollector();
    await collector.ready;
    await sleep(150);

    async function captureStep(demo, step, settleMs) {
      console.log(`Recording ${demo} step ${step}...`);
      const captured = [];
      let lastTick = Date.now();
      let doneResolve;
      let doneReject;
      const donePromise = new Promise((resolve, reject) => {
        doneResolve = resolve;
        doneReject = reject;
      });

      const unsubscribe = collector.subscribe((event) => {
        const now = Date.now();
        captured.push({ delayMs: now - lastTick, event });
        lastTick = now;
        if (event.kind === "step-done" && event.demo === demo && event.step === step) {
          doneResolve();
        } else if (event.kind === "step-error") {
          doneReject(new Error(event.error || `Step ${demo}:${step} failed`));
        }
      });

      try {
        const res = await fetch(`${baseUrl}/api/run-step?demo=${demo}&step=${step}`, {
          method: "POST",
          headers: { "X-MXC-Demo": "1" },
        });
        if (!res.ok) {
          const text = await res.text();
          throw new Error(`HTTP ${res.status} starting ${demo}:${step}: ${text}`);
        }
        await donePromise;
        if (settleMs > 0) {
          await sleep(settleMs);
        }
      } finally {
        unsubscribe();
      }

      // Ensure step-done is always the final lifecycle event while keeping trailing prompt output right before it.
      const stepDoneIdx = captured.findIndex(
        (x) => x.event.kind === "step-done" && x.event.demo === demo && x.event.step === step,
      );
      if (stepDoneIdx !== -1 && stepDoneIdx < captured.length - 1) {
        const [doneItem] = captured.splice(stepDoneIdx, 1);
        doneItem.delayMs = 25;
        captured.push(doneItem);
      }
      return compactEvents(captured);
    }

    async function captureInteractiveCommand(command, settleMs = 850) {
      console.log(`Recording interactive PTY command: ${command}`);
      const captured = [];
      let lastTick = Date.now();
      const unsubscribe = collector.subscribe((event) => {
        if (event.kind !== "output") return;
        const now = Date.now();
        captured.push({ delayMs: now - lastTick, event });
        lastTick = now;
      });
      try {
        const res = await fetch(`${baseUrl}/api/pty-input`, {
          method: "POST",
          headers: {
            "X-MXC-Demo": "1",
            "Content-Type": "text/plain;charset=utf-8",
          },
          body: command,
        });
        if (!res.ok) throw new Error(`pty-input failed: ${res.status}`);
        await sleep(settleMs);
      } finally {
        unsubscribe();
      }
      return compactEvents(captured);
    }

    const ptySteps = {};
    ptySteps["1"] = await captureStep("pty", 1, 900);
    ptySteps["2"] = await captureStep("pty", 2, 500);
    ptySteps["3"] = await captureStep("pty", 3, 500);
    ptySteps["4"] = await captureStep("pty", 4, 500);
    ptySteps["5"] = await captureStep("pty", 5, 500);

    const ptyCommands = [
      {
        label: "whoami",
        command: "whoami",
        aliases: ["whoami"],
        events: await captureInteractiveCommand("whoami", 700),
      },
      {
        label: "pwd",
        command: "pwd",
        aliases: ["pwd", "gl", "get-location"],
        events: await captureInteractiveCommand("pwd", 700),
      },
      {
        label: "ls",
        command: "Get-ChildItem -Name",
        aliases: ["ls", "dir", "gci", "get-childitem", "get-childitem -name"],
        events: await captureInteractiveCommand("Get-ChildItem -Name", 700),
      },
      {
        label: "Get-Service Spooler,W32Time",
        command: "Get-Service Spooler, W32Time | Format-Table -AutoSize",
        aliases: [
          "get-service spooler,w32time",
          "get-service spooler, w32time",
          "get-service spooler, w32time | format-table -autosize",
        ],
        events: await captureInteractiveCommand(
          "Get-Service Spooler, W32Time | Format-Table -AutoSize",
          800,
        ),
      },
      {
        label: "cat triage-evidence.txt",
        command: "Get-Content .\\triage-evidence.txt",
        aliases: [
          "cat triage-evidence.txt",
          "cat .\\triage-evidence.txt",
          "cat ./triage-evidence.txt",
          "type triage-evidence.txt",
          "type .\\triage-evidence.txt",
          "get-content .\\triage-evidence.txt",
          "get-content triage-evidence.txt",
        ],
        events: await captureInteractiveCommand("Get-Content .\\triage-evidence.txt", 750),
      },
    ];

    ptySteps["6"] = await captureStep("pty", 6, 200);

    const processSteps = {};
    for (let step = 1; step <= 5; step += 1) {
      processSteps[String(step)] = await captureStep("process", step, 120);
    }

    const microvmSteps = {};
    for (let step = 1; step <= 5; step += 1) {
      microvmSteps[String(step)] = await captureStep("microvm", step, 120);
    }

    const payload = sanitize({
      schemaVersion: 1,
      sdkPackage: "@microsoft/mxc-sdk@1.1.0",
      platform: `${os.platform()}/${os.arch()} (${os.release()})`,
      hostMeta: {
        username: rawMeta.username,
        sessionId: rawMeta.sessionId,
      },
      steps: {
        pty: ptySteps,
        process: processSteps,
        microvm: microvmSteps,
      },
      ptyCommands,
    });

    await mkdir(path.dirname(outPath), { recursive: true });
    await writeFile(outPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
    console.log(`Saved recorded MXC demo steps to ${outPath}`);
  } catch (error) {
    console.error("Recording failed:", error);
    if (serverLogs) console.error("Server logs:\n", serverLogs);
    process.exitCode = 1;
  } finally {
    collector?.close();
    server.kill();
  }
}

await main();
