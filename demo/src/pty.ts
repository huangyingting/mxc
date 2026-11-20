import { once } from "node:events";
import { copyFile, readFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  MxcError,
  deprovisionContainer,
  provisionContainer,
  spawnInContainerWithPty,
  startContainer,
  stopContainer,
  type ContainerId,
} from "@microsoft/mxc-sdk/v1";

const interactive = process.argv.includes("--interactive");
const ptyShell = process.env.MXC_PTY_COMMAND
  ?? "powershell.exe -NoLogo -NoProfile";
const demoRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

async function main(): Promise<void> {
  if (os.platform() !== "win32") {
    throw new Error("The IsolationSession PTY demo requires Windows.");
  }
  if (
    interactive
    && (!process.stdin.isTTY || !process.stdout.isTTY)
  ) {
    throw new Error("Interactive PTY mode requires a real terminal.");
  }

  let containerId: ContainerId<"isolation_session"> | undefined;
  let started = false;
  let terminal:
    | Awaited<ReturnType<typeof spawnInContainerWithPty>>
    | undefined;
  let operationError: unknown;
  const cleanupErrors: unknown[] = [];
  let stdinWasRaw = false;

  try {
    const provisioned = await provisionContainer({
      containment: "isolation_session",
      appId: "MxcLearningKit.PtyDemo",
      network: {
        egress: { default: "allow" },
        ingress: { default: "allow", hostLoopback: "allow" },
      },
    });
    containerId = provisioned.containerId;
    const workspace = provisioned.metadata?.ephemeralWorkspacePath;
    if (workspace === undefined) {
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
    console.log(
      "PTY incident staged: INC-1042, read-only diagnostics, "
      + "operator approval attached",
    );

    await startContainer(containerId);
    started = true;

    terminal = await spawnInContainerWithPty(
      containerId,
      {
        command: ptyShell,
        workingDirectory: workspace,
        inheritDefaultEnvironment: true,
        timeoutMs: interactive ? undefined : 30_000,
      },
      { size: { rows: 30, columns: 120 } },
    );

    console.log(
      `PTY connected to isolated session: ${ptyShell}; `
      + `terminal=${120}x${30}`,
    );

    if (interactive) {
      console.log(
        "Interactive task: inspect incident-request.json, verify the approval, "
        + "collect diagnostics, and save evidence in triage-evidence.txt.",
      );
      stdinWasRaw = process.stdin.isRaw;
      terminal.output.pipe(process.stdout, { end: false });
      process.stdin.pipe(terminal.input);
      process.stdin.setRawMode(true);
      process.stdin.resume();

      const outputEnded = once(terminal.output, "end");
      const result = await terminal.wait();
      await outputEnded;
      if (result.timedOut || result.exitCode !== 0) {
        throw new Error(
          `PTY shell ended with exit code ${result.exitCode}; `
          + `timedOut=${result.timedOut}`,
        );
      }
    } else {
      let transcript = "";
      const outputEnded = once(terminal.output, "end");
      terminal.output.on("data", (chunk) => {
        const text = String(chunk);
        transcript += text;
        process.stdout.write(text);
      });
      terminal.input.end(
        "$request = Get-Content .\\incident-request.json -Raw | ConvertFrom-Json\r\n"
        + "$approval = Get-Content .\\incident-approval.json -Raw | ConvertFrom-Json\r\n"
        + "if ($approval.ticket -ne $request.ticket -or "
        + "$approval.scope -ne 'read-only-diagnostics') "
        + "{ Write-Error 'Approval mismatch'; exit 2 }\r\n"
        + "$serviceLines = @($request.serviceNames | ForEach-Object { "
        + "$service = Get-Service -Name $_; "
        + "'service=' + $service.Name + ',status=' + $service.Status })\r\n"
        + "$evidence = @('ticket=' + $request.ticket, "
        + "'scope=' + $approval.scope, "
        + "'approved_by=' + $approval.approvedBy, "
        + "'operator=' + $env:USERNAME, "
        + "'session_id=' + [Diagnostics.Process]::GetCurrentProcess().SessionId); "
        + "$evidence += $serviceLines; "
        + "$evidence | Set-Content .\\triage-evidence.txt -Encoding utf8\r\n"
        + "Write-Output ('TRIAGE_COMPLETED=' + $request.ticket)\r\n"
        + "exit\r\n",
      );

      const result = await terminal.wait();
      await outputEnded;
      for (const warning of terminal.warnings) {
        console.warn(`PTY warning: ${warning}`);
      }
      if (result.timedOut || result.exitCode !== 0) {
        throw new Error(
          `PTY smoke test failed with exit code ${result.exitCode}: `
          + `timedOut=${result.timedOut}`,
        );
      }
      if (!transcript.includes("TRIAGE_COMPLETED=INC-1042")) {
        throw new Error("PTY transcript did not contain the expected assertions.");
      }

      const evidence = await readFile(
        path.join(workspace, "triage-evidence.txt"),
        "utf8",
      );
      if (
        !evidence.includes("ticket=INC-1042")
        || !evidence.includes("scope=read-only-diagnostics")
        || !evidence.includes("service=Spooler,status=")
        || !evidence.includes("session_id=")
      ) {
        throw new Error("The PTY incident evidence is incomplete.");
      }
      console.log("PTY evidence verified by the trusted host");
    }
  } catch (error) {
    operationError = error;
  } finally {
    if (interactive) {
      process.stdin.unpipe(terminal?.input);
      if (process.stdin.isTTY) {
        process.stdin.setRawMode(stdinWasRaw);
      }
      process.stdin.pause();
    }
    terminal?.dispose();

    if (containerId !== undefined) {
      if (started) {
        try {
          await stopContainer(containerId);
          console.log("PTY session stopped");
        } catch (error) {
          cleanupErrors.push(error);
          console.error("PTY session stop failed:", error);
        }
      }
      try {
        await deprovisionContainer(containerId);
        console.log("PTY session deprovisioned");
      } catch (error) {
        cleanupErrors.push(error);
        console.error("PTY session deprovision failed:", error);
      }
    }
  }

  if (operationError !== undefined) {
    throw operationError;
  }
  if (cleanupErrors.length > 0) {
    throw new AggregateError(cleanupErrors, "PTY session cleanup failed");
  }
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
