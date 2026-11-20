import { execFileSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import {
  MxcError,
  deprovisionContainer,
  provisionContainer,
  runInContainer,
  startContainer,
  stopContainer,
  type ContainerId,
  type ExecutionResult,
} from "@microsoft/mxc-sdk/v1";

const stateFileName = "session-state.txt";

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
  if (os.platform() !== "win32") {
    throw new Error("The IsolationSession demo requires Windows.");
  }

  let containerId: ContainerId<"isolation_session"> | undefined;
  let started = false;
  let operationError: unknown;
  const cleanupErrors: unknown[] = [];

  try {
    const provisioned = await provisionContainer({
      containment: "isolation_session",
      appId: "MxcLearningKit.SessionDemo",
      network: {
        egress: { default: "allow" },
        ingress: { default: "allow", hostLoopback: "allow" },
      },
    });
    containerId = provisioned.containerId;

    const metadata = provisioned.metadata;
    if (
      metadata?.ephemeralWorkspacePath === undefined
      || metadata.agentUserName === undefined
    ) {
      throw new Error("MXC did not return complete IsolationSession metadata.");
    }
    console.log(`SESSION provisioned user: ${metadata.agentUserName}`);
    console.log(`SESSION workspace: ${metadata.ephemeralWorkspacePath}`);
    console.log(`SESSION host user: ${os.userInfo().username}`);
    if (
      metadata.agentUserName.toLowerCase()
      === os.userInfo().username.toLowerCase()
    ) {
      throw new Error("IsolationSession unexpectedly reused the host user identity.");
    }

    await startContainer(containerId);
    started = true;
    console.log("SESSION started");

    const executionOptions = {
      workingDirectory: metadata.ephemeralWorkspacePath,
      inheritDefaultEnvironment: true,
      timeoutMs: 30_000,
    } as const;

    const identity = await runInContainer(containerId, {
      ...executionOptions,
      command:
        'powershell.exe -NoProfile -NonInteractive -Command "'
        + "$p = [Diagnostics.Process]::GetCurrentProcess(); "
        + "Write-Output ('IDENTITY=' + [Environment]::UserName); "
        + "Write-Output ('SESSION_ID=' + $p.SessionId)\"",
    });
    assertCompleted("session identity", identity);
    if (!identity.stdout.toLowerCase().includes(
      metadata.agentUserName.toLowerCase(),
    )) {
      throw new Error(
        "The session workload did not run under the provisioned agent user.",
      );
    }

    const hostSessionId = execFileSync(
      "powershell.exe",
      [
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        "[Diagnostics.Process]::GetCurrentProcess().SessionId",
      ],
      { encoding: "utf8" },
    ).trim();
    const isolatedSessionId = identity.stdout.match(/SESSION_ID=(\d+)/)?.[1];
    if (isolatedSessionId === undefined) {
      throw new Error("The session workload did not report its Windows session ID.");
    }
    if (isolatedSessionId === hostSessionId) {
      throw new Error(
        `IsolationSession reused the host Windows session ${hostSessionId}.`,
      );
    }
    console.log(
      `SESSION identity uses ${metadata.agentUserName} in Windows session `
      + `${isolatedSessionId}; host session is ${hostSessionId}`,
    );

    const first = await runInContainer(containerId, {
      ...executionOptions,
      command:
        `cmd.exe /d /s /c echo persisted-by-mxc > ${stateFileName} `
        + "& echo SESSION write completed",
    });
    assertCompleted("session write", first);

    const second = await runInContainer(containerId, {
      ...executionOptions,
      command:
        `cmd.exe /d /s /c type ${stateFileName} `
        + "& echo SESSION read completed",
    });
    assertCompleted("session read", second);

    const state = await readFile(
      path.join(metadata.ephemeralWorkspacePath, stateFileName),
      "utf8",
    );
    if (!state.includes("persisted-by-mxc")) {
      throw new Error("Session state was not visible after a second execution.");
    }
    console.log("SESSION state persisted across two runInContainer calls");
    console.log("SESSION host observed the shared workspace file");
  } catch (error) {
    operationError = error;
  } finally {
    if (containerId !== undefined) {
      if (started) {
        try {
          await stopContainer(containerId);
          console.log("SESSION stopped");
        } catch (error) {
          cleanupErrors.push(error);
          console.error("SESSION stop failed:", error);
        }
      }

      try {
        await deprovisionContainer(containerId);
        console.log("SESSION deprovisioned");
      } catch (error) {
        cleanupErrors.push(error);
        console.error("SESSION deprovision failed:", error);
      }
    }
  }

  if (operationError !== undefined) {
    throw operationError;
  }
  if (cleanupErrors.length > 0) {
    throw new AggregateError(cleanupErrors, "IsolationSession cleanup failed");
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
