# Real MXC SDK demos

这组 Demo 直接调用官方 `@microsoft/mxc-sdk/v1`，每个步骤都按统一结构清晰输出四部分信息（命令行 TTY 与网页 Studio 均以语义颜色区分）：

- **`[POLICY]`（蓝色）**：当前步骤生效的 MXC 隔离策略（右侧面板同步显示完整 JSON）。
- **`[RUN]`（琥珀色）**：在沙箱内实际执行的命令，以及普通标准输出日志（中性浅灰）。
- **`✓ PASS`（绿色） / `✕ BLOCKED`（红色）**：命令真实退出码（如 `exit code 0` 成功，或被操作系统内核/虚拟化边界阻断返回 `exit code 1` / `exit code 7` / `Errno 13` / `Access denied`）。
- **`Why:`（暖黄说明）**：用一句话解释为什么当前策略放行或拦截了该命令。

## 1. Demo 1 · ProcessContainer：文件与网络护栏

源码：`demo/src/process-container.ts`

运行：

```text
npm --prefix demo run demo:process
```

该 Demo 使用 ProcessContainer 分五步验证文件最小权限与方向性网络控制：

1. 探测宿主机 Backend 与 `base-container` 隔离层级。
2. Workload 只读 `demo/input/`，并向 `demo/output/report.txt` 成功写入报告。
3. 同一 Host Policy 尝试读取 `demo/protected/denied.txt` 时，MXC 拒绝访问。
4. 显式允许 `hostLoopback: "allow"` 时，成功访问本地受控回环端点（外网仍默认拒绝）。
5. 切换为 `hostLoopback: "deny"` 时，同一回环请求被立即阻断。

## 2. Demo 2 · IsolationSession：有状态 PTY 事故排查与人工接管

源码：`demo/src/pty.ts`

运行：

```text
npm --prefix demo run demo:pty
npm --prefix demo run demo:pty:interactive
npm --prefix demo run demo:web
npm --prefix demo run demo:record
```

自动模式处理只读事故工单 `INC-1042`，检查真实的 `Spooler` 和 `W32Time` 状态，并让 Host 验证 `triage-evidence.txt`。交互模式把真实键盘与输出连接到隔离 Session 中的 PowerShell，适合人工接管和现场演示。网页模式启动仅监听 `127.0.0.1` 的本机 Server；打开 `http://127.0.0.1:8787/#25`、`#26` 或 `#27` 后点击启动按钮（或按 `D`），用自定义下拉框选择 Demo 1 / Demo 2 / Demo 3 和具体步骤逐步发送。

### 静态网页（GitHub Pages）录制与交互回放

`demo:record` 会在支持 MXC 的 Windows 主机上真实运行 `IsolationSession`（6 个步骤 + 常用交互命令 `whoami`/`pwd`/`ls`/`Get-Service`/`cat triage-evidence.txt`）、`ProcessContainer`（5 个步骤）与 `Hyperlight MicroVM`（5 个步骤），脱敏本机用户名和路径后保存到 `presentation/assets/demo-recordings.json`。`python build.py --target html` 会将该录制数据内联进单文件 HTML，因此在 GitHub Pages 或离线静态打开时，同样可以逐步点击 **Send Step**、一键 **Auto-Play All** 或体验 PTY 快捷命令；而在 `http://127.0.0.1:8787` 下则自动切换为真实 OS / MicroVM 容器执行。

### 这个 Demo 在做什么

1. Host 提供工单和“仅允许只读诊断”的审批。
2. MXC Provision 独立 Windows Agent 用户和 Session。
3. `spawnInContainerWithPty()` 在 Session 中启动 PowerShell PTY。
4. 自动模式或操作员读取审批、检查服务状态并生成证据。
5. Host 验证证据，随后停止并删除 Session。

### 为什么需要 Session Isolation

运维 Agent 常常需要持续终端、多个排查步骤和人工接管，但不应共享员工当前会话的 Windows Profile、Registry、Shell 历史或个人登录环境。Session Container 提供独立用户、独立 Session 和可保持状态的 PTY。

### 边界

IsolationSession 不接受文件 Policy；网络必须声明为 egress、ingress、host loopback 全部 `allow`；窗口与 Session 自己的剪贴板可用，但 `SendInput` 被阻止。它不是凭据保险箱，也不能替代 ProcessContainer 的细粒度文件/网络隔离。工单审批是演示夹具，生产变更还需要真实审批、JIT 权限、Host-side Broker 和受控执行器。

## 3. Demo 3 · Hyperlight MicroVM：百毫秒级硬件微虚拟机（Experimental）

源码：`demo/src/microvm.ts`

运行：

```text
npm --prefix demo run demo:microvm
```

该 Demo 通过 `getPlatformSupport()` 发现宿主机的 `hyperlight` 能力，并调用随 `@microsoft/mxc-sdk` 分发的 `wxc-exec --experimental`（`1.1.0-alpha` 契约，`containment: "hyperlight"`）验证五件事：

1. 检查并在首次运行时自动拉取、预热 `bash` 与 `python` 的 Unikraft 内存快照（`%LOCALAPPDATA%\mxc-hyperlight`）。
2. 从预热快照恢复独立 `Unikraft 5.15 x86_64` 微内核（约 80–100 ms），证明不共享 Windows NT 内核。
3. 在硬件隔离的 Python 3.12 微虚拟机内安全评估不可信 AI 对账代码。
4. 验证宿主机 `C:\` 文件系统、宿主环境变量及网络套接字（`Errno 13`）在 VM 内全部不可达。
5. 验证前一次运行写入 `/tmp` 的状态在下一次启动时被彻底清空，无跨次污染。

## 构建

```text
npm --prefix demo ci
npm --prefix demo run build
```

运行全部三个 Demo：

```text
npm --prefix demo run demo:all
```

也可从仓库根目录运行 `python build.py --target demo`，仅安装依赖并执行 TypeScript 构建；真实 MXC Demo 需要 Node.js 24.21+ 和支持对应 Backend 的 Windows 主机。
