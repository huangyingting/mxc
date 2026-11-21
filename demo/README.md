# Real MXC SDK demos

这组 Demo 直接调用官方 `@microsoft/mxc-sdk/v1`，每个 Demo 都对应一个明确的 MXC 能力和可验证结果，不提供普通进程回退。

## 1. ProcessContainer：文件最小权限

源码：`demo/src/process-container.ts`

运行：

```text
npm --prefix demo run demo:process
```

该 Demo 使用 ProcessContainer 验证两件事：

1. Workload 只读 `demo/input/`，并向 `demo/output/report.txt` 成功写入报告。
2. 同一 Host Policy 尝试读取 `demo/protected/denied.txt` 时，MXC 拒绝访问。

网络入站、出站和回环保持拒绝，UI、剪贴板与输入注入禁用。Host 在执行前调用 `getPlatformSupport()` 和 Windows `probe(request)`，然后通过真实 `run(request)` 启动容器。

## 2. Session PTY：事故排查与人工接管

源码：`demo/src/pty.ts`

运行：

```text
npm --prefix demo run demo:pty
npm --prefix demo run demo:pty:interactive
```

自动模式处理只读事故工单 `INC-1042`，检查真实的 `Spooler` 和 `W32Time` 状态，并让 Host 验证 `triage-evidence.txt`。交互模式把真实键盘与输出连接到隔离 Session 中的 PowerShell，适合人工接管和现场演示。

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

## 构建

```text
npm --prefix demo ci
npm --prefix demo run build
```

运行两个 Demo：

```text
npm --prefix demo run demo:all
```

也可从仓库根目录运行 `python build.py --target demo`，仅安装依赖并执行 TypeScript 构建；真实 MXC Demo 需要 Node.js 24.21+ 和支持对应 Backend 的 Windows 主机。
