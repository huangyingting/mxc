# Real MXC SDK Demo

该 Demo 使用官方 `@microsoft/mxc-sdk/v1`，通过 MXC Backend 实际执行两个 Workload：

1. 只读 `demo/input/`，写入 `demo/output/report.txt`，预期成功。
2. 尝试读取 `demo/protected/denied.txt`，预期被 MXC 拒绝。

仓库还包含 Windows IsolationSession 示例：创建持久 Session Container，执行两次命令并验证共享状态，最后停止和删除会话。

网络入站、出站和回环均默认拒绝，UI、剪贴板与输入注入也处于禁用状态。代码通过 `getPlatformSupport()` 和 Windows `probe(request)` 先验证宿主能力，再调用真实 `run(request)`；不会模拟容器结果。

## 平台要求

- Node.js 24.21.0+（Windows 原生 stdio）；其他平台需 Node.js 24+
- MXC 支持的宿主平台和已安装的 ProcessBackend 运行时
- Windows 推荐使用支持 BaseContainer/PSEC 的 Windows 11 环境

如果宿主不支持 MXC，Demo 会报告 `getPlatformSupport()` 的实际原因并退出失败，不会降级为普通进程执行。

## 构建和运行

```text
npm --prefix demo ci
npm --prefix demo run build
npm --prefix demo run demo
npm --prefix demo run demo:session
npm --prefix demo run demo:pty
npm --prefix demo run demo:all
```

或在仓库根目录运行：

```text
python build.py --target demo
```

`build.py` 只安装依赖并执行 TypeScript 类型检查/构建，不自动启动需要本机 MXC 能力的集成 Demo。

## 观察点

- `demo/config/policy.json` 是 Host 提交给 MXC SDK 的最小权限策略。
- `demo/src/main.ts` 对两个场景使用同一文件系统、网络与 UI Policy。
- 允许场景验证 SDK 返回码，并从 Host 侧检查生成报告。
- 拒绝场景在容器内尝试读取受保护文件；若读取成功，Workload 会以失败退出。

## Session Container 边界

`demo/src/session.ts` 使用 `provisionContainer()`、`startContainer()`、`runInContainer()`、`stopContainer()` 和 `deprovisionContainer()` 演示真实状态化生命周期。

该示例验证 Session Container 的三项独特属性：

1. 容器内身份与 Provision 返回的独立 Agent 用户一致，不复用 Host 用户。
2. Workload 运行在与 Host 不同的 Windows Session ID 中，而不是仅更换进程令牌。
3. 第一次 `runInContainer()` 写入的文件可被第二次执行复用，并能由 Host 通过共享 workspace 验证。

IsolationSession 是 Windows-only 的独立用户会话，不能等同于 ProcessContainer 的文件、网络过滤能力。MXC 1.1.0 的 Provision 请求必须明确声明 egress、ingress 和 host loopback 均为 `allow`；本示例只依赖会话隔离和共享的 `ephemeralWorkspacePath`。如果这些能力不适合当前威胁模型，应选择 ProcessContainer，而不是误用 Session Container。

## 适用场景

- **有状态桌面自动化**：应用可通过应用内 API、UI Automation 等可信驱动器操作，需保留独立 Windows 用户和应用状态。
- **有状态工具 Agent**：多个命令共享会话、应用登录状态、临时工作区或交互终端。
- **多租户或并发 Agent**：每个 Agent 使用独立 Windows 用户和 Session，避免相互污染用户配置。

MXC 源码明确限制：IsolationSession 不接受文件 Policy；网络必须将 egress、ingress、host loopback 全部声明为 `allow`；窗口、GDI 和 Session 自己的剪贴板在 Session 内可用，但 `SendInput` 被阻止。因此它不能替代 ProcessContainer 的精细文件/网络隔离，也不是所有“Computer-use” Agent 都能直接使用的驱动层。

IsolationSession 也不是凭据保险箱。独立 Agent 用户保护的是 Host 用户的 Profile、Registry、Cookie 和 Token 不被直接继承；一旦 ERP 在 Agent Session 内登录，Cookie 或 Token 就可能被以该 Agent 用户运行的代码访问。应使用专用低权限机器人账号，或由 Host/可信 Broker 持有短时凭据；不要复制员工浏览器 Profile、Cookie 或 Refresh Token。

官方 `microsoft/mxc` 仓库当前提供两个直接示例：`samples/run-in-persisted-container` 对应状态化生命周期，`samples/run-with-io-pty-interactive` 对应交互式 PTY。本示例保留前者的生命周期，并增加独立 OS 用户、Session ID 和共享状态断言。公开仓库目前没有完整的 GUI Computer-use 自动化示例。

## 通用 PTY 示例

`demo/src/pty.ts` 是一个可复用的 Session PTY Host：

- `npm --prefix demo run demo:pty`：自动处理 `INC-1042` 只读诊断工单，检查真实 Windows 服务状态，并让 Host 验证 `triage-evidence.txt`。
- `npm --prefix demo run demo:pty:interactive`：必须在真实 TTY 中运行；将本机键盘与输出直接连接到隔离 Session。
- `MXC_PTY_COMMAND`：交互模式可覆盖启动程序，例如 PowerShell、`cmd.exe` 或企业开发工具的命令行入口。

这个示例对应常见的 **L1/L2 运维故障排查**：Host 将工单和审批放入共享 workspace；Operator Agent 在独立 Windows 用户/Session 的 PTY 中读取审批，只执行只读诊断；生成可审计证据后交回 Host。这样员工无需共享自己的 Profile、Registry 或 Shell 历史，且多个步骤可保持同一交互上下文。自动模式用于 CI/现场验证，交互模式用于人工介入；两种模式都在退出后停止和删除 Session。

### 这个 Demo 在做什么

1. `INC-1042` 是一个打印/时间服务异常工单。
2. Host 把工单和“仅允许只读诊断”的审批交给 Agent。
3. Agent 在独立 Windows Session 中打开 PowerShell PTY。
4. Agent 校验审批，检查真实的 `Spooler` 和 `W32Time` 服务状态。
5. Agent 生成包含工单、审批、操作员、Session ID 和服务状态的证据。
6. Host 验证证据，随后停止并删除 Session。

### 为什么需要 Session Isolation

运维 Agent 常常需要持续终端、多个排查命令和可选人工接管，但不应共享员工当前会话的 Windows Profile、Registry、Shell 历史或个人登录环境。Session Container 为运维流程提供独立操作员身份、独立 Windows Session 和跨命令状态；ProcessContainer 仍更适合一次性脚本、构建、测试与精细文件/网络隔离。

这个 Demo 中，Windows 服务状态、独立用户、不同 Session ID、PTY 以及证据文件是真实的；工单和审批是可替换的测试夹具。Demo 不读取凭据，不重启服务，也不修改配置。生产变更必须增加真实审批系统、JIT 权限和 Host-side 执行器。

示例不会执行服务重启、凭据读取或其他变更。真正的生产变更应要求额外审批、JIT 权限和 Host-side 执行器，不能仅依赖 PTY 中的文本确认。
