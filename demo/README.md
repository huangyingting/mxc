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

- **Computer-use / GUI Agent**：需要独立桌面、剪贴板和输入会话，不应共享员工当前会话。
- **有状态工具 Agent**：多个命令共享会话、应用登录状态、临时工作区或交互终端。
- **多租户或并发 Agent**：每个 Agent 使用独立 Windows 用户和 Session，避免相互污染用户配置。

不适合靠 Session Container 单独完成精细文件与网络最小权限控制；这类需求应使用 ProcessContainer，必要时再把可信 GUI 工作放入独立 Session。

官方 `microsoft/mxc` 仓库当前提供两个直接示例：`samples/run-in-persisted-container` 对应状态化生命周期，`samples/run-with-io-pty-interactive` 对应交互式 PTY。本示例保留前者的生命周期，并增加独立 OS 用户、Session ID 和共享状态断言。公开仓库目前没有完整的 GUI Computer-use 自动化示例。
