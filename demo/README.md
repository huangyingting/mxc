# Real MXC SDK Demo

该 Demo 使用官方 `@microsoft/mxc-sdk/v1`，通过 MXC Backend 实际执行两个 Workload：

1. 只读 `demo/input/`，写入 `demo/output/report.txt`，预期成功。
2. 尝试读取 `demo/protected/denied.txt`，预期被 MXC 拒绝。

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
