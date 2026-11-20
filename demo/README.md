# Policy Demo

这个 Demo 用一小段 Python 代码说明 Host 如何做默认拒绝、路径边界和网络精确放行。它**不是操作系统沙箱**；生产环境必须使用 MXC Backend 实际执行隔离。

## 运行

从仓库根目录执行：

```text
python -m demo.src.policy_demo
```

预期看到一次受控写入、一次 HTTPS 放行、一次越界读取拒绝和一次非白名单网络拒绝。

## 构建和测试

```text
python build.py --target demo
```

构建会运行 `demo/tests/` 中的标准库单元测试，不需要额外依赖。

## 演示步骤

1. 打开 `demo/config/policy.json`，说明只读输入、只写输出、拒绝优先。
2. 运行 Demo，观察相同的策略入口如何处理文件与网络请求。
3. 修改 `"ports": [443]` 或 `host`，验证放行范围立即收窄。
4. 尝试读取 `../README.md`，强调路径校验发生在实际 I/O 之前。
