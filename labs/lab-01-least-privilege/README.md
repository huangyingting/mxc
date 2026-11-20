# Lab 1：把权限收窄

## 目标

实现一个默认拒绝的策略检查器，完成三个约束：

1. 读写路径必须位于明确授权的目录内，路径前缀相似不算。
2. `denied` 规则优先于所有放行规则。
3. 网络只允许策略中精确匹配的 HTTPS host 和端口。

整个实验只实现“策略判断”，不会启动网络请求，也不是操作系统级沙箱。

## 步骤

1. 阅读 `policy.json`。
2. 在 `starter/policy.py` 中实现三个 `TODO` 函数。
3. 从仓库根目录运行验证：

```text
python labs/lab-01-least-privilege/tests/verify.py labs/lab-01-least-privilege/starter/policy.py
```

4. 修复失败项，直到全部测试通过。
5. 对照 `solution/policy.py` 中的参考实现，思考下面的问题：
   - 为什么不能用 `str(path).startswith(str(root))`？
   - 为什么 `denied` 必须优先检查？
   - 为什么只让用户选择“模式”比让用户提交任意路径更安全？

## 额外挑战

- 给 `write` 增加只允许 `*.txt` 的规则。
- 让网络规则支持多段域名白名单，但仍然拒绝相似域名。
- 将验证器改为同时验证“允许通过”和“禁止失败”两类测试。

## 验证参考答案

```text
python labs/lab-01-least-privilege/tests/verify.py labs/lab-01-least-privilege/solution/policy.py
```
