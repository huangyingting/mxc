# MXC learning kit

“Guardrails with Sandboxing” 的完整学习仓库，包含演示文稿、可部署网页、Demo 与 Hands-on lab。

## 目录

| 目录 | 用途 |
|---|---|
| `presentation/src/` | PPTX/HTML 生成器和幻灯片源代码 |
| `presentation/assets/` | 插图与图片生成提示 |
| `dist/pptx/` | 生成的 PowerPoint 文件 |
| `dist/html/` | 生成的独立 HTML 网页 |
| `docs/` | GitHub Pages 发布目录，构建时复制 HTML |
| `demo/` | 使用官方 MXC Node.js SDK 的真实沙箱 Demo |
| `labs/` | Hands-on lab 的任务、脚手架、验证器和参考答案 |
| `.github/workflows/pages.yml` | GitHub Pages 自动构建与部署 |

## 快速开始

需要 Python 3.9+。

```text
python -m pip install -r requirements.txt
python build.py --target all
```

Windows 可运行 `run.bat`，macOS/Linux 可运行 `./run.sh`。完整构建会通过 `npm` 自动安装 Demo 的锁定依赖；Demo 运行需要 Node.js 24+ 及 MXC 支持的主机。

单独构建：

```text
python build.py --target pptx
python build.py --target html
python build.py --target demo
```

- PPTX：`dist/pptx/Guardrails with Sandboxing - MXC.pptx`
- HTML：`dist/html/index.html`，同时复制到 `docs/index.html`
- ProcessContainer Demo：`npm --prefix demo run demo`
- Session Container Demo：`npm --prefix demo run demo:session`，验证独立 Agent 用户、不同 Windows Session ID 与跨执行状态
- Lab：阅读 `labs/lab-01-least-privilege/README.md`

## GitHub Pages

仓库中使用 GitHub Actions 发布 `docs/`。首次发布时，在仓库的 **Settings → Pages → Source** 中选择 **GitHub Actions**。`main` 分支更新后，工作流会重新生成 `docs/index.html` 并部署。

HTML 为包含图片、样式和脚本的单文件，画面中不显示控件或演讲者备注。使用方向键、Space、点击切换幻灯片，`F` 全屏，`Ctrl+P` 可打印为 PDF；中文讲者备注仅保留在 PPTX 中。

`.gitignore` 已排除 `dist/`、生成的 `docs/index.html`、Demo 运行输出、Python 缓存、虚拟环境及本地 IDE 文件；源码、素材、Lab 和 Pages 配置仍会纳入版本控制。

## 修改内容

- 幻灯片内容、主题和备注位于 `presentation/src/build_pptx.py`。
- HTML 布局和交互位于 `presentation/src/build_html.py`。
- 图片位于 `presentation/assets/images/`。
- 不要手工修改 `docs/index.html` 或 `dist/`，它们都是构建产物。

## Demo 的边界

Demo 直接调用官方 `@microsoft/mxc-sdk`，真实隔离由当前主机支持的 MXC Backend 执行。如果主机不支持 MXC，程序会明确失败而不会回退到普通进程。
