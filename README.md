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
- Demo 1（ProcessContainer 文件 + 网络护栏）：`npm --prefix demo run demo:process`
- Demo 2（IsolationSession 有状态 PTY）：`npm --prefix demo run demo:pty`；交互模式为 `npm --prefix demo run demo:pty:interactive`
- Demo 3（Hyperlight MicroVM 实验性硬件微虚拟机）：`npm --prefix demo run demo:microvm`
- 网页集成交互式 Demo Studio：
  - **GitHub Pages / 静态网页回放**：直接打开生成的 `dist/html/index.html` 或在线 Pages，在第 25、26、27 页点击按钮（或任意页按 `D`），即可逐步回放（或点击 `Auto-Play All` 自动连播）在 Windows 11 真机上录制的 Demo 1（`ProcessContainer`）、Demo 2（`IsolationSession PTY`）与 Demo 3（`Hyperlight MicroVM`）完整执行轨迹
  - **本机实时沙箱桥接**：运行 `npm --prefix demo run demo:web` 后打开 `http://127.0.0.1:8787/#25`、`#26` 或 `#27`，同一个界面会自动切换为 `live bridge`，驱动本机真实容器与 MicroVM（重新录制静态回放数据可运行 `npm --prefix demo run demo:record`）

Session Container 解决的是独立 Windows 用户/Session 与状态隔离，不是凭据保险箱；文件策略不可用、网络必须全开放、`SendInput` 被阻止。ERP 凭据仍应由 Host/Broker 或专用低权限账号管理。
- Lab：阅读 `labs/lab-01-least-privilege/README.md`

## GitHub Pages

仓库中使用 GitHub Actions 发布 `docs/`。首次发布时，在仓库的 **Settings → Pages → Source** 中选择 **GitHub Actions**。`main` 分支更新后，工作流会重新生成 `docs/index.html` 并部署。

HTML 为包含图片、样式、脚本与中文讲者备注的单文件。底部低干扰工具栏显示页码、`Notes` 与 `Studio`；按 `N` 或点击 `Notes` 展开/收起备注，翻页时备注同步更新，`Esc` 关闭备注，打开 Studio 时自动收起。备注与 PPTX 共用 `note()` 内容，打印时工具栏和备注均隐藏。

使用方向键、PageUp/PageDown、Space/Shift+Space 或点击切换幻灯片，Home/End 跳至首页/末页；第 6、8 页会先逐项 Appear，再进入下一页。`F` 全屏，`D` 打开 Studio，`Ctrl+P` 可打印为 PDF。画面按完整 16:9 幻灯片适配窗口，宽屏或竖屏不会裁切内容。

`.gitignore` 已排除 `dist/`、生成的 `docs/index.html`、Demo 运行输出、Python 缓存、虚拟环境及本地 IDE 文件；源码、素材、Lab 和 Pages 配置仍会纳入版本控制。

## 修改内容

- 幻灯片内容、主题和备注位于 `presentation/src/build_pptx.py`。
- HTML 布局和交互位于 `presentation/src/build_html.py`。
- 图片位于 `presentation/assets/images/`。
- 不要手工修改 `docs/index.html` 或 `dist/`，它们都是构建产物。

## 演示框架

本仓库与 `auto-routing` 保持一致的 HTML/PPTX 基础框架，以 `auto-routing` 为参照：主题字体、文本与代码间距、页眉页脚、图片 cover/contain、讲者工具栏、备注浮层、快捷键和打印样式。幻灯片内容、逐项动画与 Demo Studio 保留各自主题；两个仓库可独立构建，不依赖相邻目录。

HTML 与 PPTX 构建前都会检查元素边界、标题换行和文本框重叠。修改框架后，在两个仓库分别运行：

```text
python -m unittest discover -s presentation/tests
python build.py --target presentation
```

## 视觉系统

- 展示大标题（HTML）：**Newsreader** 衬线字体，增强编辑感与叙事张力
- 标题与正文：**Plus Jakarta Sans**（回退至 Inter / Segoe UI）
- 代码与终端：**JetBrains Mono**（回退至 Cascadia Code / Consolas）
- 配色：暖米色宣纸底、炭灰墨色、石板蓝强调色、圆角微阴影卡片与砖红警示色
- 图示与交互：PPTX 使用可编辑原生矢量形状；HTML 使用内联 SVG、自定义无原生 `<select>` 下拉组件，以及右侧实时分步架构解说栏

## Demo 的边界

Demo 直接调用官方 `@microsoft/mxc-sdk`，真实隔离由当前主机支持的 MXC Backend 执行。如果主机不支持 MXC，程序会明确失败而不会回退到普通进程。
