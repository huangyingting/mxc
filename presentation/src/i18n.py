#!/usr/bin/env python3
"""Bilingual (en + zh) slide text and English speaker notes for Guardrails with Sandboxing - MXC."""

EN_NOTES = {
    0: (
        "Welcome, everyone. Over the past couple of years, the role of AI in our stack has fundamentally shifted. "
        "Chat assistants used to sit behind a human copy-paste loop—they wrote text, and you decided what to run. "
        "Today's agents act directly on your workstation or server: reading files, running shell commands, and calling APIs.\n"
        "Once an agent can act, the security question changes completely: what happens when it misinterprets a goal or reads a poisoned README? "
        "Today we will look at Microsoft Execution Containers—MXC—and how to let agents get real work done while enforcing hard boundaries at the OS level."
    ),
    1: (
        "Over the next 45 minutes, we will follow a practical engineering arc. First, what new risks show up once agents execute commands, "
        "and why system prompts alone can't stop them. Next, what MXC is and how it maps one JSON policy onto native OS sandboxes. "
        "Then we will cover file, network, and UI rules alongside the four containment tiers, walk through coding and enterprise scenarios, "
        "see how to wire the V1 SDK into your host app in fifteen lines of code, and run three live Windows demos."
    ),
    2: (
        "Let's look at the core shift. With a traditional chat assistant, a human is always in the loop—you eyeball a command before pasting it into your terminal, "
        "so the worst case is usually just a bad suggestion. Autonomous agents remove that human gate: they browse repos, install packages, run scripts, and edit configs on their own.\n"
        "Even trickier, the exact command an agent runs next is synthesized on the fly from runtime context—you can't pre-approve a static allow-list ahead of time. "
        "If that agent runs naked under your own user login, every file you can open, every saved token, and every internal endpoint you can reach is exposed to one bad turn."
    ),
    3: (
        "What actually goes wrong in practice? Four failure modes show up again and again. First is prompt injection: an agent reads an untrusted repo or web page "
        "containing hidden instructions that trick it into running an attacker's script. Second and third are credential theft and data exfiltration—generated code "
        "reads local SSH keys, cloud tokens, or browser secrets and quietly posts them outbound.\n"
        "And fourth doesn't even require an attacker—it's pure overreach. Tell an agent to fix a failing check and it might decide the fastest shortcut is editing live production config. "
        "Whether malicious or overly 'helpful,' without a hard OS boundary underneath, the blast radius is identical.\n\n"
        "Source: Overreach scenario adapted from the Windows Developer Blog (Oct 7, 2026)."
    ),
    4: (
        "That brings us to the single most important rule in agent security: an agent cannot be its own security authority. "
        "You can write 'never touch secrets or open sockets' ten times in your system prompt, add content filters, and tune tool schemas—and you should—but all of those operate at the persuasion layer.\n"
        "If the model gets confused or spawns a subprocess that misbehaves, a prompt cannot block a syscall. Only the operating system kernel can enforce which paths open, "
        "which sockets connect, and whether clipboard or input injection is possible. Prompts guide what the model tries; the OS decides what is physically possible."
    ),
    5: (
        "On enterprise Windows, Containment, Identity, and Manageability are the three OS-level pillars built by the Windows platform team—and wired directly into Microsoft's broader enterprise control plane. "
        "Containment uses open-source MXC to lock down local file, network, and UI access; Identity connects local on-device agents to Microsoft Entra Agent ID so actions carry a distinct agent audit trail instead of borrowing an employee login; "
        "and Manageability extends Microsoft Agent 365 and Intune to endpoints to inventory agents and push MXC sandbox policies.\n"
        "Keep the rollout timeline in mind: MXC itself, MXC on Windows 365 Cloud PCs, and cloud Entra Agent ID and Agent 365 are GA today, while automatic identity binding for local CLIs and Intune policy distribution for MXC are currently in Windows preview.\n\n"
        "Source: Windows Developer Blog — Logan Iyer (Oct 7 & Nov 17, 2026)."
    ),
    6: (
        "So what is MXC? In short, it is an MIT-licensed, open-source execution layer from Microsoft that reached general availability in October. "
        "Whether you are running AI-generated scripts, third-party plugins, or untrusted tools, your host app describes what files, network endpoints, and UI access are permitted in one JSON policy, and MXC maps that policy onto native OS sandboxes.\n"
        "Despite 'Containers' in the name, there is no heavy Docker daemon to install—MXC uses lightweight kernel sandboxes built right into Windows, macOS, and Linux, with official SDKs for .NET, TypeScript, and Rust."
    ),
    7: (
        "Before looking at code, let's line up five terms you will see everywhere in MXC. Your **Host app** is the only trusted party—it picks permissions, stages inputs, and validates outputs. "
        "The **Workload** is the untrusted command or script running inside the box. The **Policy** declares allowed files, network rules, and UI limits. "
        "A **Container request** bundles that policy with the command line, working directory, timeout, and containment choice. And the **Backend** is the OS primitive that enforces it.\n"
        "Remember the bottom line: the host decides, the policy describes, and the backend enforces. An agent can ask for access, never grant it to itself."
    ),
    8: (
        "Stacking those five pieces gives you MXC's four-layer architecture: your trusted host app at the top, followed by the SDK contract, the MXC runtime, and the kernel containment backend at the bottom. "
        "Notice where the red trust boundary sits—right underneath your host app.\n"
        "A common pitfall is assuming that because code ran inside a sandbox, its output is automatically safe. Containment stops the process from escaping, "
        "but stdout or output files can still carry malformed payloads or second-order prompt injections. Always validate everything crossing back over that line."
    ),
    9: (
        "Here is how a single request flows across that boundary at runtime: the host queries the SDK to discover supported backends on this machine, builds a deny-by-default policy and request, and launches through the SDK. "
        "MXC translates that contract into native OS primitives, runs the workload, and hands back stdout, exit code, and blocked-access diagnostics while tearing down temporary state.\n"
        "One practical testing tip: different OS backends enforce slightly different edges, so don't just test that your happy path succeeds—write automated negative tests proving blocked files and denied sockets actually fail."
    ),
    10: (
        "MXC gives you four containment boxes along a clear spectrum from lightest startup to strongest kernel isolation. For everyday coding agents, test runners, and document tools, start on the far left with **Process container**—mapped to AppContainer on Windows, Seatbelt on macOS, and Bubblewrap on Linux, with millisecond startup and fine-grained file and network rules.\n"
        "If an agent drives legacy GUI apps and needs its own desktop and clipboard, pick Windows **Session container**, which creates an isolated local user session—though note it isolates the session rather than filtering per-file or per-IP rules. "
        "Use **WSL container** for Linux CLI toolchains on Windows, and step up to experimental **MicroVM** when running high-risk untrusted code that shouldn't even share the host OS kernel."
    ),
    11: (
        "When shipping cross-platform software, you write against one SDK contract, so it helps to know what runs underneath on each OS. Windows 11 defaults to `ProcessContainer` alongside `IsolationSession`, `WSLc`, and Learning Mode reports; Linux defaults to `Bubblewrap`; macOS maps to `Seatbelt`. "
        "On Windows 365 Cloud PCs, process isolation works out of the box just like a physical laptop—nested virtualization is only needed for session or VM-backed tiers.\n"
        "At the bottom, Windows Sandbox, MicroVM, and Hyperlight are experimental opt-in features via `1.1-alpha`. Always call capability discovery at startup rather than hardcoding a single backend."
    ),
    12: (
        "Once you pick a box, the next question is how long it lives. MXC supports **One-shot** and **State-aware** lifecycles. Eighty to ninety percent of agent actions—running a build, executing tests, summarizing a file—should use One-shot: spin up a fresh, narrow ProcessContainer for one command and tear it down immediately.\n"
        "Reach for State-aware sessions like `IsolationSession` or `WSLc` only when you truly need shell or desktop state across steps: call `provision` to get an opaque `ContainerId`, run commands against it, keep each backend's scope in mind, and always call `deprovision` in your cleanup block so temporary OS accounts never leak.\n\n"
        "Source: Microsoft Learn — MXC policy examples (Reuse a state-aware sandbox)."
    ),
    13: (
        "Now let's walk through the three core policy areas from the official Microsoft Learn examples, starting with files. The golden pattern is **read in, write out**: keep your source repo read-only and point all outputs to a dedicated `mxc-output` staging folder. "
        "On Windows `ProcessContainer`, if you pass an explicit `env` list for build tools, include `SystemRoot` and redirect `LOCALAPPDATA` and `TEMP` into `mxc-output` so tool caches don't force you to make the repo writable.\n"
        "Second, ungranted paths like `~/.ssh` are already blocked by default—use `deniedPaths` specifically to carve out `.env` or `secrets/` inside a granted read-only repo. Third, container exit never undoes writes to granted host folders, so inspect `mxc-output` before trusting its contents.\n\n"
        "Source: Microsoft Learn — MXC policy examples (`learn.microsoft.com/en-us/windows/apps/develop/mxc/examples`)."
    ),
    14: (
        "Locking down files isn't enough if an agent can still open any outbound socket. MXC defaults `egress`, `ingress`, and `hostLoopback` to `deny`. When a workload genuinely needs network access, there are two clean patterns—and you shouldn't mix them:\n"
        "First, direct CIDR allow-lists open specific IP ranges and TCP ports at the OS firewall—remember, firewall rules match IPs and ports, not domain names. "
        "Second, for domain-level allow-lists, run an application proxy on the host and point `runtimeConfig.networkProxy` at it. On Windows `ProcessContainer`, note two subtle details: `allowedProxyPeer` takes the proxy's Package Family Name or AppContainer profile name (not a PID), and Windows mapping requires `ingress.default: \"allow\"` while `hostLoopback: \"deny\"` keeps general localhost ports locked.\n\n"
        "Source: Microsoft Learn — MXC policy examples."
    ),
    15: (
        "Beyond files and sockets, credentials and the user's desktop are your highest-value targets. One rule eliminates ninety percent of credential leaks: never inject raw API keys or SSH private keys into sandbox env vars or disk—let your host app or external proxy attach auth headers outside the box.\n"
        "For desktop safety, the `ui` block gives you three explicit switches: default `disable: true`, `clipboard: \"none\"`, and `injection: false` (`allowInputInjection` in the V1 SDK) to block window access, clipboard snooping, and synthetic keystrokes. "
        "If a visible helper tool needs one-way clipboard paste, flip only `disable: false` and `clipboard: \"read\"`. And for full GUI desktop automation, isolate it inside a `SessionContainer` under a dedicated bot account.\n\n"
        "Source: Microsoft Learn — MXC policy examples (Allow a visible tool without broad UI access)."
    ),
    16: (
        "Here is the practical question every team hits on day one: compilers and CLI tools touch dozens of system DLLs and cache paths—how do you write a least-privilege policy without playing whack-a-mole? That is what **Learning Mode** is for.\n"
        "In production you run strict **Enforcement**; for temporary debugging there is **Permissive**. During policy authoring, **Learning Mode** still blocks ungranted access, but records every blocked path and socket attempt into a structured activity report. "
        "Run real workloads, have an engineer review that evidence to separate legitimate tool dependencies from overreach, approve the minimal grants, and flip back to Enforcement."
    ),
    17: (
        "Putting those pieces together, a secure coding agent is the clearest quick win. To edit code, run tests, and push a branch, look at how narrow its real footprint is: read/write on the single repo folder, access to `node`, `python`, and `git`, and package fetches routed through a local proxy. "
        "Other repos, personal Documents, `~/.ssh`, and arbitrary outbound IPs stay completely off-limits without slowing the agent down.\n"
        "And this isn't hypothetical—GitHub Copilot CLI uses MXC process containment in production today to box AI-generated shell commands."
    ),
    18: (
        "Beyond software engineering, three enterprise patterns come up constantly. First, **desktop automation agents** driving legacy line-of-business apps: run them in a `SessionContainer` with a low-privilege bot account so the employee works undisturbed on their own desktop while the agent processes forms in a separate session.\n"
        "Second, **sensitive-document agents** summarizing contracts or financials: lock them in an air-gapped `ProcessContainer`—read one input folder, write one summary file, `egress: deny`. Third, **agents on Windows 365 Cloud PCs**, where process containment works identically to physical Windows 11 endpoints."
    ),
    19: (
        "Importantly, MXC isn't a closed Microsoft-only component—the broader AI ecosystem is converging on it fast. GitHub Copilot, OpenAI Codex, OpenClaw, Replit, LM Studio, and Unsloth AI already support MXC today; "
        "NVIDIA OpenShell has integrated it, OpenAI and Meta are building on these primitives, and Anthropic Claude Code, Box, Perplexity, Raycast, and others have announced support. "
        "For enterprise architects, that means different coding and productivity agents are standardizing on one OS-native security layer rather than inventing incompatible sandboxes."
    ),
    20: (
        "If you are embedding MXC into your own agent platform, here is the reference architecture. When your LLM planner proposes a tool call like `npm test`, that string never goes straight to a shell—it goes to your trusted host's policy engine, which selects a pre-vetted least-privilege template and launches via the MXC SDK.\n"
        "The return arrow matters just as much: your host validates exit codes, output schemas, and kernel denial logs before feeding anything back into the agent loop. Pair that with an external credential-attaching proxy, Learning Mode during rollout, and Entra Agent ID plus Agent 365 for fleet governance."
    ),
    21: (
        "Dropping down to the SDK layer, integration is intentionally lightweight. Microsoft ships synchronized `.NET`, `TypeScript`, and `Rust` packages on NuGet, npm, and crates.io. "
        "In all three languages, you build a typed `ContainerRequest` declaring files, network, command, and timeout, and the SDK serializes the stable `1.0.0` policy contract to native OS primitives.\n"
        "Just build one habit from the start: call `getPlatformSupport()` at startup, and on Windows call `probe(request)` before launch to verify your policy achieves the exact native isolation tier you expect."
    ),
    22: (
        "In TypeScript with `@microsoft/mxc-sdk/v1`, wiring this up is the fifteen-line five-step loop on the right: first, call synchronous `getPlatformSupport()` and `probe(request)`—and fail closed if unsupported rather than silently running unsandboxed; "
        "second, build a deny-by-default `ContainerRequest` locking down read-only repo, writable staging dir, denied network, disabled UI/clipboard, and `timeoutMs`; "
        "third, iterate in Learning Mode and human-review blocked attempts; fourth, enforce and test both allowed and denied cases; and fifth, `await run(request)` with clean timeout and teardown handling."
    ),
    23: (
        "Right before we jump into the live demos, here is a six-point checklist for your host app: always set an explicit timeout and separate sandbox launch errors, non-zero workload exits, and user cancellations; "
        "run cleanup and `deprovision` inside `finally` blocks so aborted runs never orphan temporary Windows accounts;\n"
        "prefer one narrow, short-lived sandbox per action over one broad long-lived container; treat writable staging folders as untrusted since teardown doesn't roll back host writes; surface policy denials clearly; and keep a human in the loop for every Learning Mode rule."
    ),
    24: (
        "Let's click the top-right launcher to open the interactive MXC Studio and run all three Windows sandboxes. Demo 1 exercises one-shot `ProcessContainer` across five positive and negative checks.\n"
        "Step 1 loads `policy.json` and runs `getPlatformSupport()` plus `probe(request)` to confirm native Windows AppContainer file ACLs and network filters are active.\n"
        "Step 2 tests least-privilege happy-path I/O: `demo/input` is read-only and `demo/output` is writable, so reading `request.txt` and writing `report.txt` succeeds with exit code 0.\n"
        "Step 3 runs the negative file test under the exact same policy: the workload tries to read `protected/denied.txt` and is immediately blocked by the Windows kernel with Access Denied (exit 1).\n"
        "Steps 4 and 5 test the directional loopback switch while internet egress stays `deny`: with `hostLoopback: \"allow\"`, `curl` reaches our local `127.0.0.1` listener for HTTP 200; flipping only `hostLoopback` to `\"deny\"` blocks the exact same `curl` at the OS network stack (exit 7)."
    ),
    25: (
        "Demo 2 simulates front-line incident triage using a stateful `IsolationSession` paired with an interactive PowerShell PTY across six lifecycle steps.\n"
        "Step 1 stages ticket `INC-1042` and its read-only approval into a shared handoff folder, provisions a fresh temporary local Windows operator account and isolated desktop session, and opens a `120×30` PowerShell PTY—preventing any desktop or clipboard bleed into the logged-in employee session.\n"
        "Step 2 verifies SID isolation: you can see the distinct operator username and Session ID in the comparison card, and attempting to read the host employee's user profile is rejected by Windows NTFS ACLs.\n"
        "Step 3 keeps shell state alive while reading the staged ticket scope (`read-only-diagnostics`).\n"
        "Step 4 tests privilege boundaries: `Get-Service` on `Spooler` and `W32Time` succeeds, whereas `Stop-Service Spooler` is immediately denied because the temporary operator is non-admin.\n"
        "Step 5 writes `triage-evidence.txt` for independent host verification—and when running locally on port 8787, you can type live PowerShell commands right into the bottom prompt. Step 6 closes the PTY, stops the session, and calls `deprovisionContainer` to delete the temporary Windows account cleanly."
    ),
    26: (
        "Demo 3 showcases the experimental hardware-isolated tier: **Hyperlight MicroVM**. When do you reach for it? Suppose your platform executes arbitrary LLM-generated Python or third-party code where you don't even want to share the host Windows NT kernel—yet traditional VMs take seconds to cold-boot. "
        "Hyperlight runs a minimal Unikraft unikernel on Windows Hypervisor Platform and restores from a pre-warmed memory snapshot in ~90 milliseconds.\n"
        "We verify it in five steps: Step 1 checks the `1.1.0-alpha` policy and pre-warmed bash and Python 3.12 snapshots; Step 2 restores the guest and runs `uname -sr`, returning an independent `Unikraft 5.15` kernel rather than Windows; "
        "Step 3 runs untrusted invoice-reconciliation Python and returns structured JSON over stdout; Step 4 fires adversarial escape probes against `C:\\Windows`, host `USERPROFILE`, and outbound TCP sockets (`Errno 13`)—all blocked at the hardware air gap; "
        "and Step 5 proves instant memory reset: a file written to `/tmp` in Run A is completely gone (`FileNotFoundError`) when Run B restores the clean snapshot."
    ),
    27: (
        "Let's close with four takeaways: first, push guardrails down to the OS—prompts persuade, kernels enforce. Second, write one MXC policy and let it map onto native sandboxes across Windows, macOS, and Linux. "
        "Third, use Learning Mode to gather real access evidence, review with a human, and lock down least privilege. Fourth, containment isn't trust—always validate every output crossing back to the host.\n"
        "If you want a high-impact first pilot next week, start with a coding agent or an air-gapped sensitive-document workflow—wire up read-in/write-out paths and run both positive and negative tests in half a day. Thank you, and let's open it up for questions!"
    ),
}

ZH_TEXT_MAP = {
    # Slide 1 (idx 0)
    'MICROSOFT EXECUTION CONTAINERS (MXC)': 'MICROSOFT EXECUTION CONTAINERS (MXC) · 微软开源执行沙箱',
    'Guardrails with\nSandboxing': '用操作系统沙箱\n筑牢 Agent 护栏',
    'Keeping AI agents inside safe boundaries: isolating execution and controlling access to files, network and credentials on Windows': '把 AI Agent 关进安全边界：在 Windows 上隔离代码执行，精细管控文件、网络与敏感凭证访问',
    'Yingting Huang  |  Sr. Cloud Solution Architect, Microsoft': '黄颖庭 (Yingting Huang)  |  微软高级云解决方案架构师',
    'SCAN FOR LIVE SLIDES & DEMO': '扫码打开在线课件与交互 Demo',

    # Slide 2 (idx 1)
    'AGENDA': '议程 · AGENDA',
    'Our journey today': '今日分享脉络',
    'Why sandboxing': '为何需要沙箱',
    'Risks of agents that act': '自主执行带来的新风险',
    'Meet MXC': '认识 MXC',
    'What it is and how it works': '定位、概念与分层原理',
    'Features & platforms': '核心能力与平台',
    'Backends, lifecycle & rules': '隔离后端、生命周期与策略',
    'Use cases': '典型应用场景',
    'Coding, desktop & data agents': '编码、桌面与敏感数据 Agent',
    'Integration': '工程接入实战',
    'Architecture, SDKs & steps': '参考架构、三语言 SDK 与五步法',
    'Live demos': '三组真机实测',
    'Process, Session & MicroVM': '进程、独立会话与微虚拟机',
    'Takeaways & Q&A': '总结与问答',
    'What to do next': '落地清单与现场交流',

    # Slide 3 (idx 2)
    '01 · WHY SANDBOXING': '01 · 为什么必须上沙箱',
    "Agents don't just talk. They act.": 'Agent 不只是陪聊，它会直接动手执行',
    'HUMAN IN THE LOOP': '人在回路把关',
    'Chat assistant': '传统对话助手',
    'Answers questions': '只生成文字建议，不碰系统',
    '• Generates text and code suggestions': '• 仅输出文本回答与代码片段建议',
    '• A human reviews, copies and runs commands': '• 由真人审阅、复制并决定是否执行命令',
    'Worst case: a wrong answer': '最坏后果：答非所问',
    'DIRECT OS EXECUTION': '直接操作系统执行',
    'Autonomous agent': '自主执行 Agent',
    'Takes actions on your machine': '在你的电脑或服务器上直接动手',
    '• Reads and writes local workspace files': '• 自主读写本地工作区与项目文件',
    '• Runs shell commands & installs packages': '• 自己敲命令行、拉取并安装第三方包',
    '• Calls APIs and drives desktop apps': '• 调用外部接口并操控桌面软件',
    'Worst case: your whole session compromised': '最坏后果：当前账号权限被全面突破',
    'Agents execute code they did not write beforehand — and cannot fully predict at runtime.': 'Agent 运行的是临时拼装、无法提前穷举预测的动态命令 —— 绝不能直接顶着员工本人权限裸跑。',

    # Slide 4 (idx 3)
    'What can go wrong?': '一旦失控，具体会出什么乱子？',
    'Prompt injection': '提示词注入 (Prompt Injection)',
    "A poisoned README or web page tricks the agent into running an attacker's command.": '恶意 README 或网页夹带隐藏指令，诱骗 Agent 顺手执行攻击者的外部脚本。',
    'Credential theft': '本机凭证窃取 (Credential Theft)',
    'Generated code reads cloud tokens, SSH keys and saved browser secrets.': '生成的代码翻取本机云平台 Token、SSH 私钥或浏览器保存的登录凭证。',
    'Data exfiltration': '敏感数据外传 (Data Exfiltration)',
    'Sensitive local files quietly leave over the network to an unknown host.': '本机敏感文档与源码通过未受控的网络连接悄悄发往外部未知服务器。',
    'Overreach': '好心办坏事的越权操作 (Overreach)',
    'A "helpful" shortcut edits production config it was never meant to touch.': '没有外部攻击者，模型为了“最快修好报错”自作主张去改线上生产配置。',

    # Slide 5 (idx 4)
    'Prompts persuade. The OS enforces.': '提示词只能“劝导”，操作系统才能“强制执行”',
    '"An agent cannot be its own security authority."': '“绝不能让 Agent 自己当自己的安全管理员。”',
    'Windows Developer Blog  ·  October 7, 2026': 'Windows 开发者博客  ·  2026 年 10 月 7 日',
    'AI-level': 'AI 模型层护栏',
    'Persuasion only': '概率性劝导（防君子）',
    'System prompts': '系统提示词约束',
    'Content filters': '上下文内容过滤',
    'Tool rules': '工具调用规则',
    'Output checks': '模型输出复核',
    'Execution-level': '系统执行层护栏',
    'Kernel-enforced': '内核级硬隔离（兜底线）',
    'File allow / deny': '文件读写白/黑名单',
    'Network rules': '网络出站与回环管控',
    'Process isolation': '进程与桌面会话隔离',
    'Secrets outside': '密钥凭证隔离在箱外',
    'Defense in depth: the model decides what to try; the operating system decides what is possible.': '纵深防御铁律：模型决定它“想尝试什么”，操作系统内核决定它“物理上能做成什么”。',

    # Slide 6 (idx 5)
    '01 · WHY SANDBOXING · WINDOWS PLATFORM': '01 · 为什么必须上沙箱 · WINDOWS 平台底座',
    'Three Windows pillars for governed agents': 'Windows 受控 Agent 的三大操作系统级支柱',
    'Containment': '沙箱隔离',
    'Limit what an agent can access & do': '管住 Agent 能碰哪些资源、能做什么',
    'OS-enforced file, network, process & UI boundaries on Windows 11 and Windows 365 via open-source MXC (also macOS & Linux).': '通过开源 MXC 在 Windows 11 与 Windows 365 上强制执行文件、网络、进程与 UI 边界（同步支持 macOS / Linux）。',
    'THIS SESSION · GA': '本场核心 · 已正式发布',
    'Identity': '独立身份',
    'Tell agent activity apart from a person': '把 Agent 操作与真人操作彻底区分开',
    'Windows connects local on-device agents to Microsoft Entra Agent ID so actions carry an agent audit trail instead of a human login.': 'Windows 将本机 Agent 接入 Microsoft Entra Agent ID，赋予独立审计身份，不再借用员工个人账号。',
    'GA + Windows preview': '云端 GA + Windows 预览',
    'Manageability': '终端可管理性',
    'Govern access and monitor agents': '设备端盘点、审计与策略统一下发',
    'Windows extends Microsoft Agent 365 and Intune to endpoints to inventory local agents and distribute MXC sandbox policies.': '将 Microsoft Agent 365 与 Intune 延伸至终端设备，统一盘点本地 Agent 并下发 MXC 沙箱策略。',
    'Windows OS primitives wired into Microsoft’s enterprise plane: local MXC containment + Entra Agent ID + Agent 365 & Intune.': 'Windows 原生底座打通微软企业治理平面：本地 MXC 沙箱隔离 + Entra Agent ID + Agent 365 与 Intune。',

    # Slide 7 (idx 6)
    '02 · UNDERSTANDING MXC': '02 · 深入理解 MXC',
    'An open-source, policy-driven execution layer for untrusted or AI-generated workloads.': '面向不可信代码与 AI 生成任务的微软开源、策略驱动沙箱执行层。',
    'Open source on GitHub': 'GitHub MIT 协议开源',
    'Generally available since Oct 7, 2026': '2026 年 10 月 7 日正式发布',
    '3 SDKs': '3 套官方 SDK',
    '.NET, TypeScript & Rust': '.NET、TypeScript 与 Rust',
    '3 OSes': '3 大主流系统',
    'Windows, macOS & Linux': 'Windows、macOS 与 Linux',
    'One JSON policy, mapped to the right OS sandbox on every host.': '写一份统一的 JSON 权限策略，自动映射为各操作系统自带的原生轻量沙箱。',

    # Slide 8 (idx 7)
    'Five concepts to remember': '上手前先对齐五个核心概念',
    'Host app': '宿主应用 (Host)',
    'Your trusted app or agent runtime. Picks permissions, supplies inputs, handles results.': '你自己编写的唯一可信调度方：负责定权限、喂输入、校验产出。',
    'Workload': '工作负载 (Workload)',
    'The command, script or process that runs under containment.': '真正被关进沙箱里运行的那条命令、脚本或子进程。',
    'Policy': '权限策略 (Policy)',
    'The access you permit: readable files, writable folders, network destinations, UI.': '明确声明放行的权限清单：可读路径、可写目录、网络目标与 UI 规则。',
    'Container request': '容器请求 (Request)',
    'Combines the policy with the command, working directory, timeout and sandbox choice.': '把权限策略与具体命令行、工作目录、超时时间及沙箱类型打包成一次请求。',
    'Backend': '隔离后端 (Backend)',
    'The OS mechanism that enforces it: process sandbox, separate session or VM.': '底层真正立起隔离墙的系统机制：进程沙箱、独立会话或微虚拟机。',
    'The host decides. The policy describes. The backend enforces.': '核心铁律：宿主拍板定权 · 策略声明边界 · 内核强制执行。',

    # Slide 9 (idx 8)
    'Four layers, one boundary': '四层架构与唯一信任分界线',
    '01 · TRUSTED': '01 · 可信层',
    'Host app / agent runtime': '宿主应用 / Agent 运行时',
    'Classifies the task, chooses trusted policy defaults, and validates outputs & approvals.': '识别任务类型、选择受信任的最小权限策略模板，并严格校验所有返回结果与人工审批。',
    '02 · CONTRACT': '02 · 契约层',
    'MXC SDK & policy contract': 'MXC SDK 与标准策略契约',
    'Discovers platform support, serializes versioned policy config, and manages lifecycle.': '探测当前机器支持的后端，序列化版本化策略配置，并管理容器生命周期。',
    '03 · RUNTIME': '03 · 执行层',
    'MXC execution layer': 'MXC 核心执行引擎',
    'Validates the config, resolves a compatible backend, and maps to OS operations.': '校验策略合法性，匹配兼容的隔离后端，并将配置翻译为操作系统原生调用。',
    '04 · KERNEL': '04 · 内核层',
    'Containment backend': '操作系统隔离后端',
    'Creates the isolated environment and enforces file, network & session boundaries.': '创建隔离运行环境，在内核层强制拦截越界的文件、网络与桌面会话访问。',
    'TRUST BOUNDARY  ·  UNTRUSTED BELOW': '信任分界线  ·  红线以下皆为不可信执行区',
    'Rule #1: an agent may suggest a permission, but only trusted host logic or a human can grant it.': '第一铁律：Agent 只能“申请”权限，唯有红线之上的可信宿主逻辑或真人才能“批准”权限。',

    # Slide 10 (idx 9)
    'How one request flows': '单次沙箱请求的完整流转闭环',
    '1 · Discover': '1 · 探测能力',
    'Ask the SDK which backends this machine supports.': '调用 SDK 查询当前机器支持哪些原生隔离后端。',
    '2 · Describe': '2 · 声明策略',
    'Build a policy & execution request, deny by default.': '按“默认全部拒绝”原则组装最小权限策略与执行请求。',
    '3 · Launch': '3 · 启动容器',
    'Start the contained workload through the SDK.': '通过 SDK 拉起受控沙箱并投入工作负载。',
    '4 · Translate': '4 · 内核映射',
    'Map the request to native OS sandbox primitives.': 'MXC 将统一请求翻译为对应系统的内核级隔离原语。',
    '5 · Return': '5 · 验真清理',
    'Validate output, exit code, denials & clean up.': '宿主核验输出、退出码与拦截记录，并彻底清理环境。',
    'Backends have different security properties. Always test both the allowed and the denied operations.': '不同平台后端的隔离脾气各有差异 —— 自动化测试必须同时覆盖“放行成功”与“越权被拦”正反两条路径。',

    # Slide 11 (idx 10)
    '03 · FEATURES & PLATFORMS': '03 · 核心特性与跨平台支撑',
    'Pick the right box': '按业务风险挑选合适分量的隔离盒子',
    '←  LIGHTER  ·  FASTER STARTUP': '←  更轻量  ·  毫秒级极速启动',
    'STRONGER KERNEL ISOLATION  →': '更强硬件级内核隔离  →',
    'Process container': '进程级容器',
    'Windows 11 · macOS · Linux': 'Windows 11 · macOS · Linux',
    'AppContainer, Seatbelt or Bubblewrap. Millisecond startup and fine-grained file, network & UI rules — default for coding agents.': '映射 AppContainer / Seatbelt / Bubblewrap。毫秒级启动，文件、网络与 UI 规则最细 —— 编码 Agent 默认首选。',
    'Shared kernel · Process ACL': '共享内核 · 进程级 ACL',
    'GA · Default': '正式版 · 默认',
    'Session container': '独立会话容器',
    'Windows 11 only': '仅限 Windows 11',
    'Separate local Windows login and desktop for GUI apps; blocks clipboard/input bleed (does not filter paths or IPs).': '新建独立本地账号与桌面会话跑 GUI 软件，阻断剪贴板与键鼠串扰（不做细粒度路径或 IP 过滤）。',
    'Shared kernel · Own desktop': '共享内核 · 独立桌面',
    'GA · GUI': '正式版 · 桌面',
    'WSL container': 'WSL 容器',
    'Windows 11': 'Windows 11',
    'Runs Linux CLI workloads and toolchains on a Windows host; networking is all-isolated or bridged (no per-IP rules).': '在 Windows 宿主上隔离运行 Linux 命令行与工具链；网络仅支持全断网或桥接（无逐 IP 规则）。',
    'Isolated Linux subsystem': '隔离 Linux 子系统',
    'GA · Linux': '正式版 · Linux',
    'MicroVM': '硬件微虚拟机',
    'Windows 11 · Linux': 'Windows 11 · Linux',
    'Hardware-virtualized unikernel boundary with ~90 ms warm snapshot restore; never shares the host OS kernel.': '硬件虚拟化微内核边界，预热内存快照约 90 ms 恢复启动，绝不与宿主机共享操作系统内核。',
    'Dedicated VM kernel': '独立虚拟机内核',
    'Experimental': '实验特性',

    # Slide 12 (idx 11)
    'Where MXC runs today': 'MXC 当前跨平台后端支持矩阵',
    'HOST PLATFORM': '宿主平台环境',
    'DEFAULT PROCESS BACKEND': '默认进程隔离后端',
    'ADDITIONAL OPTIONS & CLOUD PC': '扩展隔离选项与云电脑支持',
    'Also supports IsolationSession, WSLc & Learning Mode reports': '同时支持 IsolationSession、WSLc 与 Learning Mode 报告',
    'Also supports LXC containers': '同时支持 LXC 容器后端',
    'Enforces native macOS sandbox profiles': '底层映射 macOS 原生 Seatbelt 沙箱配置',
    'Windows 365 Cloud PC': 'Windows 365 云电脑',
    'Works out of the box; session or VM backends need nested virtualization': '进程隔离直接开箱即用；会话或虚拟机后端需开启嵌套虚拟化',
    'Experimental (explicit opt-in via 1.1-alpha): Windows Sandbox, MicroVM & Hyperlight unikernel.': '实验性后端（需通过 1.1-alpha 显式开启）：Windows Sandbox、MicroVM 与 Hyperlight 微内核。',

    # Slide 13 (idx 12)
    'One-shot or state-aware': '生命周期选择：用完即焚 vs. 保留状态',
    'One-shot': '单次即焚 (One-shot)',
    'DEFAULT CHOICE': '80%+ 场景默认首选',
    'Create': '创建沙箱',
    'Run one workload': '执行单次任务',
    'Tear down': '立即销毁',
    'Builds, tests & generated commands: spin up a fresh, narrow ProcessContainer for every action.': '编译、测试与单条生成命令：每次动作现拉一个全新的极窄权限 ProcessContainer，跑完即焚。',
    'State-aware': '有状态会话 (State-aware)',
    'PERSISTENT STATE': '跨步保留现场',
    'Provision': '分配账号',
    'Start': '启动会话',
    'Execute': '多步执行',
    'Stop': '停止运行',
    'Deprovision': '注销清理',
    'ProvisionRequest → ContainerId on IsolationSession (session/UI only) or WSLc — always deprovision.': '先调 Provision 拿 ContainerId 跑多步任务（IsolationSession 或 WSLc）—— 结束务必调 deprovision 销号。',
    'CAPTURED RESULT': '同步收集输出',
    'Returns stdout, stderr & exit code once the workload finishes.': '等任务执行完毕后一次性拿回 stdout、stderr 与退出码。',
    'LIVE STREAMING': '实时流式输出',
    'Streams stdin, stdout and stderr in real time while running.': '运行过程中实时流式读写标准输入、标准输出与错误流。',
    'INTERACTIVE SHELL': '交互式伪终端',
    'Full pseudo-terminal session for stateful shells & human handoff.': '完整持久化伪终端会话，支持跨步保留上下文与真人工程师随时接管。',

    # Slide 14 (idx 13)
    'Files: read in, write out': '文件护栏：只读进、专口出',
    'mxc-config.1.0.0.json  ·  repo-scoped tool': 'mxc-config.1.0.0.json  ·  项目目录级受控工具',
    '01 · Read in, write out (& redirect TEMP)': '01 · 只读进、专口出（并重定向 TEMP）',
    'Keep source read-only; route outputs, tool caches, TEMP & LOCALAPPDATA to a dedicated mxc-output folder.': '源码目录锁死只读；将产出物、编译缓存、TEMP 与 LOCALAPPDATA 全部指向专用 mxc-output 目录。',
    '02 · Carve out secrets inside grants': '02 · 在放行目录内精准挖掉机密子路径',
    'Ungranted paths (~/.ssh) stay blocked by default; use deniedPaths for .env or secrets/ inside a granted repo.': '未授权路径（如 ~/.ssh）默认本就不可见；用 deniedPaths 在已放行只读的仓库里单独封死 .env 或 secrets/。',
    '03 · Teardown is not undo': '03 · 销毁沙箱不等于撤回宿主机写入',
    'Writes to granted host folders survive container exit — validate mxc-output as an untrusted staging zone.': '落入已授权宿主机目录的文件在沙箱退出后依然存在 —— 必须把 mxc-output 当作隔离待检区先验再收。',

    # Slide 15 (idx 14)
    'Network: deny by default': '网络护栏：默认断网，按需开白名单',
    'SANDBOXED WORKLOAD': '沙箱内受控进程',
    'Agent process': 'Agent 子进程',
    'Only explicit CIDR or proxy rules open a path.': '默认全面断网，仅显式网段或受控代理可放行。',
    'Direct egress.allow CIDR + port (no proxy)': 'egress.allow 直连受信任网段 + 端口',
    'Any other IP / loopback': '任何其他外网 IP / 本机端口',
    'Blocked at OS firewall (default: deny)': '在操作系统防火墙层直接阻断 (default: deny)',
    'App-started proxy via networkProxy': '经宿主启动的受控代理转发 (networkProxy)',
    'NETWORK POLICY PATTERNS': '官方网络放行设计要点',
    '• Direct CIDR: egress.allow[{ to:[{cidr}], ports }] opens TCP port, not one hostname — don’t mix with networkProxy.': '• 直连网段：egress.allow 按 IP 段和 TCP 端口放行（不认域名），切勿与 networkProxy 混配。',
    '• Hostname rules: start your host proxy first and point runtimeConfig.networkProxy to it.': '• 域名级白名单：先在宿主机外启动代理，再将 runtimeConfig.networkProxy 指向该代理。',
    '• Windows proxy peer: allowedProxyPeer binds a package family or AppContainer profile name (not a PID).': '• Windows 代理身份：allowedProxyPeer 填代理的包家族名或 AppContainer 配置名（而非 PID）。',
    '• Windows mapping: proxy path needs ingress.default: "allow" while hostLoopback: "deny" blocks general loopback.': '• Windows 映射细节：代理模式需设 ingress.default: "allow"，同时靠 hostLoopback: "deny" 封死其他本地端口。',

    # Slide 16 (idx 15)
    'Keys and UI stay outside': '密钥与员工桌面严格留在沙箱之外',
    'INSIDE THE SANDBOX': '沙箱内部 · 最小够用资源',
    '✓  Repo files (granted)': '✓  当前授权项目仓库文件',
    '✓  Build & test toolchain': '✓  编译与测试工具链',
    '✓  Approved API calls': '✓  经代理审核的外部接口',
    'OUTSIDE · OUT OF REACH': '沙箱外部 · 内核硬隔离禁区',
    '✕  API tokens & cloud keys': '✕  API Token 与云平台密钥',
    '✕  ~/.ssh private keys': '✕  ~/.ssh 本机私钥',
    '✕  Saved browser secrets': '✕  浏览器保存的密码与 Cookie',
    '✕  Employee clipboard': '✕  员工剪贴板内容',
    '✕  Host desktop & input': '✕  员工桌面窗口与键鼠输入',
    'Credentials stay on the host': '原始密钥绝不进沙箱',
    'Keep raw keys out of sandbox env & disk. Let your trusted host or proxy attach credentials outside.': '严禁把明文 Key 塞进沙箱环境变量或磁盘；由箱外可信宿主或代理验完请求后代挂凭证。',
    'Visible UI vs. full GUI desktop': '弹窗小工具 vs. 完整桌面软件',
    'Use SessionContainer + bot login for full desktop apps, or scope ProcessContainer UI explicitly:': '跑完整 GUI 客户端请用 SessionContainer + 专用账号；普通进程则通过 ui 字段精准锁死：',

    # Slide 17 (idx 16)
    'ui policy  ·  1.0.0 JSON / V1 SDK': 'ui 策略配置  ·  1.0.0 JSON / V1 SDK',
    '"ui": {\n  "disable": true,     // false = visible tool\n  "clipboard": "none", // or "read" only\n  "injection": false   // allowInputInjection\n}': '"ui": {\n  "disable": true,     // false = 允许弹窗工具\n  "clipboard": "none", // 或单向只读 "read"\n  "injection": false   // 禁止注入键鼠输入\n}',
    'Observe first, then enforce': '先观察探路，再开启强制拦截',
    'Enforcement': 'Enforcement 强制',
    'Applies the policy strictly; blocks any ungranted access.': '生产环境标准态：严格执行最小权限策略，任何未授权访问当场拦截。',
    'Learning': 'Learning 学习',
    'Blocks ungranted access AND records every attempt in an activity report.': '策略打磨利器：照常拦截越权访问，同时把每次被拒尝试完整记入活动报告。',
    'Permissive': 'Permissive 宽容',
    'Allows and logs ungranted access — temporary diagnostics only.': '放行未授权访问并记录日志 —— 仅用于临时疑难排障，切勿用于生产。',
    '01 · Observe': '01 · 实跑采集报告',
    '02 · Human Review': '02 · 工程师逐条甄别',
    '03 · Enforce': '03 · 锁死强制上线',
    'Learning Mode provides evidence, not auto-grants. A human approves every rule.': 'Learning Mode 只负责收集证据，绝不自动扩权 —— 每一条放行规则都必须由真人拍板。',

    # Slide 18 (idx 17)
    '04 · USE CASES': '04 · 企业典型落地场景',
    'Secure coding agents': '安全受控的编码 Agent',
    'Build, test and push — without touching anything outside the repo.': '放心编译、跑单测、提代码 —— 不触碰仓库外的任何文件与凭证。',
    'Project repo  C:\\src\\app': '当前项目仓库  C:\\src\\app',
    'Read / write': '允许读写',
    'Toolchain: node, python, git': '开发工具链：node、python、git',
    'Allowed': '允许调用',
    'Package registry via local proxy': '经受控本地代理拉取依赖包',
    '~/.ssh, cloud tokens & Documents': '~/.ssh 私钥、云凭证与个人文档',
    'Denied': '内核拒绝',
    'Any other network destination': '任何未经审批的外部网络地址',
    'Blocked': '防火墙阻断',
    'In production today: GitHub Copilot CLI enforces command isolation via MXC.': '大规模生产验证：GitHub Copilot CLI 底层正是通过 MXC 进程隔离来约束动态生成的命令。',

    # Slide 19 (idx 18)
    'Desktop and data agents': '桌面自动化与敏感文档处理 Agent',
    'Desktop automation agent': '桌面 GUI 自动化 Agent',
    'Keeps GUI app state in a separate Windows account and session, isolated from the employee desktop.': '在独立的 Windows 本地账号与桌面会话中操控传统软件，与员工前台办公桌面彻底隔离。',
    'Best fit: Session container + dedicated bot login': '最佳匹配：Session 容器 + 低权限专用机器人账号',
    'Sensitive-document agent': '敏感合同与财报处理 Agent',
    'Reads one approved input folder, writes one summary report, with network completely disabled.': '只读一个指定输入文件夹、只写一份汇总报告，网络全程彻底封死，物理掐断外传风险。',
    'Best fit: Process container (read-in / write-out, egress deny)': '最佳匹配：Process 容器（只读进/专口出 + 全程断网）',
    'Agents on Windows 365 Cloud PCs': '运行在 Windows 365 云电脑上的 Agent',
    'Runs alongside existing Cloud PC workloads; process isolation works out of the box just like physical PCs.': '与员工云桌面并排运行，ProcessContainer 进程隔离无需额外虚拟化配置即可开箱即用。',
    'Best fit: Process container on Windows 365 (GA)': '最佳匹配：Windows 365 上的 Process 容器（已正式发布）',

    # Slide 20 (idx 19)
    'Already in the ecosystem': '快速收敛的行业工具链生态',
    'Industry convergence: coding CLIs, local model runners & enterprise SaaS agents standardize on OS-native containment.': '行业标准收敛：主流编程 CLI、本地模型运行器与企业 SaaS Agent 正统一接入操作系统原生隔离底座。',
    'SUPPORTS MXC TODAY': '今日已原生支持 MXC',
    'INTEGRATED / BUILDING': '已集成 / 深度共建中',
    'NVIDIA OpenShell\nOpenAI and Meta': 'NVIDIA OpenShell\nOpenAI 与 Meta',
    'SUPPORT ANNOUNCED': '已官宣接入支持',

    # Slide 21 (idx 20)
    '05 · INTEGRATING MXC': '05 · 工程接入与架构落地',
    'Reference architecture': '企业级 Agent 平台接入参考架构',
    '01 · PLANNER': '01 · 规划大脑',
    'Agent planner': 'Agent 规划器',
    'LLM proposes next action': '大模型提议下一步动作',
    '02 · INTENT': '02 · 调用意图',
    'Tool call': '工具调用请求',
    '03 · TRUSTED HOST': '03 · 可信宿主',
    'Policy engine': '宿主策略引擎',
    'Selects least-privilege policy': '匹配预审的最小权限模板',
    '04 · CONTRACT': '04 · 标准契约',
    'Lifecycle & diagnostics': '生命周期管理与拦截诊断',
    '05 · OS KERNEL': '05 · 系统内核',
    'Sandbox backend': '原生隔离后端',
    'Enforces hard boundary': '内核层强制执行硬边界',
    '←   RETURN PATH  ·  Stdout, exit code & kernel denials validated by trusted host before feeding back to agent': '←   回传校验闭环  ·  标准输出、退出码与内核拦截记录均须经可信宿主验真后，才允许喂回给大模型',
    'Egress proxy': '箱外受控出站代理',
    'Attaches credentials and enforces domain allow-lists strictly outside the sandbox.': '在沙箱外部代挂 API 凭证并强制执行域名级白名单过滤。',
    'Learning Mode': 'Learning Mode 权限收敛',
    'Captures blocked-access evidence so humans can tighten policies before rollout.': '在上线前捕获真实的被拒访问证据，辅助工程师精准收敛最小权限。',
    'Enterprise governance': '企业级身份与全局治理',
    'Entra Agent ID & Agent 365 govern agent identity, audit trails, and fleet policy.': '结合 Entra Agent ID 与 Agent 365 统一管理 Agent 身份、审计轨迹与终端策略下发。',

    # Slide 22 (idx 21)
    'Three SDKs, one contract': '三套官方语言 SDK，同一份标准契约',
    'CAPABILITY DISCOVERY': '平台能力探测入口',
    'Write one ContainerRequest in code — the SDK maps it to the native OS sandbox across Windows, macOS & Linux.': '在业务代码里编写统一的 ContainerRequest —— SDK 自动将其映射为 Windows、macOS 与 Linux 原生沙箱。',
    'Requirements: Windows 11 24H2+ (physical or Windows 365), macOS, or Linux  ·  .NET 8+ / Node.js 18+ / Rust.': '环境要求：Windows 11 24H2+（物理机或 Windows 365）、macOS 或 Linux  ·  .NET 8+ / Node.js 18+ / Rust。',

    # Slide 23 (idx 22)
    'const result = await run(request); // validate output next!': 'const result = await run(request); // 下一步务必校验输出！',
    'Integrate in five steps': '五步闭环把 MXC 接进你的项目',
    'Discover': '1 · 探测环境 (Discover)',
    'Call getPlatformSupport() and probe(request) before launch.': '启动前先调 getPlatformSupport() 与 probe(request)，不支持则报错熔断。',
    'Describe': '2 · 声明策略 (Describe)',
    'Build a V1 ContainerRequest: files, network, ui & timeoutMs.': '按默认拒绝组装 V1 ContainerRequest：锁死文件、网络、UI 与超时。',
    'Learn': '3 · 学习探路 (Learn)',
    'Run in Learning Mode first; review the blocked-access report.': '开发期先开 Learning Mode 实跑，由人工逐条审阅被拒访问报告。',
    'Enforce': '4 · 强制验证 (Enforce)',
    'Switch to Enforcement; verify allowed AND denied cases.': '切回 Enforcement 强制模式，正反测试“合规放行”与“越权拦截”。',
    'Operate': '5 · 稳健运维 (Operate)',
    'Handle timeouts & cleanup across success, failure & cancel.': '在成功、失败与取消所有分支中兜住超时控制与临时资源清理。',

    # Slide 24 (idx 23)
    "Your host's job list": '可信宿主程序的六项必修职责',
    'Set an explicit timeout; handle launch error, workload failure & cancellation separately.': '每次调用必须设显式超时，并将“沙箱启动失败”、“内部脚本报错”与“用户取消”分开处理。',
    'Clean up even when a later step throws — always deprovision persistent sessions.': '无论哪一步抛异常都必须走清理逻辑 —— 持久化会话务必调 deprovision 注销临时账号。',
    'Spin up one narrow sandbox per action rather than sharing one broad policy.': '坚持“一次动作起一个极窄短命沙箱”，切忌图省事共用一个大权限长寿容器。',
    'Write outputs to a dedicated staging folder; teardown never undoes host writes.': '产出统一写入专用隔离目录；记住销毁沙箱不会撤回已落盘宿主机目录的文件。',
    'Explain policy denials clearly to users and agents; never fail silently.': '遇到策略拦截时向用户和 Agent 清晰说明被拒原因，绝不静默吞掉错误。',
    'Require a human to approve every grant suggested by Learning Mode.': 'Learning Mode 跑出的任何放行建议，在合入生产策略前都必须经过真人审批。',

    # Slide 25 (idx 24)
    '06 · LIVE DEMO 1': '06 · 现场实测 DEMO 1',
    'ProcessContainer: files & network': 'ProcessContainer：文件与网络内核闸门',
    '✓  Scoped read-in / write-out': '✓  最小权限合规读写放行',
    'Workload reads demo/input and writes report.txt into its dedicated output folder.': '工作负载按策略读取 demo/input 并将 report.txt 写入专用输出目录，返回码 0 顺利通过。',
    '✕  Protected file denied': '✕  机密文件被内核当场拦截',
    'The same workload attempting to read a secret file is blocked by the OS kernel.': '同一份策略下试图偷读 protected/denied.txt 机密文件，被 Windows 内核 ACL 直接拒绝。',
    '⇄  Directional network leash': '⇄  外网封死下的回环精准开关',
    'Reaches an explicitly permitted loopback endpoint; blocks unapproved outbound traffic.': '公网全程断网；开启 hostLoopback 时打通本地监听端口 (200)，关闭后同命令立即被拒 (exit 7)。',

    # Slide 26 (idx 25)
    '06 · LIVE DEMO 2': '06 · 现场实测 DEMO 2',
    'Isolated incident triage': '一线运维隔离排障实战',
    '01 · Stage': '01 · 投放工单与只读授权',
    'Host stages INC-1042 and its read-only approval inside the ephemeral workspace.': '可信宿主将故障工单 INC-1042 与只读诊断授权放入专用交接目录。',
    '02 · Open PTY': '02 · 创建临时账号与独立 PTY',
    'Provisions a separate local Windows operator account & isolated Session.': '现场新建独立的本地 Windows 操作员账号与隔离会话，拉起持久化 PowerShell 终端。',
    '03 · Triage': '03 · 服务只读排障与越权拦截',
    'Checks Spooler & W32Time; blocks cross-user reads and service-stop attempts.': '正常查询 Spooler 与 W32Time 状态；偷读员工个人目录与 Stop-Service 停服务操作均被系统拒绝。',
    '04 · Teardown': '04 · 宿主验真与销号清理',
    'Host validates evidence, stops the session, and deprovisions the temp account.': '宿主独立核验排查报告无误后停止会话，并调用 deprovision 将临时操作员账号彻底注销删除。',

    # Slide 27 (idx 26)
    '06 · LIVE DEMO 3': '06 · 现场实测 DEMO 3',
    'Untrusted code in a MicroVM': 'Hyperlight MicroVM：微虚拟机跑陌生代码',
    'Warm snapshot restore: ~90 ms · Unikraft 5.15': '预热内存快照恢复：~90 ms · Unikraft 5.15 微内核',
    '✓  Hardware VM kernel': '✓  硬件虚拟化独立微内核',
    'Unikraft 5.15 on Hypervisor': 'Hypervisor 之上的 Unikraft 5.15',
    'Runs inside a hardware-virtualized partition, never sharing the Windows NT kernel.': '运行在硬件虚拟化隔离分区内，绝不与宿主机共用 Windows NT 内核。',
    '✕  Zero host surface': '✕  零宿主机暴露面（三路逃逸全拦）',
    'No host FS, env, or sockets': '无宿主磁盘、无宿主环境变量、无网络栈',
    'Host C:\\, environment variables, and TCP sockets (Errno 13) remain unreachable.': '恶意探针试探宿主机 C:\\、环境变量与外部 TCP 连接 (Errno 13) 全部撞墙失败。',
    '↺  Instant snapshot wipe': '↺  毫秒级快照恢复与瞬时清零',
    '<100 ms restore & reset': '<100 ms 恢复启动，退出即蒸发',
    'Every call restores a warmed memory snapshot; /tmp writes vanish on exit.': '每次调用均从干净的预热内存快照恢复；上一次写入 /tmp 的残留文件在下次调用时瞬间归零。',

    # Slide 28 (idx 27)
    'KEY TAKEAWAYS': '核心总结 · KEY TAKEAWAYS',
    'Box the agent,\nnot the ambition': '管住 Agent 的边界，\n释放 Agent 的潜能',
    "1   Push guardrails down to the OS. Prompts alone can't enforce.": '1   把护栏下沉到操作系统内核，光靠提示词防不住。',
    '2   MXC: one policy, the right sandbox on Windows, macOS & Linux.': '2   MXC 统一策略：自动对接 Win / macOS / Linux 沙箱。',
    '3   Learn first, then enforce least privilege.': '3   先开 Learning Mode 探路审阅，再锁死最小权限。',
    "4   Containment isn't trust. Validate every result.": '4   沙箱隔离不等于结果可信，返回宿主后务必再验。',
    'Thank you  ·  Q&A': '谢谢大家  ·  问答交流 (Q&A)',
}

ZH_DEMO_LAUNCHERS = {
    24: ('process', '1', '启动 Demo 1 · ProcessContainer'),
    25: ('pty', '1', '启动 Demo 2 · IsolationSession PTY'),
    26: ('microvm', '1', '启动 Demo 3 · Hyperlight MicroVM'),
}
