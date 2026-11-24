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
    'MICROSOFT EXECUTION CONTAINERS (MXC)': 'MICROSOFT EXECUTION CONTAINERS (MXC) · 开源沙箱执行层',
    'Guardrails with\nSandboxing': '用系统级沙箱\n隔离 Agent 执行风险',
    'Keeping AI agents inside safe boundaries: isolating execution and controlling access to files, network and credentials on Windows': '在 Windows 上隔离 AI Agent 的命令执行，按策略控制文件、网络与凭证访问',
    'SCAN FOR LIVE SLIDES & DEMO': '扫码查看在线幻灯片与交互演示',

    # Slide 2 (idx 1)
    'AGENDA': '目录 · AGENDA',
    'Our journey today': '今天的内容安排',
    'Why sandboxing': '为什么需要沙箱',
    'Risks of agents that act': '自主执行带来的风险',
    'Meet MXC': '认识 MXC',
    'What it is and how it works': '定位、概念与分层结构',
    'Features & platforms': '主要功能与平台',
    'Backends, lifecycle & rules': '隔离方式、生命周期与规则',
    'Use cases': '常见应用场景',
    'Coding, desktop & data agents': '编码、桌面与文档处理场景',
    'Integration': '接入方式',
    'Architecture, SDKs & steps': '参考架构、SDK 与接入步骤',
    'Live demos': '现场演示',
    'Process, Session & MicroVM': '进程、独立会话与微虚拟机',
    'Takeaways & Q&A': '总结与问答',
    'What to do next': '落地建议与交流',

    # Slide 3 (idx 2)
    '01 · WHY SANDBOXING': '01 · 为什么需要沙箱',
    "Agents don't just talk. They act.": 'Agent 不只是回答问题，还会直接执行操作',
    'HUMAN IN THE LOOP': '由人确认后执行',
    'Chat assistant': '普通对话助手',
    'Answers questions': '只生成文字建议',
    '• Generates text and code suggestions': '• 只输出文字回复和代码建议',
    '• A human reviews, copies and runs commands': '• 由人先看一遍，再决定要不要复制运行',
    'Worst case: a wrong answer': '最坏结果：给错回答',
    'DIRECT OS EXECUTION': '直接调用系统执行',
    'Autonomous agent': '自主执行的 Agent',
    'Takes actions on your machine': '在本机或服务器上直接执行动作',
    '• Reads and writes local workspace files': '• 直接读写本地目录里的文件',
    '• Runs shell commands & installs packages': '• 自动运行命令行、安装第三方包',
    '• Calls APIs and drives desktop apps': '• 调用外部接口、操作桌面软件',
    'Worst case: your whole session compromised': '最坏结果：当前用户权限下的资源受损或泄露',
    'Agents execute code they did not write beforehand — and cannot fully predict at runtime.': 'Agent 运行的命令往往是在执行过程中动态生成的，无法提前逐条审核。',

    # Slide 4 (idx 3)
    'What can go wrong?': '没有隔离边界时常见的问题',
    'Prompt injection': '提示词注入 (Prompt Injection)',
    "A poisoned README or web page tricks the agent into running an attacker's command.": '外部 README 或网页里夹带隐藏指令，诱导 Agent 运行非预期的命令。',
    'Credential theft': '读取本机凭证 (Credential Theft)',
    'Generated code reads cloud tokens, SSH keys and saved browser secrets.': '生成的脚本读取本机保存的云平台 Token、SSH 私钥或浏览器登录信息。',
    'Data exfiltration': '敏感文件外发 (Data Exfiltration)',
    'Sensitive local files quietly leave over the network to an unknown host.': '本地敏感文档或源码通过网络被发送到未授权的外部地址。',
    'Overreach': '超出预期的越界修改 (Overreach)',
    'A "helpful" shortcut edits production config it was never meant to touch.': '没有外部攻击，模型为了尽快修好报错而擅自修改了不该动的配置。',

    # Slide 5 (idx 4)
    'Prompts persuade. The OS enforces.': '提示词只能引导行为，操作系统才能强制限制',
    '"An agent cannot be its own security authority."': '“Agent 不能自己充当自己的权限管理员。”',
    'Windows Developer Blog  ·  October 7, 2026': 'Windows 开发者博客  ·  2026 年 10 月 7 日',
    'AI-level': 'AI 模型层措施',
    'Persuasion only': '行为引导（减少出错概率）',
    'System prompts': '系统提示词约定',
    'Content filters': '输入输出内容过滤',
    'Tool rules': '工具调用说明',
    'Output checks': '结果检查',
    'Execution-level': '系统执行层隔离',
    'Kernel-enforced': '系统级强制限制（权限边界）',
    'File allow / deny': '文件读写路径控制',
    'Network rules': '网络连接规则',
    'Process isolation': '进程与桌面会话隔离',
    'Secrets outside': '密钥保留在沙箱外',
    'Defense in depth: the model decides what to try; the operating system decides what is possible.': '分层防护：模型决定它想尝试什么操作，操作系统决定哪些操作真正被允许。',

    # Slide 6 (idx 5)
    '01 · WHY SANDBOXING · WINDOWS PLATFORM': '01 · 为什么需要沙箱 · WINDOWS 平台能力',
    'Three Windows pillars for governed agents': 'Windows 管控本地 Agent 的三项基础能力',
    'Containment': '运行隔离',
    'Limit what an agent can access & do': '限制 Agent 能访问哪些本地资源',
    'OS-enforced file, network, process & UI boundaries on Windows 11 and Windows 365 via open-source MXC (also macOS & Linux).': '通过开源 MXC 在 Windows 11 与 Windows 365 上限制文件、网络、进程和界面访问（同时支持 macOS 与 Linux）。',
    'THIS SESSION · GA': '本次重点 · 已正式发布',
    'Identity': '独立身份',
    'Tell agent activity apart from a person': '区分 Agent 操作与员工本人操作',
    'Windows connects local on-device agents to Microsoft Entra Agent ID so actions carry an agent audit trail instead of a human login.': '将本机运行的 Agent 关联到 Microsoft Entra Agent ID，单独记录审计日志，避免混用员工个人账号。',
    'GA + Windows preview': '云端已发布 · Windows 预览中',
    'Manageability': '统一管理',
    'Govern access and monitor agents': '在终端侧清点 Agent 并下发策略',
    'Windows extends Microsoft Agent 365 and Intune to endpoints to inventory local agents and distribute MXC sandbox policies.': '结合 Microsoft Agent 365 与 Intune 查看设备上的本地 Agent，并集中下发 MXC 沙箱规则。',
    'Windows OS primitives wired into Microsoft’s enterprise plane: local MXC containment + Entra Agent ID + Agent 365 & Intune.': '将 Windows 本地执行层接入企业管理体系：本地 MXC 隔离 + Entra Agent ID + Agent 365 与 Intune。',

    # Slide 7 (idx 6)
    '02 · UNDERSTANDING MXC': '02 · 认识 MXC',
    'An open-source, policy-driven execution layer for untrusted or AI-generated workloads.': '微软开源的轻量沙箱执行层，用于按策略安全运行不可信或 AI 生成的代码。',
    'Open source on GitHub': 'GitHub MIT 协议开源',
    'Generally available since Oct 7, 2026': '2026 年 10 月 7 日正式发布',
    '3 SDKs': '3 种语言 SDK',
    '.NET, TypeScript & Rust': '.NET、TypeScript 与 Rust',
    '3 OSes': '3 类操作系统',
    'Windows, macOS & Linux': 'Windows、macOS 与 Linux',
    'One JSON policy, mapped to the right OS sandbox on every host.': '使用统一的 JSON 策略描述权限，由 MXC 自动调用当前操作系统自带的沙箱机制。',

    # Slide 8 (idx 7)
    'Five concepts to remember': '五个基本概念',
    'Host app': '宿主应用 (Host)',
    'Your trusted app or agent runtime. Picks permissions, supplies inputs, handles results.': '你编写的可信主程序：负责设定权限、准备输入文件并检查返回结果。',
    'Workload': '工作负载 (Workload)',
    'The command, script or process that runs under containment.': '放在沙箱里实际运行的那条命令、脚本或进程。',
    'Policy': '权限策略 (Policy)',
    'The access you permit: readable files, writable folders, network destinations, UI.': '明确允许的访问范围：可读文件、可写目录、网络地址与界面权限。',
    'Container request': '容器请求 (Request)',
    'Combines the policy with the command, working directory, timeout and sandbox choice.': '把权限策略与具体的命令行、工作目录、超时时间和沙箱类型组合在一起。',
    'Backend': '隔离后端 (Backend)',
    'The OS mechanism that enforces it: process sandbox, separate session or VM.': '操作系统底层实际执行隔离的机制：进程沙箱、独立会话或微虚拟机。',
    'The host decides. The policy describes. The backend enforces.': '基本分工：宿主程序决定权限，策略文件描述范围，系统后端负责执行。',

    # Slide 9 (idx 8)
    'Four layers, one boundary': '四层结构与信任边界',
    '01 · TRUSTED': '01 · 可信层',
    'Host app / agent runtime': '宿主应用 / Agent 主程序',
    'Classifies the task, chooses trusted policy defaults, and validates outputs & approvals.': '根据任务类型选择预设的最小权限模板，并在执行后校验返回结果和审批状态。',
    '02 · CONTRACT': '02 · 接口层',
    'MXC SDK & policy contract': 'MXC SDK 与策略定义',
    'Discovers platform support, serializes versioned policy config, and manages lifecycle.': '查询当前系统的支持情况，生成标准格式的策略配置，并管理容器启停。',
    '03 · RUNTIME': '03 · 运行时',
    'MXC execution layer': 'MXC 执行层',
    'Validates the config, resolves a compatible backend, and maps to OS operations.': '检查配置格式，选择匹配的隔离后端，并转换为对应的操作系统调用。',
    '04 · KERNEL': '04 · 系统层',
    'Containment backend': '操作系统隔离后端',
    'Creates the isolated environment and enforces file, network & session boundaries.': '创建隔离环境，在系统层面限制文件、网络和桌面会话访问。',
    'TRUST BOUNDARY  ·  UNTRUSTED BELOW': '信任边界  ·  下方为受限的不可信执行环境',
    'Rule #1: an agent may suggest a permission, but only trusted host logic or a human can grant it.': '原则：Agent 可以请求某项权限，但只能由可信的宿主代码或人工确认来批准。',

    # Slide 10 (idx 9)
    'How one request flows': '一次沙箱请求的执行过程',
    '1 · Discover': '1 · 查询环境',
    'Ask the SDK which backends this machine supports.': '通过 SDK 查询当前机器支持哪些沙箱后端。',
    '2 · Describe': '2 · 配置策略',
    'Build a policy & execution request, deny by default.': '按默认拒绝的原则配置所需权限和执行参数。',
    '3 · Launch': '3 · 启动执行',
    'Start the contained workload through the SDK.': '调用 SDK 在沙箱内启动目标命令。',
    '4 · Translate': '4 · 系统映射',
    'Map the request to native OS sandbox primitives.': 'MXC 将请求转换为当前操作系统的原生隔离机制。',
    '5 · Return': '5 · 校验清理',
    'Validate output, exit code, denials & clean up.': '宿主程序检查输出、退出码与拦截记录，并清理临时资源。',
    'Backends have different security properties. Always test both the allowed and the denied operations.': '不同平台的沙箱细节存在差异，测试时建议同时验证“允许的操作能跑通”和“禁止的操作确实被拦截”。',

    # Slide 11 (idx 10)
    '03 · FEATURES & PLATFORMS': '03 · 主要功能与支持平台',
    'Pick the right box': '根据场景选择合适的隔离方式',
    '←  LIGHTER  ·  FASTER STARTUP': '←  开销更小  ·  启动更快',
    'STRONGER KERNEL ISOLATION  →': '隔离程度更高  →',
    'Process container': '进程级容器',
    'Windows 11 · macOS · Linux': 'Windows 11 · macOS · Linux',
    'AppContainer, Seatbelt or Bubblewrap. Millisecond startup and fine-grained file, network & UI rules — default for coding agents.': '对应 AppContainer、Seatbelt 或 Bubblewrap。毫秒级启动，支持细致的文件、网络和界面规则，适合大多数编码 Agent。',
    'Shared kernel · Process ACL': '共享内核 · 进程级权限控制',
    'GA · Default': '正式版 · 默认',
    'Session container': '会话级容器',
    'Windows 11 only': '仅支持 Windows 11',
    'Separate local Windows login and desktop for GUI apps; blocks clipboard/input bleed (does not filter paths or IPs).': '创建独立的本地 Windows 账号和桌面来运行 GUI 程序，隔离剪贴板与键鼠输入（不做细粒度路径或 IP 限制）。',
    'Shared kernel · Own desktop': '共享内核 · 独立桌面会话',
    'GA · GUI': '正式版 · 桌面',
    'WSL container': 'WSL 容器',
    'Windows 11': 'Windows 11',
    'Runs Linux CLI workloads and toolchains on a Windows host; networking is all-isolated or bridged (no per-IP rules).': '在 Windows 上隔离运行 Linux 命令行工具链；网络支持完全断网或桥接模式（不支持按单个 IP 过滤）。',
    'Isolated Linux subsystem': '独立 Linux 子系统',
    'GA · Linux': '正式版 · Linux',
    'MicroVM': '微虚拟机',
    'Windows 11 · Linux': 'Windows 11 · Linux',
    'Hardware-virtualized unikernel boundary with ~90 ms warm snapshot restore; never shares the host OS kernel.': '基于硬件虚拟化运行独立微内核，从预热内存快照恢复约需 90 ms，不与宿主机共用系统内核。',
    'Dedicated VM kernel': '独立虚拟机内核',
    'Experimental': '实验特性',

    # Slide 12 (idx 11)
    'Where MXC runs today': '各操作系统平台支持的隔离后端',
    'HOST PLATFORM': '操作系统平台',
    'DEFAULT PROCESS BACKEND': '默认进程沙箱后端',
    'ADDITIONAL OPTIONS & CLOUD PC': '其他可选后端与云电脑说明',
    'Also supports IsolationSession, WSLc & Learning Mode reports': '同时支持 IsolationSession、WSLc 与学习模式报告',
    'Also supports LXC containers': '同时支持 LXC 容器后端',
    'Enforces native macOS sandbox profiles': '调用 macOS 原生 Seatbelt 沙箱机制',
    'Windows 365 Cloud PC': 'Windows 365 云电脑',
    'Works out of the box; session or VM backends need nested virtualization': '进程隔离可直接使用；会话或虚拟机后端需开启嵌套虚拟化',
    'Experimental (explicit opt-in via 1.1-alpha): Windows Sandbox, MicroVM & Hyperlight unikernel.': '实验特性（需通过 1.1-alpha 显式开启）：Windows Sandbox、MicroVM 与 Hyperlight 微内核。',

    # Slide 13 (idx 12)
    'One-shot or state-aware': '容器生命周期：单次运行与保留状态',
    'One-shot': '单次运行 (One-shot)',
    'DEFAULT CHOICE': '大多数场景的默认选择',
    'Create': '创建容器',
    'Run one workload': '运行单条命令',
    'Tear down': '销毁清理',
    'Builds, tests & generated commands: spin up a fresh, narrow ProcessContainer for every action.': '适合编译、测试等独立任务：每次执行新建一个最小权限的 ProcessContainer，运行完立即销毁。',
    'State-aware': '有状态会话 (State-aware)',
    'PERSISTENT STATE': '跨步骤保留环境状态',
    'Provision': '创建环境',
    'Start': '启动会话',
    'Execute': '多次执行',
    'Stop': '停止会话',
    'Deprovision': '注销清理',
    'ProvisionRequest → ContainerId on IsolationSession (session/UI only) or WSLc — always deprovision.': '先通过 Provision 获取 ContainerId 再执行多步命令（适用于 IsolationSession 或 WSLc），用完务必调用 deprovision 清理。',
    'CAPTURED RESULT': '等待执行结果',
    'Returns stdout, stderr & exit code once the workload finishes.': '命令运行结束后，一次性返回标准输出、错误输出和退出码。',
    'LIVE STREAMING': '实时流式读写',
    'Streams stdin, stdout and stderr in real time while running.': '在运行过程中实时读写标准输入、标准输出和错误信息。',
    'INTERACTIVE SHELL': '交互式终端',
    'Full pseudo-terminal session for stateful shells & human handoff.': '提供完整的伪终端会话，便于多步排查或在需要时交由人工继续操作。',

    # Slide 14 (idx 13)
    'Files: read in, write out': '文件访问控制：源码只读，输出放到单独目录',
    'mxc-config.1.0.0.json  ·  repo-scoped tool': 'mxc-config.1.0.0.json  ·  仓库范围只读与单独输出目录示例',
    '01 · Read in, write out (& redirect TEMP)': '01 · 输入只读，输出与临时文件写入专用目录',
    'Keep source read-only; route outputs, tool caches, TEMP & LOCALAPPDATA to a dedicated mxc-output folder.': '把代码目录设为只读；将输出结果、工具缓存、TEMP 和 LOCALAPPDATA 指向单独的 mxc-output 目录。',
    '02 · Carve out secrets inside grants': '02 · 在已开放的目录里排除敏感文件',
    'Ungranted paths (~/.ssh) stay blocked by default; use deniedPaths for .env or secrets/ inside a granted repo.': '未开放的路径（如 ~/.ssh）默认无法访问；如果只读仓库里包含 .env 或 secrets/，可用 deniedPaths 单独禁止。',
    '03 · Teardown is not undo': '03 · 销毁沙箱不会回滚已写入宿主目录的文件',
    'Writes to granted host folders survive container exit — validate mxc-output as an untrusted staging zone.': '写进宿主机授权目录的文件在容器退出后仍会保留，宿主程序使用前应先检查 mxc-output 中的内容。',

    # Slide 15 (idx 14)
    'Network: deny by default': '网络访问控制：默认禁止联网，按需放行',
    'SANDBOXED WORKLOAD': '沙箱内进程',
    'Agent process': 'Agent 进程',
    'Only explicit CIDR or proxy rules open a path.': '默认阻断网络连接，仅放行明确配置的网段或代理。',
    'Direct egress.allow CIDR + port (no proxy)': '通过 egress.allow 放行指定网段与端口',
    'Any other IP / loopback': '其他外部地址与本地端口',
    'Blocked at OS firewall (default: deny)': '由系统防火墙默认拦截 (default: deny)',
    'App-started proxy via networkProxy': '通过宿主启动的本地代理访问 (networkProxy)',
    'NETWORK POLICY PATTERNS': '网络规则配置要点',
    '• Direct CIDR: egress.allow[{ to:[{cidr}], ports }] opens TCP port, not one hostname — don’t mix with networkProxy.': '• 直连网段：egress.allow 按 IP 段和端口放行（不按域名匹配），不要与 networkProxy 同时混用。',
    '• Hostname rules: start your host proxy first and point runtimeConfig.networkProxy to it.': '• 按域名控制：先在沙箱外启动代理服务，再把 runtimeConfig.networkProxy 指向该代理。',
    '• Windows proxy peer: allowedProxyPeer binds a package family or AppContainer profile name (not a PID).': '• Windows 代理标识：allowedProxyPeer 填写代理的包家族名或 AppContainer 名称，不是进程 PID。',
    '• Windows mapping: proxy path needs ingress.default: "allow" while hostLoopback: "deny" blocks general loopback.': '• Windows 配置细节：代理模式需设 ingress.default: "allow"，同时用 hostLoopback: "deny" 拦住其他本地端口。',

    # Slide 16 (idx 15)
    'Keys and UI stay outside': '把密钥和员工桌面留在沙箱外部',
    'INSIDE THE SANDBOX': '沙箱内部 · 允许访问的资源',
    '✓  Repo files (granted)': '✓  授权范围内的项目文件',
    '✓  Build & test toolchain': '✓  编译与测试工具链',
    '✓  Approved API calls': '✓  经代理转发的受控接口',
    'OUTSIDE · OUT OF REACH': '沙箱外部 · 禁止访问的内容',
    '✕  API tokens & cloud keys': '✕  API Token 与云平台密钥',
    '✕  ~/.ssh private keys': '✕  ~/.ssh 私钥文件',
    '✕  Saved browser secrets': '✕  浏览器保存的登录信息',
    '✕  Employee clipboard': '✕  员工剪贴板内容',
    '✕  Host desktop & input': '✕  宿主机桌面窗口与键鼠输入',
    'Credentials stay on the host': '不要把原始密钥放进沙箱',
    'Keep raw keys out of sandbox env & disk. Let your trusted host or proxy attach credentials outside.': '避免把明文密钥写进沙箱环境变量或文件；需要鉴权时，由沙箱外的宿主程序或代理附加凭证。',
    'Visible UI vs. full GUI desktop': '普通界面限制与桌面软件隔离',
    'Use SessionContainer + bot login for full desktop apps, or scope ProcessContainer UI explicitly:': '操作完整桌面软件建议使用 SessionContainer 搭配专用账号；命令行进程可通过 ui 字段限制界面权限：',

    # Slide 17 (idx 16)
    'ui policy  ·  1.0.0 JSON / V1 SDK': 'ui 策略示例  ·  1.0.0 JSON / V1 SDK',
    '"ui": {\n  "disable": true,     // false = visible tool\n  "clipboard": "none", // or "read" only\n  "injection": false   // allowInputInjection\n}': '"ui": {\n  "disable": true,     // false = 允许显示窗口\n  "clipboard": "none", // 或设为只读 "read"\n  "injection": false   // 禁止模拟键鼠输入\n}',
    'Observe first, then enforce': '先梳理所需权限，再开启强制限制',
    'Enforcement': 'Enforcement 强制',
    'Applies the policy strictly; blocks any ungranted access.': '正式运行使用：严格按策略限制，未授权的访问直接拦截。',
    'Learning': 'Learning 学习',
    'Blocks ungranted access AND records every attempt in an activity report.': '调试策略使用：拦截未授权访问的同时，把每次被拒的路径和连接记录到报告里。',
    'Permissive': 'Permissive 宽容',
    'Allows and logs ungranted access — temporary diagnostics only.': '允许访问并记录日志 —— 仅适合临时排查问题，不要用于正式环境。',
    '01 · Observe': '01 · 运行并收集记录',
    '02 · Human Review': '02 · 人工核对所需权限',
    '03 · Enforce': '03 · 开启强制模式运行',
    'Learning Mode provides evidence, not auto-grants. A human approves every rule.': '学习模式只负责整理访问记录，不会自动放开权限，新增规则仍需人工确认。',

    # Slide 18 (idx 17)
    '04 · USE CASES': '04 · 常见应用场景',
    'Secure coding agents': '编码类 Agent 的权限控制',
    'Build, test and push — without touching anything outside the repo.': '正常编译、跑测试与提交代码，不接触仓库以外的文件和凭证。',
    'Project repo  C:\\src\\app': '当前项目目录  C:\\src\\app',
    'Read / write': '允许读写',
    'Toolchain: node, python, git': '开发工具：node、python、git',
    'Allowed': '允许运行',
    'Package registry via local proxy': '通过本地代理下载依赖包',
    '~/.ssh, cloud tokens & Documents': '~/.ssh 私钥、云凭证与个人文档',
    'Denied': '禁止访问',
    'Any other network destination': '其他未授权的外部网络地址',
    'Blocked': '默认阻断',
    'In production today: GitHub Copilot CLI enforces command isolation via MXC.': '实际应用案例：GitHub Copilot CLI 已使用 MXC 隔离模型生成的命令行执行。',

    # Slide 19 (idx 18)
    'Desktop and data agents': '桌面操作与敏感文档处理场景',
    'Desktop automation agent': '操作桌面软件的 Agent',
    'Keeps GUI app state in a separate Windows account and session, isolated from the employee desktop.': '在独立的本地 Windows 账号和桌面会话中操作界面程序，不干扰员工当前使用的桌面。',
    'Best fit: Session container + dedicated bot login': '适用方式：Session 容器 + 专用低权限本地账号',
    'Sensitive-document agent': '处理敏感文档的 Agent',
    'Reads one approved input folder, writes one summary report, with network completely disabled.': '只读指定的输入文件夹、将汇总结果写入输出目录，并全程关闭网络连接。',
    'Best fit: Process container (read-in / write-out, egress deny)': '适用方式：Process 容器（输入只读、输出分离 + 禁用网络）',
    'Agents on Windows 365 Cloud PCs': '运行在 Windows 365 云电脑上的 Agent',
    'Runs alongside existing Cloud PC workloads; process isolation works out of the box just like physical PCs.': '与云电脑桌面环境配合运行，ProcessContainer 进程隔离用法与本地物理机一致。',
    'Best fit: Process container on Windows 365 (GA)': '适用方式：Windows 365 上的 Process 容器（已正式发布）',

    # Slide 20 (idx 19)
    'Already in the ecosystem': '正在接入 MXC 的工具生态',
    'Industry convergence: coding CLIs, local model runners & enterprise SaaS agents standardize on OS-native containment.': '越来越多编码工具、本地模型运行器与企业级 Agent 开始采用统一的操作系统原生沙箱接口。',
    'SUPPORTS MXC TODAY': '目前已支持 MXC',
    'INTEGRATED / BUILDING': '已集成 / 持续合作中',
    'NVIDIA OpenShell\nOpenAI and Meta': 'NVIDIA OpenShell\nOpenAI 与 Meta',
    'SUPPORT ANNOUNCED': '已宣布即将支持',

    # Slide 21 (idx 20)
    '05 · INTEGRATING MXC': '05 · 在项目中接入 MXC',
    'Reference architecture': 'Agent 平台接入参考架构',
    '01 · PLANNER': '01 · 任务规划',
    'Agent planner': 'Agent 规划模块',
    'LLM proposes next action': '模型提出下一步操作请求',
    '02 · INTENT': '02 · 工具请求',
    'Tool call': '调用工具意图',
    '03 · TRUSTED HOST': '03 · 宿主程序',
    'Policy engine': '权限策略模块',
    'Selects least-privilege policy': '选择对应的最小权限配置',
    '04 · CONTRACT': '04 · SDK 接口',
    'Lifecycle & diagnostics': '管理容器生命周期与日志',
    '05 · OS KERNEL': '05 · 系统隔离',
    'Sandbox backend': '操作系统沙箱后端',
    'Enforces hard boundary': '限制文件、网络与会话范围',
    '←   RETURN PATH  ·  Stdout, exit code & kernel denials validated by trusted host before feeding back to agent': '←   返回路径  ·  宿主程序先检查标准输出、退出码和拦截记录，确认正常后再传回给 Agent',
    'Egress proxy': '外部出站代理',
    'Attaches credentials and enforces domain allow-lists strictly outside the sandbox.': '在沙箱外负责附加接口凭证，并按域名列表控制外部访问。',
    'Learning Mode': '学习模式辅助配置',
    'Captures blocked-access evidence so humans can tighten policies before rollout.': '记录测试运行时的被拒访问明细，方便开发人员整理所需的最小权限。',
    'Enterprise governance': '企业统一管理',
    'Entra Agent ID & Agent 365 govern agent identity, audit trails, and fleet policy.': '配合 Entra Agent ID 与 Agent 365 管理身份记录、审计日志和终端策略。',

    # Slide 22 (idx 21)
    'Three SDKs, one contract': '三种语言 SDK，统一的策略格式',
    'CAPABILITY DISCOVERY': '查询平台支持能力',
    'Write one ContainerRequest in code — the SDK maps it to the native OS sandbox across Windows, macOS & Linux.': '在代码中构造统一的 ContainerRequest，由 SDK 负责调用 Windows、macOS 或 Linux 的原生沙箱。',
    'Requirements: Windows 11 24H2+ (physical or Windows 365), macOS, or Linux  ·  .NET 8+ / Node.js 18+ / Rust.': '运行环境：Windows 11 24H2+（物理机或 Windows 365）、macOS 或 Linux  ·  .NET 8+ / Node.js 18+ / Rust。',

    # Slide 23 (idx 22)
    'const result = await run(request); // validate output next!': 'const result = await run(request); // 随后检查返回结果',
    'Integrate in five steps': '接入 MXC 的五个步骤',
    'Discover': '1 · 检查环境 (Discover)',
    'Call getPlatformSupport() and probe(request) before launch.': '启动前先调用 getPlatformSupport() 和 probe(request) 确认系统支持。',
    'Describe': '2 · 配置权限 (Describe)',
    'Build a V1 ContainerRequest: files, network, ui & timeoutMs.': '按默认拒绝配置 ContainerRequest：明确文件、网络、界面和超时参数。',
    'Learn': '3 · 测试梳理 (Learn)',
    'Run in Learning Mode first; review the blocked-access report.': '开发阶段先用学习模式试跑，根据被拒记录整理实际需要的路径。',
    'Enforce': '4 · 开启限制 (Enforce)',
    'Switch to Enforcement; verify allowed AND denied cases.': '切换到强制模式，同时测试正常操作和越界拦截是否符合预期。',
    'Operate': '5 · 超时与清理 (Operate)',
    'Handle timeouts & cleanup across success, failure & cancel.': '处理好正常结束、运行报错和手动取消等情况下的超时与资源释放。',

    # Slide 24 (idx 23)
    "Your host's job list": '宿主程序需要处理好的六件事',
    'Set an explicit timeout; handle launch error, workload failure & cancellation separately.': '每次运行都设置超时时间，并区分“沙箱启动失败”、“内部命令报错”和“用户主动取消”。',
    'Clean up even when a later step throws — always deprovision persistent sessions.': '无论中间哪一步出错都要执行清理逻辑，有状态会话结束后务必调用 deprovision 删除临时账号。',
    'Spin up one narrow sandbox per action rather than sharing one broad policy.': '尽量每次操作单独启动一个权限最小的新容器，不要让多个任务共用一个宽泛的长期容器。',
    'Write outputs to a dedicated staging folder; teardown never undoes host writes.': '将输出写到单独的暂存目录；容器销毁并不会自动删除已经写入宿主机目录的文件。',
    'Explain policy denials clearly to users and agents; never fail silently.': '遇到权限拦截时返回清晰的提示信息，方便用户或上层程序理解原因，不要静默失败。',
    'Require a human to approve every grant suggested by Learning Mode.': '学习模式记录下的未授权访问仅作参考，正式放开权限前都应经过人工确认。',

    # Slide 25 (idx 24)
    '06 · LIVE DEMO 1': '06 · 演示 1',
    'ProcessContainer: files & network': 'ProcessContainer：文件与网络权限测试',
    '✓  Scoped read-in / write-out': '✓  授权范围内的读写正常执行',
    'Workload reads demo/input and writes report.txt into its dedicated output folder.': '沙箱内程序读取 demo/input 并将 report.txt 写入指定的 output 目录，正常返回。',
    '✕  Protected file denied': '✕  受保护的敏感文件被拦截',
    'The same workload attempting to read a secret file is blocked by the OS kernel.': '保持相同配置尝试读取 protected/denied.txt，被系统权限规则直接拒绝。',
    '⇄  Directional network leash': '⇄  本地回环网络按配置放行或阻断',
    'Reaches an explicitly permitted loopback endpoint; blocks unapproved outbound traffic.': '禁掉外网的前提下：允许本地回环时可正常访问本机端口，关闭后同一请求被拦截。',

    # Slide 26 (idx 25)
    '06 · LIVE DEMO 2': '06 · 演示 2',
    'Isolated incident triage': '独立会话下的故障排查演示',
    '01 · Stage': '01 · 准备工单文件',
    'Host stages INC-1042 and its read-only approval inside the ephemeral workspace.': '宿主程序将工单 INC-1042 和只读排查说明放入共享工作目录。',
    '02 · Open PTY': '02 · 创建临时账号与终端',
    'Provisions a separate local Windows operator account & isolated Session.': '创建独立的本地 Windows 临时账号与会话，打开 PowerShell 交互终端。',
    '03 · Triage': '03 · 查询服务状态与权限验证',
    'Checks Spooler & W32Time; blocks cross-user reads and service-stop attempts.': '允许查询 Spooler 与 W32Time 状态；读取宿主机员工目录和停止系统服务均被拒绝。',
    '04 · Teardown': '04 · 校验结果并清理账号',
    'Host validates evidence, stops the session, and deprovisions the temp account.': '宿主程序核对输出的排查记录后关闭会话，并删除临时创建的本地账号。',

    # Slide 27 (idx 26)
    '06 · LIVE DEMO 3': '06 · 演示 3',
    'Untrusted code in a MicroVM': '在微虚拟机里运行不可信代码',
    'Warm snapshot restore: ~90 ms · Unikraft 5.15': '从预热快照恢复：约 90 ms · Unikraft 5.15 微内核',
    '✓  Hardware VM kernel': '✓  独立的微虚拟机内核',
    'Unikraft 5.15 on Hypervisor': '运行在 Hypervisor 上的 Unikraft 5.15',
    'Runs inside a hardware-virtualized partition, never sharing the Windows NT kernel.': '在硬件虚拟化分区内运行独立微内核，不与宿主机共用 Windows 内核。',
    '✕  Zero host surface': '✕  无法访问宿主机磁盘与网络',
    'No host FS, env, or sockets': '隔离宿主文件系统、环境变量与网络连接',
    'Host C:\\, environment variables, and TCP sockets (Errno 13) remain unreachable.': '测试脚本尝试读取宿主机 C 盘、环境变量和建立外部网络连接 (Errno 13) 均被阻断。',
    '↺  Instant snapshot wipe': '↺  每次运行自动恢复初始状态',
    '<100 ms restore & reset': '<100 ms 快照恢复，退出不留残留',
    'Every call restores a warmed memory snapshot; /tmp writes vanish on exit.': '每次调用都从干净的内存快照启动；上一次写进 /tmp 的文件在下一次运行时自动清空。',

    # Slide 28 (idx 27)
    'KEY TAKEAWAYS': '总结 · KEY TAKEAWAYS',
    'Box the agent,\nnot the ambition': '管住执行边界，\n放心使用 Agent',
    "1   Push guardrails down to the OS. Prompts alone can't enforce.": '1   在操作系统层落实权限边界，不能只靠提示词约束。',
    '2   MXC: one policy, the right sandbox on Windows, macOS & Linux.': '2   用一份 MXC 策略对接 Windows、macOS 与 Linux 原生沙箱。',
    '3   Learn first, then enforce least privilege.': '3   先用学习模式摸清所需路径，人工确认后再开启强制模式。',
    "4   Containment isn't trust. Validate every result.": '4   沙箱隔离不代表输出一定安全，返回宿主后仍需校验结果。',
    'Thank you  ·  Q&A': '谢谢大家  ·  问答交流',
}

ZH_DEMO_LAUNCHERS = {
    24: ('process', '1', '启动 Demo 1 · ProcessContainer'),
    25: ('pty', '1', '启动 Demo 2 · IsolationSession PTY'),
    26: ('microvm', '1', '启动 Demo 3 · Hyperlight MicroVM'),
}
