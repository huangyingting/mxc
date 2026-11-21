#!/usr/bin/env python3
"""
build_pptx.py - rebuilds "Guardrails with Sandboxing - MXC" (28 slides, 16:9,
hand-drawn comic style, Chinese speaker notes) as a standalone .pptx.

Usage:   pip install python-pptx pillow
         python -m presentation.src.build_pptx [--out dist/pptx/deck.pptx]

How to reuse
  1. THEME   - change the colours / fonts in the THEME block below.
  2. CONTENT - each slide is one "# ---------- N TITLE" block: edit the strings,
               lists and note() text, or copy a block to make a new slide.
  3. HELPERS - tb() text box, rect() shape, img() picture, header()/footer()
               title & page furniture, note() speaker notes. Units are EMU
               (914400 = 1 inch). Text boxes auto-size their height.
  4. IMAGES  - put the illustrations in presentation/assets/images.
"""
import argparse, os
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from PIL import Image

# ===================== THEME =====================
# ===================== LAYOUT ====================
W,H=12192000,6858000; M=457200; CW=W-2*M
BG='F3EEE4'; CARD='E9E2D4'; CARD2='E9E2D4'; EDGE='C9C0AE'; INK='2B2B2B'; MUTED='57534C'; FOOT='6E685F'
CY='3F607E'; AMB='A84A35'; CODEBG='FAF7F1'; CODE='2F4B66'
HF='Plus Jakarta Sans'; BF='Plus Jakarta Sans'; MF='JetBrains Mono'
PANEL=4572000   # width of the side illustration panel
# ===================== HELPERS ===================
slides={}  # idx -> list of patches
notes={}
def P(s,op): slides.setdefault(s,[]).append(op)
WF={'Plus Jakarta Sans':0.52,'Newsreader':0.56,'JetBrains Mono':0.6,'Aptos Display':0.55,'Aptos':0.51,'Cascadia Code':0.6,'Georgia':0.6,'Arial':0.56,'Calibri':0.5,'Courier New':0.6}
def nlines(text,size,cx,font,bold):
    f=WF[font]*(1.07 if bold else 1)
    cpl=max(1,int((cx-182880)/(size*12700*f)))
    n=0
    for para in text.split('\n'):
        words=para.split(' '); line=0; cnt=1
        for w_ in words:
            L=len(w_)+(1 if line else 0)
            if line+L>cpl and line>0: cnt+=1; line=len(w_)
            else: line+=L
        n+=cnt
    return n
def hgt(paras,cx,font):
    t=0
    for p in paras:
        t+=nlines(p['text'],p['sizePt'],cx,font,p.get('bold',False))*p['sizePt']*12700*1.2
    return int(t+182880)
def tb(s,x,y,cx,text=None,size=16,color=INK,font=BF,bold=False,align='left',paras=None,ls=None,cy=None,name=None):
    if paras is None:
        paras=[{'text':text,'sizePt':size,'bold':bold,'color':color}]
    else:
        paras=[dict({'sizePt':size,'bold':bold,'color':color},**p) for p in paras]
    h=cy or hgt(paras,cx,font)
    op={'op':'add_text_box','slide':s,'x':int(x),'y':int(y),'cx':int(cx),'cy':int(h),'fontFamily':font,'align':align,'color':color}
    if len(paras)==1 and ls is None:
        p=paras[0]; op.update(text=p['text'],sizePt=p['sizePt'],bold=p['bold'],color=p['color'])
    else:
        op['paragraphs']=[{k:v for k,v in p.items()} for p in paras]
        if ls is not None:
            if len(paras)==1:
                p=paras[0]; del op['paragraphs']; op.update(text=p['text'],sizePt=p['sizePt'],bold=p['bold'],color=p['color'])
            op['letterSpacingPt']=ls
    if name: op['name']=name
    P(s,op); return int(y+h)
def rect(s,x,y,cx,cy,fill=CARD,line=None,lw=None,preset='rect',text=None,size=16,color=INK,bold=False,font=BF,align='center',shadow=None,rot=None):
    op={'op':'add_shape','slide':s,'preset':preset,'x':int(x),'y':int(y),'cx':int(cx),'cy':int(cy),'fill':fill}
    if line: op['line']=line; op['lineWidth']=lw or 12700
    if text is not None:
        op.update(text=text,sizePt=size,color=color,bold=bold,fontFamily=font,align=align)
    if shadow: op['shadow']=shadow
    if rot: op['rot']=rot
    P(s,op)
def img(s,path,x,y,cx,cy,alt):
    P(s,{'op':'add_image','slide':s,'path':path,'x':int(x),'y':int(y),'cx':int(cx),'cy':int(cy),'fit':'cover','altText':alt})
def header(s,eyebrow,title,x=M,cx=CW,title_cy=None):
    tb(s,x,400000,cx,eyebrow,size=13,color=CY,font=HF,bold=True,ls=2)
    tb(s,x,760000,cx,title,size=32,color=INK,font=HF,bold=True,cy=title_cy)
def footer(s,src,num,x=M,cx=CW):
    if src: tb(s,x,6180000,cx-900000,src,size=12,color=FOOT)
    tb(s,x+cx-800000,6180000,800000,str(num),size=12,color=FOOT,align='right')
def note(s,t): notes[s]=t

# ===================== SLIDES ====================
# ---------- 0 COVER
s=0
img(s,'hd-cover.jpg',0,0,W,H,'Hand-drawn sketch illustration: a friendly robot AI agent standing inside a glass cube, with files, cloud and keys kept outside the boundary')
x=M+150000; cx=5900000
tb(s,x,1250000,cx,'MICROSOFT EXECUTION CONTAINERS (MXC)',size=13,color=CY,font=HF,bold=True,ls=2)
y=tb(s,x,1650000,cx,'Guardrails with\nSandboxing',size=46,color=INK,font=HF,bold=True)
y=tb(s,x,y+150000,cx,'Keeping AI agents inside safe boundaries: isolating execution and controlling access to files, network and credentials on Windows',size=18,color=MUTED)
y=tb(s,x,y+350000,cx,'Yingting Huang  |  Sr. Cloud Solution Architect, Microsoft',size=16,color=INK,bold=True)
tb(s,x,y+30000,cx,'November 2026  ·  45 minutes  ·  L300 → L400',size=14,color=MUTED)
note(s,'AI Agent 已从内容生成工具演进为具备执行能力的系统，可以读取文件、运行命令、修改状态并调用外部服务。因此，安全目标不仅是减少错误回答，还必须限制错误或恶意动作的影响范围。\nMicrosoft Execution Containers，简称 MXC，为不完全可信或 AI 生成的工作负载提供操作系统级执行边界。其核心思想是保留必要的任务能力，同时将文件、网络、凭据和用户界面访问限制在明确授权范围内。')

# ---------- 1 AGENDA
s=1
header(s,'AGENDA','Our journey today')
items=[('01','Why sandboxing','Risks of agents that act'),('02','Meet MXC','What it is and how it works'),('03','Features & platforms','Backends, files, network, modes'),('04','Use cases','Coding, desktop and data agents'),('05','Integration','SDKs and reference architecture'),('06','Live demos','Process, Session & MicroVM'),('07','Takeaways & Q&A','What to do next')]
col=CW/7; ry=3150000
rect(s,M,ry-19050,CW,38100,fill=EDGE)
for i,(n,t,d) in enumerate(items):
    cx0=M+col*i; bw=col-80000
    tb(s,cx0,1900000,bw,n,size=30,color=CY,font=HF,bold=True)
    rect(s,cx0+60000,ry-120000,240000,240000,fill=CY,preset='ellipse')
    y=tb(s,cx0,3450000,bw,t,size=17,color=INK,bold=True,cy=hgt([{'text':'Features & platforms','sizePt':17,'bold':True}],bw,BF))
    tb(s,cx0,y+40000,bw,d,size=14,color=MUTED)
footer(s,'Level: L300 → L400  ·  Audience: architects, developers, security engineers',2)
note(s,'内容分为三个层次：风险与设计原则、MXC 的架构和平台能力、集成方法与验证场景。最终需要形成一套可落地的决策：工作负载需要何种文件和网络权限、凭据应由哪一层持有、如何选择隔离 Backend，以及如何同时验证允许路径与拒绝路径。')

# ---------- 2 AGENTS ACT
s=2
header(s,'01 · WHY SANDBOXING','Agents don\'t just talk. They act.')
cw_=4950000; cyy=1650000; ch=3550000
rect(s,M,cyy,cw_,ch,fill=CARD)
y=tb(s,M+300000,cyy+250000,cw_-600000,'Chat assistant',size=24,bold=True,color=MUTED)
y=tb(s,M+300000,y+50000,cw_-600000,'Answers questions',size=16,color=CY)
tb(s,M+300000,y+150000,cw_-600000,None,paras=[{'text':'• Generates text'},{'text':'• A human copies and runs the result'},{'text':'• Worst case: a wrong answer'}],size=17,color=MUTED)
rx=W-M-cw_
rect(s,rx,cyy,cw_,ch,fill=CARD2,line=CY,lw=25400)
y=tb(s,rx+300000,cyy+250000,cw_-600000,'Autonomous agent',size=24,bold=True)
y=tb(s,rx+300000,y+50000,cw_-600000,'Takes actions on your machine',size=16,color=CY)
tb(s,rx+300000,y+150000,cw_-600000,None,paras=[{'text':'• Reads and writes files'},{'text':'• Runs shell commands, installs packages'},{'text':'• Calls APIs and drives apps'},{'text':'• Worst case: your whole session'}],size=17)
rect(s,M+cw_+200000,cyy+ch/2-300000,W-2*M-2*cw_-400000,600000,fill=CY,preset='rightArrow')
tb(s,M,5400000,CW,'Agents run code they didn\'t write, and can\'t fully predict.',size=22,bold=True,color=INK)
footer(s,'',3)
note(s,'传统聊天机器人只返回内容，执行动作仍由人触发；Agent 则可能直接读取仓库、安装依赖、运行 shell、修改文件或调用 API。其命令通常在运行时由模型生成，难以逐项预审。\n若 Agent 继承交互用户的完整权限，一次错误规划或注入攻击就可能影响该用户能访问的全部资源。风险性质因此从内容错误扩展为系统级操作风险，必须通过执行层限制可达范围。')

# ---------- 3 THREATS (image right)
s=3
img(s,'hd-chaos.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a hooded villain slips a poisoned note to a robot agent while files fly out and a drone steals a key')
cxl=W-PANEL-2*M
header(s,'01 · WHY SANDBOXING','What can go wrong?',cx=cxl)
rows=[('Prompt injection','A poisoned README tells the agent to run the attacker\'s command.'),('Credential theft','Generated code reads tokens, SSH keys and browser secrets.'),('Data exfiltration','Files quietly leave for an unknown host.'),('Overreach','A "reasonable" shortcut edits production config it was never meant to touch.')]
y=1650000
for i,(h_,d) in enumerate(rows):
    tb(s,M,y,700000,f'0{i+1}',size=24,color=AMB,font=HF,bold=True)
    yy=tb(s,M+800000,y,cxl-800000,None,paras=[{'text':h_,'sizePt':19,'bold':True,'color':INK},{'text':d,'sizePt':15,'color':MUTED}])
    y=yy+130000
footer(s,'Scenario 04 adapted from the Windows Developer Blog, Oct 7, 2026',4,cx=cxl)
note(s,'主要威胁包括四类。Prompt Injection 来自工作负载读取的不可信文档、网页或仓库内容，可诱导模型执行额外指令。凭据窃取源于 Agent 对 token、SSH key 和浏览器资料的读取能力。数据外传要求同时考虑文件可读范围和网络出口，而不仅检查输入内容。越权操作则表现为任务语义上相似、但超出授权边界的动作，例如为更新网站直接修改生产配置。\n这些风险既可能由攻击者触发，也可能由模型选择错误路径导致，因此不能仅依赖内容分类或提示词约束。')

# ---------- 4 PROMPTS PERSUADE
s=4
header(s,'01 · WHY SANDBOXING','Prompts persuade. The OS enforces.')
y=tb(s,M,1650000,9500000,'"An agent cannot be its own security authority."',size=32,color=INK,font=HF,bold=True)
y=tb(s,M,y+20000,9500000,'Windows Developer Blog, October 7, 2026',size=14,color=MUTED)
lab=1900000; chw=2240000; gap=120000
def chiprow(yv,label,chips,lifted):
    tb(s,M,yv+70000,lab-100000,label,size=16,bold=True,color=CY if lifted else MUTED)
    for i,c in enumerate(chips):
        x=M+lab+i*(chw+gap)
        if lifted: rect(s,x,yv,chw,560000,fill=CARD2,line=CY,lw=19050,text=c,size=15,color=INK)
        else: rect(s,x,yv,chw,560000,fill=CARD,text=c,size=15,color=MUTED)
chiprow(3500000,'AI-level',['System prompts','Content filters','Tool rules','Output checks'],False)
chiprow(4250000,'Execution-level',['File allow / deny','Network rules','Process isolation','Secrets outside'],True)
tb(s,M,5150000,CW,'You need both: the model decides what to try; the operating system decides what is possible.',size=17,color=INK)
footer(s,'',5)
note(s,'系统提示词、内容过滤和工具规则属于 AI 层控制，用于降低不安全行为发生的概率；操作系统权限、进程隔离和网络策略属于执行层控制，用于限制行为是否可实施。\n模型可以决定尝试读取文件或建立连接，但只有执行层能可靠拒绝未授权操作。若仅依赖提示词，模型在遇到注入、理解偏差或工具错误时仍可能越界。Agent 安全设计应同时保留模型层约束和不可绕过的执行层边界。')

# ---------- 5 THREE PILLARS
s=5
header(s,'01 · WHY SANDBOXING','Three pillars for agents on Windows')
cw3=(CW-2*238800)/3; y0=1650000; ch=3900000
pill=[('Containment','Limit what an agent can access and do','Microsoft Execution Containers (MXC). Windows 365 support for MXC is also generally available.','Generally available'),('Identity','Tell agent activity apart from a person','Microsoft Entra Agent ID is GA; identity for supported local agents is in preview.','GA + preview'),('Manageability','Govern access and monitor agents','Agent 365 is GA; Intune/MXC controls for local agents are in preview.','GA + preview')]
for i,(t,sub,body,state) in enumerate(pill):
    x=M+i*(cw3+238800)
    rect(s,x,y0,cw3,ch,fill=CARD)
    pw=2100000 if t=='Containment' else 1500000
    rect(s,x+300000,y0+300000,pw,400000,fill=EDGE,preset='roundRect',text=state,size=13,color=INK,bold=True)
    y=tb(s,x+300000,y0+900000,cw3-600000,t,size=26,bold=True,font=HF)
    y=tb(s,x+300000,y+60000,cw3-600000,sub,size=17,color=CY)
    tb(s,x+300000,y+150000,cw3-600000,body,size=15,color=MUTED)
footer(s,'Sources: Microsoft Ignite 2026 · Agent 365 · Entra Agent ID · Microsoft Intune',6)
note(s,'Windows Agent 安全模型包含三个支柱。Containment 限制 Agent 对文件、网络、进程和 UI 的访问；Identity 区分 Agent 活动与用户活动；Manageability 提供发现、策略、审计和持续治理。\n发布状态需要分层说明。MXC 与 Windows 365 对 MXC 的支持已 GA。Microsoft Entra Agent ID 已 GA，但 Ignite 2026 宣布的受支持本地 Agent 自动身份仍处于 Preview；例如先覆盖 GitHub Copilot CLI 和 Claude Code。Microsoft Agent 365 已于 2026 年 5 月 GA；Agent 365 的本地/影子 Agent 治理以及 Intune 对 MXC ProcessContainer 的文件/网络边界策略仍是 Preview。不要把“产品 GA”误解成“所有本地 Agent 控制都已 GA”。')

# ---------- 6 MEET MXC (image left)
s=6
img(s,'hd-guard.jpg',0,0,PANEL,H,'Hand-drawn sketch panel: a robot agent works inside a force-field box while a guardian with a rulebook decides which items pass in')
xl=PANEL+M; cxr=W-xl-M
header(s,'02 · UNDERSTANDING MXC','Meet MXC',x=xl,cx=cxr)
y=tb(s,xl,1550000,cxr,'An open-source, policy-driven execution layer for untrusted or AI-generated workloads.',size=19,color=INK)
tw=(cxr-200000)/2; th=1100000; ty=y+200000
facts=[('MIT','Open source on GitHub'),('GA','Generally available since Oct 7, 2026'),('3 SDKs','.NET, TypeScript, Rust'),('3 OSes','Windows, macOS, Linux')]
for i,(v,l) in enumerate(facts):
    x=xl+(i%2)*(tw+200000); yy=ty+(i//2)*(th+150000)
    rect(s,x,yy,tw,th,fill=CARD)
    tb(s,x+250000,yy+90000,tw-400000,None,paras=[{'text':v,'sizePt':32,'bold':True,'color':CY},{'text':l,'sizePt':14,'color':INK}],font=BF)
tb(s,xl,ty+2*th+150000+250000,cxr,'One JSON policy, mapped to the right OS sandbox on every host.',size=17,bold=True,color=INK)
footer(s,'Sources: Microsoft Learn; Windows Developer Blog, Oct 2026',7,x=xl,cx=cxr)
note(s,'MXC 是开源、策略驱动的执行层，用于运行 AI 生成代码、第三方工具、插件或其他不完全可信的工作负载。Host 提交文件、网络、UI、命令和生命周期配置，MXC 将统一请求映射到当前平台支持的 Backend。\nContainer 不等同于 Docker 镜像或虚拟机。底层实现可以是进程沙箱、独立会话、WSL 环境或实验性虚拟化机制。MXC 采用 MIT 许可证，提供 .NET、TypeScript 和 Rust SDK，并支持 Windows、macOS 与 Linux。')

# ---------- 7 FIVE CONCEPTS
s=7
header(s,'02 · UNDERSTANDING MXC','Five concepts to remember')
gap=150000; cw5=(CW-4*gap)/5; y0=1650000; ch=3600000
cons=[('Host app','Your trusted app or agent runtime. Picks permissions, supplies inputs, handles results.'),('Workload','The command, script or process that runs under containment.'),('Policy','The access you permit: readable files, writable folders, network destinations, UI.'),('Execution config','The policy plus command, working directory, runtime settings and containment choice.'),('Backend','The OS mechanism that enforces it: process sandbox, separate session or VM.')]
for i,(t,d) in enumerate(cons):
    x=M+i*(cw5+gap)
    rect(s,x,y0,cw5,ch,fill=CARD)
    y=tb(s,x+200000,y0+250000,cw5-400000,f'0{i+1}',size=28,color=CY,font=HF,bold=True)
    y=tb(s,x+200000,y+80000,cw5-400000,t,size=19,bold=True,cy=hgt([{'text':'Execution config','sizePt':19,'bold':True}],cw5-400000,BF))
    tb(s,x+200000,y+100000,cw5-400000,d,size=14,color=MUTED)
tb(s,M,5450000,CW,'The host decides. The policy describes. The backend enforces.',size=20,bold=True)
footer(s,'Source: Microsoft Learn, Microsoft Execution Containers overview',8)
note(s,'MXC 包含五个核心对象。Host app 是可信决策方，负责分类任务和选择权限。Workload 是实际执行的命令、脚本或进程。Policy 以默认拒绝方式描述文件、网络与 UI 权限。Execution config 将 Policy、命令、工作目录、超时和隔离选择组合为一次请求。Backend 是实施限制的操作系统机制。\n权限必须由 Host 或人工审批授予。Workload 只能建议需求，不能自行选择或扩大 Policy；否则隔离边界可以被工作负载主动覆盖。')

# ---------- 8 REQUEST FLOW
s=8
header(s,'02 · UNDERSTANDING MXC','How one request flows')
steps=[('HOST','Discover','Ask the SDK which backends this machine supports.'),('HOST','Describe','Build a policy and execution request, deny by default.'),('SDK','Launch','Start the workload through the SDK.'),('MXC','Translate','Map the request to the selected backend\'s operations.'),('HOST','Return','Handle output, exit status, warnings and cleanup.')]
sw=2200000; step=(CW-sw)/4; yc=2150000
for i,(who,t,d) in enumerate(steps):
    x=M+i*step
    tb(s,x+150000,1650000,sw-300000,who,size=13,color=CY,font=HF,bold=True,ls=2)
    rect(s,x,yc,sw,850000,fill=CY,preset='chevron' if i else 'homePlate',text=t,size=18,bold=True,color=BG)
    tb(s,x+150000,yc+1050000,sw-350000,d,size=15,color=MUTED)
rect(s,M,4750000,CW,800000,fill=CARD,line=AMB,lw=19050)
tb(s,M+300000,4750000,CW-600000,'Backends have different security properties. Always test both the allowed and the denied operations.',size=17,color=INK,cy=800000)
footer(s,'Source: Microsoft Learn, How MXC works',9)
note(s,'一次执行请求包含五个阶段：Host 发现平台能力，构造默认拒绝的 Policy 和 Execution config，通过 SDK 启动 Workload，MXC 将配置转换为 Backend 操作，Host 最终处理输出、退出码、拒绝事件和清理。\n验证必须覆盖正向与负向路径。允许的操作应成功，越权操作应产生明确拒绝。不同 Backend 的能力与安全语义并不完全一致，因此跨平台部署不能只复用成功路径测试。')

# ---------- 9 FOUR LAYERS
s=9
header(s,'02 · UNDERSTANDING MXC','Four layers, one boundary')
lays=[('Host app / agent runtime','Classifies the task, chooses trusted policy defaults, handles output and approval.'),('MXC SDK & policy contract','Discovers support, turns intent into a versioned config, manages lifecycle.'),('MXC execution layer','Validates the config, resolves a backend, translates to platform operations.'),('Containment backend','Creates the environment and enforces what the platform supports.')]
y=1600000; lh=820000; lw_=3400000
for i,(t,d) in enumerate(lays):
    rect(s,M,y,lw_,lh,fill=CARD2,text=t,size=17,bold=True,align='left')
    rect(s,M+lw_+100000,y,CW-lw_-100000,lh,fill=CARD)
    tb(s,M+lw_+400000,y,CW-lw_-700000,d,size=16,color=MUTED,cy=lh)
    y+=lh+110000
tb(s,M,y+120000,CW,'Rule #1: an agent may suggest a permission, but only trusted host logic or a human can grant it.',size=18,bold=True,color=INK)
footer(s,'Source: Microsoft Learn, MXC architecture and containment models',10)
note(s,'架构分为 Host app、MXC SDK 与策略契约、MXC 执行层和 Containment Backend 四层。Host 选择授权；SDK 处理能力发现、配置转换和生命周期；执行层校验请求并选择 Backend；Backend 创建环境并执行平台限制。\n信任边界位于 Host 与 Backend 之间。Agent 可以提出权限需求，但只有 Host 的策略逻辑或人工审批可以批准。Host 自身位于容器外，因此仍需验证输入、输出和工具调用，其实现不受 MXC 自动保护。')

# ---------- 10 PICK THE RIGHT BOX
s=10
header(s,'03 · FEATURES & PLATFORMS','Pick the right box')
ry=1900000
tb(s,M,1550000,4000000,'Lighter and faster',size=14,color=MUTED)
tb(s,W-M-4000000,1550000,4000000,'Stronger isolation',size=14,color=MUTED,align='right')
rect(s,M,ry-80000,CW,220000,fill=EDGE,preset='rightArrow')
gap=180000; cw4=(CW-3*gap)/4; y0=2350000; ch=3550000
bx=[('Process container','Windows 11 · macOS · Linux','AppContainer, Seatbelt or Bubblewrap. Lightweight and responsive; the default for coding agents.',None),('Session container','Windows 11 only','Separate Windows account and OS session with session-local UI. No file or network filtering; SendInput is blocked.',None),('WSL container','Windows 11','Runs Linux workloads on a Windows host through WSL.',None),('MicroVM','Windows 11 · Linux','Hardware-enforced virtual machine boundary; the strongest option.','Experimental')]
for i,(t,p,d,st) in enumerate(bx):
    x=M+i*(cw4+gap)
    rect(s,x+cw4/2-110000,ry-80000,220000,220000,fill=CY,line=BG,lw=25400,preset='ellipse')
    rect(s,x,y0,cw4,ch,fill=CARD)
    yy=y0+250000
    if st:
        rect(s,x+250000,yy,1500000,380000,fill=AMB,preset='roundRect',text=st,size=13,bold=True,color=BG); yy+=500000
    yy=tb(s,x+250000,yy,cw4-500000,t,size=20,bold=True)
    yy=tb(s,x+250000,yy+40000,cw4-500000,p,size=14,color=CY,bold=True)
    tb(s,x+250000,yy+120000,cw4-500000,d,size=15,color=MUTED)
footer(s,'Sources: Windows Developer Blog, Oct 7, 2026; Microsoft Learn, MXC architecture',11)
note(s,'隔离方式应按工作负载需求选择。Process container 启动开销较低，适合 Coding Agent；Windows、macOS 和 Linux 分别映射到 AppContainer、Seatbelt 和 Bubblewrap。\nSession container 创建独立 Windows 用户和 OS Session，Session 内窗口、GDI 与剪贴板可工作，但 Host 的 UI 和用户 Profile 不属于该 Session。官方实现拒绝文件策略，要求网络明确为 egress、ingress、host loopback 全部 allow，并阻止 SendInput。因此它解决身份、会话和桌面分离，不解决精细文件/网络最小权限，也不是通用凭据隔离器。')

# ---------- 11 WHERE MXC RUNS
s=11
header(s,'03 · FEATURES & PLATFORMS','Where MXC runs today')
c1=M; w1=3000000; c2=c1+w1+150000; w2=3000000; c3=c2+w2+150000; w3=W-M-c3
hy=1600000
for x,w,t in [(c1,w1,'HOST'),(c2,w2,'DEFAULT PROCESS BACKEND'),(c3,w3,'ALSO IN THE V1 SDK / NOTES')]:
    tb(s,x+200000,hy,w-200000,t,size=12,color=CY,font=HF,bold=True,ls=1)
rowsd=[('Windows 11','ProcessContainer','IsolationSession and WSLc; Learning Mode activity reports'),('Linux','Bubblewrap','LXC'),('macOS','Seatbelt','Check each release for other options'),('Windows 365 Cloud PC','ProcessContainer','Same as a physical PC; session or VM backends need nested virtualization')]
y=1980000; rh=720000
for i,(a,b,c) in enumerate(rowsd):
    rect(s,c1,y,CW,rh,fill=CARD)
    tb(s,c1+200000,y,w1-200000,a,size=18,bold=True,cy=rh)
    tb(s,c2+200000,y,w2-200000,b,size=17,color=CY,font=MF,bold=True,cy=rh)
    tb(s,c3+200000,y,w3-300000,c,size=15,color=MUTED,cy=rh)
    y+=rh+100000
rect(s,M,y+100000,CW,640000,fill=CARD,line=AMB,lw=19050)
tb(s,M+300000,y+100000,CW-600000,'Experimental, explicit opt-in only: Windows Sandbox, MicroVM and Hyperlight (not creatable through V1 SDK APIs).',size=15,color=INK,cy=640000)
footer(s,'Source: Microsoft Learn, MXC containment options and platform support (MXC v1.0.0); Windows 365 support for MXC',12)
note(s,'平台接口统一不代表 Backend 能力完全一致。Windows 11 默认使用 ProcessContainer，V1 SDK 还覆盖特定的 IsolationSession 和 WSLc 场景；Linux 默认使用 Bubblewrap；macOS 使用 Seatbelt。\nWindows 365 上的进程隔离与物理 Windows 设备行为一致，不需要嵌套虚拟化；Session 或虚拟化类 Backend 则需要确认额外环境要求。Windows Sandbox、MicroVM 和 Hyperlight 仍为实验能力，需要显式启用且不能通过 V1 SDK 常规 API 创建。应用应在运行时发现和预检能力，而非硬编码 Backend。')

# ---------- 12 FILES
s=12
header(s,'03 · FEATURES & PLATFORMS','Files: read in, write out')
cx_=6000000; cy0=1600000; chh=4300000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+250000,cy0+120000,cx_-500000,'policy.json',size=12,color=FOOT,font=MF)
code='{\n  "version": "1.0.0",\n  "containment": "processcontainer",\n  "process": {\n    "commandLine": "python summarize.py"\n  },\n  "filesystem": {\n    "readwritePaths": ["C:\\\\demo\\\\output"],\n    "deniedPaths": ["C:\\\\Users\\\\me\\\\.ssh"]\n  }\n}'
tb(s,M+250000,cy0+550000,cx_-500000,None,paras=[{'text':l} for l in code.split('\n')],size=14,color=CODE,font=MF)
rx=M+cx_+400000; rw=W-M-rx
rules=[('Read in, write out','Separate readable inputs from one dedicated writable output folder.'),('Deny the crown jewels','Explicitly deny secrets, system and profile paths.'),('Teardown is not undo','Writes to granted host folders survive the sandbox. Choose output paths deliberately.')]
y=1650000
for i,(h_,d) in enumerate(rules):
    y=tb(s,rx,y,rw,None,paras=[{'text':h_,'sizePt':19,'bold':True,'color':CY if i==0 else INK},{'text':d,'sizePt':15,'color':MUTED}])+250000
footer(s,'Source: github.com/microsoft/mxc, docs/examples.md; Microsoft Learn, MXC architecture',13)
note(s,'文件策略应区分只读输入与专用可写输出，并使用 deniedPaths 显式保护敏感目录。授权范围应使用具体路径，而不是仅把当前工作目录视为隐式权限。\nTeardown 只销毁执行环境，不会回滚已写入主机目录的数据，也不会撤销通过允许网络连接完成的外部操作。因此输出目录必须被视为受控交接区，Host 需要在消费生成文件前验证格式、类型、大小和语义。')

# ---------- 13 NETWORK
s=13
header(s,'03 · FEATURES & PLATFORMS','Network: deny by default')
bx0=M; by0=1900000; bw=2400000; bh=3400000
rect(s,bx0,by0,bw,bh,fill=CARD2,line=CY,lw=28575)
tb(s,bx0+200000,by0+1000000,bw-400000,None,paras=[{'text':'Agent workload','sizePt':19,'bold':True,'color':INK},{'text':'egress: deny','sizePt':15,'color':CY},{'text':'ingress: deny','sizePt':15,'color':CY}])
dx=bx0+bw+1250000; dw=2900000; dh=780000
dests=[('140.82.114.6 : 443','Allowed by CIDR + port',CY,'✓'),('Any other address','Blocked',AMB,'✕'),('Local proxy 127.0.0.1:8080','Domain allow-list lives here',CY,'→')]
for i,(t,d,c,g) in enumerate(dests):
    yy=by0+120000+i*(dh+380000)
    rect(s,bx0+bw+120000,yy+dh/2-160000,1000000,320000,fill=c,preset='rightArrow')
    rect(s,dx,yy,dw,dh,fill=CARD,line=c,lw=19050)
    tb(s,dx+180000,yy,dw-300000,None,paras=[{'text':g+'  '+t,'sizePt':15,'bold':True,'color':INK},{'text':d,'sizePt':13,'color':MUTED}],cy=dh)
fx=dx+dw+350000; fw=W-M-fx
facts=['Rules are directional: egress and ingress, each with its own default.','Allow by IP/CIDR, protocol and port, not hostnames.','Need domains? Route through an HTTP(S) proxy with runtimeConfig.networkProxy.','In production, identify a packaged proxy with allowedProxyPeer.']
tb(s,fx,1750000,fw,None,paras=[{'text':f} for f in facts],size=15,color=MUTED)
footer(s,'Source: github.com/microsoft/mxc, docs/examples.md',14)
note(s,'网络策略应分别配置 egress、ingress 和 host loopback，默认全部拒绝。允许规则由地址、CIDR、协议和端口共同构成，且底层规则通常不直接执行域名匹配。\n需要域名级控制时，应让 Workload 只连接可信本地代理，由代理实施域名白名单、身份验证和审计。生产环境可通过 allowedProxyPeer 限定允许的代理进程，避免任意回环服务冒用网络代理权限。')

# ---------- 14 KEYS & UI
s=14
header(s,'03 · FEATURES & PLATFORMS','Keys and UI stay outside')
bx0=M; by0=1750000; bw=3600000; bh=3700000
rect(s,bx0,by0,bw,bh,fill=CARD2,line=CY,lw=28575)
tb(s,bx0+250000,by0+200000,bw-500000,'INSIDE THE SANDBOX',size=12,color=CY,font=HF,bold=True,ls=2)
for i,t in enumerate(['Repo files (granted)','Build tools','Approved API calls']):
    rect(s,bx0+250000,by0+750000+i*900000,bw-500000,650000,fill=CARD,text=t,size=16,color=INK)
ox=bx0+bw+300000; ow=2600000
tb(s,ox,by0,ow,'OUTSIDE, OUT OF REACH',size=12,color=AMB,font=HF,bold=True,ls=1)
for i,t in enumerate(['API tokens','SSH keys','Browser secrets','Your clipboard','Your desktop & input']):
    rect(s,ox,by0+450000+i*660000,ow,540000,fill=CARD,line=AMB,lw=12700,text='✕  '+t,size=15,color=INK,align='left')
ex=ox+ow+350000; ew=W-M-ex
y=tb(s,ex,1750000,ew,None,paras=[{'text':'Credentials','sizePt':19,'bold':True,'color':INK},{'text':'Keep secrets out of the workload\'s reach. Let the trusted host or proxy attach them to approved calls.','sizePt':15,'color':MUTED}])
y=tb(s,ex,y+200000,ew,None,paras=[{'text':'UI and input','sizePt':19,'bold':True,'color':INK},{'text':'Restrict windows, clipboard and synthetic keyboard or pointer input.','sizePt':15,'color':MUTED}])
tb(s,ex,y+200000,ew,None,paras=[{'text':'Need a desktop?','sizePt':19,'bold':True,'color':CY},{'text':'A session has its own desktop and clipboard; synthetic input is blocked.','sizePt':15,'color':MUTED}])
footer(s,'Sources: Microsoft Learn, MXC overview; Windows Developer Blog, Oct 7, 2026',15)
note(s,'API token、SSH key、浏览器凭据和用户私有资料不应复制到 Workload 可见的环境。外部服务认证应由可信 Host 或代理完成，凭据只附加到经过策略验证的请求中。\nSession container 的源码语义是独立 Windows 用户和 Session：Session 内窗口、GDI 和 Session 自己的剪贴板可用，但无法访问 Host 的对应对象，且 SendInput 被阻止。若 Session 内已经登录 ERP，ERP Cookie 或 Token 对该 Agent 用户可见；Session 并不会把这些凭据变成不可读取。此时必须使用专用低权限账号、短期登录和可信认证 Broker。')

# ---------- 15 LEARNING MODE (image right)
s=15
img(s,'hd-learning.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: an owl detective records the footprints of a robot agent walking inside a box')
cxl=W-PANEL-2*M
header(s,'03 · FEATURES & PLATFORMS','Observe first, then enforce',cx=cxl)
modes=[('Enforcement','Applies the policy; blocks anything not granted.'),('Learning','Blocks ungranted access and records each attempt in an activity report.'),('Permissive','Allows and records ungranted access, for diagnosis only.')]
y=1600000; mh=900000
for i,(t,d) in enumerate(modes):
    rect(s,M,y,cxl,mh,fill=CARD)
    tb(s,M+250000,y,2100000,t,size=18,bold=True,color=CY if t=='Learning' else INK,cy=mh)
    tb(s,M+2400000,y,cxl-2600000,d,size=15,color=MUTED,cy=mh)
    y+=mh+110000
sw=(cxl-2*120000)/3
for i,t in enumerate(['Observe','Review','Enforce']):
    rect(s,M+i*(sw+120000),y+100000,sw,560000,fill=CY if i==2 else CARD2,preset='chevron' if i else 'homePlate',text=t,size=16,bold=True,color=BG if i==2 else INK)
tb(s,M,y+800000,cxl,'Learning Mode gives evidence, not decisions. A human approves every grant.',size=15,color=INK)
footer(s,'Sources: Microsoft Learn, Learning Mode; Windows 365 support for MXC',16,cx=cxl)
note(s,'Enforcement Mode 拒绝未授权访问并严格执行策略。Learning Mode 同样拒绝访问，但会将拒绝事件写入活动报告，用于分析正常业务需要哪些权限。Permissive Mode 会允许并记录未授权访问，只适用于受控诊断环境。\n活动报告不等同于自动审批。每项权限建议都应结合任务目的、数据敏感度和攻击面进行人工审核。正式上线前需移除学习或审计宽松配置，并再次验证允许和拒绝行为。')

# ---------- 16 LIFECYCLE
s=16
header(s,'03 · FEATURES & PLATFORMS','One-shot or state-aware')
lx=M; lw0=2300000; ax=M+lw0+200000; aw=W-M-ax
def lane(y,label,sub,stepsL,desc):
    tb(s,lx,y,lw0,None,paras=[{'text':label,'sizePt':20,'bold':True,'color':INK},{'text':sub,'sizePt':14,'color':CY}])
    n=len(stepsL); g=330000; bw=(aw-(n-1)*g)/n
    for i,t in enumerate(stepsL):
        x=ax+i*(bw+g)
        rect(s,x,y,bw,600000,fill=CARD2,text=t,size=15,bold=True)
        if i<n-1: rect(s,x+bw+50000,y+190000,g-100000,220000,fill=CY,preset='rightArrow')
    tb(s,ax,y+680000,aw,desc,size=15,color=MUTED)
lane(1600000,'One-shot','Default choice',['Create','Run one workload','Tear down'],'Builds, tests, generated commands: a fresh, narrow sandbox for every action.')
lane(3000000,'State-aware','When state must persist',['Provision','Start','Execute','Stop','Deprovision'],'Agent profiles, installed tools. In MXC 1.0: IsolationSession and WSLc.')
ry=4450000; g=150000; cw3=(CW-2*g)/3
for i,(t,d) in enumerate([('run','Returns captured output when the workload finishes.'),('spawn','Live stdin, stdout and stderr while it runs.'),('PTY','A pseudoterminal the host controls, including its size.')]):
    x=M+i*(cw3+g)
    rect(s,x,ry,cw3,1350000,fill=CARD)
    tb(s,x+250000,ry,cw3-500000,None,paras=[{'text':t,'sizePt':20,'bold':True,'color':CY},{'text':d,'sizePt':15,'color':MUTED}],cy=1350000)
footer(s,'Source: Microsoft Learn, MXC architecture (lifecycles) and platform support',17)
note(s,'One-shot 生命周期按“创建、执行、清理”组织，适合构建、测试、生成命令和报告处理。每次动作使用新的窄权限环境，可以避免为多个任务保留一个过宽的长期容器。\nState-aware 生命周期用于必须保留状态的工具环境或交互式会话，包含 provision、start、execute、stop 和 deprovision。stop 不等于 deprovision，即使停止失败也应尝试释放资源。run 返回最终结果，spawn 支持实时流，PTY 用于终端与屏幕尺寸控制。')

# ---------- 17 CODING AGENT (image left)
s=17
img(s,'hd-coding.jpg',0,0,PANEL,H,'Hand-drawn sketch panel: a robot coding inside a fenced project folder while a locked vault with a key sits outside')
xl=PANEL+M; cxr=W-xl-M
header(s,'04 · USE CASES','Secure coding agents',x=xl,cx=cxr)
y=tb(s,xl,1550000,cxr,'Build, test and push, without touching anything else.',size=18,color=INK)
perm=[('Project repo  C:\\src\\app','Read / write',CY),('Toolchain: node, python, git','Allowed',CY),('Package registry via proxy','Allowed',CY),('~/.ssh, tokens, Documents','Denied',AMB),('Any other network destination','Blocked',AMB)]
y+=200000; rh=560000
for t,st,c in perm:
    rect(s,xl,y,cxr,rh,fill=CARD)
    tb(s,xl+250000,y,cxr-2200000,t,size=15,color=INK,cy=rh)
    rect(s,xl+cxr-1750000,y+90000,1550000,rh-180000,fill=c,preset='roundRect',text=st,size=13,bold=True,color=BG)
    y+=rh+90000
tb(s,xl,y+150000,cxr,'In production: GitHub Copilot CLI uses MXC process isolation.',size=17,bold=True,color=CY)
footer(s,'Source: Windows Developer Blog, Jun 2026 · example policy',18,x=xl,cx=cxr)
note(s,'Coding Agent 需要访问当前仓库、执行工具链、调用 git 并可能下载依赖，但通常不需要读取 SSH key、其他仓库、个人文档或云凭据。Policy 应围绕当前开发任务授予权限，而不是继承开发者会话的完整权限。\n包管理器访问可通过可信代理限制域名和凭据。GitHub Copilot CLI 使用 MXC 进程隔离来约束动态生成并执行的代码，验证了“任务可执行、主机访问受限”的设计模型。具体路径与工具列表仍需按项目生成，而不是复用固定模板。')

# ---------- 18 DESKTOP & DATA (image right)
s=18
img(s,'hd-desktop.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a human relaxes at his own desk while a robot agent processes invoices at its own separate desktop behind a glass wall')
cxl=W-PANEL-2*M
header(s,'04 · USE CASES','Desktop and data agents',cx=cxl)
cards=[('Desktop automation agent','Keeps app state in a separate Windows account and session, driven through app APIs or UI automation.','Fit: session container; its network is open and SendInput is blocked'),('Sensitive-document agent','Reads one approved folder, writes one report, no network.','Fit: process container with a deny-by-default policy'),('Agents on Cloud PCs','Run beside existing work on Windows 365; process isolation works as on a physical PC.','Fit: Windows 365 + MXC, generally available')]
y=1600000; ch=1380000
for t,d,f in cards:
    rect(s,M,y,cxl,ch,fill=CARD)
    tb(s,M+250000,y,cxl-500000,None,paras=[{'text':t,'sizePt':18,'bold':True,'color':INK},{'text':d,'sizePt':14,'color':MUTED},{'text':f,'sizePt':14,'color':CY,'bold':True}],cy=ch)
    y+=ch+110000
footer(s,'Sources: Windows Developer Blog; Microsoft Learn',19,cx=cxl)
note(s,'Session container 适合需要独立 Windows 用户、应用状态和长生命周期的桌面自动化，例如使用 UI Automation、应用内 API 或可信自动化驱动器。它不能泛化为“模型直接使用鼠标点击 Host 桌面”：源码明确阻止 SendInput，Session 内窗口与剪贴板也只属于 Agent Session。\nERP 等应用的身份需要单独设计。Session 不继承员工 SSO；若在其中登录，凭据就位于该 Agent 用户的 Profile 中。应使用专用低权限机器人账号或 Host-side credential broker，而不是复制员工 Cookie。Windows 365 的 Process container 与物理 PC 行为一致；Session/虚拟化 Backend 仍需额外验证。')

# ---------- 19 ECOSYSTEM
s=19
header(s,'04 · USE CASES','Already in the ecosystem')
img(s,'partners-building-on-mxc-grid.png',M,1550000,7600000,2916211,'Partner logo reference: Anthropic, Box, Egnyte, Heidi, LM Studio, manus, Nous, NVIDIA, OpenAI, OpenClaw, Perplexity, Raycast, Replit, Simular and Unsloth')
rect(s,8290000,1500000,24000,4610000,fill=EDGE)
panel_x=8460000; panel_w=W-M-panel_x
status_rows=[
    ('SUPPORTS MXC TODAY','GitHub Copilot · OpenAI Codex\nOpenClaw · Replit\nLM Studio · Unsloth AI',1250000),
    ('INTEGRATED / BUILDING','NVIDIA OpenShell\nOpenAI and Meta',1200000),
    ('SUPPORT ANNOUNCED','Anthropic Claude Code · Box · Egnyte\nHeidi Health · Hermes Agent\nManus · Perplexity · Raycast · Simular',1800000),
]
y=1550000
for title,body,h in status_rows:
    rect(s,panel_x,y,panel_w,h,fill=CARD)
    tb(s,panel_x+220000,y+100000,panel_w-440000,title,size=11,color=CY,font=HF,bold=True,ls=1)
    tb(s,panel_x+220000,y+420000,panel_w-440000,body,size=13,color=INK,cy=h-500000)
    y+=h+180000
footer(s,'Ecosystem image: partner logo reference  ·  Status: verify each product before presenting',20)
note(s,'这张页把生态信息分成两部分。左侧伙伴 Logo 图用于展示 MXC 覆盖范围，右侧保留明确的支持状态，避免只看 Logo 产生误解。\n“Supports MXC today”包括 GitHub Copilot、OpenAI Codex、OpenClaw、Replit、LM Studio 和 Unsloth AI；NVIDIA OpenShell 已集成，OpenAI 和 Meta 正在基于新的隔离能力构建；其余产品属于已宣布支持或仍在推进的生态。支持阶段会持续变化，Logo 图用于说明广度，文字栏用于说明当前成熟度，正式演讲前仍应逐项核实。')

# ---------- 20 SDKs
s=20
header(s,'05 · INTEGRATING MXC','Three SDKs, one contract')
g=238800; cw3=(CW-2*g)/3; y0=1650000; ch=2700000
sdks=[('.NET','Microsoft.Mxc.Sdk','NuGet.org','GetPlatformSupport()'),('TypeScript','@microsoft/mxc-sdk','npm','getPlatformSupport()'),('Rust','mxc-sdk','crates.io','v1::platform_support()')]
for i,(t,pkg,reg,api) in enumerate(sdks):
    x=M+i*(cw3+g)
    rect(s,x,y0,cw3,ch,fill=CARD)
    yy=tb(s,x+250000,y0+250000,cw3-500000,t,size=22,bold=True,font=HF)
    yy=tb(s,x+250000,yy+100000,cw3-500000,pkg,size=15,bold=True,color=CY,font=MF)
    yy=tb(s,x+250000,yy+20000,cw3-500000,'from '+reg,size=14,color=MUTED)
    yy=tb(s,x+250000,yy+250000,cw3-500000,'DISCOVERY',size=12,color=FOOT,font=HF,bold=True,ls=1)
    tb(s,x+250000,yy,cw3-500000,api,size=12,color=INK,font=MF)
rect(s,M,4600000,CW,1200000,fill=CARD2,line=CY,lw=19050)
tb(s,M+300000,4600000,CW-600000,None,paras=[{'text':'The SDK converts your ContainerRequest into the stable 1.0.0 schema, so you never pick a schema version.','sizePt':17,'bold':True,'color':INK},{'text':'Windows 365 prerequisites: Windows 11 24H2 or later; .NET 8+ or Node.js 18+.','sizePt':15,'color':MUTED}],cy=1200000)
footer(s,'Sources: Microsoft Learn, MXC overview, platform support and Windows 365 support for MXC',21)
note(s,'MXC 提供 .NET、TypeScript 和 Rust SDK，使用统一的 ContainerRequest 描述文件、网络、UI、命令、超时和生命周期。SDK 将请求转换为稳定的 1.0.0 schema，开发者无需直接调用平台沙箱实现。\n集成应先执行平台能力发现，Windows 可结合 probe 确认请求将获得的隔离级别。SDK 属性名与序列化 JSON 字段不一定完全相同，例如 timeoutMs 映射到 process.timeout。Rust 使用 IsolationSession 或 WSLc 时还需启用对应 feature。')

# ---------- 21 REFERENCE ARCHITECTURE
s=21
header(s,'05 · INTEGRATING MXC','Reference architecture')
boxes=[('Agent planner','LLM decides the next step'),('Tool call','"run npm test"'),('Host policy engine','Trusted: picks the policy'),('MXC SDK','Request, lifecycle, diagnostics'),('Container backend','Enforces the boundary')]
g=382000; bw=(CW-4*g)/5; y0=1700000; bh=1250000
for i,(t,d) in enumerate(boxes):
    x=M+i*(bw+g)
    rect(s,x,y0,bw,bh,fill=CARD)
    tb(s,x+120000,y0,bw-240000,None,paras=[{'text':t,'sizePt':16,'bold':True,'color':INK},{'text':d,'sizePt':13,'color':MUTED}],align='center',cy=bh)
    if i<4: rect(s,x+bw+30000,y0+bh/2-170000,g-60000,340000,fill=CY,preset='chevron')
rect(s,M+bw/2,3150000,CW-bw,520000,fill=CARD2,line=CY,lw=19050,preset='leftArrow',text='Output, exit code and denials, validated by the host before the agent acts',size=14,color=INK)
g2=150000; cw3=(CW-2*g2)/3; y1=4000000; ch=1800000
side=[('Egress proxy','Attaches credentials and enforces a domain allow-list, outside the box.'),('Learning Mode','Activity reports tighten the policy before enforcement.'),('Enterprise control','Entra Agent ID and Agent 365 are GA; local-agent and Intune/MXC controls are preview.')]
for i,(t,d) in enumerate(side):
    x=M+i*(cw3+g2)
    rect(s,x,y1,cw3,ch,fill=CARD)
    tb(s,x+250000,y1,cw3-500000,None,paras=[{'text':t,'sizePt':18,'bold':True,'color':CY},{'text':d,'sizePt':15,'color':MUTED}],cy=ch)
footer(s,'Architecture pattern based on Microsoft Learn and the Windows Developer Blog; adapt to your agent runtime',22)
note(s,'Agent 规划器产生工具调用后，请求首先进入 Host 的可信策略引擎。Host 按任务类型选择预定义 Policy，通过 MXC SDK 创建容器，再由 Backend 在边界内执行。\n返回路径同样重要：Host 必须检查退出码、输出格式、拒绝事件和生成文件，验证通过后才允许 Agent 使用结果。出站代理负责网络和凭据，Learning Mode 用于权限收敛。企业控制面中，Entra Agent ID 和 Agent 365 已 GA；本地 Agent 自动身份、Agent 365 本地治理以及 Intune/MXC 容器边界仍处于 Preview。隔离解决未授权执行，结果可信度仍由应用层保证。')

# ---------- 22 FIVE STEPS + CODE
s=22
header(s,'05 · INTEGRATING MXC','Integrate in five steps')
st=[('Discover','getPlatformSupport() and probe() before launching.'),('Describe','Build a ContainerRequest: files, network, UI, command, timeout.'),('Learn','Run in Learning Mode; review the activity report.'),('Enforce','Remove learning settings; test allowed and denied cases.'),('Operate','Handle timeouts and failures; explain denials to users.')]
lw0=5300000; y=1600000; rh=820000
for i,(t,d) in enumerate(st):
    tb(s,M,y,700000,str(i+1),size=30,color=CY,font=HF,bold=True,cy=rh)
    tb(s,M+700000,y,lw0-700000,None,paras=[{'text':t,'sizePt':18,'bold':True,'color':INK},{'text':d,'sizePt':14,'color':MUTED}],cy=rh)
    y+=rh+60000
cx0=M+lw0+300000; cw_=W-M-cx0; cy0=1600000; chh=4300000
rect(s,cx0,cy0,cw_,chh,fill=CODEBG,line=EDGE)
tb(s,cx0+250000,cy0+120000,cw_-500000,'host.ts  ·  illustrative pseudocode',size=12,color=FOOT,font=MF)
code=['const support = await getPlatformSupport();','if (!support.isSupported) throw support.reason;','','const request = {','  command: "npm test",','  workingDirectory: repo,','  filesystem: { readwritePaths: [repo] },','  network: { egress: { default: "deny" } },','  timeoutMs: 120_000,','};','const result = await runContained(request);','validate(result); // never trust blindly']
tb(s,cx0+250000,cy0+520000,cw_-400000,None,paras=[{'text':l if l else ' '} for l in code],size=13,color=CODE,font=MF)
footer(s,'Code is illustrative; check the MXC SDK guides for exact type and method names',23)
note(s,'集成分五步执行：运行时发现平台能力；构造默认拒绝的 ContainerRequest；使用 Learning Mode 收集访问需求；人工审核后切换 Enforcement；在运维阶段处理失败、取消、超时和清理。\n请求必须覆盖命令、工作目录、文件权限、网络、超时和隔离方式。拒绝事件应作为显式结果返回 Host，由应用解释原因或发起审批，不应转换为静默失败或无限重试。示例代码仅说明结构，具体 API 以选用 SDK 的版本和类型定义为准。')

# ---------- 23 HOST JOB LIST
s=23
header(s,'05 · INTEGRATING MXC','Your host\'s job list')
items=['Set a timeout; treat launch failure, workload failure and cancellation separately.','Clean up even when a later phase fails; always attempt to deprovision.','One narrow sandbox per action beats one broad policy.','Write outputs to a dedicated folder; teardown does not undo host writes.','Explain denials to the user; never fail silently.','A human approves every policy change Learning Mode suggests.']
g=150000; tw=(CW-g)/2; th=1180000; y0=1650000
for i,t in enumerate(items):
    x=M+(i%2)*(tw+g); y=y0+(i//2)*(th+130000)
    rect(s,x,y,tw,th,fill=CARD)
    rect(s,x+250000,y+th/2-240000,480000,480000,fill=CY,preset='ellipse',text='✓',size=18,bold=True,color=BG)
    tb(s,x+950000,y,tw-1200000,t,size=17,color=INK,cy=th)
footer(s,'Source: Microsoft Learn, MXC architecture (host responsibilities)',24)
note(s,'Host 需要设置超时，并区分启动失败、Workload 失败、用户取消和清理失败。状态型容器即使在停止阶段失败，也应继续尝试 deprovision。\n权限设计应采用一次动作一份窄策略，避免长期容器累积宽权限。结果必须写入专用输出目录并在消费前验证。拒绝事件应返回结构化原因，Learning Mode 的权限建议必须经过人工审批。')

# ---------- 24 DEMO 1 (PROCESSCONTAINER: FILES + NETWORK)
s=24
header(s,'06 · LIVE DEMO 1','ProcessContainer: files & network')
cx_=5550000; cy0=1620000; chh=4300000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+250000,cy0+140000,cx_-500000,'policy.json  ·  @microsoft/mxc-sdk/v1',size=12,color=FOOT,font=MF,cy=340000)
pc_code='{\n  "containment": { "type": "processcontainer" },\n  "filesystem": {\n    "readonlyPaths": ["demo/input"],\n    "readwritePaths": ["demo/output"],\n    "deniedPaths": ["demo/protected/denied.txt"]\n  },\n  "network": {\n    "egress": { "default": "deny" },\n    "ingress": { "hostLoopback": "allow" }\n  }\n}'
tb(s,M+250000,cy0+480000,cx_-500000,None,paras=[{'text':l} for l in pc_code.split('\n')],size=12.5,color=CODE,font=MF)
rect(s,M+250000,cy0+3340000,cx_-500000,720000,fill=CARD2,line=CY,lw=12700)
tb(s,M+420000,cy0+3340000,cx_-840000,None,paras=[
    {'text':'PS> npm --prefix demo run demo:process','sizePt':12,'bold':True,'color':INK},
    {'text':'getPlatformSupport() → probe() → run()','sizePt':12,'color':CY},
],font=MF,cy=720000)
rx=M+cx_+320000; rw=W-M-rx; ry0=1620000; rcard_h=1320000; rgap=170000
pc_cards=[
    ('✓  Scoped read-in / write-out','readonlyPaths + readwritePaths','Workload reads demo/input and writes report.txt into its dedicated output folder.',CY),
    ('✕  Protected file denied','deniedPaths: ["protected/denied.txt"]','The same workload attempting to read a secret file is blocked by the OS kernel.',AMB),
    ('⇄  Directional network leash','egress: deny · hostLoopback: allow / deny','Reaches an explicitly permitted loopback/proxy endpoint; blocks unapproved access.',CY),
]
for idx,(t,sub,desc,col_) in enumerate(pc_cards):
    yy=ry0+idx*(rcard_h+rgap)
    rect(s,rx,yy,rw,rcard_h,fill=CARD,line=col_,lw=19050)
    tb(s,rx+240000,yy,rw-480000,None,paras=[
        {'text':t,'sizePt':17,'bold':True,'color':INK},
        {'text':sub,'sizePt':12.5,'bold':True,'color':col_},
        {'text':desc,'sizePt':13.5,'color':MUTED},
    ],cy=rcard_h)
footer(s,'Live host app uses @microsoft/mxc-sdk/v1  ·  5 interactive steps in Demo Studio',25)
note(s,'Demo 1 将文件最小权限与方向性网络控制合并为一个完整的 ProcessContainer 演示，共五个步骤。Host 首先调用 getPlatformSupport 与 Windows probe 校验 base-container 能力，确保所有访问控制可在内核层执行，绝不回退到普通进程。\n文件维度同时验证正向与负向路径：允许读取 demo/input 并向专用目录 demo/output/report.txt 写入结果，而同一策略下读取 demo/protected/denied.txt 则被操作系统直接拒绝。网络维度将 egress 默认设为拒绝，分别对比 hostLoopback 为 allow 与 deny 时对本地受控端点的连接结果；若需域名级白名单，可配合 runtimeConfig.networkProxy 与 allowedProxyPeer 交由可信代理执行。')

# ---------- 25 DEMO 2 (SESSION PTY, image right)
s=25
img(s,'hd-blocked.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a sneaky amber tentacle pulls a robot agent toward a door, but a force field blocks it')
cxl=W-PANEL-2*M
header(s,'06 · LIVE DEMO 2','Incident triage in isolation',cx=cxl,title_cy=600000)
beats=[('Stage','Host places INC-1042 and its read-only approval in the shared workspace.'),('Open PTY','A dedicated Windows operator session launches PowerShell with MXC-owned I/O.'),('Triage','Check Spooler and W32Time, retain context, and write auditable evidence.'),('Close','Host validates the report, stops the session and deprovisions the operator.')]
railx=M+150000; y=1650000
ys=[]
for i,(t,d) in enumerate(beats):
    ys.append(y)
    y=tb(s,M+600000,y,cxl-600000,None,paras=[{'text':t,'sizePt':19,'bold':True,'color':CY if i!=3 else AMB},{'text':d,'sizePt':15,'color':MUTED}])+200000
rect(s,railx+100000-12700,ys[0]+200000,25400,ys[-1]-ys[0],fill=EDGE)
for i,yy in enumerate(ys):
    rect(s,railx,yy+100000,200000,200000,fill=CY if i!=3 else AMB,preset='ellipse')
footer(s,'PS> npm --prefix demo run demo:pty  ·  6 stateful PTY steps in Demo Studio',26,cx=cxl)
note(s,'Demo 2 展示有状态的 IsolationSession 与交互式 PTY，对应常见的 L1/L2 运维事故排查。INC-1042 表示打印和时间服务异常：Host 先把工单与“只读诊断”审批写入共享 workspace，再创建独立 Windows 操作员账号与 Session ID，启动 120×30 PowerShell PTY。Agent 校验工单范围、查询真实 Spooler 与 W32Time 状态并生成审计证据；期间人工操作员可随时接管终端输入命令，最后由 Host 验证证据并停止、清理临时账号。\n在交互式 Demo Studio 中，观众可直观看到员工当前登录账号/Session ID 与隔离操作员账号/Session ID 的并排对比。')

# ---------- 26 DEMO 3 (HYPERLIGHT MICROVM)
s=26
header(s,'06 · LIVE DEMO 3','Untrusted code in a MicroVM')
cx_=5550000; cy0=1620000; chh=4300000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
rect(s,M+cx_-1700000,cy0+140000,1450000,340000,fill=AMB,preset='roundRect',text='Experimental',size=11,bold=True,color=BG)
tb(s,M+250000,cy0+140000,cx_-2000000,'hyperlight-policy.json',size=12,color=FOOT,font=MF,cy=340000)
vm_code='{\n  "version": "1.1.0-alpha",\n  "containment": "hyperlight",\n  "hyperlight": {\n    "runtime": "python"\n  },\n  "process": {\n    "commandLine": "python eval_untrusted.py",\n    "timeout": 30000\n  }\n}'
tb(s,M+250000,cy0+580000,cx_-500000,None,paras=[{'text':l} for l in vm_code.split('\n')],size=13.5,color=CODE,font=MF)
rect(s,M+250000,cy0+3320000,cx_-500000,740000,fill=CARD2,line=CY,lw=12700)
tb(s,M+420000,cy0+3320000,cx_-840000,None,paras=[
    {'text':'PS> npm --prefix demo run demo:microvm','sizePt':12,'bold':True,'color':INK},
    {'text':'Warm snapshot restore: ~90 ms · Unikraft 5.15','sizePt':12,'color':CY},
],font=MF,cy=740000)
rx=M+cx_+320000; rw=W-M-rx; ry0=1620000; rcard_h=1320000; rgap=170000
vm_cards=[
    ('✓  Hardware VM kernel','Unikraft 5.15 on Hypervisor','Runs inside a hardware-virtualized partition, never sharing the Windows NT kernel.',CY),
    ('✕  Zero host surface','No host FS, env, or sockets','Host C:\\, environment variables, and TCP sockets (Errno 13) remain unreachable.',AMB),
    ('↺  Instant snapshot wipe','<100 ms restore & reset','Every call restores a warmed memory snapshot; /tmp writes vanish on exit.',CY),
]
for idx,(t,sub,desc,col_) in enumerate(vm_cards):
    yy=ry0+idx*(rcard_h+rgap)
    rect(s,rx,yy,rw,rcard_h,fill=CARD,line=col_,lw=19050)
    tb(s,rx+240000,yy,rw-480000,None,paras=[
        {'text':t,'sizePt':17,'bold':True,'color':INK},
        {'text':sub,'sizePt':12.5,'bold':True,'color':col_},
        {'text':desc,'sizePt':13.5,'color':MUTED},
    ],cy=rcard_h)
footer(s,'Experimental Hyperlight MicroVM backend via MXC 1.1.0-alpha contract',27)
note(s,'Demo 3 展示最强的硬件虚拟化隔离边界——Hyperlight MicroVM。当 Agent 需要直接执行模型临时生成的 Python 代码、未知第三方插件或高风险解析逻辑时，仅靠进程级 ACL 仍共享宿主操作系统内核；而传统完整虚拟机冷启动又往往需要数秒。\nMXC 内置的实验性 hyperlight Backend 基于 Windows Hypervisor Platform 启动轻量级 Unikraft unikernel 微型虚拟机，并通过预热内存快照在约 90 毫秒内完成恢复与代码执行。演示依次验证五件事：探测并校验预热快照、确认 Guest 运行在独立 Unikraft 内核而非 Windows 内核、在百毫秒内完成不可信 Python 对账计算、证明宿主机 C 盘与环境变量及网络套接字全部不可达，以及每次调用结束后内存状态彻底重置。该能力必须通过 1.1.0-alpha 契约与 --experimental 显式启用。')

# ---------- 27 CLOSER
s=27
img(s,'hd-closer.jpg',0,0,W,H,'Hand-drawn sketch illustration: three robot agents in bubbles and an engineer giving a thumbs up at sunrise over a city')
x=M+150000; cx=5900000
tb(s,x,650000,cx,'KEY TAKEAWAYS',size=13,color=CY,font=HF,bold=True,ls=2)
y=tb(s,x,1000000,cx,'Box the agent,\nnot the ambition',size=38,color=INK,font=HF,bold=True)
tk=['1   Push guardrails down to the OS. Prompts alone can\'t enforce.','2   MXC: one policy, the right sandbox on Windows, macOS and Linux.','3   Learn first, then enforce least privilege.','4   Containment isn\'t trust. Validate every result.']
y=tb(s,x,y+150000,cx,None,paras=[{'text':t} for t in tk],size=16,color=INK)
y=tb(s,x,y+150000,cx,None,paras=[{'text':'learn.microsoft.com/windows/apps/develop/mxc'},{'text':'github.com/microsoft/mxc'}],size=14,color=CY)
tb(s,x,y+150000,cx,'Thank you  ·  Q&A',size=24,color=INK,font=HF,bold=True)
note(s,'实施 Agent 执行隔离应遵循四项原则：AI 层约束与操作系统层强制控制并用；采用统一 Policy，由平台选择合适 Backend；先观察并人工审核，再收紧为最小权限 Enforcement；隔离后仍需验证输出和副作用。\n建议从 Coding Agent 或敏感文档处理场景开始 PoC，先定义精确的文件与网络权限，再覆盖正常执行、越权拒绝、异常清理和审计路径。稳定后再扩展身份、企业治理和桌面自动化能力。')


# ===================== WRITER (python-pptx) =====================
PRESETS={'rect':MSO_SHAPE.RECTANGLE,'roundRect':MSO_SHAPE.ROUNDED_RECTANGLE,'ellipse':MSO_SHAPE.OVAL,
         'rightArrow':MSO_SHAPE.RIGHT_ARROW,'leftArrow':MSO_SHAPE.LEFT_ARROW,
         'chevron':MSO_SHAPE.CHEVRON,'homePlate':MSO_SHAPE.PENTAGON}
ALIGN={'left':PP_ALIGN.LEFT,'center':PP_ALIGN.CENTER,'right':PP_ALIGN.RIGHT}
rgb=lambda h:RGBColor.from_string(h)

def write_text(tf,o):
    tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    paras=o.get('paragraphs') or [{'text':o['text']}]
    for i,p in enumerate(paras):
        para=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        para.alignment=ALIGN[p.get('align',o.get('align','center'))]
        lines=p['text'].split('\n')
        for j,line in enumerate(lines):
            if j: para.add_line_break()
            r=para.add_run(); r.text=line; f=r.font
            f.name=o.get('fontFamily',BF); f.size=Pt(p.get('sizePt',o.get('sizePt',16)))
            f.bold=p.get('bold',o.get('bold',False)); f.color.rgb=rgb(p.get('color',o.get('color',INK)))
            if o.get('letterSpacingPt'): r._r.get_or_add_rPr().set('spc',str(int(o['letterSpacingPt']*100)))

def add_picture(sl,o,img_dir):
    path=os.path.join(img_dir,os.path.basename(o['path']))
    pic=sl.shapes.add_picture(path,Emu(o['x']),Emu(o['y']),Emu(o['cx']),Emu(o['cy']))
    iw,ih=Image.open(path).size; fr=o['cx']/o['cy']; ir=iw/ih
    if ir>fr: c=(1-fr/ir)/2; pic.crop_left=pic.crop_right=c        # centre-crop ("cover")
    else:     c=(1-ir/fr)/2; pic.crop_top=pic.crop_bottom=c
    pic._element.nvPicPr.cNvPr.set('descr',o.get('altText',''))

def add_appear_animation(sl,item_shapes):
    next_id=3
    item_nodes=[]
    for shapes in item_shapes:
        outer_id=next_id; inner_id=next_id+1; next_id+=2
        effects=[]
        for index,shape in enumerate(shapes):
            effect_id=next_id; behavior_id=next_id+1; next_id+=2
            node_type='clickEffect' if index==0 else 'withEffect'
            effects.append(
                f'<p:par><p:cTn id="{effect_id}" presetID="1" presetClass="entr" '
                f'presetSubtype="0" fill="hold" grpId="0" nodeType="{node_type}">'
                '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                '<p:childTnLst><p:set><p:cBhvr>'
                f'<p:cTn id="{behavior_id}" dur="1" fill="hold">'
                '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{shape.shape_id}"/></p:tgtEl>'
                '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                '</p:childTnLst></p:cTn></p:par>'
            )
        item_nodes.append(
            f'<p:par><p:cTn id="{outer_id}" fill="hold">'
            '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            '<p:childTnLst><p:par>'
            f'<p:cTn id="{inner_id}" fill="hold">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(effects)}</p:childTnLst>'
            '</p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
        )
    build_nodes=[]
    for shapes in item_shapes:
        for index,shape in enumerate(shapes):
            anim_bg=' animBg="1"' if index==0 else ''
            build_nodes.append(
                f'<p:bldP spid="{shape.shape_id}" grpId="0"{anim_bg}/>'
            )
    timing=parse_xml(
        f'<p:timing {nsdecls("p")}>'
        '<p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst>{"".join(item_nodes)}</p:childTnLst>'
        '</p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0">'
        '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0">'
        '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn>'
        '</p:par></p:tnLst>'
        f'<p:bldLst>{"".join(build_nodes)}</p:bldLst>'
        '</p:timing>'
    )
    sl._element.insert_element_before(timing,'p:extLst')

def build(out,img_dir):
    out_dir=os.path.dirname(os.path.abspath(out))
    os.makedirs(out_dir,exist_ok=True)
    prs=Presentation(); prs.slide_width,prs.slide_height=Emu(W),Emu(H)
    blank=prs.slide_layouts[6]
    for k in sorted(slides):
        sl=prs.slides.add_slide(blank)
        sl.background.fill.solid(); sl.background.fill.fore_color.rgb=rgb(BG)
        rendered=[]
        for o in slides[k]:
            op=o['op']
            if op=='add_image': add_picture(sl,o,img_dir)
            elif op=='add_text_box':
                sh=sl.shapes.add_textbox(Emu(o['x']),Emu(o['y']),Emu(o['cx']),Emu(o['cy'])); write_text(sh.text_frame,o)
                rendered.append(sh)
            elif op=='add_shape':
                sh=sl.shapes.add_shape(PRESETS[o['preset']],Emu(o['x']),Emu(o['y']),Emu(o['cx']),Emu(o['cy']))
                sh.shadow.inherit=False
                sh.fill.solid(); sh.fill.fore_color.rgb=rgb(o['fill'])
                if o.get('line'): sh.line.color.rgb=rgb(o['line']); sh.line.width=Emu(o.get('lineWidth',12700))
                else: sh.line.fill.background()
                if o.get('rot'): sh.rotation=o['rot']/60000
                if 'text' in o: write_text(sh.text_frame,o)
                rendered.append(sh)
        if k==5:
            add_appear_animation(sl,[rendered[start:start+5] for start in (2,7,12)])
        elif k==7:
            add_appear_animation(sl,[rendered[start:start+4] for start in (2,6,10,14,18)])
        if k in notes: sl.notes_slide.notes_text_frame.text=notes[k]
    prs.save(out); print(f'Saved {out}  ({len(prs.slides)} slides)')

if __name__=='__main__':
    here=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    root=os.path.dirname(here)
    ap=argparse.ArgumentParser()
    ap.add_argument('--out',default=os.path.join(root,'dist','pptx','Guardrails with Sandboxing - MXC.pptx'))
    ap.add_argument('--images',default=os.path.join(here,'assets','images'))
    a=ap.parse_args(); build(a.out,a.images)
