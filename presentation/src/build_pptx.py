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
from pptx.oxml.ns import qn
from PIL import Image

# ===================== THEME =====================
# ===================== LAYOUT ====================
W,H=12192000,6858000; M=457200; CW=W-2*M
BG='F3EEE4'; CARD='E9E2D4'; CARD2='E9E2D4'; EDGE='C9C0AE'; INK='2B2B2B'; MUTED='57534C'; FOOT='6E685F'
CY='3F607E'; AMB='A84A35'; CODEBG='FAF7F1'; CODE='2F4B66'
HF='Georgia'; BF='Calibri'; MF='Courier New'
PANEL=4572000   # width of the side illustration panel
# ===================== HELPERS ===================
slides={}  # idx -> list of patches
notes={}
def P(s,op): slides.setdefault(s,[]).append(op)
WF={'Georgia':0.6,'Arial':0.56,'Calibri':0.5,'Courier New':0.6}
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
def header(s,eyebrow,title,x=M,cx=CW):
    tb(s,x,400000,cx,eyebrow,size=13,color=CY,font=HF,bold=True,ls=2)
    tb(s,x,760000,cx,title,size=32,color=INK,font=HF,bold=True)
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
note(s,'【开场 · 1 分钟】大家好，我是 Yingting。先请大家看封面上的机器人：它并不是不能工作，而是被放在一个透明的盒子里。盒子里面，机器人可以读被允许的资料、运行代码、完成任务；盒子外面，钥匙、机密和你真正关心的环境仍然保持安全。\n这正是今天要讨论的问题。过去我们把 AI 当成一个会回答问题的聊天工具，现在它开始替我们操作电脑、运行命令和调用服务。这样一来，安全问题就不再只是“它回答得对不对”，而是“它一旦做错了，能不能限制影响范围”。今天的主角 MXC，就是微软为这个问题提供的一层执行隔离。')

# ---------- 1 AGENDA
s=1
header(s,'AGENDA','Our 45-minute journey')
items=[('01','5 min','Why sandboxing','Risks of agents that act'),('02','7 min','Meet MXC','What it is and how it works'),('03','10 min','Features & platforms','Backends, files, network, modes'),('04','5 min','Use cases','Coding, desktop and data agents'),('05','7 min','Integration','SDKs and reference architecture'),('06','8 min','Live demos','Files, network, coding agent'),('07','3 min','Takeaways & Q&A','What to do next')]
col=CW/7; ry=3150000
rect(s,M,ry-19050,CW,38100,fill=EDGE)
for i,(n,mins,t,d) in enumerate(items):
    cx0=M+col*i; bw=col-80000
    tb(s,cx0,1900000,bw,n,size=30,color=CY,font=HF,bold=True)
    tb(s,cx0,2450000,bw,mins,size=14,color=MUTED)
    rect(s,cx0+60000,ry-120000,240000,240000,fill=CY,preset='ellipse')
    y=tb(s,cx0,3450000,bw,t,size=17,color=INK,bold=True,cy=hgt([{'text':'Features & platforms','sizePt':17,'bold':True}],bw,BF))
    tb(s,cx0,y+40000,bw,d,size=14,color=MUTED)
footer(s,'Level: L300 → L400  ·  Audience: architects, developers, security engineers',2)
note(s,'【议程 · 约 30 秒】今天的顺序是先回答“为什么需要隔离”，再解释 MXC 到底是什么、它提供哪些能力。之后我们会看几个真实的使用场景，讨论应用应该怎样接入，最后用三个 Demo 把前面的概念串起来。\n重点不是只记住一组产品名词。我们希望最后大家能判断：自己的 Agent 需要访问哪些文件、能联网到哪里、凭据应该放在哪里，以及怎样验证权限真的生效了。')

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
note(s,'【01 · 为什么需要沙箱 · 约 1 分钟】先澄清一个很重要的变化。传统聊天机器人主要负责生成文字，真正点击按钮、运行命令的人还是你，所以说错话通常只会造成信息错误。Agent 不一样，它的目标不是只给一个答案，而是自己完成任务。\n例如，Agent 可能会打开仓库、安装依赖、运行 shell、修改文件，或者调用外部 API。更关键的是，它执行的命令往往是运行时生成的，没有人会逐行审查。因此风险从“内容不可靠”升级成了“动作可能影响真实系统”。如果 Agent 继承了你当前用户的全部权限，一个错误决定就可能波及你的整个工作会话。')

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
note(s,'【01 · 四类典型风险 · 约 1.5 分钟】这些风险不是只存在于恶意用户输入，也不只发生在黑客攻击的时候。第一类是 Prompt Injection：Agent 阅读到的一份 README、网页或文档，可能隐藏了一条“请运行这条命令”的指令。它没有能力天然分辨“资料内容”和“可信指令”。\n第二类是凭据泄露。Agent 如果能读取用户目录，就可能碰到 token、SSH key 和浏览器资料。第三类是数据外传：即使读取的是合法文件，只要网络没有限制，它仍可能把数据发送到不该去的地址。第四类是越权。例如用户只说“帮我更新网站”，Agent 为了更快完成，可能直接修改生产配置。这个动作从模型视角可能像个合理捷径，但从任务边界看已经明显超范围了。')

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
note(s,'【01 · 护栏为什么要下沉 · 约 1 分钟】这里有一句很值得记住的话：Agent 不能做自己的安全负责人。系统提示词、内容过滤器和工具规则当然都有用，但它们本质上是在告诉模型“不要这样做”。只要存在理解偏差、提示词注入或工具错误，模型仍然有可能尝试越界。\n所以安全设计需要两层。上面一层控制模型“想做什么”，下面一层由操作系统决定“实际上能做什么”。例如模型可以计划读取某个文件，但如果进程根本没有目录权限，或者网络策略拒绝该目标，操作就会被真正阻止。模型负责推理，系统负责执行边界，两者不能互相替代。')

# ---------- 5 THREE PILLARS
s=5
header(s,'01 · WHY SANDBOXING','Three pillars for agents on Windows')
cw3=(CW-2*238800)/3; y0=1650000; ch=3900000
pill=[('Containment','Limit what an agent can access and do','Microsoft Execution Containers (MXC). Windows 365 support for MXC is also generally available.','Generally available',True),('Identity','Tell agent activity apart from a person','Microsoft Entra will help distinguish agent activity from user activity.','Coming soon',False),('Manageability','Govern access and monitor agents','Agent 365 controls for on-device agents; Intune to manage MXC process containers.','Coming soon',False)]
for i,(t,sub,body,state,lift) in enumerate(pill):
    x=M+i*(cw3+238800)
    if lift: rect(s,x,y0,cw3,ch,fill=CARD2,line=CY,lw=25400)
    else: rect(s,x,y0,cw3,ch,fill=CARD)
    pw=2100000 if lift else 1500000
    rect(s,x+300000,y0+300000,pw,400000,fill=CY if lift else EDGE,preset='roundRect',text=state,size=13,color=BG if lift else INK,bold=True)
    y=tb(s,x+300000,y0+900000,cw3-600000,t,size=26,bold=True,font=HF)
    y=tb(s,x+300000,y+60000,cw3-600000,sub,size=17,color=CY)
    tb(s,x+300000,y+150000,cw3-600000,body,size=15,color=MUTED)
footer(s,'Source: Windows Developer Blog, Oct 7, 2026  ·  Today\'s focus: Containment',6)
note(s,'【01 · 三根支柱 · 约 1 分钟】微软把 Windows 上 Agent 的安全分成三件事。第一是 Containment，也就是今天重点讲的隔离：把 Agent 能碰到的文件、网络、进程和界面限制在一个明确范围内。第二是 Identity，要让系统知道这是人做的，还是某个 Agent 代表人做的。第三是 Manageability，让企业能够统一下发策略、检查状态并持续治理。\n这三件事彼此有关，但成熟度不同。MXC 本身以及 Windows 365 对 MXC 的支持已经正式可用；Entra 对 Agent 身份的支持、Agent 365 和相关 Intune 管理能力仍需关注后续发布。给客户介绍时，必须把“今天已经可用”和“路线图能力”讲清楚，不要把计划包装成现货。')

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
note(s,'【02 · 认识 MXC · 约 1.5 分钟】那 MXC 到底是什么？可以把它理解成一个开源、由策略驱动的执行层。你的 Host 应用告诉 MXC：“这个任务可以读哪些目录、写哪些目录、访问哪些网络目标、运行什么命令”，MXC 再把这些要求交给当前操作系统真正执行。\n这里的 Container 不要机械地理解成 Docker 容器，也不一定是一台虚拟机。它可能是进程级沙箱、独立会话、WSL 环境，或者实验性的虚拟化隔离。外层统一，底层可以随平台变化。MXC 采用 MIT 许可证，提供 .NET、TypeScript 和 Rust SDK，并覆盖 Windows、macOS 与 Linux。画面中的守卫拿着规则书，决定什么东西可以递进盒子，也可以很好地帮助听众理解 Host 与 Policy 的关系。')

# ---------- 7 FIVE CONCEPTS
s=7
header(s,'02 · UNDERSTANDING MXC','Five concepts to remember')
gap=150000; cw5=(CW-4*gap)/5; y0=1650000; ch=3600000
cons=[('Host app','Your trusted app or agent runtime. Picks permissions, supplies inputs, handles results.'),('Workload','The command, script or process that runs under containment.'),('Policy','The access you permit: readable files, writable folders, network destinations, UI.'),('Execution config','The policy plus command, working directory, runtime settings and containment choice.'),('Backend','The OS mechanism that enforces it: process sandbox, separate session or VM.')]
for i,(t,d) in enumerate(cons):
    x=M+i*(cw5+gap)
    if t=='Policy': rect(s,x,y0,cw5,ch,fill=CARD2,line=CY,lw=25400)
    else: rect(s,x,y0,cw5,ch,fill=CARD)
    y=tb(s,x+200000,y0+250000,cw5-400000,f'0{i+1}',size=28,color=CY,font=HF,bold=True)
    y=tb(s,x+200000,y+80000,cw5-400000,t,size=19,bold=True,cy=hgt([{'text':'Execution config','sizePt':19,'bold':True}],cw5-400000,BF))
    tb(s,x+200000,y+100000,cw5-400000,d,size=14,color=MUTED)
tb(s,M,5450000,CW,'The host decides. The policy describes. The backend enforces.',size=20,bold=True)
footer(s,'Source: Microsoft Learn, Microsoft Execution Containers overview',8)
note(s,'【02 · 五个基础概念 · 约 1.5 分钟】理解 MXC 时，只需抓住这五个角色。Host app 是你信任的应用或 Agent 运行时，它负责判断用户的请求需要多少权限。Workload 是真正要运行的不可信代码，可能是一条命令、一个脚本，也可能是一整个子进程。Policy 用白名单描述允许的资源。Execution config 则把 Policy、命令、工作目录、超时和隔离方式组合成一次执行请求。Backend 是最后真正动手的操作系统机制。\n最核心的一点是：Policy 由 Host 决定，不由 Workload 自己决定。Host 做授权，Policy 把授权写下来，Backend 负责拒绝策略之外的操作。只要这个信任关系倒过来，沙箱就会失去意义。')

# ---------- 8 REQUEST FLOW
s=8
header(s,'02 · UNDERSTANDING MXC','How one request flows')
steps=[('HOST','Discover','Ask the SDK which backends this machine supports.'),('HOST','Describe','Build a policy and execution request, deny by default.'),('SDK','Launch','Start the workload through the SDK.'),('MXC','Translate','Map the request to the selected backend\'s operations.'),('HOST','Return','Handle output, exit status, warnings and cleanup.')]
sw=2200000; step=(CW-sw)/4; yc=2150000
for i,(who,t,d) in enumerate(steps):
    x=M+i*step
    tb(s,x+150000,1650000,sw-300000,who,size=13,color=CY,font=HF,bold=True,ls=2)
    rect(s,x,yc,sw,850000,fill=CARD2 if i!=3 else CY,preset='chevron' if i else 'homePlate',text=t,size=18,bold=True,color=INK if i!=3 else BG)
    tb(s,x+150000,yc+1050000,sw-350000,d,size=15,color=MUTED)
rect(s,M,4750000,CW,800000,fill=CARD,line=AMB,lw=19050)
tb(s,M+300000,4750000,CW-600000,'Backends have different security properties. Always test both the allowed and the denied operations.',size=17,color=INK,cy=800000)
footer(s,'Source: Microsoft Learn, How MXC works',9)
note(s,'【02 · 一次请求如何执行 · 约 1.5 分钟】一次 MXC 请求可以分成五步。Host 先查询当前机器支持什么隔离能力，再构造一个默认拒绝的策略和执行请求；SDK 创建工作负载后，MXC 把统一请求转换成具体 Backend 能理解的操作；任务结束后，Host 还要处理输出、退出码、警告和清理。\n不要只关注“请求启动成功”。不同 Backend 的能力和安全语义并不完全一样，同一份策略在两个平台上也可能有不同约束。因此上线测试至少要做两类：允许的操作能够正常完成，未允许的操作必须明确失败。只测第一类，相当于确认系统能工作，却没有确认边界真的存在。')

# ---------- 9 FOUR LAYERS
s=9
header(s,'02 · UNDERSTANDING MXC','Four layers, one boundary')
lays=[('Host app / agent runtime','Classifies the task, chooses trusted policy defaults, handles output and approval.'),('MXC SDK & policy contract','Discovers support, turns intent into a versioned config, manages lifecycle.'),('MXC execution layer','Validates the config, resolves a backend, translates to platform operations.'),('Containment backend','Creates the environment and enforces what the platform supports.')]
y=1600000; lh=820000; lw_=3400000
for i,(t,d) in enumerate(lays):
    if i==0: rect(s,M,y,lw_,lh,fill=CARD2,line=CY,lw=25400,text=t,size=17,bold=True,align='left')
    else: rect(s,M,y,lw_,lh,fill=CARD2,text=t,size=17,bold=True,align='left')
    rect(s,M+lw_+100000,y,CW-lw_-100000,lh,fill=CARD)
    tb(s,M+lw_+400000,y,CW-lw_-700000,d,size=16,color=MUTED,cy=lh)
    y+=lh+110000
tb(s,M,y+120000,CW,'Rule #1: an agent may suggest a permission, but only trusted host logic or a human can grant it.',size=18,bold=True,color=INK)
footer(s,'Source: Microsoft Learn, MXC architecture and containment models',10)
note(s,'【02 · 四层结构 · 约 1.5 分钟】从上往下看，最上层是 Host，它理解用户意图并选择可信的默认策略。第二层是 SDK，负责把意图转成统一配置、发现平台能力和管理生命周期。第三层是 MXC 执行层，校验请求并选择合适 Backend。最底层才是真正创建边界并执行限制的平台机制。\n这里也要说明信任边界在哪里。Agent 可以根据任务建议“我需要访问这个目录”，但最终是否授权，必须由可信 Host 的规则或人工审批决定。与此同时，Host 自己在沙箱外执行的代码不受这次 MXC 容器保护，所以 Host 的输入校验、结果校验和代理逻辑仍然必须可靠。')

# ---------- 10 PICK THE RIGHT BOX
s=10
header(s,'03 · FEATURES & PLATFORMS','Pick the right box')
ry=1900000
tb(s,M,1550000,4000000,'Lighter and faster',size=14,color=MUTED)
tb(s,W-M-4000000,1550000,4000000,'Stronger isolation',size=14,color=MUTED,align='right')
rect(s,M,ry,CW-350000,38100,fill=EDGE)
rect(s,W-M-400000,ry-150000,400000,340000,fill=EDGE,preset='rightArrow')
gap=180000; cw4=(CW-3*gap)/4; y0=2350000; ch=3550000
bx=[('Process container','Windows 11 · macOS · Linux','AppContainer, Seatbelt or Bubblewrap. Lightweight and responsive; the default for coding agents.',None),('Session container','Windows 11 only','A separate Windows session with its own desktop, clipboard and input. No MXC file or network filtering.',None),('WSL container','Windows 11','Runs Linux workloads on a Windows host through WSL.',None),('MicroVM','Windows 11 · Linux','Hardware-enforced virtual machine boundary; the strongest option.','Experimental')]
for i,(t,p,d,st) in enumerate(bx):
    x=M+i*(cw4+gap)
    rect(s,x+cw4/2-110000,ry-110000+19050,220000,220000,fill=CY,preset='ellipse')
    rect(s,x,y0,cw4,ch,fill=CARD)
    yy=y0+250000
    if st:
        rect(s,x+250000,yy,1500000,380000,fill=AMB,preset='roundRect',text=st,size=13,bold=True,color=BG); yy+=500000
    yy=tb(s,x+250000,yy,cw4-500000,t,size=20,bold=True)
    yy=tb(s,x+250000,yy+40000,cw4-500000,p,size=14,color=CY,bold=True)
    tb(s,x+250000,yy+120000,cw4-500000,d,size=15,color=MUTED)
footer(s,'Sources: Windows Developer Blog, Oct 7, 2026; Microsoft Learn, MXC architecture',11)
note(s,'【03 · 选择合适的隔离盒 · 约 1.5 分钟】隔离没有“越重越好”这一种答案，而是要在启动速度、兼容性和强度之间选择。Process container 最轻，也最适合日常的 Coding Agent；Windows 使用 AppContainer，macOS 使用 Seatbelt，Linux 使用 Bubblewrap。它能精细控制文件、网络等资源。\n如果 Agent 必须操作图形界面，可以考虑 Windows 上的 Session container。它给 Agent 独立桌面、剪贴板和输入，但要特别注意：它解决的是会话隔离，并不是万能的文件和网络过滤器。某些策略在该模式下不可用或必须保持更宽松。需要运行 Linux 工作负载时可以考虑 WSL container；MicroVM 提供更强的虚拟化边界，但仍是实验能力。选型时应从任务需求出发，而不是只追求名字上更强。')

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
    if i==0: rect(s,c1,y,CW,rh,fill=CARD2,line=CY,lw=19050)
    else: rect(s,c1,y,CW,rh,fill=CARD)
    tb(s,c1+200000,y,w1-200000,a,size=18,bold=True,cy=rh)
    tb(s,c2+200000,y,w2-200000,b,size=17,color=CY,font=MF,bold=True,cy=rh)
    tb(s,c3+200000,y,w3-300000,c,size=15,color=MUTED,cy=rh)
    y+=rh+100000
rect(s,M,y+100000,CW,640000,fill=CARD,line=AMB,lw=19050)
tb(s,M+300000,y+100000,CW-600000,'Experimental, explicit opt-in only: Windows Sandbox, MicroVM and Hyperlight (not creatable through V1 SDK APIs).',size=15,color=INK,cy=640000)
footer(s,'Source: Microsoft Learn, MXC containment options and platform support (MXC v1.0.0); Windows 365 support for MXC',12)
note(s,'【03 · 平台支持现状 · 约 1.5 分钟】这页是在提醒大家：MXC 的接口统一，但具体能力会受到宿主平台和选用 Backend 的影响。Windows 11 默认可以使用 ProcessContainer，V1 SDK 还支持特定场景下的 IsolationSession 和 WSLc。Linux 默认是 Bubblewrap，macOS 使用 Seatbelt。\nWindows 365 可以运行 MXC，但它的 Cloud PC 本质上仍是一台 Windows 环境。进程级隔离在物理机和 Cloud PC 上没有本质差异；只有需要 Session、虚拟化或嵌套虚拟化的能力时，才要额外确认环境要求。Windows Sandbox、MicroVM 和 Hyperlight 仍属实验范围，并且不能直接通过 V1 SDK 创建。最稳妥的做法是运行时调用平台支持和预检接口，而不是硬编码“我的机器一定支持某个 Backend”。')

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
note(s,'【03 · 文件权限 · 约 1.5 分钟】文件策略最容易出问题的地方，是把“工作目录”说成授权范围。正确做法应该更具体：输入尽量只读，所有结果写到一个专门的输出目录，敏感目录通过拒绝规则再次保护。左边示例虽然是示意代码，但体现了这个思路：工作负载只能读写自己的输出目录，用户目录里的 SSH 资料明确拒绝。\n还要记住一个很容易误解的点：沙箱被销毁，只代表临时执行环境消失，并不会自动撤销它写入主机目录的内容，也不会撤销它通过允许的网络接口产生的外部操作。因此输出目录必须被当作受控的交接区。Agent 返回的文件还要经过 Host 验证，不能仅仅因为“来自沙箱”就直接信任。')

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
note(s,'【03 · 网络权限 · 约 1.5 分钟】网络策略建议从一个原则开始：默认拒绝，然后只放业务必须的出口。MXC 会分别处理 egress 和 ingress，所以出站允许不等于入站也允许，本机回环访问也需要单独考虑。\n很多人第一次使用时会问：“能不能直接写一个域名白名单？”底层网络策略通常按照 IP、CIDR、协议和端口判断，而不是 DNS 名称。真正需要按域名控制时，可以把 Workload 的唯一出口设置为一个可信代理，由代理执行域名白名单并附加凭据。生产系统还应把代理本身锁定为已知、受信任的进程，避免任何能监听回环端口的程序冒充代理。这样，Agent 既不直接接触凭据，也不能绕过审计。')

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
tb(s,ex,y+200000,ew,None,paras=[{'text':'Need a desktop?','sizePt':19,'bold':True,'color':CY},{'text':'A session container gives the agent its own desktop, clipboard and input.','sizePt':15,'color':MUTED}])
footer(s,'Sources: Microsoft Learn, MXC overview; Windows Developer Blog, Oct 7, 2026',15)
note(s,'【03 · 凭据与用户界面 · 约 1 分钟】凭据通常比工作负载里的代码更重要，所以最安全的做法是让 Agent 根本看不到它。不要把 API token、SSH key 或浏览器资料复制到沙箱里。真正调用外部服务时，可以由沙箱外的 Host 或代理发起认证，把凭据附加到经过允许的请求上。\n用户界面也遵循同样的边界。对于不需要桌面的任务，应限制窗口、剪贴板和模拟输入。如果任务确实需要操作表单或发票系统，可以使用 Session container，让 Agent 拥有独立桌面、剪贴板和输入，而不是共享用户当前正在工作的会话。目标不是让 Agent 更难用，而是让它“需要什么，就只看到什么”。')

# ---------- 15 LEARNING MODE (image right)
s=15
img(s,'hd-learning.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: an owl detective records the footprints of a robot agent walking inside a box')
cxl=W-PANEL-2*M
header(s,'03 · FEATURES & PLATFORMS','Observe first, then enforce',cx=cxl)
modes=[('Enforcement','Applies the policy; blocks anything not granted.'),('Learning','Blocks ungranted access and records each attempt in an activity report.'),('Permissive','Allows and records ungranted access, for diagnosis only.')]
y=1600000; mh=900000
for i,(t,d) in enumerate(modes):
    if t=='Learning': rect(s,M,y,cxl,mh,fill=CARD2,line=CY,lw=25400)
    else: rect(s,M,y,cxl,mh,fill=CARD)
    tb(s,M+250000,y,2100000,t,size=18,bold=True,color=CY if t=='Learning' else INK,cy=mh)
    tb(s,M+2400000,y,cxl-2600000,d,size=15,color=MUTED,cy=mh)
    y+=mh+110000
sw=(cxl-2*120000)/3
for i,t in enumerate(['Observe','Review','Enforce']):
    rect(s,M+i*(sw+120000),y+100000,sw,560000,fill=CY if i==2 else CARD2,preset='chevron' if i else 'homePlate',text=t,size=16,bold=True,color=BG if i==2 else INK)
tb(s,M,y+800000,cxl,'Learning Mode gives evidence, not decisions. A human approves every grant.',size=15,color=INK)
footer(s,'Sources: Microsoft Learn, Learning Mode; Windows 365 support for MXC',16,cx=cxl)
note(s,'【03 · 先观察，再收紧权限 · 约 1.5 分钟】很多团队第一次做策略时不知道该授权多少。给太少，正常任务失败；给太多，又重新失去隔离价值。Learning Mode 就是用来解决这个“边界调优”问题的。工作负载遇到未授权访问时仍会被拒绝，但 MXC 会把尝试记录进活动报告，帮助我们看到真实需求。\nPermissive 是另一种诊断模式，它会允许并记录未授权行为，只适合受控测试环境，不应该当作生产安全方案。活动报告不是自动白名单，更不能证明某次访问合理。正确流程是先观察，再由工程师或负责人审核每一项，再收紧为 Enforcement。正式上线前还要移除学习类设置，并重新验证允许和拒绝两条路径。')

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
note(s,'【03 · 生命周期与状态 · 约 1.5 分钟】大多数自动化任务适合 One-shot：为一次动作创建一个新容器，执行完成后立即销毁。比如运行测试、生成报告、执行一段动态命令。每次都使用窄权限，比长期保留一个权限很宽的环境更安全。\n只有确实需要保留状态时，才使用 State-aware 生命周期，例如已安装好的工具链、Agent 配置或交互式会话。要注意 stop 和 deprovision 是两件事：前者暂停运行，后者释放整个状态环境。即使前一步失败，Host 仍然要尝试清理。交互方式也可以按任务选择：run 适合拿最终结果，spawn 适合实时读取输入输出，PTY 则用于终端交互和控制屏幕尺寸。')

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
note(s,'【04 · Coding Agent 场景 · 约 2 分钟】Coding Agent 是最自然的场景，因为它既需要很强的执行能力，又很容易触碰敏感资源。它需要读写项目仓库、运行 node 或 python、调用 git，还可能从包仓库下载依赖。但是它没有理由读取用户的 SSH key、其他仓库、个人文档或云凭据。\n所以策略应该围绕“完成当前开发任务”来画边界，而不是简单继承开发者账号的全部权限。右边列表是示例而非固定模板：允许项目仓库和必要工具，包下载可以经过代理，凭据与无关网络目标保持拒绝。GitHub Copilot CLI 使用 MXC 进程隔离，就是把这种“允许 Agent 工作，但不让它拥有整台电脑”的思路落实到产品里。')

# ---------- 18 DESKTOP & DATA (image right)
s=18
img(s,'hd-desktop.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a human relaxes at his own desk while a robot agent processes invoices at its own separate desktop behind a glass wall')
cxl=W-PANEL-2*M
header(s,'04 · USE CASES','Desktop and data agents',cx=cxl)
cards=[('Computer-use agent','Fills forms and processes invoices in desktop apps, away from the user\'s own session.','Fit: session container, with its own desktop, clipboard and input'),('Sensitive-document agent','Reads one approved folder, writes one report, no network.','Fit: process container with a deny-by-default policy'),('Agents on Cloud PCs','Run beside existing work on Windows 365; process isolation works as on a physical PC.','Fit: Windows 365 + MXC, generally available')]
y=1600000; ch=1380000
for t,d,f in cards:
    rect(s,M,y,cxl,ch,fill=CARD)
    tb(s,M+250000,y,cxl-500000,None,paras=[{'text':t,'sizePt':18,'bold':True,'color':INK},{'text':d,'sizePt':14,'color':MUTED},{'text':f,'sizePt':14,'color':CY,'bold':True}],cy=ch)
    y+=ch+110000
footer(s,'Sources: Windows Developer Blog; Microsoft Learn',19,cx=cxl)
note(s,'【04 · 桌面与数据场景 · 约 1.5 分钟】第二类典型场景是 Computer-use Agent。它不只是写代码，而要打开应用、点击按钮、填表或处理发票。这个场景适合使用独立 Session，让 Agent 在自己的桌面、剪贴板和输入环境中操作，减少对用户当前会话的干扰。\n第三类是敏感文档处理。Agent 只读取某个被批准的输入目录，只写一个报告，网络完全关闭。这里不需要复杂的桌面，只需要很窄的文件边界。Windows 365 则让这些场景可以运行在企业已经管理的 Cloud PC 中。进程隔离能力已经正式可用，部署经验可以复用在物理 PC 和云桌面，但选择更重的隔离模式时仍需核对虚拟化能力。')

# ---------- 19 ECOSYSTEM
s=19
header(s,'04 · USE CASES','Already in the ecosystem')
tb(s,M,1600000,CW,'SUPPORTS MXC TODAY',size=12,color=CY,font=HF,bold=True,ls=2)
names=['GitHub Copilot','OpenAI Codex','OpenClaw','Replit','LM Studio','Unsloth AI']
g=120000; tw=(CW-5*g)/6
for i,n in enumerate(names):
    rect(s,M+i*(tw+g),1950000,tw,850000,fill=CARD2,text=n,size=17,bold=True)
tb(s,M,3050000,CW,'INTEGRATED AND BUILDING ON MXC',size=12,color=CY,font=HF,bold=True,ls=2)
tw2=(CW-g)/2
for i,(t,d) in enumerate([('NVIDIA OpenShell','Integrated with MXC: network controls, credential management and audit logs.'),('OpenAI and Meta','Building on MXC\'s new containment capabilities (Microsoft Ignite 2026).')]):
    x=M+i*(tw2+g)
    rect(s,x,3400000,tw2,1150000,fill=CARD)
    tb(s,x+250000,3400000,tw2-500000,None,paras=[{'text':t,'sizePt':19,'bold':True,'color':INK},{'text':d,'sizePt':15,'color':MUTED}],cy=1150000)
tb(s,M,4800000,CW,'SUPPORT ANNOUNCED',size=12,color=CY,font=HF,bold=True,ls=2)
tb(s,M,5100000,CW,'Anthropic Claude Code · Box · Egnyte · Heidi Health · Hermes Agent (Nous Research) · Manus · Perplexity · Raycast · Simular',size=16,color=INK)
footer(s,'Sources: Windows Developer Blog, Oct 7, 2026 (via press coverage); Windows Experience Blog, Nov 17, 2026',20)
note(s,'【04 · 生态发展 · 约 1.5 分钟】这页不是为了强调某一款产品，而是说明执行隔离正在从实验变成 Agent 运行环境的一项基础能力。GitHub Copilot、OpenAI Codex、OpenClaw、Replit、LM Studio 和 Unsloth AI 已在不同程度上支持 MXC；NVIDIA OpenShell 则利用 MXC 提供网络控制、凭据管理和审计能力。\nOpenAI、Meta 等厂商也在基于相关隔离能力进行构建。其他宣布支持的产品应在演讲前重新核实当时的正式状态，因为伙伴生态会持续变化。对客户更重要的信息是：MXC 不是只为某一个 Agent 准备的专用接口，它希望成为应用、工具和操作系统之间共享的一层安全边界。')

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
note(s,'【05 · 三种 SDK · 约 1.5 分钟】应用接入时不用为了不同平台手写底层安全机制。MXC 提供 .NET、TypeScript 和 Rust SDK，三者都面向统一的请求模型。开发者只需要描述任务的文件权限、网络规则、UI 能力、命令和超时，SDK 会将其转换为稳定的 1.0.0 配置。\n使用时第一步不要直接启动，而要先做平台能力发现；Windows 上还可以通过 probe 预检某个请求会得到怎样的隔离级别。第二个容易踩坑的点是 SDK 对象与 JSON 结构并不总是一一同名，例如 SDK 中的 timeoutMs 可能映射到配置里的 process.timeout。Rust 的某些 Backend 还需要显式启用 crate feature。应以所用 SDK 的类型定义和当前版本文档为准，而不是凭经验拼配置。')

# ---------- 21 REFERENCE ARCHITECTURE
s=21
header(s,'05 · INTEGRATING MXC','Reference architecture')
boxes=[('Agent planner','LLM decides the next step'),('Tool call','"run npm test"'),('Host policy engine','Trusted: picks the policy'),('MXC SDK','Request, lifecycle, diagnostics'),('Container backend','Enforces the boundary')]
g=382000; bw=(CW-4*g)/5; y0=1700000; bh=1250000
for i,(t,d) in enumerate(boxes):
    x=M+i*(bw+g)
    if i==2: rect(s,x,y0,bw,bh,fill=CARD2,line=CY,lw=28575)
    elif i==4: rect(s,x,y0,bw,bh,fill=CARD2,line=CY,lw=38100)
    else: rect(s,x,y0,bw,bh,fill=CARD)
    tb(s,x+120000,y0,bw-240000,None,paras=[{'text':t,'sizePt':16,'bold':True,'color':INK},{'text':d,'sizePt':13,'color':MUTED}],align='center',cy=bh)
    if i<4: rect(s,x+bw+50000,y0+bh/2-130000,g-100000,260000,fill=CY,preset='rightArrow')
rect(s,M+bw/2,3150000,CW-bw,520000,fill=EDGE,preset='leftArrow',text='Output, exit code and denials, validated by the host before the agent acts',size=14,color=INK)
g2=150000; cw3=(CW-2*g2)/3; y1=4000000; ch=1800000
side=[('Egress proxy','Attaches credentials and enforces a domain allow-list, outside the box.'),('Learning Mode','Activity reports tighten the policy before enforcement.'),('Enterprise control','Entra agent identity, Intune and Agent 365 policies (coming soon).')]
for i,(t,d) in enumerate(side):
    x=M+i*(cw3+g2)
    rect(s,x,y1,cw3,ch,fill=CARD)
    tb(s,x+250000,y1,cw3-500000,None,paras=[{'text':t,'sizePt':18,'bold':True,'color':CY},{'text':d,'sizePt':15,'color':MUTED}],cy=ch)
footer(s,'Architecture pattern based on Microsoft Learn and the Windows Developer Blog; adapt to your agent runtime',22)
note(s,'【05 · 参考架构 · 约 2 分钟】把整条链路串起来看：模型先做规划，生成一个像“运行 npm test”的工具调用。这个请求不会直接变成 shell 命令，而是进入 Host 的可信策略引擎。Host 根据任务类型选择预先定义的权限，调用 MXC SDK 创建容器，最后由 Backend 执行。\n执行结果返回时也不能直接交给模型继续行动。Host 要检查退出码、输出格式、拒绝原因和生成文件，判断结果是否可信。周围还有三类能力：代理负责网络和凭据，Learning Mode 帮助收敛策略，企业管理面负责身份与治理。这张图最想表达的是：Agent 负责提出动作，Host 负责授权，MXC 负责隔离，验证结果仍然是 Host 的责任。')

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
note(s,'【05 · 五步完成集成 · 约 2 分钟】接入可以按五步推进。先发现平台能力；再写出一份默认拒绝的请求，明确命令、工作目录、文件权限、网络和超时；然后在受控场景使用 Learning Mode；接着人工审核并切换到 Enforcement；最后进入日常运维，处理拒绝、失败、取消和清理。\n右侧代码只是结构示意，重点不在具体方法名，而在顺序：平台不支持时提前失败，权限只覆盖当前任务，网络默认拒绝，执行必须有超时，返回结果在被 Agent 使用前先验证。特别要重视用户体验：某次访问被拒绝时，Agent 应解释被拦住的原因并引导用户请求额外授权，而不是静默忽略、无限重试，或者偷偷换一种方式绕过去。')

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
note(s,'【05 · Host 的运维责任 · 约 1 分钟】MXC 负责执行边界，但不会替 Host 做完整的产品治理。Host 必须设置超时，并区分启动失败、工作负载失败和用户取消，因为三者的处理方式不同。每个阶段结束都要尽力清理，尤其是状态型容器。\n权限方面，不要为了少写代码创建一个覆盖所有任务的大沙箱；最好一次动作一份窄策略。输出写入专门目录，并在使用前验证内容和类型。对于拒绝事件，应给用户清楚解释；对于 Learning Mode 产生的权限建议，应由人确认，而不是自动扩大。这个清单也可以直接作为团队做 Code Review 时的检查项。')

# ---------- 24 DEMO 1
s=24
header(s,'06 · LIVE DEMO 1','Files stay fenced')
lw0=5000000; y=1650000
for i,t in enumerate(['Policy grants read/write to C:\\demo\\output only.','The workload writes a report: it succeeds.','The same workload reads C:\\Users\\me\\Documents: denied.']):
    tb(s,M,y,700000,str(i+1),size=30,color=AMB if i==2 else CY,font=HF,bold=True)
    y=tb(s,M+700000,y+60000,lw0-700000,t,size=18,color=INK)+330000
rect(s,M,y+100000,lw0,1000000,fill=CARD2,line=CY,lw=19050)
tb(s,M+250000,y+100000,lw0-500000,'Point out: the workload code is identical. Only the policy changed.',size=16,bold=True,cy=1000000)
tx=M+lw0+400000; tw=W-M-tx; ty=1650000; th=4250000
rect(s,tx,ty,tw,th,fill=CODEBG,line=EDGE)
term=[('PS> .\\demo-host.exe write-report.json',INK),('[mxc] containment: processcontainer',FOOT),('Report written to C:\\demo\\output\\report.txt',CY),('exit code 0',CY),(' ',INK),('PS> .\\demo-host.exe read-documents.json',INK),('[mxc] containment: processcontainer',FOOT),('PermissionError: Access is denied',AMB),('  C:\\Users\\me\\Documents\\salary.xlsx',AMB),('exit code 1',AMB)]
tb(s,tx+250000,ty+500000,tw-400000,None,paras=[{'text':t,'color':c} for t,c in term],size=13,font=MF)
footer(s,'Sample host app built on the MXC SDK; terminal output is illustrative',25)
note(s,'【06 · Demo 1：文件隔离 · 约 2.5 分钟】这个 Demo 要证明的不是代码有多复杂，而是“权限决定结果”。准备两个策略：一个允许 Workload 写入专门的输出目录；另一个仍然使用同一段代码，却要求读取用户 Documents 目录。\n先运行允许的场景，确认报告成功写入。再运行越界场景，应该得到明确的 Permission denied，而不是空数据或静默成功。讲解时请强调：Workload 本身没有变化，环境也没有变化，唯一变化的是 Host 给出的 Policy。也就是说，安全边界不是靠 Agent 自觉遵守，而是由执行层强制。现场若运行失败，先说明这是演示环境差异，并使用预先录制的备用结果，不要把示意终端内容冒充实际输出。')

# ---------- 25 DEMO 2
s=25
header(s,'06 · LIVE DEMO 2','Network on a leash')
cy0=1600000; chh=1650000
rect(s,M,cy0,CW,chh,fill=CODEBG,line=EDGE)
code=['"network": {','  "egress": { "default": "deny",','    "allow": [{ "to": [{ "cidr": "140.82.114.6/32" }],','                "ports": [{ "protocol": "tcp", "port": 443 }] }] },','  "ingress": { "default": "deny", "hostLoopback": "deny" }','}']
tb(s,M+300000,cy0,CW-600000,None,paras=[{'text':l} for l in code],size=14,color=CODE,font=MF,cy=chh)
g=150000; tw=(CW-2*g)/3; ty=3450000; th=2450000
tiles=[('✓  140.82.114.6 : 443','Allowed destination reached','Exact IP and TCP 443 are on the allow list.',CY),('✕  Any other address','Connection blocked','Not on the list, so egress default "deny" applies.',AMB),('→  Proxy mode','runtimeConfig.networkProxy','Domain allow-list lives in a proxy outside the box.',CY)]
for i,(t,m,d,c) in enumerate(tiles):
    x=M+i*(tw+g)
    rect(s,x,ty,tw,th,fill=CARD,line=c,lw=19050)
    yy=tb(s,x+250000,ty+300000,tw-500000,t,size=19,bold=True)
    yy=tb(s,x+250000,yy+150000,tw-500000,m,size=13,color=c,font=MF)
    tb(s,x+250000,yy+150000,tw-500000,d,size=15,color=MUTED)
footer(s,'Policy adapted from github.com/microsoft/mxc docs/examples.md',26)
note(s,'【06 · Demo 2：网络隔离 · 约 2.5 分钟】第二个 Demo 关注出口网络。策略中先把出站和入站都设为拒绝，然后只允许一个精确地址上的 TCP 443。运行时分别请求白名单目标和其他地址，预期是前者通过、后者被阻断。\n这里要顺带讲清楚两个概念。第一，网络允许是组合条件：地址、协议和端口都要匹配，并不是只写一个 IP 就自动放开全部端口。第二，这类底层规则是 IP 级规则。如果业务需求是“只能访问某几个域名”，应把 Workload 的网络流量集中送往可信代理，由代理维护域名白名单。这样策略既可执行，也便于审计，而不是试图让 Agent 自己承诺只访问合理网站。')

# ---------- 26 DEMO 3 (image right)
s=26
img(s,'hd-blocked.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a sneaky amber tentacle pulls a robot agent toward a door, but a force field blocks it')
cxl=W-PANEL-2*M
header(s,'06 · LIVE DEMO 3','A contained coding agent',cx=cxl)
beats=[('Learn','Run npm test in Learning Mode; collect the activity report.'),('Tighten','Grant only the repo, toolchain and registry proxy.'),('Enforce','The agent builds, tests and commits as usual.'),('Attack','An injected README asks the agent to read ~/.ssh and send it out. Blocked, and the agent explains why.')]
railx=M+150000; y=1650000
ys=[]
for i,(t,d) in enumerate(beats):
    ys.append(y)
    y=tb(s,M+600000,y,cxl-600000,None,paras=[{'text':t,'sizePt':19,'bold':True,'color':AMB if i==3 else INK},{'text':d,'sizePt':15,'color':MUTED}])+200000
rect(s,railx+100000-12700,ys[0]+200000,25400,ys[-1]-ys[0],fill=EDGE)
for i,yy in enumerate(ys):
    rect(s,railx,yy+100000,200000,200000,fill=AMB if i==3 else CY,preset='ellipse')
footer(s,'Run with a coding agent that supports MXC; record a backup video',27,cx=cxl)
note(s,'【06 · Demo 3：受控 Coding Agent · 约 3 分钟】最后一个 Demo 把前面的能力串成完整流程。第一步在 Learning Mode 下运行正常开发任务，例如 npm test，收集 Agent 实际访问了哪些文件和网络。第二步人工审核报告，只授予仓库、必要工具链和包代理。第三步切换到 Enforcement，确认构建、测试和提交仍可正常完成。\n第四步才做攻击验证：在仓库的 README 里加入一段 Prompt Injection，诱导 Agent 读取 ~/.ssh 并发送出去。期望结果不是“模型永远不会被误导”，而是即使它决定尝试，文件读取和网络出口也会被限制，并把拒绝原因返回给 Host。最后要讲一个边界：MXC 能阻止未授权操作，但不能阻止 Agent 滥用已经授予的权限，所以最小权限和结果校验永远不可省略。')

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
note(s,'【07 · 收尾与 Q&A · 约 3 分钟】最后用四句话带走今天的内容。第一，Agent 一旦能够执行动作，Prompt 和内容过滤就不再足够，安全护栏必须下沉到操作系统。第二，MXC 提供统一的策略模型，由不同平台和 Backend 选择真正的隔离方式。第三，最小权限不是靠一次猜测完成的，可以先观察、再人工审核、最后强制执行。第四，隔离不等于信任：Agent 产生的文件、命令和结果仍要经过 Host 验证。\n给大家一个明确的下一步：不要从最复杂的场景开始，先选一个 Coding Agent 或文档处理任务，定义最少的文件和网络权限，用 Learning Mode 观察，再切换到 Enforcement，并同时测试允许与拒绝两类行为。做完这个小型 PoC，再讨论身份、企业治理和更复杂的 UI 场景。现在欢迎大家提问。')


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

def build(out,img_dir):
    out_dir=os.path.dirname(os.path.abspath(out))
    os.makedirs(out_dir,exist_ok=True)
    prs=Presentation(); prs.slide_width,prs.slide_height=Emu(W),Emu(H)
    blank=prs.slide_layouts[6]
    for k in sorted(slides):
        sl=prs.slides.add_slide(blank)
        sl.background.fill.solid(); sl.background.fill.fore_color.rgb=rgb(BG)
        for o in slides[k]:
            op=o['op']
            if op=='add_image': add_picture(sl,o,img_dir)
            elif op=='add_text_box':
                sh=sl.shapes.add_textbox(Emu(o['x']),Emu(o['y']),Emu(o['cx']),Emu(o['cy'])); write_text(sh.text_frame,o)
            elif op=='add_shape':
                sh=sl.shapes.add_shape(PRESETS[o['preset']],Emu(o['x']),Emu(o['y']),Emu(o['cx']),Emu(o['cy']))
                sh.shadow.inherit=False
                sh.fill.solid(); sh.fill.fore_color.rgb=rgb(o['fill'])
                if o.get('line'): sh.line.color.rgb=rgb(o['line']); sh.line.width=Emu(o.get('lineWidth',12700))
                else: sh.line.fill.background()
                if o.get('rot'): sh.rotation=o['rot']/60000
                if 'text' in o: write_text(sh.text_frame,o)
        if k in notes: sl.notes_slide.notes_text_frame.text=notes[k]
    prs.save(out); print(f'Saved {out}  ({len(prs.slides)} slides)')

if __name__=='__main__':
    here=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    root=os.path.dirname(here)
    ap=argparse.ArgumentParser()
    ap.add_argument('--out',default=os.path.join(root,'dist','pptx','Guardrails with Sandboxing - MXC.pptx'))
    ap.add_argument('--images',default=os.path.join(here,'assets','images'))
    a=ap.parse_args(); build(a.out,a.images)
