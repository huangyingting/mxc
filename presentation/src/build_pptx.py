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
from pptx.oxml.ns import nsdecls
from PIL import Image

# ===================== THEME & LAYOUT =====================
W,H=12192000,6858000; M=457200; CW=W-2*M
BG='F3EEE4'; CARD='E9E2D4'; CARD2='E9E2D4'; EDGE='C9C0AE'; INK='2B2B2B'; MUTED='57534C'; FOOT='6E685F'
CY='3F607E'; AMB='A84A35'; CODEBG='FAF7F1'; CODE='2F4B66'
HF='Plus Jakarta Sans'; BF='Plus Jakarta Sans'; MF='JetBrains Mono'
PANEL=4572000   # width of the side illustration panel
# ===================== HELPERS ===================
slides={}  # idx -> list of patches
notes={}
def P(s,op): slides.setdefault(s,[]).append(op)
WF={'Plus Jakarta Sans':0.55,'Newsreader':0.58,'JetBrains Mono':0.62,'Aptos Display':0.56,'Aptos':0.53,'Cascadia Code':0.62,'Georgia':0.6,'Arial':0.56,'Calibri':0.52,'Courier New':0.62}
def nlines(text,size,cx,font,bold):
    f=WF[font]*(1.08 if bold else 1)
    cpl=max(1,int((cx-240000)/(size*12700*f)))
    n=0
    for para in text.split('\n'):
        words=para.split(' '); line=0; cnt=1
        for w_ in words:
            L=len(w_)+(1 if line else 0)
            if line+L>cpl and line>0: cnt+=1; line=len(w_)
            else: line+=L
        n+=cnt
    return n
def hgt(paras,cx,font,space_after=6.5):
    t=0
    for idx,p in enumerate(paras):
        lh=1.48 if font==MF else (1.16 if p['sizePt']>=29 else 1.42)
        t+=nlines(p['text'],p['sizePt'],cx,font,p.get('bold',False))*p['sizePt']*12700*lh
        if idx<len(paras)-1:
            t+=p.get('spaceAfterPt',space_after)*12700
    return int(t+240000)
def tb(s,x,y,cx,text=None,size=16,color=INK,font=BF,bold=False,align='left',paras=None,ls=None,cy=None,name=None,space_after=None):
    default_sa = (2.0 if font==MF else 6.5) if space_after is None else space_after
    if paras is None:
        paras=[{'text':text,'sizePt':size,'bold':bold,'color':color}]
    else:
        paras=[dict({'sizePt':size,'bold':bold,'color':color},**p) for p in paras]
    h=cy or hgt(paras,cx,font,default_sa)
    op={'op':'add_text_box','slide':s,'x':int(x),'y':int(y),'cx':int(cx),'cy':int(h),'fontFamily':font,'align':align,'color':color,'spaceAfterPt':default_sa}
    if len(paras)==1 and ls is None and 'href' not in paras[0]:
        p=paras[0]; op.update(text=p['text'],sizePt=p['sizePt'],bold=p['bold'],color=p['color'])
    else:
        op['paragraphs']=[{k:v for k,v in p.items()} for p in paras]
        if ls is not None:
            if len(paras)==1 and 'href' not in paras[0]:
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
def img(s,path,x,y,cx,cy,alt,fit='cover'):
    P(s,{'op':'add_image','slide':s,'path':path,'x':int(x),'y':int(y),'cx':int(cx),'cy':int(cy),'fit':fit,'altText':alt})
def header(s,eyebrow,title,x=M,cx=CW,title_cy=None):
    tb(s,x,310000,cx,eyebrow,size=12,color=CY,font=HF,bold=True,ls=2,cy=280000,name='eyebrow')
    tb(s,x,640000,cx,title,size=29.5,color=INK,font=HF,bold=True,cy=title_cy or 640000,name='slide_title')
def footer(s,src,num,x=M,cx=CW):
    if src: tb(s,x,6260000,cx-900000,src,size=11.5,color=FOOT,cy=340000,name='footer_src')
    tb(s,x+cx-800000,6260000,800000,f'{num:02d}',size=11.5,color=FOOT,align='right',cy=340000,name='footer_num')
def note(s,t): notes[s]=t

def card_box(s,x,y,cx,cy,paras,fill=CARD,line=None,lw=19050,pad_x=240000,pad_y=120000,font=BF,space_after=6.5,align='left'):
    rect(s,x,y,cx,cy,fill=fill,line=line,lw=lw if line else None)
    tb(s,x+pad_x,y+pad_y,cx-2*pad_x,None,paras=paras,font=font,align=align,cy=cy-2*pad_y,space_after=space_after)

# ===================== SLIDES ====================
# ---------- 0 COVER
s=0
img(s,'hd-cover.jpg',0,0,W,H,'Hand-drawn sketch illustration: a friendly robot AI agent standing inside a glass cube, with files, cloud and keys kept outside the boundary')
x=M+150000; cx=5900000
tb(s,x,1250000,cx,'MICROSOFT EXECUTION CONTAINERS (MXC)',size=12,color=CY,font=HF,bold=True,ls=2,cy=300000)
tb(s,x,1650000,cx,'Guardrails with\nSandboxing',size=45,color=INK,font=HF,bold=True,cy=1480000)
tb(s,x,3280000,cx,'Keeping AI agents inside safe boundaries: isolating execution and controlling access to files, network and credentials on Windows',size=17,color=MUTED,cy=1300000)
tb(s,x,4850000,cx,'Yingting Huang  |  Sr. Cloud Solution Architect, Microsoft',size=15,color=INK,bold=True,cy=500000)
tb(s,x,5400000,cx,'November 2026  ·  45 minutes  ·  L300 → L400',size=13.5,color=MUTED,cy=340000)
note(s,'大家好，欢迎来到今天的分享。这两年大家都在做 Agent，你会发现它的角色已经变了——以前大模型只是陪你聊天、写段代码让你自己复制粘贴；现在的 Agent 是真正在你的电脑或服务器上干活，自己读文件、跑命令、调接口。\n能力变强之后，安全问题也跟着变了：万一它理解错了意图，或者读到了别人埋了坑的文档，怎么保证它不会把系统搞乱、把密钥传出去？今天我们就来聊微软开源的沙箱执行层 MXC，看看怎么在放手让 Agent 干活的同时，在操作系统这一层把安全边界牢牢守住。')

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
note(s,'今天这 45 分钟，我们顺着实际工程落地的脉络往下讲。先看看当 Agent 开始动手执行命令之后，到底会带来哪些新风险，为什么光靠写提示词根本防不住；接着拆解 MXC 是什么，以及它怎么把一份权限策略变成操作系统底层的硬隔离；然后聊大家最关心的文件、网络、密钥和四种不同轻重的隔离容器该怎么选；最后结合写代码、桌面操作和敏感数据处理，看看怎么用几行 SDK 代码把它接进你自己的系统，并现场跑三个真机 Demo。')

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
note(s,'先看最本质的变化。过去我们用聊天助手，它只负责生成文字，中间永远隔着一个人——命令安不安全，你扫一眼再决定跑不跑，最坏也就是答非所问。但自主 Agent 把中间这道人工把关拿掉了，它会自己翻代码库、装依赖、跑脚本、改配置。\n更棘手的是，它下一步要敲什么命令，往往是运行时根据上下文临时写出来的，你根本没法提前把所有命令都审一遍。如果让它直接顶着你本人的账号权限裸跑，那你能打开的所有文件、保存的所有凭证、能连上的所有内网服务，其实都暴露在它的一次误判或者一次外部注入之下。')

# ---------- 3 THREATS (image right)
s=3
img(s,'hd-chaos.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a hooded villain slips a poisoned note to a robot agent while files fly out and a drone steals a key')
cxl=W-PANEL-2*M
header(s,'01 · WHY SANDBOXING','What can go wrong?',cx=cxl)
rows=[('Prompt injection','A poisoned README tells the agent to run the attacker\'s command.'),('Credential theft','Generated code reads tokens, SSH keys and browser secrets.'),('Data exfiltration','Files quietly leave for an unknown host.'),('Overreach','A "reasonable" shortcut edits production config it was never meant to touch.')]
y=1650000
for i,(h_,d) in enumerate(rows):
    tb(s,M,y,700000,f'0{i+1}',size=24,color=AMB,font=HF,bold=True)
    yy=tb(s,M+800000,y,cxl-800000,None,paras=[{'text':h_,'sizePt':18,'bold':True,'color':INK},{'text':d,'sizePt':14,'color':MUTED}],cy=960000)
    y=yy+130000
footer(s,'Scenario 04 adapted from the Windows Developer Blog, Oct 7, 2026',4,cx=cxl)
note(s,'具体会出什么乱子呢？最典型的就是这四种情况。比如 Agent 去读一个开源仓库或者网页，里面藏了一句恶意指令，哄着它顺手跑一段外部脚本，这就是提示词注入。再比如它写出来的代码去翻你本机的 SSH 私钥、云平台 Token 或者浏览器里存的登录态，然后悄悄发往外部地址，几秒钟就完成了凭证窃取和数据外传。\n还有一种情况连攻击者都不需要，纯粹是模型太想帮你把事办成了——比如你让它修个网页报错，它觉得直接改线上生产配置最快，结果好心办了坏事。不管是被人下套，还是模型自己抄了近道，只要底下没有硬边界拦着，后果是一样的。')

# ---------- 4 PROMPTS PERSUADE
s=4
header(s,'01 · WHY SANDBOXING','Prompts persuade. The OS enforces.')
y=tb(s,M,1650000,9500000,'"An agent cannot be its own security authority."',size=32,color=INK,font=HF,bold=True)
y=tb(s,M,y+20000,9500000,'Windows Developer Blog, October 7, 2026',size=14,color=MUTED)
lab=1900000; chw=2240000; gap=120000
def chiprow(yv,label,chips,lifted):
    tb(s,M,yv+70000,lab-100000,label,size=15,bold=True,color=CY if lifted else MUTED)
    for i,c in enumerate(chips):
        x=M+lab+i*(chw+gap)
        if lifted: rect(s,x,yv,chw,560000,fill=CARD2,line=CY,lw=19050,text=c,size=15,color=INK)
        else: rect(s,x,yv,chw,560000,fill=CARD,text=c,size=15,color=MUTED)
chiprow(3500000,'AI-level',['System prompts','Content filters','Tool rules','Output checks'],False)
chiprow(4250000,'Execution-level',['File allow / deny','Network rules','Process isolation','Secrets outside'],True)
tb(s,M,5150000,CW,'You need both: the model decides what to try; the operating system decides what is possible.',size=17,color=INK)
footer(s,'',5)
note(s,'所以做 Agent 架构一定要想明白一句话：不能让 Agent 自己当自己的安全管理员。你在系统提示词里写十遍“绝对不准碰密钥、不准连外网”，再叠加上下文过滤和工具调用规则，这些当然要做，但它们本质上都是在劝模型守规矩，只能降低出错的概率。\n一旦模型被绕晕了，或者它调起的脚本行为超出了预期，提示词是拦不住系统调用的。真正能兜底的只有操作系统：哪个目录能读、哪个端口能通、能不能碰剪贴板，必须由系统内核强制拍板。模型决定它想尝试什么，操作系统决定它实际上能做成什么，这两层缺一不可。')

# ---------- 5 THREE PILLARS
s=5
header(s,'01 · WHY SANDBOXING','Three pillars for agents on Windows')
cw3=(CW-2*238800)/3; y0=1650000; ch=3900000
pill=[('Containment','Limit what an agent can access and do','Isolates file, network, process and UI access at the OS layer via MXC — on physical PCs and Windows 365 Cloud PCs.','Generally available'),('Identity','Tell agent activity apart from a person','Microsoft Entra Agent ID gives each agent its own identity and audit trail instead of sharing a human login.','GA + preview'),('Manageability','Govern access and monitor agents','Microsoft Agent 365 and Intune discover agents, monitor activity, and roll out enterprise sandbox policies.','GA + preview')]
for i,(t,sub,body,state) in enumerate(pill):
    x=M+i*(cw3+238800)
    rect(s,x,y0,cw3,ch,fill=CARD)
    pw=2100000 if t=='Containment' else 1500000
    rect(s,x+300000,y0+300000,pw,400000,fill=EDGE,preset='roundRect',text=state,size=13,color=INK,bold=True)
    y=tb(s,x+300000,y0+900000,cw3-600000,t,size=26,bold=True,font=HF)
    y=tb(s,x+300000,y+60000,cw3-600000,sub,size=17,color=CY)
    tb(s,x+300000,y+150000,cw3-600000,body,size=15,color=MUTED)
footer(s,'Sources: Microsoft Ignite 2026 · Agent 365 · Entra Agent ID · Microsoft Intune',6)
note(s,'放到企业级 Windows 环境里，管好 Agent 靠的是三块能力配合：用 MXC 做执行隔离，限制它能碰哪些本地资源；用 Entra Agent ID 给 Agent 发独立身份，别让它借用员工账号，否则真出了事看审计日志都分不清是人干的还是机器干的；再用 Agent 365 和 Intune 做统一的发现、监控和策略下发。\n大家排落地计划的时候可以留意一下进度：MXC 本身、Windows 365 云电脑上的 MXC，还有云端的 Entra Agent ID 和 Agent 365 都已经正式发布了；而给本机跑的 GitHub Copilot CLI、Claude Code 自动发身份，以及通过 Intune 直接给本地 MXC 容器下发文件网络规则，目前还在预览阶段。')

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
    tb(s,x+250000,yy,tw-400000,None,paras=[{'text':v,'sizePt':30,'bold':True,'color':CY},{'text':l,'sizePt':13,'color':INK}],font=BF,cy=th)
tb(s,xl,ty+2*th+150000+250000,cxr,'One JSON policy, mapped to the right OS sandbox on every host.',size=17,bold=True,color=INK)
footer(s,'Sources: Microsoft Learn; Windows Developer Blog, Oct 2026',7,x=xl,cx=cxr)
note(s,'那 MXC 到底是什么？简单讲，它是微软用 MIT 协议开源的一套轻量级沙箱执行层，今年十月已经正式发布。不管你要跑的是大模型现写的代码、第三方插件，还是来路不明的脚本，宿主程序只需要用同一份 JSON 策略说清楚允许读哪儿、写哪儿、连哪儿，MXC 就会自动调用当前操作系统原生的隔离机制去执行。\n注意它虽然名字叫 Execution Containers，但不是让你去装一套笨重的 Docker 镜像——它直接复用 Windows、macOS 和 Linux 内核自带的轻量沙箱能力，官方同时维护了 .NET、TypeScript 和 Rust 三个语言的 SDK。')

# ---------- 7 FIVE CONCEPTS
s=7
header(s,'02 · UNDERSTANDING MXC','Five concepts to remember')
gap=150000; cw5=(CW-4*gap)/5; y0=1650000; ch=3900000
cons=[('Host app','Your trusted app or agent runtime. Picks permissions, supplies inputs, handles results.'),('Workload','The command, script or process that runs under containment.'),('Policy','The access you permit: readable files, writable folders, network destinations, UI.'),('Container request','Combines the policy with the command, working directory, timeout and sandbox choice.'),('Backend','The OS mechanism that enforces it: process sandbox, separate session or VM.')]
for i,(t,d) in enumerate(cons):
    x=M+i*(cw5+gap)
    rect(s,x,y0,cw5,ch,fill=CARD)
    tb(s,x+200000,y0+180000,cw5-400000,f'0{i+1}',size=26,color=CY,font=HF,bold=True,cy=480000)
    tb(s,x+200000,y0+700000,cw5-400000,t,size=18,bold=True,cy=850000)
    tb(s,x+200000,y0+1590000,cw5-400000,d,size=14,color=MUTED,cy=2160000)
tb(s,M,5660000,CW,'The host decides. The policy describes. The backend enforces.',size=19,bold=True,cy=500000)
footer(s,'Source: Microsoft Learn, Microsoft Execution Containers overview',8)
note(s,'看 MXC 的接口之前，把这五个概念对上号就一目了然了。最外层是你自己写的宿主程序 Host，它是唯一可信的调度方，负责定权限、喂输入、收结果；真正被关进沙箱里跑的那条命令或脚本叫 Workload；允许读哪些文件、连哪些地址，写下来就是 Policy；把 Policy 跟具体的命令行、工作目录、超时和隔离方式打包到一块儿，就是一次完整的执行配置；最后在底层真正把隔离墙立起来的系统机制，就是 Backend。\n整套设计里最核心的规矩就是底下这句话：权限只能由 Host 或真人来拍板。Agent 可以提申请说它想读哪个目录，但绝不能让沙箱里的代码自己给自己扩权。')

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
note(s,'发起一次受控执行的过程非常顺畅。Host 先问一下 SDK 当前这台机器支持哪些沙箱后端，然后按默认拒绝的原则拼好策略和请求交给 SDK；MXC 在中间把这份统一请求翻译成对应操作系统的底层调用，等任务跑完，再把标准输出、退出码、拦截记录收集回来，顺手把临时环境清理干净。\n这里有个特别实在的工程经验：不同平台底层沙箱的脾气不完全一样，大家写测试的时候千万别只测正常路径能不能跑通，一定要专门写反向测试，去验证不该读的敏感文件是不是真被拦住了、没放行的网络是不是真断了。只有允许和拒绝两边都验过，护栏才算立稳了。')

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
tb(s,M,y+120000,CW,'Rule #1: an agent may suggest a permission, but only trusted host logic or a human can grant it.',size=16,bold=True,color=INK,cy=780000)
footer(s,'Source: Microsoft Learn, MXC architecture and containment models',10)
note(s,'把刚才的调用链按架构职责切开，从上往下就是宿主应用、SDK 与策略契约、MXC 执行层，以及最底层的系统隔离后端。真正的信任分界线，就画在宿主应用和底下被隔离的执行环境之间。\n我特别想提醒一点：很多团队接完沙箱容易产生错觉，觉得代码既然在沙箱里跑，拿回来的结果就绝对安全。其实沙箱只管不让里面的进程越界搞破坏，但它打印出来的文本、生成出来的文件里照样可能夹带脏数据或者二次注入。宿主程序站在信任边界外面，收回来的任何结果还是得先校验再往下游传。')

# ---------- 10 PICK THE RIGHT BOX
s=10
header(s,'03 · FEATURES & PLATFORMS','Pick the right box')
ry=1900000
tb(s,M,1550000,4000000,'Lighter and faster',size=14,color=MUTED)
tb(s,W-M-4000000,1550000,4000000,'Stronger isolation',size=14,color=MUTED,align='right')
rect(s,M,ry-80000,CW,220000,fill=EDGE,preset='rightArrow')
gap=180000; cw4=(CW-3*gap)/4; y0=2350000; ch=3550000
bx=[('Process container','Windows 11 · macOS · Linux','AppContainer, Seatbelt or Bubblewrap. Fast startup and fine-grained file/network rules; default for coding agents.',None),('Session container','Windows 11 only','Separate Windows login and desktop for GUI apps. Protects your screen & clipboard (open network; blocks fake keystrokes).',None),('WSL container','Windows 11','Runs Linux workloads and toolchains on a Windows host through WSL.',None),('MicroVM','Windows 11 · Linux','Hardware-virtualized unikernel boundary; never shares the host OS kernel.','Experimental')]
for i,(t,p,d,st) in enumerate(bx):
    x=M+i*(cw4+gap)
    rect(s,x+cw4/2-110000,ry-80000,220000,220000,fill=CY,line=BG,lw=25400,preset='ellipse')
    rect(s,x,y0,cw4,ch,fill=CARD)
    yy=y0+250000
    if st:
        rect(s,x+250000,yy,1500000,380000,fill=AMB,preset='roundRect',text=st,size=13,bold=True,color=BG); yy+=500000
    yy=tb(s,x+250000,yy,cw4-500000,t,size=18,bold=True,cy=700000)
    yy=tb(s,x+250000,yy+40000,cw4-500000,p,size=12.5,color=CY,bold=True,cy=550000)
    tb(s,x+250000,yy+120000,cw4-500000,d,size=13,color=MUTED,cy=y0+ch-yy-240000)
footer(s,'Sources: Windows Developer Blog, Oct 7, 2026; Microsoft Learn, MXC architecture',11)
note(s,'MXC 提供了四种不同分量的隔离盒子，怎么选就看你的业务场景。平时做代码生成、跑测试、处理本地文档，首选最左边的 Process Container，它在 Windows、macOS、Linux 上分别映射成 AppContainer、Seatbelt 和 Bubblewrap，毫秒级启动，文件和网络权限都能卡得很细。\n但如果你的 Agent 需要操作传统桌面软件、得有自己的窗口和剪贴板，那就选 Windows 专属的 Session Container。它会建一个独立的本地账号和登录会话，跟员工正在用的桌面完全隔开，还封死了模拟键鼠输入；不过它不做细粒度的文件和网络限制。再往右，要在 Windows 上跑 Linux 工具链可以用 WSL 容器；要跑高风险的陌生代码、连操作系统内核都不想共享，那就上最右边实验性的硬件级 MicroVM。')

# ---------- 11 WHERE MXC RUNS
s=11
header(s,'03 · FEATURES & PLATFORMS','Where MXC runs today')
c1=M; w1=3000000; c2=c1+w1+150000; w2=3000000; c3=c2+w2+150000; w3=W-M-c3
hy=1500000
for x,w,t in [(c1,w1,'HOST'),(c2,w2,'DEFAULT PROCESS BACKEND'),(c3,w3,'ADDITIONAL OPTIONS & CLOUD PC')]:
    tb(s,x+200000,hy,w-200000,t,size=11,color=CY,font=HF,bold=True,ls=1,cy=400000)
rowsd=[('Windows 11','ProcessContainer','Also supports IsolationSession, WSLc & Learning Mode reports'),('Linux','Bubblewrap','Also supports LXC containers'),('macOS','Seatbelt','Enforces native macOS sandbox profiles'),('Windows 365 Cloud PC','ProcessContainer','Works out of the box; session or VM backends need nested virtualization')]
y=1980000; rh=720000
for i,(a,b,c) in enumerate(rowsd):
    rect(s,c1,y,CW,rh,fill=CARD)
    tb(s,c1+200000,y,w1-200000,a,size=18,bold=True,cy=rh)
    tb(s,c2+200000,y,w2-200000,b,size=17,color=CY,font=MF,bold=True,cy=rh)
    tb(s,c3+200000,y,w3-300000,c,size=15,color=MUTED,cy=rh)
    y+=rh+100000
rect(s,M,y+100000,CW,640000,fill=CARD,line=AMB,lw=19050)
tb(s,M+300000,y+100000,CW-600000,'Experimental (explicit opt-in via 1.1-alpha): Windows Sandbox, MicroVM, and Hyperlight unikernel.',size=15,color=INK,cy=640000)
footer(s,'Source: Microsoft Learn, MXC containment options and platform support (MXC v1.0.0); Windows 365 support for MXC',12)
note(s,'做跨平台应用时，统一 SDK 省去了重写业务代码的麻烦，但每个平台底层落到哪儿我们得心里有数。Windows 11 默认走 ProcessContainer，同时支持独立会话、WSL 容器和后面要讲的学习模式；Linux 默认用 Bubblewrap，macOS 用 Seatbelt。如果团队平时在 Windows 365 云电脑上干活，进程级隔离跟物理机完全一致，连嵌套虚拟化都不用开；只有用到会话或虚拟化后端时才需要额外确认环境。\n至于最底下的 Windows Sandbox、MicroVM 和 Hyperlight，目前还属于实验特性，需要显式打开实验开关才能用。所以代码里最稳妥的做法，永远是在启动前先查一次平台支持情况，别硬编码死某一种后端。')

# ---------- 12 FILES
s=12
header(s,'03 · FEATURES & PLATFORMS','Files: read in, write out')
cx_=6000000; cy0=1600000; chh=4300000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+250000,cy0+120000,cx_-500000,'policy.json',size=12,color=FOOT,font=MF,cy=340000)
code='{\n  "version": "1.0.0",\n  "containment": "processcontainer",\n  "process": {\n    "commandLine": "python summarize.py"\n  },\n  "filesystem": {\n    "readonlyPaths": ["C:\\\\demo\\\\input"],\n    "readwritePaths": ["C:\\\\demo\\\\output"],\n    "deniedPaths": ["C:\\\\Users\\\\me\\\\.ssh"]\n  }\n}'
tb(s,M+250000,cy0+520000,cx_-500000,code,size=13,color=CODE,font=MF,cy=3560000)
rx=M+cx_+400000; rw=W-M-rx
rules=[('Read in, write out','Separate readable inputs from one dedicated writable output folder.'),('Deny the crown jewels','Explicitly deny secrets, system and profile paths.'),('Teardown is not undo','Writes to granted host folders survive the sandbox. Choose output paths deliberately.')]
y=1650000
for i,(h_,d) in enumerate(rules):
    y=tb(s,rx,y,rw,None,paras=[{'text':h_,'sizePt':19,'bold':True,'color':CY if i==0 else INK},{'text':d,'sizePt':15,'color':MUTED}])+250000
footer(s,'Source: github.com/microsoft/mxc, docs/examples.md; Microsoft Learn, MXC architecture',13)
note(s,'我们具体看看文件权限怎么管最干净。核心思路就六个字：读入、写出分离。给 Agent 看的参考资料一律挂成只读，再单独建一个空的专用目录让它写产出物；同时用 deniedPaths 把 .ssh、各类密钥和用户隐私目录明确拉黑，千万别图省事把整个用户主目录直接放开。\n还有个细节特别容易踩坑：销毁沙箱不等于撤销落盘。容器退出时清掉的是它自己的临时运行环境，只要你允许它往宿主机的输出目录写了文件，那些改动在容器销毁后是实打实留下的。所以一定要把输出目录当成待检区，宿主程序先验完格式和内容，再拿去正式使用。')

# ---------- 13 NETWORK
s=13
header(s,'03 · FEATURES & PLATFORMS','Network: deny by default')
bx0=M; by0=1900000; bw=2400000; bh=3400000
rect(s,bx0,by0,bw,bh,fill=CARD2,line=CY,lw=28575)
tb(s,bx0+200000,by0+1000000,bw-400000,None,paras=[{'text':'Agent workload','sizePt':19,'bold':True,'color':INK},{'text':'egress: deny','sizePt':15,'color':CY},{'text':'ingress: deny','sizePt':15,'color':CY}])
dx=bx0+bw+1250000; dw=2900000; dh=840000
dests=[('140.82.114.6 : 443','Allowed by CIDR + port',CY,'✓'),('Any other address','Blocked',AMB,'✕'),('Local proxy 127.0.0.1:8080','Domain allow-list lives here',CY,'→')]
for i,(t,d,c,g) in enumerate(dests):
    yy=by0+120000+i*(dh+380000)
    rect(s,bx0+bw+120000,yy+dh/2-160000,1000000,320000,fill=c,preset='rightArrow')
    rect(s,dx,yy,dw,dh,fill=CARD,line=c,lw=19050)
    tb(s,dx+180000,yy,dw-300000,None,paras=[{'text':g+'  '+t,'sizePt':15,'bold':True,'color':INK},{'text':d,'sizePt':13,'color':MUTED}],cy=dh)
fx=dx+dw+350000; fw=W-M-fx
facts=['Rules are directional: egress and ingress, each with its own default.','Allow by IP/CIDR, protocol and port, not hostnames.','Need domains? Route through an HTTP(S) proxy with runtimeConfig.networkProxy.','In production, lock proxy traffic to your trusted process via allowedProxyPeer.']
tb(s,fx,1750000,fw,None,paras=[{'text':f} for f in facts],size=15,color=MUTED)
footer(s,'Source: github.com/microsoft/mxc, docs/examples.md',14)
note(s,'光锁住文件还不够，网络出口要是不管，Agent 读到的任何一点东西都能顺着网线发走。MXC 把网络拆成出站、入站和本机回环三个方向，上手第一步就是把默认规则全设成拒绝。\n这里要注意，操作系统底层的防火墙规则认的是 IP、网段和端口号，它没法直接认域名。如果业务上必须按域名放行，比如只允许连公司的内部网关或者指定的包管理源，标准做法是在宿主机起一个受控的本地代理，通过 networkProxy 把沙箱流量引过去，由代理来做域名白名单和日志审计；在生产环境里还可以加上 allowedProxyPeer 校验代理进程身份，防止本机其他程序冒名顶替。')

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
note(s,'文件与网络之外，最敏感的就是密钥和用户的桌面环境了。有一条铁律可以省掉九成的泄密风险：永远别把 API Key、SSH 私钥直接塞进沙箱的环境变量或配置文件里。沙箱里只放干活必需的代码和工具；真要调受保护的外部接口，让沙箱外的宿主程序或者刚刚说的本地代理在审完请求后，在外面替它把凭证签上去。\n用户界面也是一样，普通沙箱碰不到你的窗口和剪贴板，也没法伪造鼠标键盘敲击。就算用了带独立桌面的 Session Container，它的窗口和剪贴板也只属于那个临时账号自己；而且一旦在那个会话里登录了 ERP 等业务系统，里面的 Agent 是能摸到当前登录态的，所以务必给它配专门的低权限机器人账号，别拿员工本人的账号去登。')

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
note(s,'聊到最小权限，大家落地时肯定会遇到一个头疼问题：跑一个编译工具或者脚本动不动要读几十个系统路径，我一开始哪猜得到该放行哪些目录？全靠手写肯定跑不通。MXC 专门做了一个 Learning Mode 来帮我们探路。\n平时生产上线开的是严格拦截的 Enforcement，排查疑难杂症可以临时开只记日志不拦截的 Permissive；而在策略打磨阶段，最实用的就是 Learning Mode——它碰到没授权的访问照样拦住，但会把每一次被拒的尝试完整记进活动报告里。你可以先拿真实任务跑几轮，拿到这份证据清单，由工程师逐条甄别哪些真是正常依赖、哪些是越界试探，人工拍板加进策略之后，再切回强制模式上线。')

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
lane(3000000,'State-aware','When state must persist',['Provision','Start','Execute','Stop','Deprovision'],'Persistent shells and multi-step triage via IsolationSession or WSLc.')
ry=4450000; g=150000; cw3=(CW-2*g)/3
for i,(t,d) in enumerate([('run','Returns captured output when the workload finishes.'),('spawn','Live stdin, stdout and stderr while it runs.'),('PTY','Interactive terminal session for stateful shells and human handoff.')]):
    x=M+i*(cw3+g)
    rect(s,x,ry,cw3,1350000,fill=CARD)
    tb(s,x+250000,ry,cw3-500000,None,paras=[{'text':t,'sizePt':20,'bold':True,'color':CY},{'text':d,'sizePt':15,'color':MUTED}],cy=1350000)
footer(s,'Source: Microsoft Learn, MXC architecture (lifecycles) and platform support',17)
note(s,'再看容器该活多久。MXC 支持用完即焚的 One-shot 和保留现场的 State-aware 两种生命周期。日常百分之八九十的动作，像跑一次编译、执行一遍测试、生成一份报告，都应该优先选 One-shot：每次调用现拉一个全新的窄权限沙箱，干完立刻销毁，不给潜在问题留任何驻留空间。\n只有当你确实需要跨多步保留上下文，比如预装好一套复杂环境，或者开着终端连续排查故障时，才用 IsolationSession 或 WSLc 这种有状态容器。用的时候一定记得，停止运行 stop 不代表释放资源，不管中间成没成功，最后都要调 deprovision 把临时账号和环境清理干净。底下的三种执行接口也很清晰：等完整结果用 run，盯实时输出流用 spawn，做交互式命令行终端就用 PTY。')

# ---------- 17 CODING AGENT (image left)
s=17
img(s,'hd-coding.jpg',0,0,PANEL,H,'Hand-drawn sketch panel: a robot coding inside a fenced project folder while a locked vault with a key sits outside')
xl=PANEL+M; cxr=W-xl-M
header(s,'04 · USE CASES','Secure coding agents',x=xl,cx=cxr)
y=tb(s,xl,1550000,cxr,'Build, test and push, without touching anything else.',size=18,color=INK,cy=600000)
perm=[('Project repo  C:\\src\\app','Read / write',CY),('Toolchain: node, python, git','Allowed',CY),('Package registry via proxy','Allowed',CY),('~/.ssh, tokens, Documents','Denied',AMB),('Any other network destination','Blocked',AMB)]
y+=200000; rh=560000
for t,st,c in perm:
    rect(s,xl,y,cxr,rh,fill=CARD)
    tb(s,xl+250000,y,cxr-2200000,t,size=15,color=INK,cy=rh)
    rect(s,xl+cxr-1750000,y+90000,1550000,rh-180000,fill=c,preset='roundRect',text=st,size=13,bold=True,color=BG)
    y+=rh+90000
tb(s,xl,y+150000,cxr,'In production: GitHub Copilot CLI uses MXC process isolation.',size=14,bold=True,color=CY,cy=500000)
footer(s,'Source: Windows Developer Blog, Jun 2026 · example policy',18,x=xl,cx=cxr)
note(s,'把前面讲的拼到一起，最直观的例子就是写代码的 Coding Agent。你想让它改代码、跑测试、提分支，仔细算算它真正需要的权限其实非常窄：当前这一个仓库目录可读写，node、python、git 这些工具链可调用，再加上通过本地代理去拉依赖包。至于电脑上别的项目、你的个人文档、~/.ssh 密钥，还有随便连外网的权限，全给它关掉，一点不耽误干活。\n这套模式已经在真实大规模生产里跑通了——大家平时用的 GitHub Copilot CLI，底层就是靠 MXC 的进程隔离来兜底约束模型动态生成和执行的命令。')

# ---------- 18 DESKTOP & DATA (image right)
s=18
img(s,'hd-desktop.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a human relaxes at his own desk while a robot agent processes invoices at its own separate desktop behind a glass wall')
cxl=W-PANEL-2*M
header(s,'04 · USE CASES','Desktop and data agents',cx=cxl)
cards=[('Desktop automation agent','Keeps app state in a separate Windows account and session, driven through app APIs or UI automation.','Best fit: Session container (isolated desktop, dedicated bot login)'),('Sensitive-document agent','Reads one approved folder, writes one report, no network.','Best fit: Process container (read-in / write-out, network denied)'),('Agents on Cloud PCs','Run beside existing work on Windows 365; process isolation works as on a physical PC.','Best fit: Process container on Windows 365 Cloud PC (GA)')]
y=1600000; ch=1380000
for t,d,f in cards:
    rect(s,M,y,cxl,ch,fill=CARD)
    tb(s,M+250000,y,cxl-500000,None,paras=[{'text':t,'sizePt':17.5,'bold':True,'color':INK},{'text':d,'sizePt':13.5,'color':MUTED},{'text':f,'sizePt':13.5,'color':CY,'bold':True}],cy=ch)
    y+=ch+110000
footer(s,'Sources: Windows Developer Blog; Microsoft Learn',19,cx=cxl)
note(s,'除了研发写代码，企业业务里还有几个场景特别对路。比如操作传统客户端或内网系统的桌面 Agent，给它开一个 Session Container，让它在独立的本地账号和会话里通过 UI Automation 或应用接口跑流程——员工在前台正常办公，Agent 在后台独立桌面里处理单据，两边互不干扰，再配上专门的低权限业务账号就稳妥了。\n再比如审合同、做财报汇总的敏感文档 Agent，直接关进完全断网的 Process Container，只读一个输入目录、只写一个结果目录，物理上断掉外传可能。还有部署在 Windows 365 云电脑上的 Agent，跟员工的云桌面并排运行，进程隔离的用法和体验跟本地电脑完全一致。')

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
    tb(s,panel_x+220000,y+40000,panel_w-440000,title,size=11,color=CY,font=HF,bold=True,ls=1,cy=280000)
    tb(s,panel_x+220000,y+320000,panel_w-440000,body,size=12.5,color=INK,cy=h-340000)
    y+=h+180000
footer(s,'Sources: Windows Developer Blog & Microsoft Ignite 2026 partner announcements',20)
note(s,'而且 MXC 不是微软关起门来自用的组件，整个 AI 工具链生态接得非常快。像 GitHub Copilot、OpenAI Codex、OpenClaw、Replit、LM Studio 还有 Unsloth AI 现在都已经直接支持 MXC 了；NVIDIA 的 OpenShell 已经集成进来，OpenAI 和 Meta 也在基于这套新的隔离机制往下做深；再加上 Anthropic Claude Code、Box、Perplexity、Raycast 等一众厂商也陆续官宣支持。对企业架构师来说是个好消息：不同工具底层正在往同一套操作系统级安全基础设施收敛，不用每引一个工具就重搭一套孤立的沙箱。')

# ---------- 20 SDKs
s=20
header(s,'05 · INTEGRATING MXC','Three SDKs, one contract')
g=238800; cw3=(CW-2*g)/3; y0=1650000; ch=2700000
sdks=[('.NET','Microsoft.Mxc.Sdk','NuGet.org','GetPlatformSupport()'),('TypeScript','@microsoft/mxc-sdk','npm','getPlatformSupport()'),('Rust','mxc-sdk','crates.io','v1::platform_support()')]
for i,(t,pkg,reg,api) in enumerate(sdks):
    x=M+i*(cw3+g)
    rect(s,x,y0,cw3,ch,fill=CARD)
    yy=tb(s,x+250000,y0+150000,cw3-500000,t,size=22,bold=True,font=HF,cy=520000)
    yy=tb(s,x+250000,yy+40000,cw3-500000,pkg,size=15,bold=True,color=CY,font=MF,cy=650000)
    yy=tb(s,x+250000,yy+20000,cw3-500000,'from '+reg,size=14,color=MUTED,cy=350000)
    yy=tb(s,x+250000,yy+120000,cw3-500000,'DISCOVERY',size=12,color=FOOT,font=HF,bold=True,ls=1,cy=280000)
    tb(s,x+250000,yy+40000,cw3-500000,api,size=12,color=INK,font=MF,cy=400000)
rect(s,M,4600000,CW,1200000,fill=CARD2,line=CY,lw=19050)
tb(s,M+300000,4600000,CW-600000,None,paras=[{'text':'Write one ContainerRequest in code — the SDK maps it to the right OS sandbox on Windows, macOS, and Linux.','sizePt':17,'bold':True,'color':INK},{'text':'Requirements: Windows 11 24H2+ (physical or Windows 365), macOS, or Linux  ·  .NET 8+ / Node.js 18+ / Rust.','sizePt':15,'color':MUTED}],cy=1200000)
footer(s,'Sources: Microsoft Learn, MXC overview, platform support and Windows 365 support for MXC',21)
note(s,'如果你准备把它接进自家的 Agent 运行时，工程接入非常轻。官方在 NuGet、npm 和 crates.io 上同步维护了 .NET、TypeScript 和 Rust 三套 SDK。不管你用哪门语言，写法都是一致的：在代码里填好统一的 ContainerRequest，声明清楚文件、网络、命令和超时，SDK 会自动帮你序列化成稳定的 1.0.0 标准契约，不用你操心底层系统的差异。\n接入时只要养成一个好习惯：启动任务前先调 getPlatformSupport 确认当前机器的支撑情况；如果在 Windows 上跑，还可以顺手调一下 probe，提前确认你配的这份策略能不能拿到预期的原生隔离级别。')

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
side=[('Egress proxy','Attaches credentials and enforces a domain allow-list, outside the box.'),('Learning Mode','Activity reports tighten the policy before enforcement.'),('Enterprise control','Entra Agent ID & Agent 365 govern agent identity, audit trails, and fleet policy.')]
for i,(t,d) in enumerate(side):
    x=M+i*(cw3+g2)
    rect(s,x,y1,cw3,ch,fill=CARD)
    tb(s,x+250000,y1,cw3-500000,None,paras=[{'text':t,'sizePt':18,'bold':True,'color':CY},{'text':d,'sizePt':15,'color':MUTED}],cy=ch)
footer(s,'Sources: Microsoft Learn & Windows Developer Blog',22)
note(s,'把视角拉高一点，这就是一套可以直接落地的生产参考架构。当大模型规划器决定调工具，比如想执行一句 npm test 时，指令绝不能直通操作系统，而是先交到宿主侧的可信策略引擎手里；宿主按任务类型选好现成的最小权限模板，再通过 MXC SDK 拉起对应的隔离后端去跑。\n跑完往回走的那条路径同样不能省：退出码正不正常、产出格式对不对、中间有没有触发越权拒绝，全由宿主验过关之后再把结果喂回给 Agent。外围再配上管域名和挂凭证的出站代理、帮着收敛权限的 Learning Mode，以及做身份与全局治理的 Entra Agent ID 和 Agent 365，整套防线就严丝合缝了。')

# ---------- 22 FIVE STEPS + CODE
s=22
header(s,'05 · INTEGRATING MXC','Integrate in five steps')
st=[('Discover','getPlatformSupport() and probe() before launching.'),('Describe','Build a ContainerRequest: files, network, UI, command, timeout.'),('Learn','Run in Learning Mode; review the activity report.'),('Enforce','Remove learning settings; test allowed and denied cases.'),('Operate','Handle timeouts and failures; explain denials to users.')]
lw0=5300000; y=1600000; rh=820000
for i,(t,d) in enumerate(st):
    tb(s,M,y,700000,str(i+1),size=30,color=CY,font=HF,bold=True,cy=rh)
    tb(s,M+700000,y,lw0-700000,None,paras=[{'text':t,'sizePt':17,'bold':True,'color':INK},{'text':d,'sizePt':13.5,'color':MUTED}],cy=rh)
    y+=rh+60000
cx0=M+lw0+300000; cw_=W-M-cx0; cy0=1600000; chh=4300000
rect(s,cx0,cy0,cw_,chh,fill=CODEBG,line=EDGE)
tb(s,cx0+250000,cy0+120000,cw_-500000,'host.ts  ·  @microsoft/mxc-sdk',size=12,color=FOOT,font=MF,cy=340000)
code=['const support = await getPlatformSupport();','if (!support.isSupported) throw support.reason;','','const request = {','  command: "npm test",','  workingDirectory: repo,','  filesystem: { readwritePaths: [repo] },','  network: { egress: { default: "deny" } },','  timeoutMs: 120_000,','};','const result = await runContained(request);','validate(result); // never trust blindly']
tb(s,cx0+250000,cy0+520000,cw_-400000,None,paras=[{'text':l if l else ' '} for l in code],size=13,color=CODE,font=MF)
footer(s,'Source: Microsoft Learn, MXC SDK integration guide',23)
note(s,'具体到研发团队写代码，接进项目里基本就是右边这十几行的五步闭环：先调 getPlatformSupport 和 probe 摸清当前机器能力，不支持就明确报错，绝不静默降级成裸跑；接着按默认拒绝写好 ContainerRequest，把命令、工作目录、读写边界和超时定死；开发期打开 Learning Mode 收集真实访问需求；人工审完策略后关掉学习开关，把允许路径和越权拦截路径都测通过；最后在运行期把超时、异常和权限被拒的原因清晰返给上层，别让沙箱正常拦住了危险动作，外面的 Agent 却以为出了故障在那儿盲目重试。')

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
note(s,'进 Demo 之前，我们最后把宿主程序自己该扛的职责做个小结。跑任务一定要设超时，并且把沙箱没拉起来、里面命令自己跑错了、还有用户主动取消这几类情况区分开；不管哪一步出了异常，收尾的资源回收和 deprovision 都得保证执行，别留下一堆悬空的临时账号。\n分配权限时，坚持一次动作配一个极窄的短命沙箱，远比偷懒复用一个大权限长寿容器安全得多。产出物单独落到专用目录，记住销毁沙箱不会撤回已经写进宿主机目录的文件。遇到越权拦截要大方地把原因告诉用户，而 Learning Mode 跑出来的任何放行建议，上线前都必须经过真人点头。')

# ---------- 24 DEMO 1 (PROCESSCONTAINER: FILES + NETWORK)
s=24
header(s,'06 · LIVE DEMO 1','ProcessContainer: files & network')
cx_=5550000; cy0=1620000; chh=4300000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+250000,cy0+140000,cx_-500000,'policy.json  ·  @microsoft/mxc-sdk/v1',size=12,color=FOOT,font=MF,cy=340000)
pc_code='{\n  "containment": { "type": "processcontainer" },\n  "filesystem": {\n    "readonlyPaths": ["demo/input"],\n    "readwritePaths": ["demo/output"],\n    "deniedPaths": ["demo/protected/denied.txt"]\n  },\n  "network": {\n    "egress": { "default": "deny" },\n    "ingress": { "hostLoopback": "allow" }\n  }\n}'
tb(s,M+250000,cy0+480000,cx_-500000,pc_code,size=11,color=CODE,font=MF,cy=2960000)
rect(s,M+250000,cy0+3480000,cx_-500000,720000,fill=CARD2,line=CY,lw=12700)
tb(s,M+420000,cy0+3480000,cx_-840000,None,paras=[
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
note(s,'原理和架构聊透了，我们直接来看三个真机实跑的 Demo。\n第一个 Demo 跑的是大家最常用的 ProcessContainer。我们先调 getPlatformSupport 和 probe 验证当前 Windows 内核的原生容器能力；接着在只读 input、只写 output 的最小权限下顺利生成报告；随后保持策略一字不改，让同一个脚本去偷读 protected/denied.txt 里的机密内容，你会看到操作系统内核当场拦截；最后在完全封死外网的前提下，对比开启和关闭本机回环时对本地监听端口的访问结果，亲眼确认文件和网络两道闸门都在内核层生效。')

# ---------- 25 DEMO 2 (SESSION PTY, image right)
s=25
img(s,'hd-blocked.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a sneaky amber tentacle pulls a robot agent toward a door, but a force field blocks it')
cxl=W-PANEL-2*M
header(s,'06 · LIVE DEMO 2','Isolated incident triage',cx=cxl)
beats=[('Stage','Host places INC-1042 and its read-only approval in the shared workspace.'),('Open PTY','Spins up a separate local Windows account & session, isolated from the employee desktop.'),('Triage','Check Spooler and W32Time, retain context, and write auditable evidence.'),('Close','Host validates the report, stops the session and deprovisions the operator.')]
railx=M+150000; y=1550000
ys=[]
for i,(t,d) in enumerate(beats):
    ys.append(y)
    y=tb(s,M+600000,y,cxl-600000,None,paras=[{'text':t,'sizePt':18,'bold':True,'color':CY if i!=3 else AMB},{'text':d,'sizePt':13.5,'color':MUTED}],cy=1060000)+80000
rect(s,railx+100000-12700,ys[0]+200000,25400,ys[-1]-ys[0],fill=EDGE)
for i,yy in enumerate(ys):
    rect(s,railx,yy+100000,200000,200000,fill=CY if i!=3 else AMB,preset='ellipse')
footer(s,'PS> npm --prefix demo run demo:pty  ·  6 stateful PTY steps in Demo Studio',26,cx=cxl)
note(s,'第二个 Demo 模拟很真实的运维一线排障，用的是保留状态的 IsolationSession 配合交互式 PTY 终端。假设来了一张打印和时间服务异常的工单 INC-1042，宿主程序先把工单和只读诊断授权放进共享目录，然后现场建一个独立的 Windows 临时操作员账号和专属 Session，拉起 PowerShell 终端。\n大家一会儿看的时候可以重点盯一眼账号对比：终端里的用户名和 Session ID 跟当前登录电脑的员工完全隔离，它查它的 Spooler 和 W32Time 服务、写它的排查证据，需要时真人工程师还能直接接管这个终端敲命令，全程绝对碰不到员工自己的桌面和剪贴板。等宿主核验完生成的排查报告，立刻关闭会话并把临时账号彻底注销删干净。')

# ---------- 26 DEMO 3 (HYPERLIGHT MICROVM)
s=26
header(s,'06 · LIVE DEMO 3','Untrusted code in a MicroVM')
cx_=5550000; cy0=1620000; chh=4300000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
rect(s,M+cx_-1700000,cy0+140000,1450000,340000,fill=AMB,preset='roundRect',text='Experimental',size=11,bold=True,color=BG)
tb(s,M+250000,cy0+140000,cx_-2000000,'hyperlight-policy.json',size=12,color=FOOT,font=MF,cy=340000)
vm_code='{\n  "version": "1.1.0-alpha",\n  "containment": "hyperlight",\n  "hyperlight": {\n    "runtime": "python"\n  },\n  "process": {\n    "commandLine": "python eval_untrusted.py",\n    "timeout": 30000\n  }\n}'
tb(s,M+250000,cy0+580000,cx_-500000,vm_code,size=12.5,color=CODE,font=MF,cy=2840000)
rect(s,M+250000,cy0+3480000,cx_-500000,740000,fill=CARD2,line=CY,lw=12700)
tb(s,M+420000,cy0+3480000,cx_-840000,None,paras=[
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
note(s,'第三个 Demo 看隔离级别拉满的实验性后端：Hyperlight MicroVM。什么时候非用它不可？比如你的系统得直接运行大模型现写的复杂 Python 逻辑，或者来路不明的第三方代码，你连宿主机的 Windows 内核都不放心跟它共用；可传统虚拟机一开机就要好几秒，在线业务根本等不起。\nHyperlight 走的是另一条路：它在 Windows Hypervisor 硬件虚拟化之上跑极简的 Unikraft 微内核，每次直接从预热好的内存快照恢复，90 毫秒左右就能跑完返回。演示里我们会顺着走完整个闭环：先检查预热快照，确认里面跑的是独立的 Unikraft 5.15 内核而不是 Windows；接着在百毫秒内算完一段不可信的 Python 对账脚本；再故意跑探针代码去摸宿主机 C 盘、环境变量和网络套接字，看它们被硬件边界全部挡死；最后验证上一次运行留在 /tmp 里的文件，在下一次快照恢复时瞬间清零。')

# ---------- 27 CLOSER
s=27
img(s,'hd-closer.jpg',0,0,W,H,'Hand-drawn sketch illustration: three robot agents in bubbles and an engineer giving a thumbs up at sunrise over a city')
x=M+150000; cx=5900000
tb(s,x,650000,cx,'KEY TAKEAWAYS',size=12,color=CY,font=HF,bold=True,ls=2,cy=280000)
tb(s,x,1000000,cx,'Box the agent,\nnot the ambition',size=38,color=INK,font=HF,bold=True,cy=1300000)
tk=['1   Push guardrails down to the OS. Prompts alone can\'t enforce.','2   MXC: one policy, the right sandbox on Windows, macOS and Linux.','3   Learn first, then enforce least privilege.','4   Containment isn\'t trust. Validate every result.']
tb(s,x,2500000,cx,None,paras=[{'text':t} for t in tk],size=14.5,color=INK,cy=2200000)
tb(s,x,4850000,cx,None,paras=[{'text':'learn.microsoft.com/windows/apps/develop/mxc'},{'text':'github.com/microsoft/mxc'}],size=14,color=CY,cy=850000)
tb(s,x,5850000,cx,'Thank you  ·  Q&A',size=24,color=INK,font=HF,bold=True,cy=700000)
note(s,'最后用四句话给今天的分享做个收尾：把安全护栏下沉到操作系统去强制执行，别让提示词独自扛雷；用好 MXC 这一份通用策略，让它在 Windows、macOS 和 Linux 上各自匹配最顺手的原生沙箱；落地时先用 Learning Mode 看清真实访问足迹，人工审过再锁死最小权限；同时永远记得，关进沙箱不等于结果可信，拿回来的输出一定要再验一道。\n如果大家回去想挑第一个场景试水，我最推荐从 Coding Agent 或者断网的敏感文档处理开始，花半天时间把输入输出目录和拒绝列表配出来，正反两条路径一测，马上就能看到实实在在的效果。今天的内容就到这里，谢谢大家，看看大家有什么问题！')


# ===================== BUILD-TIME LAYOUT VALIDATOR =====================
def validate_layout():
    errors=[]
    for k in sorted(slides):
        ops=slides[k]
        for o in ops:
            if o['cx']<=0 or o['cy']<=0 or o['x']<0 or o['y']<0 or o['x']+o['cx']>W or o['y']+o['cy']>H:
                errors.append(f'Slide {k+1}: {o["op"]} has invalid bounds')
        tbs=[o for o in ops if o['op']=='add_text_box']
        for o in tbs:
            if o.get('name')=='slide_title':
                txt=o.get('text') or o['paragraphs'][0]['text']
                if nlines(txt,o.get('sizePt',29.5),o['cx'],'Newsreader',True)>1:
                    errors.append(f'Slide {k+1}: header title wraps >1 line ({txt!r})')
        for i in range(len(tbs)):
            for j in range(i+1,len(tbs)):
                a,b=tbs[i],tbs[j]
                x_ov=min(a['x']+a['cx'],b['x']+b['cx'])-max(a['x'],b['x'])
                y_ov=min(a['y']+a['cy'],b['y']+b['cy'])-max(a['y'],b['y'])
                if x_ov>20000 and y_ov>0:
                    ta=(a.get('text') or a['paragraphs'][0]['text'])[:30]
                    tb_=(b.get('text') or b['paragraphs'][0]['text'])[:30]
                    errors.append(f'Slide {k+1}: textbox overlap ({y_ov} EMU) between {ta!r} and {tb_!r}')
    if errors:
        raise ValueError('Layout validation failed:\n  ' + '\n  '.join(errors))

# ===================== WRITER (python-pptx) =====================
PRESETS={'rect':MSO_SHAPE.RECTANGLE,'roundRect':MSO_SHAPE.ROUNDED_RECTANGLE,'ellipse':MSO_SHAPE.OVAL,
         'rightArrow':MSO_SHAPE.RIGHT_ARROW,'leftArrow':MSO_SHAPE.LEFT_ARROW,'downArrow':MSO_SHAPE.DOWN_ARROW,
         'chevron':MSO_SHAPE.CHEVRON,'homePlate':MSO_SHAPE.PENTAGON}
ALIGN={'left':PP_ALIGN.LEFT,'center':PP_ALIGN.CENTER,'right':PP_ALIGN.RIGHT}
rgb=lambda h:RGBColor.from_string(h)

def write_text(tf,o):
    tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    tf.margin_left=Emu(76200); tf.margin_right=Emu(76200)
    tf.margin_top=Emu(57150); tf.margin_bottom=Emu(57150)
    paras=o.get('paragraphs') or [{'text':o['text']}]
    raw_font=o.get('fontFamily',BF)
    default_sa=o.get('spaceAfterPt', 2.0 if raw_font==MF else 6.5)
    for i,p in enumerate(paras):
        para=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        para.alignment=ALIGN[p.get('align',o.get('align','center'))]
        sz=p.get('sizePt',o.get('sizePt',16))
        para.line_spacing=1.35 if raw_font==MF else (1.12 if sz>=29 else 1.24)
        if i<len(paras)-1:
            para.space_after=Pt(p.get('spaceAfterPt',default_sa))
        lines=p['text'].split('\n')
        for j,line in enumerate(lines):
            if j: para.add_line_break()
            r=para.add_run(); r.text=line
            if p.get('href'): r.hyperlink.address=p['href']
            f=r.font
            f.name=raw_font; f.size=Pt(sz)
            f.bold=p.get('bold',o.get('bold',False)); f.color.rgb=rgb(p.get('color',o.get('color',INK)))
            if o.get('letterSpacingPt'): r._r.get_or_add_rPr().set('spc',str(int(o['letterSpacingPt']*100)))

def add_picture(sl,o,img_dir):
    path=os.path.join(img_dir,os.path.basename(o['path']))
    if path.lower().endswith('.svg'):
        path=path[:-4]+'.png'
    with Image.open(path) as image:
        iw,ih=image.size
    x,y,cx,cy=o['x'],o['y'],o['cx'],o['cy']
    fr=cx/cy; ir=iw/ih
    if o.get('fit')=='contain':
        if ir>fr:
            nh=int(cx/ir); y=int(y+(cy-nh)/2); cy=nh
        else:
            nw=int(cy*ir); x=int(x+(cx-nw)/2); cx=nw
        pic=sl.shapes.add_picture(path,Emu(x),Emu(y),Emu(cx),Emu(cy))
    else:
        pic=sl.shapes.add_picture(path,Emu(x),Emu(y),Emu(cx),Emu(cy))
        if abs(ir-fr)>1e-4:
            if ir>fr: c=(1-fr/ir)/2; pic.crop_left=pic.crop_right=c
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
    validate_layout()
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
