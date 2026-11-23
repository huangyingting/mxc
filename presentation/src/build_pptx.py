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
BG='F3EEE4'; CARD='E9E2D4'; CARD2='FAF7F1'; EDGE='C9C0AE'; INK='2B2B2B'; MUTED='57534C'; FOOT='6E685F'
CY='3F607E'; AMB='A84A35'; CODEBG='FAF7F1'; CODE='2F4B66'
HF='Segoe UI Variable Display'; BF='Calibri'; MF='Consolas'
PANEL=4572000   # width of the side illustration panel
# ===================== HELPERS ===================
slides={}  # idx -> list of patches
notes={}
def P(s,op): slides.setdefault(s,[]).append(op)
WF={'Inter Tight':0.50,'Segoe UI Variable Display':0.51,'Segoe UI Variable Text':0.50,'Calibri':0.50,'Consolas':0.55,'Inconsolata':0.50,'Plus Jakarta Sans':0.55,'Newsreader':0.58,'JetBrains Mono':0.62,'Aptos Display':0.56,'Aptos':0.53,'Cascadia Code':0.62,'Georgia':0.6,'Arial':0.56,'Courier New':0.62}
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
    tb(s,x,310000,cx,eyebrow,size=12,color=CY,font=HF,bold=True,ls=0.9,cy=280000,name='eyebrow')
    tb(s,x,640000,cx,title,size=29.5,color=INK,font=HF,bold=True,cy=title_cy or 640000,name='slide_title')
def footer(s,src,num,x=M,cx=CW):
    tb(s,W-940000,H-430000,700000,f'{num:02d}',size=11.5,color=FOOT,font=MF,bold=True,align='right',cy=290000,name='footer_num')
def note(s,t): notes[s]=t

def card_box(s,x,y,cx,cy,paras,fill=CARD,line=EDGE,lw=12700,pad_x=240000,pad_y=120000,font=BF,space_after=6.5,align='left'):
    rect(s,x,y,cx,cy,fill=fill,line=line,lw=lw if line else None)
    tb(s,x+pad_x,y+pad_y,cx-2*pad_x,None,paras=paras,font=font,align=align,cy=cy-2*pad_y,space_after=space_after)

def banner_box(s,x,y,cx,cy,text,accent=CY,fill=CODEBG,color=INK,size=15.5,bold=True,align='left'):
    rect(s,x,y,cx,cy,fill=fill,line=EDGE,lw=12700)
    rect(s,x+95000,y+110000,46000,cy-220000,fill=accent,preset='roundRect')
    tb(s,x+240000,y,cx-480000,text,size=size,bold=bold,color=color,align=align,cy=cy)

def medallion(s,cx_mid,cy_mid,diam=360000,glyph='→',fill=CY,color=BG,size=14):
    r=diam//2
    rect(s,cx_mid-r,cy_mid-r,diam,diam,fill=fill,preset='ellipse',text=glyph,size=size,bold=True,color=color,font=MF)

# ===================== SLIDES ====================
# ---------- 0 COVER
s=0
img(s,'hd-cover.jpg',0,0,W,H,'Hand-drawn sketch illustration: a friendly robot AI agent standing inside a glass cube, with files, cloud and keys kept outside the boundary')
x=M+150000; cx=5900000
rect(s,x,480000,560000,40000,fill=CY,preset='roundRect')
tb(s,x,590000,cx,'MICROSOFT EXECUTION CONTAINERS (MXC)',size=12,color=CY,font=HF,bold=True,ls=0.9,cy=260000)
tb(s,x,920000,cx,'Guardrails with\nSandboxing',size=44,color=INK,font=HF,bold=True,cy=1380000)
tb(s,x,2400000,cx,'Keeping AI agents inside safe boundaries: isolating execution and controlling access to files, network and credentials on Windows',size=16.2,color=MUTED,cy=1100000)
tb(s,x,3620000,cx,'Yingting Huang  |  Sr. Cloud Solution Architect, Microsoft',size=15,color=INK,bold=True,cy=360000)
footer(s,'',1)
note(s,'大家好，欢迎来到今天的分享。这两年大家都在做 Agent，你会发现它的角色已经变了——以前大模型只是陪你聊天、写段代码让你自己复制粘贴；现在的 Agent 是真正在你的电脑或服务器上干活，自己读文件、跑命令、调接口。\n能力变强之后，安全问题也跟着变了：万一它理解错了意图，或者读到了别人埋了坑的文档，怎么保证它不会把系统搞乱、把密钥传出去？今天我们就来聊微软开源的沙箱执行层 MXC，看看怎么在放手让 Agent 干活的同时，在操作系统这一层把安全边界牢牢守住。')

# ---------- 1 AGENDA
s=1
header(s,'AGENDA','Our journey today')
items=[('01','Why sandboxing','Risks of agents that act'),('02','Meet MXC','What it is and how it works'),('03','Features & platforms','Backends, lifecycle & rules'),('04','Use cases','Coding, desktop & data agents'),('05','Integration','Architecture, SDKs & steps'),('06','Live demos','Process, Session & MicroVM'),('07','Takeaways & Q&A','What to do next')]
gap_a=115000; cw7=(CW-6*gap_a)/7; y0_a=1960000; ch_a=3260000
rect(s,M+200000,1680000,CW-400000,26000,fill=EDGE,preset='roundRect')
for i,(n,t,d) in enumerate(items):
    cx0=M+i*(cw7+gap_a)
    rect(s,cx0+cw7/2-70000,1623000,140000,140000,fill=CY,preset='ellipse')
    rect(s,cx0,y0_a,cw7,ch_a,fill=CARD,line=EDGE,lw=12700)
    pad_a=135000
    rect(s,cx0+pad_a,y0_a+210000,660000,300000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=n,size=12,bold=True,color=CY,font=MF)
    tb(s,cx0+pad_a,y0_a+660000,cw7-2*pad_a,t,size=16,color=INK,bold=True,cy=780000)
    rect(s,cx0+pad_a,y0_a+1520000,cw7-2*pad_a,14000,fill=EDGE)
    tb(s,cx0+pad_a,y0_a+1640000,cw7-2*pad_a,d,size=13.2,color=MUTED,cy=1240000)
footer(s,'',2)
note(s,'今天这 45 分钟，我们顺着实际工程落地的脉络往下讲。先看看当 Agent 开始动手执行命令之后，到底会带来哪些新风险，为什么光靠写提示词根本防不住；接着拆解 MXC 是什么，以及它怎么把一份权限策略变成操作系统底层的硬隔离；然后聊大家最关心的文件、网络、密钥和四种不同轻重的隔离容器该怎么选；最后结合写代码、桌面操作和敏感数据处理，看看怎么用几行 SDK 代码把它接进你自己的系统，并现场跑三个真机 Demo。')

# ---------- 2 AGENTS ACT
s=2
header(s,'01 · WHY SANDBOXING','Agents don\'t just talk. They act.')
cw_=5180000; cyy=1540000; ch=3680000
for idx_c,(xx,badge,title_c,sub_c,bullets,risk_lbl,is_act) in enumerate([
    (M,'HUMAN IN THE LOOP','Chat assistant','Answers questions',[
        '• Generates text and code suggestions',
        '• A human reviews, copies and runs commands',
    ],'Worst case: a wrong answer',False),
    (W-M-cw_,'DIRECT OS EXECUTION','Autonomous agent','Takes actions on your machine',[
        '• Reads and writes local workspace files',
        '• Runs shell commands & installs packages',
        '• Calls APIs and drives desktop apps',
    ],'Worst case: your whole session compromised',True),
]):
    rect(s,xx,cyy,cw_,ch,fill=CARD2 if is_act else CARD,line=CY if is_act else EDGE,lw=19050 if is_act else 12700)
    pad=260000
    rect(s,xx+pad,cyy+220000,2100000,290000,fill=CY if is_act else CODEBG,line=None if is_act else EDGE,lw=12700,preset='roundRect',text=badge,size=10,bold=True,color=BG if is_act else MUTED,font=MF)
    tb(s,xx+pad,cyy+620000,cw_-2*pad,title_c,size=23,bold=True,color=INK if is_act else MUTED,font=HF,cy=440000)
    tb(s,xx+pad,cyy+1080000,cw_-2*pad,sub_c,size=14.5,bold=True,color=CY,cy=340000)
    rect(s,xx+pad,cyy+1480000,cw_-2*pad,14000,fill=EDGE)
    tb(s,xx+pad,cyy+1600000,cw_-2*pad,None,paras=[{'text':b,'spaceAfterPt':8} for b in bullets],size=15,color=INK if is_act else MUTED,cy=1360000)
    rect(s,xx+pad,cyy+ch-560000,cw_-2*pad,360000,fill=AMB if is_act else CODEBG,line=None if is_act else EDGE,lw=12700,preset='roundRect',text=risk_lbl,size=11.5,bold=True,color=BG if is_act else MUTED,font=MF)
medallion(s,W//2,cyy+ch//2,diam=460000,glyph='→',fill=CY,size=18)
banner_box(s,M,5420000,CW,620000,'Agents execute code they did not write beforehand — and cannot fully predict at runtime.',size=16.5)
footer(s,'',3)
note(s,'先看最本质的变化。过去我们用聊天助手，它只负责生成文字，中间永远隔着一个人——命令安不安全，你扫一眼再决定跑不跑，最坏也就是答非所问。但自主 Agent 把中间这道人工把关拿掉了，它会自己翻代码库、装依赖、跑脚本、改配置。\n更棘手的是，它下一步要敲什么命令，往往是运行时根据上下文临时写出来的，你根本没法提前把所有命令都审一遍。如果让它直接顶着你本人的账号权限裸跑，那你能打开的所有文件、保存的所有凭证、能连上的所有内网服务，其实都暴露在它的一次误判或者一次外部注入之下。')

# ---------- 3 THREATS (image right)
s=3
img(s,'hd-chaos.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a hooded villain slips a poisoned note to a robot agent while files fly out and a drone steals a key')
cxl=W-PANEL-2*M
header(s,'01 · WHY SANDBOXING','What can go wrong?',cx=cxl)
rows=[('Prompt injection','A poisoned README or web page tricks the agent into running an attacker\'s command.'),('Credential theft','Generated code reads cloud tokens, SSH keys and saved browser secrets.'),('Data exfiltration','Sensitive local files quietly leave over the network to an unknown host.'),('Overreach','A "helpful" shortcut edits production config it was never meant to touch.')]
y=1560000; rh_t=1020000; gap_t=115000
for i,(h_,d) in enumerate(rows):
    rect(s,M,y,cxl,rh_t,fill=CARD,line=EDGE,lw=12700)
    rect(s,M+200000,y+rh_t//2-210000,540000,420000,fill=AMB,preset='roundRect',text=f'0{i+1}',size=14,bold=True,color=BG,font=MF)
    tb(s,M+900000,y+90000,cxl-1100000,None,paras=[{'text':h_,'sizePt':17.5,'bold':True,'color':INK,'spaceAfterPt':4},{'text':d,'sizePt':13.6,'color':MUTED}],cy=rh_t-180000)
    y+=rh_t+gap_t
footer(s,'',4,cx=cxl)
note(s,'具体会出什么乱子呢？最典型的就是这四种情况。比如 Agent 去读一个开源仓库或者网页，里面藏了一句恶意指令，哄着它顺手跑一段外部脚本，这就是提示词注入。再比如它写出来的代码去翻你本机的 SSH 私钥、云平台 Token 或者浏览器里存的登录态，然后悄悄发往外部地址，几秒钟就完成了凭证窃取和数据外传。\n还有一种情况连攻击者都不需要，纯粹是模型太想帮你把事办成了——比如你让它修个网页报错，它觉得直接改线上生产配置最快，结果好心办了坏事。不管是被人下套，还是模型自己抄了近道，只要底下没有硬边界拦着，后果是一样的。\n\n参考来源：第四个越权案例改写自 Windows Developer Blog（2026 年 10 月 7 日）。')

# ---------- 4 PROMPTS PERSUADE
s=4
header(s,'01 · WHY SANDBOXING','Prompts persuade. The OS enforces.')
rect(s,M,1540000,CW,1240000,fill=CODEBG,line=EDGE,lw=12700)
rect(s,M+95000,1540000+140000,46000,1240000-280000,fill=CY,preset='roundRect')
tb(s,M+280000,1620000,CW-520000,None,paras=[
    {'text':'"An agent cannot be its own security authority."','sizePt':26,'bold':True,'color':INK,'spaceAfterPt':6},
    {'text':'Windows Developer Blog  ·  October 7, 2026','sizePt':13,'color':MUTED},
],cy=1080000)
lab=2350000; chw=(CW-lab-3*120000)/4; gap=120000
def chiprow(yv,label,sub_lbl,chips,lifted):
    rect(s,M,yv,lab-140000,720000,fill=CARD2 if lifted else CARD,line=CY if lifted else EDGE,lw=19050 if lifted else 12700)
    tb(s,M+160000,yv+70000,lab-460000,None,paras=[
        {'text':label,'sizePt':14.5,'bold':True,'color':INK if lifted else MUTED,'spaceAfterPt':2},
        {'text':sub_lbl,'sizePt':11,'bold':lifted,'color':CY if lifted else MUTED},
    ],cy=580000)
    for i,c in enumerate(chips):
        x=M+lab+i*(chw+gap)
        rect(s,x,yv,chw,720000,fill=CARD2 if lifted else CARD,line=CY if lifted else EDGE,lw=19050 if lifted else 12700,text=c,size=14.5,bold=lifted,color=INK if lifted else MUTED)
chiprow(3040000,'AI-level','Persuasion only',['System prompts','Content filters','Tool rules','Output checks'],False)
chiprow(3940000,'Execution-level','Kernel-enforced',['File allow / deny','Network rules','Process isolation','Secrets outside'],True)
banner_box(s,M,4960000,CW,660000,'Defense in depth: the model decides what to try; the operating system decides what is possible.',size=16)
footer(s,'',5)
note(s,'所以做 Agent 架构一定要想明白一句话：不能让 Agent 自己当自己的安全管理员。你在系统提示词里写十遍“绝对不准碰密钥、不准连外网”，再叠加上下文过滤和工具调用规则，这些当然要做，但它们本质上都是在劝模型守规矩，只能降低出错的概率。\n一旦模型被绕晕了，或者它调起的脚本行为超出了预期，提示词是拦不住系统调用的。真正能兜底的只有操作系统：哪个目录能读、哪个端口能通、能不能碰剪贴板，必须由系统内核强制拍板。模型决定它想尝试什么，操作系统决定它实际上能做成什么，这两层缺一不可。')

# ---------- 5 THREE PILLARS
s=5
header(s,'01 · WHY SANDBOXING · WINDOWS PLATFORM','Three Windows pillars for governed agents')
cw3=(CW-2*220000)/3; y0=1520000; ch=3680000
pill=[
    ('Containment','Limit what an agent can access & do',
     'OS-enforced file, network, process & UI boundaries on Windows 11 and Windows 365 via open-source MXC (also macOS & Linux).',
     'THIS SESSION · GA'),
    ('Identity','Tell agent activity apart from a person',
     'Windows connects local on-device agents to Microsoft Entra Agent ID so actions carry an agent audit trail instead of a human login.',
     'GA + Windows preview'),
    ('Manageability','Govern access and monitor agents',
     'Windows extends Microsoft Agent 365 and Intune to endpoints to inventory local agents and distribute MXC sandbox policies.',
     'GA + Windows preview'),
]
for i,(t,sub,body,state) in enumerate(pill):
    x=M+i*(cw3+220000)
    is_main=(i==0)
    rect(s,x,y0,cw3,ch,fill=CARD2 if is_main else CARD,line=CY if is_main else EDGE,lw=19050 if is_main else 12700)
    pw=2050000 if is_main else 2260000
    rect(s,x+240000,y0+220000,pw,310000,fill=CY if is_main else CODEBG,line=None if is_main else EDGE,lw=12700,preset='roundRect',text=state,size=10.5,color=BG if is_main else CY,bold=True,font=MF)
    tb(s,x+240000,y0+660000,cw3-480000,t,size=23,bold=True,font=HF,color=INK,cy=460000)
    tb(s,x+240000,y0+1140000,cw3-480000,sub,size=14,bold=True,color=CY,cy=440000)
    rect(s,x+240000,y0+1660000,cw3-480000,14000,fill=EDGE)
    tb(s,x+240000,y0+1780000,cw3-480000,body,size=13.8,color=INK if is_main else MUTED,cy=1760000)
banner_box(s,M,5380000,CW,620000,'Windows OS primitives wired into Microsoft’s enterprise plane: local MXC containment + Entra Agent ID + Agent 365 & Intune.',size=14.8)
footer(s,'',6)
note(s,'放到企业级 Windows 平台上，这三大支柱（隔离 Containment、身份 Identity、可管理性 Manageability）是 Windows 团队明确定义的三大操作系统级底座能力——但它不是把 Windows 做成一座孤岛，而是把本机执行层和微软整体的企业治理平面打通：本机隔离靠我们今天重点讲的开源 MXC，限制它能碰哪些本地资源；身份这一层把微软全局的 Entra Agent ID 延伸到 Windows 本机进程上，别让本地 Agent 借用员工账号；管理这一层则把微软全局的 Agent 365 控制平面和 Intune 终端管理下沉到 Windows 设备上，统一做盘点、审计和沙箱策略下发。\n大家排落地计划的时候也可以留意一下进度：MXC 本身、Windows 365 云电脑上的 MXC，还有云端的 Entra Agent ID 和 Agent 365 都已经正式发布了；而给本机跑的 GitHub Copilot CLI、Claude Code 自动绑定身份，以及通过 Intune 直接给本地 MXC 容器下发文件网络规则，目前还在 Windows 预览阶段。\n\n参考来源：Windows Developer Blog — Logan Iyer（2026 年 10 月 7 日）及《Enterprise Deployment of Governed Agents on Windows》（2026 年 11 月 17 日）。')

# ---------- 6 MEET MXC (image left)
s=6
img(s,'hd-guard.jpg',0,0,PANEL,H,'Hand-drawn sketch panel: a robot agent works inside a force-field box while a guardian with a rulebook decides which items pass in')
xl=PANEL+M; cxr=W-xl-M
header(s,'02 · UNDERSTANDING MXC','Meet MXC',x=xl,cx=cxr)
tb(s,xl,1500000,cxr,'An open-source, policy-driven execution layer for untrusted or AI-generated workloads.',size=17.5,color=INK,cy=560000)
tw=(cxr-180000)/2; th=1180000; ty=2220000
facts=[('MIT','Open source on GitHub'),('GA','Generally available since Oct 7, 2026'),('3 SDKs','.NET, TypeScript & Rust'),('3 OSes','Windows, macOS & Linux')]
for i,(v,l) in enumerate(facts):
    x=xl+(i%2)*(tw+180000); yy=ty+(i//2)*(th+150000)
    rect(s,x,yy,tw,th,fill=CARD,line=EDGE,lw=12700)
    rect(s,x+220000,yy+170000,440000,32000,fill=CY,preset='roundRect')
    tb(s,x+220000,yy+240000,tw-380000,None,paras=[{'text':v,'sizePt':27,'bold':True,'color':CY,'spaceAfterPt':4},{'text':l,'sizePt':13,'color':INK}],font=BF,cy=th-280000)
banner_box(s,xl,5080000,cxr,680000,'One JSON policy, mapped to the right OS sandbox on every host.',size=15.5)
footer(s,'',7,x=xl,cx=cxr)
note(s,'那 MXC 到底是什么？简单讲，它是微软用 MIT 协议开源的一套轻量级沙箱执行层，今年十月已经正式发布。不管你要跑的是大模型现写的代码、第三方插件，还是来路不明的脚本，宿主程序只需要用同一份 JSON 策略说清楚允许读哪儿、写哪儿、连哪儿，MXC 就会自动调用当前操作系统原生的隔离机制去执行。\n注意它虽然名字叫 Execution Containers，但不是让你去装一套笨重的 Docker 镜像——它直接复用 Windows、macOS 和 Linux 内核自带的轻量沙箱能力，官方同时维护了 .NET、TypeScript 和 Rust 三个语言的 SDK。')

# ---------- 7 FIVE CONCEPTS
s=7
header(s,'02 · UNDERSTANDING MXC','Five concepts to remember')
gap=145000; cw5=(CW-4*gap)/5; y0=1560000; ch=3680000
cons=[('Host app','Your trusted app or agent runtime. Picks permissions, supplies inputs, handles results.'),('Workload','The command, script or process that runs under containment.'),('Policy','The access you permit: readable files, writable folders, network destinations, UI.'),('Container request','Combines the policy with the command, working directory, timeout and sandbox choice.'),('Backend','The OS mechanism that enforces it: process sandbox, separate session or VM.')]
for i,(t,d) in enumerate(cons):
    x=M+i*(cw5+gap)
    rect(s,x,y0,cw5,ch,fill=CARD,line=EDGE,lw=12700)
    rect(s,x+180000,y0+220000,620000,290000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=f'0{i+1}',size=12,bold=True,color=CY,font=MF)
    tb(s,x+180000,y0+660000,cw5-360000,t,size=17.5,bold=True,color=INK,cy=680000)
    rect(s,x+180000,y0+1420000,cw5-360000,14000,fill=EDGE)
    tb(s,x+180000,y0+1540000,cw5-360000,d,size=13.5,color=MUTED,cy=1960000)
banner_box(s,M,5420000,CW,600000,'The host decides. The policy describes. The backend enforces.',size=16.5)
footer(s,'',8)
note(s,'看 MXC 的接口之前，把这五个概念对上号就一目了然了。最外层是你自己写的宿主程序 Host，它是唯一可信的调度方，负责定权限、喂输入、收结果；真正被关进沙箱里跑的那条命令或脚本叫 Workload；允许读哪些文件、连哪些地址，写下来就是 Policy；把 Policy 跟具体的命令行、工作目录、超时和隔离方式打包到一块儿，就是一次完整的执行配置；最后在底层真正把隔离墙立起来的系统机制，就是 Backend。\n整套设计里最核心的规矩就是底下这句话：权限只能由 Host 或真人来拍板。Agent 可以提申请说它想读哪个目录，但绝不能让沙箱里的代码自己给自己扩权。')

# ---------- 8 FOUR LAYERS
s=8
header(s,'02 · UNDERSTANDING MXC','Four layers, one boundary')
lays=[
    ('01 · TRUSTED','Host app / agent runtime','Classifies the task, chooses trusted policy defaults, and validates outputs & approvals.',True),
    ('02 · CONTRACT','MXC SDK & policy contract','Discovers platform support, serializes versioned policy config, and manages lifecycle.',False),
    ('03 · RUNTIME','MXC execution layer','Validates the config, resolves a compatible backend, and maps to OS operations.',False),
    ('04 · KERNEL','Containment backend','Creates the isolated environment and enforces file, network & session boundaries.',False),
]
y=1520000; lh=760000; lw_=3650000
for i,(tag_l,t,d,trusted) in enumerate(lays):
    rect(s,M,y,lw_,lh,fill=CARD2 if trusted else CARD,line=CY if trusted else EDGE,lw=19050 if trusted else 12700)
    tb(s,M+220000,y+70000,lw_-400000,None,paras=[
        {'text':tag_l,'sizePt':10,'bold':True,'color':CY,'spaceAfterPt':2},
        {'text':t,'sizePt':16,'bold':True,'color':INK},
    ],cy=lh-140000)
    rect(s,M+lw_+110000,y,CW-lw_-110000,lh,fill=CARD2 if trusted else CARD,line=CY if trusted else EDGE,lw=19050 if trusted else 12700)
    tb(s,M+lw_+340000,y,CW-lw_-560000,d,size=14.8,color=INK if trusted else MUTED,cy=lh)
    if i==0:
        by_b=y+lh+55000
        rect(s,M,by_b+95000,CW,22000,fill=AMB)
        rect(s,W//2-1550000,by_b,3100000,210000,fill=AMB,preset='roundRect',text='TRUST BOUNDARY  ·  UNTRUSTED BELOW',size=9.8,bold=True,color=BG,font=MF)
        y+=lh+320000
    else:
        y+=lh+95000
banner_box(s,M,5420000,CW,600000,'Rule #1: an agent may suggest a permission, but only trusted host logic or a human can grant it.',size=15.5)
footer(s,'',9)
note(s,'把刚才这五个角色按上下层级叠起来，就是 MXC 的四层架构：从上往下依次是宿主应用、SDK 与策略契约、MXC 执行层，以及最底层的系统隔离后端。真正的信任分界线，就画在宿主应用和底下被隔离的执行环境之间。\n我特别想提醒一点：很多团队接完沙箱容易产生错觉，觉得代码既然在沙箱里跑，拿回来的结果就绝对安全。其实沙箱只管不让里面的进程越界搞破坏，但它打印出来的文本、生成出来的文件里照样可能夹带脏数据或者二次注入。宿主程序站在信任边界外面，收回来的任何结果还是得先校验再往下游传。')

# ---------- 9 REQUEST FLOW
s=9
header(s,'02 · UNDERSTANDING MXC','How one request flows')
steps=[('HOST','1 · Discover','Ask the SDK which backends this machine supports.'),('HOST','2 · Describe','Build a policy & execution request, deny by default.'),('SDK','3 · Launch','Start the contained workload through the SDK.'),('MXC','4 · Translate','Map the request to native OS sandbox primitives.'),('HOST','5 · Return','Validate output, exit code, denials & clean up.')]
gap_rf=210000; sw=(CW-4*gap_rf)/5; yc=1620000; ch_rf=3200000
for i,(who,t,d) in enumerate(steps):
    x=M+i*(sw+gap_rf)
    rect(s,x,yc,sw,ch_rf,fill=CARD,line=EDGE,lw=12700)
    rect(s,x+170000,yc+210000,820000,270000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=who,size=10.5,bold=True,color=CY,font=MF)
    tb(s,x+170000,yc+640000,sw-340000,t,size=17.5,bold=True,color=INK,font=HF,cy=440000)
    rect(s,x+170000,yc+1160000,sw-340000,14000,fill=EDGE)
    tb(s,x+170000,yc+1300000,sw-340000,d,size=13.8,color=MUTED,cy=1660000)
    if i<4:
        medallion(s,x+sw+gap_rf//2,yc+ch_rf//2,diam=280000,glyph='→',fill=CY,size=12)
banner_box(s,M,5160000,CW,780000,'Backends have different security properties. Always test both the allowed and the denied operations.',size=16)
footer(s,'',10)
note(s,'看清了静态分层，我们再看一次真实请求是怎么穿过这道边界的。Host 先问一下 SDK 当前这台机器支持哪些沙箱后端，然后按默认拒绝的原则拼好策略和请求交给 SDK；MXC 在中间把这份统一请求翻译成对应操作系统的底层调用，等任务跑完，再把标准输出、退出码、拦截记录收集回来，顺手把临时环境清理干净。\n这里有个特别实在的工程经验：不同平台底层沙箱的脾气不完全一样，大家写测试的时候千万别只测正常路径能不能跑通，一定要专门写反向测试，去验证不该读的敏感文件是不是真被拦住了、没放行的网络是不是真断了。那底层到底有哪些沙箱可选？我们接着往下看。')

# ---------- 10 PICK THE RIGHT BOX
s=10
header(s,'03 · FEATURES & PLATFORMS','Pick the right box')

# Clean spectrum bar: left pill ───▸ right pill (zero stretched arrow polygon or text collision)
spec_y=1470000; spec_h=380000; pill_w=3250000
rect(s,M,spec_y,pill_w,spec_h,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',
     text='←  LIGHTER  ·  FASTER STARTUP',size=10.8,bold=True,color=CY,font=MF)
rect(s,W-M-pill_w,spec_y,pill_w,spec_h,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',
     text='STRONGER KERNEL ISOLATION  →',size=10.8,bold=True,color=CY,font=MF)
line_x1=M+pill_w+160000; line_x2=W-M-pill_w-160000; line_w=line_x2-line_x1; mid_y=spec_y+spec_h//2
seg_gap=90000; seg_w=(line_w-2*seg_gap)//3
for seg_i in range(3):
    sx=line_x1+seg_i*(seg_w+seg_gap)
    rect(s,sx,mid_y-20000,seg_w,40000,fill=CY,preset='roundRect')

gap=180000; cw4=(CW-3*gap)/4; y0=2040000; ch=3960000
bx=[
    ('Process container','Windows 11 · macOS · Linux',
     'AppContainer, Seatbelt or Bubblewrap. Millisecond startup and fine-grained file & network rules — default for coding agents.',
     'Shared kernel · Process ACL','GA · Default',CY,1),
    ('Session container','Windows 11 only',
     'Separate local Windows login and desktop for GUI apps. Protects your screen & clipboard and blocks synthetic input.',
     'Shared kernel · Own desktop','GA · GUI',CY,2),
    ('WSL container','Windows 11',
     'Runs Linux CLI workloads, compilers and toolchains on a Windows host inside an isolated WSL environment.',
     'Isolated Linux subsystem','GA · Linux',CY,3),
    ('MicroVM','Windows 11 · Linux',
     'Hardware-virtualized unikernel boundary with ~90 ms warm snapshot restore; never shares the host OS kernel.',
     'Dedicated VM kernel','Experimental',AMB,4),
]
for i,(t,p,d,boundary,badge,accent_col,level) in enumerate(bx):
    x=M+i*(cw4+gap)
    is_default=(level==1)
    is_exp=(level==4)
    rect(s,x,y0,cw4,ch,fill=CARD,line=EDGE,lw=12700)
    pad=210000
    meter_y=y0+230000; bar_w=175000; bar_g=48000
    for m_idx in range(4):
        bx_m=x+pad+m_idx*(bar_w+bar_g)
        active=m_idx<level
        rect(s,bx_m,meter_y+55000,bar_w,95000,
             fill=CY if active else EDGE,preset='roundRect')
    pill_bw=1180000
    pill_fill=CY if is_default else (AMB if is_exp else CODEBG)
    pill_line=None if (is_default or is_exp) else EDGE
    pill_col=BG if (is_default or is_exp) else CY
    rect(s,x+cw4-pad-pill_bw,meter_y,pill_bw,210000,
         fill=pill_fill,line=pill_line,lw=12700,preset='roundRect',
         text=badge,size=9.8,bold=True,color=pill_col,font=MF)
    tb(s,x+pad,y0+560000,cw4-2*pad,t,size=18.5,bold=True,color=INK,font=HF,cy=420000)
    tb(s,x+pad,y0+1010000,cw4-2*pad,p,size=12,color=CY,bold=True,cy=340000)
    rect(s,x+pad,y0+1430000,cw4-2*pad,14000,fill=EDGE)
    tb(s,x+pad,y0+1540000,cw4-2*pad,d,size=13,color=MUTED,cy=1720000)
    rect(s,x+pad,y0+ch-560000,cw4-2*pad,360000,fill=CODEBG,line=EDGE,lw=12700)
    tb(s,x+pad+60000,y0+ch-560000,cw4-2*pad-120000,boundary,size=10.8,bold=True,color=CY,font=MF,align='center',cy=360000)
footer(s,'',11)
note(s,'MXC 提供了四种不同分量的隔离盒子，怎么选就看你的业务场景。平时做代码生成、跑测试、处理本地文档，首选最左边的 Process Container，它在 Windows、macOS、Linux 上分别映射成 AppContainer、Seatbelt 和 Bubblewrap，毫秒级启动，文件和网络权限都能卡得很细。\n但如果你的 Agent 需要操作传统桌面软件、得有自己的窗口和剪贴板，那就选 Windows 专属的 Session Container。它会建一个独立的本地账号和登录会话，跟员工正在用的桌面完全隔开，还封死了模拟键鼠输入；不过它不做细粒度的文件和网络限制。再往右，要在 Windows 上跑 Linux 工具链可以用 WSL 容器；要跑高风险的陌生代码、连操作系统内核都不想共享，那就上最右边实验性的硬件级 MicroVM。')

# ---------- 11 WHERE MXC RUNS
s=11
header(s,'03 · FEATURES & PLATFORMS','Where MXC runs today')
c1=M; w1=2850000; c2=c1+w1+150000; w2=3050000; c3=c2+w2+150000; w3=W-M-c3
hy=1500000
rect(s,M,hy,CW,420000,fill=CODEBG,line=EDGE,lw=12700)
for x,w,t in [(c1,w1,'HOST PLATFORM'),(c2,w2,'DEFAULT PROCESS BACKEND'),(c3,w3,'ADDITIONAL OPTIONS & CLOUD PC')]:
    tb(s,x+200000,hy,w-240000,t,size=10.8,color=CY,font=MF,bold=True,ls=0.8,cy=420000)
rowsd=[('Windows 11','ProcessContainer','Also supports IsolationSession, WSLc & Learning Mode reports'),('Linux','Bubblewrap','Also supports LXC containers'),('macOS','Seatbelt','Enforces native macOS sandbox profiles'),('Windows 365 Cloud PC','ProcessContainer','Works out of the box; session or VM backends need nested virtualization')]
y=2020000; rh=740000
for i,(a,b,c) in enumerate(rowsd):
    rect(s,c1,y,CW,rh,fill=CARD,line=EDGE,lw=12700)
    tb(s,c1+200000,y,w1-220000,a,size=16.5,bold=True,color=INK,cy=rh)
    rect(s,c2+180000,y+rh//2-180000,2350000,360000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=b,size=13.5,bold=True,color=CY,font=MF)
    tb(s,c3+200000,y,w3-300000,c,size=14.2,color=MUTED,cy=rh)
    y+=rh+95000
banner_box(s,M,5440000,CW,580000,'Experimental (explicit opt-in via 1.1-alpha): Windows Sandbox, MicroVM & Hyperlight unikernel.',accent=AMB,size=14.8)
footer(s,'',12)
note(s,'做跨平台应用时，统一 SDK 省去了重写业务代码的麻烦，但每个平台底层落到哪儿我们得心里有数。Windows 11 默认走 ProcessContainer，同时支持独立会话、WSL 容器和后面要讲的学习模式；Linux 默认用 Bubblewrap，macOS 用 Seatbelt。如果团队平时在 Windows 365 云电脑上干活，进程级隔离跟物理机完全一致，连嵌套虚拟化都不用开；只有用到会话或虚拟化后端时才需要额外确认环境。\n至于最底下的 Windows Sandbox、MicroVM 和 Hyperlight，目前还属于实验特性，需要显式打开实验开关才能用。所以代码里最稳妥的做法，永远是在启动前先查一次平台支持情况，别硬编码死某一种后端。')

# ---------- 12 LIFECYCLE
s=12
header(s,'03 · FEATURES & PLATFORMS','One-shot or state-aware')
def lane(y,label,sub,stepsL,desc,is_default):
    rect(s,M,y,CW,1300000,fill=CARD2 if is_default else CARD,line=CY if is_default else EDGE,lw=19050 if is_default else 12700)
    lw0=2350000
    tb(s,M+220000,y+120000,lw0-260000,None,paras=[
        {'text':label,'sizePt':19,'bold':True,'color':INK,'spaceAfterPt':4},
        {'text':sub,'sizePt':12.5,'bold':True,'color':CY},
    ],cy=1060000)
    rect(s,M+lw0,y+180000,14000,940000,fill=EDGE)
    ax=M+lw0+200000; aw=CW-lw0-420000
    n=len(stepsL); g=220000; bw=(aw-(n-1)*g)/n
    for i,t in enumerate(stepsL):
        x=M+lw0+200000+i*(bw+g)
        rect(s,x,y+190000,bw,480000,fill=CARD if is_default else CODEBG,line=EDGE,lw=12700,preset='roundRect',text=t,size=13.5,bold=True,color=INK)
        if i<n-1:
            medallion(s,x+bw+g//2,y+430000,diam=170000,glyph='→',fill=CY,size=9.5)
    tb(s,ax,y+760000,aw,desc,size=13.8,color=INK if is_default else MUTED,cy=420000)
lane(1520000,'One-shot','DEFAULT CHOICE',['Create','Run one workload','Tear down'],'Builds, tests & generated commands: spin up a fresh, narrow sandbox for every action.',True)
lane(2960000,'State-aware','PERSISTENT STATE',['Provision','Start','Execute','Stop','Deprovision'],'Interactive shells & multi-step triage via IsolationSession or WSLc — always deprovision.',False)
ry=4440000; g=160000; cw3=(CW-2*g)/3
for i,(t,api_lbl,d) in enumerate([('run()','CAPTURED RESULT','Returns stdout, stderr & exit code once the workload finishes.'),('spawn()','LIVE STREAMING','Streams stdin, stdout and stderr in real time while running.'),('PTY','INTERACTIVE SHELL','Full pseudo-terminal session for stateful shells & human handoff.')]):
    x=M+i*(cw3+g)
    rect(s,x,ry,cw3,1560000,fill=CARD,line=EDGE,lw=12700)
    rect(s,x+220000,ry+190000,1580000,260000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=api_lbl,size=9.5,bold=True,color=CY,font=MF)
    tb(s,x+220000,ry+530000,cw3-440000,None,paras=[
        {'text':t,'sizePt':18.5,'bold':True,'color':INK,'spaceAfterPt':5},
        {'text':d,'sizePt':13.5,'color':MUTED},
    ],cy=900000)
footer(s,'',13)
note(s,'选好了跑哪种容器、在哪套系统上跑，下一步就是决定这个容器该活多久。MXC 支持用完即焚的 One-shot 和保留现场的 State-aware 两种生命周期。日常百分之八九十的动作，像跑一次编译、执行一遍测试、生成一份报告，都应该优先选 One-shot：每次调用现拉一个全新的窄权限沙箱，干完立刻销毁，不给潜在问题留任何驻留空间。\n只有当你确实需要跨多步保留上下文，比如预装好一套复杂环境，或者开着终端连续排查故障时，才用 IsolationSession 或 WSLc 这种有状态容器。用的时候一定记得，停止运行 stop 不代表释放资源，不管中间成没成功，最后都要调 deprovision 把临时账号和环境清理干净。底下的三种执行接口也很清晰：等完整结果用 run，盯实时输出流用 spawn，做交互式命令行终端就用 PTY。')

# ---------- 13 FILES
s=13
header(s,'03 · FEATURES & PLATFORMS','Files: read in, write out')
cx_=5850000; cy0=1540000; chh=4460000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+240000,cy0+120000,cx_-480000,'policy.json  ·  filesystem guardrails',size=11.5,color=FOOT,font=MF,cy=300000)
rect(s,M+240000,cy0+450000,cx_-480000,14000,fill=EDGE)
code='{\n  "version": "1.0.0",\n  "containment": "processcontainer",\n  "process": {\n    "commandLine": "python summarize.py"\n  },\n  "filesystem": {\n    "readonlyPaths": ["C:\\\\demo\\\\input"],\n    "readwritePaths": ["C:\\\\demo\\\\output"],\n    "deniedPaths": ["C:\\\\Users\\\\me\\\\.ssh"]\n  }\n}'
tb(s,M+240000,cy0+540000,cx_-480000,code,size=13,color=CODE,font=MF,cy=3700000)
rx=M+cx_+280000; rw=W-M-rx
rules=[('01 · Read in, write out','Separate readable input folders from one dedicated, empty writable output directory.',CY),('02 · Deny the crown jewels','Explicitly block secrets, ~/.ssh, cloud credentials, system and user profile paths.',AMB),('03 · Teardown is not undo','Writes to granted host folders survive container exit — treat output folders as a staging zone.',CY)]
for i,(h_,d,col_) in enumerate(rules):
    yy=1540000+i*1540000
    card_box(s,rx,yy,rw,1380000,[
        {'text':h_,'sizePt':17,'bold':True,'color':col_,'spaceAfterPt':7},
        {'text':d,'sizePt':13.8,'color':MUTED},
    ],fill=CARD,line=EDGE,lw=12700,pad_x=220000,pad_y=110000)
footer(s,'',14)
note(s,'容器的形态和生命周期定下来之后，我们来看里面的安全策略具体管哪三件事——首先是文件。核心思路就六个字：读入、写出分离。给 Agent 看的参考资料一律挂成只读，再单独建一个空的专用目录让它写产出物；同时用 deniedPaths 把 .ssh、各类密钥和用户隐私目录明确拉黑，千万别图省事把整个用户主目录直接放开。\n还有个细节特别容易踩坑：销毁沙箱不等于撤销落盘。容器退出时清掉的是它自己的临时运行环境，只要你允许它往宿主机的输出目录写了文件，那些改动在容器销毁后是实打实留下的。所以一定要把输出目录当成待检区，宿主程序先验完格式和内容，再拿去正式使用。\n\n参考来源：github.com/microsoft/mxc 官方策略示例（docs/examples.md）。')

# ---------- 14 NETWORK
s=14
header(s,'03 · FEATURES & PLATFORMS','Network: deny by default')
bx0=M; by0=1620000; bw=2550000; bh=4180000
rect(s,bx0,by0,bw,bh,fill=CARD2,line=CY,lw=19050)
rect(s,bx0+200000,by0+260000,bw-400000,290000,fill=CY,preset='roundRect',text='SANDBOXED WORKLOAD',size=9.8,bold=True,color=BG,font=MF)
tb(s,bx0+200000,by0+760000,bw-400000,'Agent process',size=20,bold=True,color=INK,align='center',cy=480000)
rect(s,bx0+220000,by0+1420000,bw-440000,520000,fill=CARD,line=EDGE,lw=12700,preset='roundRect',text='egress: "deny"',size=12.5,bold=True,color=CY,font=MF)
rect(s,bx0+220000,by0+2080000,bw-440000,520000,fill=CARD,line=EDGE,lw=12700,preset='roundRect',text='ingress: "deny"',size=12.5,bold=True,color=CY,font=MF)
tb(s,bx0+180000,by0+2840000,bw-360000,'Only explicit rules punch a narrow hole.',size=13,color=MUTED,align='center',cy=1000000)

conn_w=880000; dx=bx0+bw+conn_w; dw=3100000; dh=1160000; dgap=350000
dests=[('140.82.114.6 : 443','Explicitly allowed by IP/CIDR + port',CY,'✓'),('Any other address','Blocked at the OS firewall layer',AMB,'✕'),('127.0.0.1 : 8080 proxy','Domain allow-list & audit live here',CY,'→')]
for i,(t,d,c,g) in enumerate(dests):
    yy=by0+i*(dh+dgap)
    cy_mid=yy+dh//2
    rect(s,bx0+bw,cy_mid-14000,conn_w,28000,fill=c)
    medallion(s,bx0+bw+conn_w//2,cy_mid,diam=300000,glyph=g,fill=c,size=13)
    rect(s,dx,yy,dw,dh,fill=CARD,line=EDGE,lw=12700)
    tb(s,dx+200000,yy+110000,dw-360000,None,paras=[
        {'text':t,'sizePt':15,'bold':True,'color':INK,'spaceAfterPt':5},
        {'text':d,'sizePt':12.8,'bold':True,'color':c},
    ],cy=dh-220000)
fx=dx+dw+260000; fw=W-M-fx
card_box(s,fx,by0,fw,bh,[
    {'text':'NETWORK GUARDRAIL RULES','sizePt':10.5,'bold':True,'color':CY,'spaceAfterPt':10},
    {'text':'• Directional rules: separate defaults for egress, ingress & loopback.','sizePt':13.8,'color':INK,'spaceAfterPt':10},
    {'text':'• OS firewall filters IP/CIDR, protocol & port — not hostnames.','sizePt':13.8,'color':INK,'spaceAfterPt':10},
    {'text':'• Need domain filtering? Route via runtimeConfig.networkProxy.','sizePt':13.8,'color':INK,'spaceAfterPt':10},
    {'text':'• Lock proxy traffic to your host PID with allowedProxyPeer.','sizePt':13.8,'bold':True,'color':CY},
],fill=CARD,line=EDGE,lw=12700,pad_x=220000,pad_y=160000)
footer(s,'',15)
note(s,'光锁住文件还不够，网络出口要是不管，Agent 读到的任何一点东西都能顺着网线发走。MXC 把网络拆成出站、入站和本机回环三个方向，上手第一步就是把默认规则全设成拒绝。\n这里要注意，操作系统底层的防火墙规则认的是 IP、网段和端口号，它没法直接认域名。如果业务上必须按域名放行，比如只允许连公司的内部网关或者指定的包管理源，标准做法是在宿主机起一个受控的本地代理，通过 networkProxy 把沙箱流量引过去，由代理来做域名白名单和日志审计；在生产环境里还可以加上 allowedProxyPeer 校验代理进程身份，防止本机其他程序冒名顶替。')

# ---------- 15 KEYS & UI
s=15
header(s,'03 · FEATURES & PLATFORMS','Keys and UI stay outside')
bx0=M; by0=1560000; bw=3450000; bh=4400000
rect(s,bx0,by0,bw,bh,fill=CARD2,line=CY,lw=19050)
rect(s,bx0+220000,by0+220000,bw-440000,300000,fill=CY,preset='roundRect',text='INSIDE THE SANDBOX',size=10.5,bold=True,color=BG,font=MF)
for i,t in enumerate(['✓  Repo files (granted)','✓  Build & test toolchain','✓  Approved API calls']):
    rect(s,bx0+220000,by0+820000+i*1080000,bw-440000,760000,fill=CARD,line=EDGE,lw=12700,text=t,size=15,bold=True,color=INK,align='left')
ox=bx0+bw+220000; ow=3050000
rect(s,ox,by0,ow,bh,fill=CARD,line=AMB,lw=19050)
rect(s,ox+200000,by0+220000,ow-400000,300000,fill=AMB,preset='roundRect',text='OUTSIDE · OUT OF REACH',size=10.5,bold=True,color=BG,font=MF)
for i,t in enumerate(['✕  API tokens & cloud keys','✕  ~/.ssh private keys','✕  Saved browser secrets','✕  Employee clipboard','✕  Host desktop & input']):
    rect(s,ox+200000,by0+720000+i*700000,ow-400000,540000,fill=CODEBG,line=EDGE,lw=12700,text=t,size=13.8,color=INK,align='left')
ex=ox+ow+220000; ew=W-M-ex
for idx_k,(kt,kd,kcol) in enumerate([
    ('Credentials stay on the host','Keep raw keys out of sandbox env & disk. Let your trusted host or proxy attach credentials to approved calls.',CY),
    ('UI & synthetic input blocked','Process sandboxes cannot read employee windows, scrape the clipboard, or inject fake mouse/keystroke events.',CY),
    ('Need a GUI desktop?','Session containers get their own separate Windows desktop & clipboard — use a dedicated bot account.',CY),
]):
    ky=by0+idx_k*1510000
    card_box(s,ex,ky,ew,1380000,[
        {'text':kt,'sizePt':16,'bold':True,'color':kcol,'spaceAfterPt':6},
        {'text':kd,'sizePt':13.2,'color':MUTED},
    ],fill=CARD,line=EDGE,lw=12700,pad_x=200000,pad_y=100000)
footer(s,'',16)
note(s,'文件与网络之外，最敏感的就是密钥和用户的桌面环境了。有一条铁律可以省掉九成的泄密风险：永远别把 API Key、SSH 私钥直接塞进沙箱的环境变量或配置文件里。沙箱里只放干活必需的代码和工具；真要调受保护的外部接口，让沙箱外的宿主程序或者刚刚说的本地代理在审完请求后，在外面替它把凭证签上去。\n用户界面也是一样，普通沙箱碰不到你的窗口和剪贴板，也没法伪造鼠标键盘敲击。就算用了带独立桌面的 Session Container，它的窗口和剪贴板也只属于那个临时账号自己；而且一旦在那个会话里登录了 ERP 等业务系统，里面的 Agent 是能摸到当前登录态的，所以务必给它配专门的低权限机器人账号，别拿员工本人的账号去登。')

# ---------- 16 LEARNING MODE (image right)
s=16
img(s,'hd-learning.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: an owl detective records the footprints of a robot agent walking inside a box')
cxl=W-PANEL-2*M
header(s,'03 · FEATURES & PLATFORMS','Observe first, then enforce',cx=cxl)
modes=[('Enforcement','Applies the policy strictly; blocks any ungranted access.',False),('Learning','Blocks ungranted access AND records every attempt in an activity report.',True),('Permissive','Allows and logs ungranted access — temporary diagnostics only.',False)]
y=1540000; mh=820000
for i,(t,d,hi) in enumerate(modes):
    rect(s,M,y,cxl,mh,fill=CARD2 if hi else CARD,line=CY if hi else EDGE,lw=19050 if hi else 12700)
    rect(s,M+180000,y+mh//2-170000,1780000,340000,fill=CY if hi else CODEBG,line=None if hi else EDGE,lw=12700,preset='roundRect',text=t,size=13.5,bold=True,color=BG if hi else INK)
    tb(s,M+2120000,y,cxl-2300000,d,size=14.2,color=INK if hi else MUTED,cy=mh)
    y+=mh+100000
gap_lm=220000; sw=(cxl-2*gap_lm)/3; py_lm=y+90000
for i,t in enumerate(['01 · Observe','02 · Human Review','03 · Enforce']):
    sx=M+i*(sw+gap_lm)
    rect(s,sx,py_lm,sw,520000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=t,size=13.5,bold=True,color=CY,font=MF)
    if i<2:
        medallion(s,sx+sw+gap_lm//2,py_lm+260000,diam=200000,glyph='→',fill=CY,size=10)
banner_box(s,M,py_lm+680000,cxl,580000,'Learning Mode provides evidence, not auto-grants. A human approves every rule.',size=14.2)
footer(s,'',17,cx=cxl)
note(s,'聊到这三道最小权限护栏，大家落地时肯定会遇到一个头疼问题：跑一个编译工具或者脚本动不动要读几十个系统路径，我一开始哪猜得到该放行哪些目录？全靠手写肯定跑不通。MXC 专门做了一个 Learning Mode 来帮我们探路。\n平时生产上线开的是严格拦截的 Enforcement，排查疑难杂症可以临时开只记日志不拦截的 Permissive；而在策略打磨阶段，最实用的就是 Learning Mode——它碰到没授权的访问照样拦住，但会把每一次被拒的尝试完整记进活动报告里。你可以先拿真实任务跑几轮，拿到这份证据清单，由工程师逐条甄别哪些真是正常依赖、哪些是越界试探，人工拍板加进策略之后，再切回强制模式上线。')

# ---------- 17 CODING AGENT (image left)
s=17
img(s,'hd-coding.jpg',0,0,PANEL,H,'Hand-drawn sketch panel: a robot coding inside a fenced project folder while a locked vault with a key sits outside')
xl=PANEL+M; cxr=W-xl-M
header(s,'04 · USE CASES','Secure coding agents',x=xl,cx=cxr)
tb(s,xl,1480000,cxr,'Build, test and push — without touching anything outside the repo.',size=17,color=INK,cy=440000)
perm=[('Project repo  C:\\src\\app','Read / write',CY),('Toolchain: node, python, git','Allowed',CY),('Package registry via local proxy','Allowed',CY),('~/.ssh, cloud tokens & Documents','Denied',AMB),('Any other network destination','Blocked',AMB)]
y=2020000; rh=560000
for t,st,c in perm:
    rect(s,xl,y,cxr,rh,fill=CARD,line=EDGE,lw=12700)
    tb(s,xl+220000,y,cxr-2150000,t,size=14.8,color=INK,cy=rh)
    rect(s,xl+cxr-1750000,y+95000,1550000,rh-190000,fill=c,preset='roundRect',text=st,size=12.5,bold=True,color=BG,font=MF)
    y+=rh+85000
banner_box(s,xl,5340000,cxr,600000,'In production today: GitHub Copilot CLI enforces command isolation via MXC.',size=14.5)
footer(s,'',18,x=xl,cx=cxr)
note(s,'把前面讲的拼到一起，最直观的例子就是写代码的 Coding Agent。你想让它改代码、跑测试、提分支，仔细算算它真正需要的权限其实非常窄：当前这一个仓库目录可读写，node、python、git 这些工具链可调用，再加上通过本地代理去拉依赖包。至于电脑上别的项目、你的个人文档、~/.ssh 密钥，还有随便连外网的权限，全给它关掉，一点不耽误干活。\n这套模式已经在真实大规模生产里跑通了——大家平时用的 GitHub Copilot CLI，底层就是靠 MXC 的进程隔离来兜底约束模型动态生成和执行的命令。')

# ---------- 18 DESKTOP & DATA (image right)
s=18
img(s,'hd-desktop.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a human relaxes at his own desk while a robot agent processes invoices at its own separate desktop behind a glass wall')
cxl=W-PANEL-2*M
header(s,'04 · USE CASES','Desktop and data agents',cx=cxl)
cards=[('Desktop automation agent','Keeps GUI app state in a separate Windows account and session, isolated from the employee desktop.','Best fit: Session container + dedicated bot login'),('Sensitive-document agent','Reads one approved input folder, writes one summary report, with network completely disabled.','Best fit: Process container (read-in / write-out, egress deny)'),('Agents on Windows 365 Cloud PCs','Runs alongside existing Cloud PC workloads; process isolation works out of the box just like physical PCs.','Best fit: Process container on Windows 365 (GA)')]
y=1540000; ch=1400000
for idx_u,(t,d,f) in enumerate(cards):
    rect(s,M,y,cxl,ch,fill=CARD,line=EDGE,lw=12700)
    tb(s,M+240000,y+110000,cxl-480000,None,paras=[
        {'text':t,'sizePt':17.2,'bold':True,'color':INK,'spaceAfterPt':5},
        {'text':d,'sizePt':13.4,'color':MUTED},
    ],cy=820000)
    rect(s,M+240000,y+ch-420000,cxl-480000,300000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=f,size=11.5,bold=True,color=CY,font=MF)
    y+=ch+110000
footer(s,'',19,cx=cxl)
note(s,'除了研发写代码，企业业务里还有几个场景特别对路。比如操作传统客户端或内网系统的桌面 Agent，给它开一个 Session Container，让它在独立的本地账号和会话里通过 UI Automation 或应用接口跑流程——员工在前台正常办公，Agent 在后台独立桌面里处理单据，两边互不干扰，再配上专门的低权限业务账号就稳妥了。\n再比如审合同、做财报汇总的敏感文档 Agent，直接关进完全断网的 Process Container，只读一个输入目录、只写一个结果目录，物理上断掉外传可能。还有部署在 Windows 365 云电脑上的 Agent，跟员工的云桌面并排运行，进程隔离的用法和体验跟本地电脑完全一致。')

# ---------- 19 ECOSYSTEM
s=19
header(s,'04 · USE CASES','Already in the ecosystem')
ix_e=M; iy_e=1540000; iw_e=7520000; ih_e=3460000
rect(s,ix_e,iy_e,iw_e,ih_e,fill='FFFFFF',line=EDGE,lw=12700)
img(s,'partners-building-on-mxc-grid.png',ix_e+80000,iy_e+80000,iw_e-160000,ih_e-160000,'Partner logo reference: Anthropic, Box, Egnyte, Heidi, LM Studio, manus, Nous, NVIDIA, OpenAI, OpenClaw, Perplexity, Raycast, Replit, Simular and Unsloth',fit='contain')
banner_box(s,ix_e,5180000,iw_e,800000,'Industry convergence: coding CLIs, local model runners & enterprise SaaS agents standardize on OS-native containment.',size=14)
panel_x=ix_e+iw_e+260000; panel_w=W-M-panel_x
status_rows=[
    ('SUPPORTS MXC TODAY','GitHub Copilot · OpenAI Codex\nOpenClaw · Replit\nLM Studio · Unsloth AI',1340000,CY),
    ('INTEGRATED / BUILDING','NVIDIA OpenShell\nOpenAI and Meta',1200000,CY),
    ('SUPPORT ANNOUNCED','Anthropic Claude Code · Box · Egnyte\nHeidi Health · Hermes Agent\nManus · Perplexity · Raycast · Simular',1620000,CY),
]
y=1540000
for title,body,h,col_e in status_rows:
    rect(s,panel_x,y,panel_w,h,fill=CARD,line=EDGE,lw=12700)
    tb(s,panel_x+200000,y+100000,panel_w-400000,title,size=10.8,color=col_e,font=MF,bold=True,ls=0.8,cy=260000)
    rect(s,panel_x+200000,y+390000,panel_w-400000,14000,fill=EDGE)
    tb(s,panel_x+200000,y+440000,panel_w-400000,body,size=12.8,color=INK,cy=h-500000)
    y+=h+140000
footer(s,'',20)
note(s,'而且 MXC 不是微软关起门来自用的组件，整个 AI 工具链生态接得非常快。像 GitHub Copilot、OpenAI Codex、OpenClaw、Replit、LM Studio 还有 Unsloth AI 现在都已经直接支持 MXC 了；NVIDIA 的 OpenShell 已经集成进来，OpenAI 和 Meta 也在基于这套新的隔离机制往下做深；再加上 Anthropic Claude Code、Box、Perplexity、Raycast 等一众厂商也陆续官宣支持。对企业架构师来说是个好消息：不同工具底层正在往同一套操作系统级安全基础设施收敛，不用每引一个工具就重搭一套孤立的沙箱。')

# ---------- 20 REFERENCE ARCHITECTURE
s=20
header(s,'05 · INTEGRATING MXC','Reference architecture')
boxes=[('01 · PLANNER','Agent planner','LLM proposes next action'),('02 · INTENT','Tool call','"run npm test"'),('03 · TRUSTED HOST','Policy engine','Selects least-privilege policy'),('04 · CONTRACT','MXC SDK','Lifecycle & diagnostics'),('05 · OS KERNEL','Sandbox backend','Enforces hard boundary')]
g=240000; bw=(CW-4*g)/5; y0=1560000; bh=1420000
for i,(k_lbl,t,d) in enumerate(boxes):
    x=M+i*(bw+g)
    hi=(i==2)
    rect(s,x,y0,bw,bh,fill=CARD2 if hi else CARD,line=CY if hi else EDGE,lw=19050 if hi else 12700)
    tb(s,x+120000,y0+110000,bw-240000,None,paras=[
        {'text':k_lbl,'sizePt':9.5,'bold':True,'color':CY,'spaceAfterPt':4,'align':'center'},
        {'text':t,'sizePt':15.5,'bold':True,'color':INK,'spaceAfterPt':4,'align':'center'},
        {'text':d,'sizePt':12.2,'color':MUTED,'align':'center'},
    ],align='center',cy=bh-220000)
    if i<4:
        medallion(s,x+bw+g//2,y0+bh//2,diam=220000,glyph='→',fill=CY,size=11)
rect(s,M,3160000,CW,540000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',
     text='←   RETURN PATH  ·  Stdout, exit code & kernel denials validated by trusted host before feeding back to agent',
     size=13,bold=True,color=INK)
g2=170000; cw3=(CW-2*g2)/3; y1=3920000; ch=2040000
side=[('Egress proxy','Attaches credentials and enforces domain allow-lists strictly outside the sandbox.',CY),('Learning Mode','Captures blocked-access evidence so humans can tighten policies before rollout.',CY),('Enterprise governance','Entra Agent ID & Agent 365 govern agent identity, audit trails, and fleet policy.',CY)]
for i,(t,d,col_s) in enumerate(side):
    x=M+i*(cw3+g2)
    card_box(s,x,y1,cw3,ch,[
        {'text':t,'sizePt':17.5,'bold':True,'color':col_s,'spaceAfterPt':8},
        {'text':d,'sizePt':14.2,'color':MUTED},
    ],fill=CARD,line=EDGE,lw=12700,pad_x=240000,pad_y=140000)
footer(s,'',21)
note(s,'看完了典型场景和行业生态，如果我们要把 MXC 接进自家的 Agent 平台，整体该怎么搭？这就是一套可以直接落地的生产参考架构。当大模型规划器决定调工具，比如想执行一句 npm test 时，指令绝不能直通操作系统，而是先交到宿主侧的可信策略引擎手里；宿主按任务类型选好现成的最小权限模板，再通过 MXC SDK 拉起对应的隔离后端去跑。\n跑完往回走的那条路径同样不能省：退出码正不正常、产出格式对不对、中间有没有触发越权拒绝，全由宿主验过关之后再把结果喂回给 Agent。外围再配上管域名和挂凭证的出站代理、帮着收敛权限的 Learning Mode，以及做身份与全局治理的 Entra Agent ID 和 Agent 365，整套防线就严丝合缝了。')

# ---------- 21 SDKs
s=21
header(s,'05 · INTEGRATING MXC','Three SDKs, one contract')
g=220000; cw3=(CW-2*g)/3; y0=1560000; ch=2960000
sdks=[('.NET','Microsoft.Mxc.Sdk','NuGet.org','GetPlatformSupport()'),('TypeScript','@microsoft/mxc-sdk','npm','getPlatformSupport()'),('Rust','mxc-sdk','crates.io','v1::platform_support()')]
for i,(t,pkg,reg,api) in enumerate(sdks):
    x=M+i*(cw3+g)
    rect(s,x,y0,cw3,ch,fill=CARD,line=EDGE,lw=12700)
    pad_s=240000
    rect(s,x+cw3-pad_s-1150000,y0+220000,1150000,280000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=reg,size=11,bold=True,color=CY,font=MF)
    tb(s,x+pad_s,y0+180000,cw3-2*pad_s-1200000,t,size=23,bold=True,font=HF,color=INK,cy=420000)
    tb(s,x+pad_s,y0+700000,cw3-2*pad_s,pkg,size=15,bold=True,color=CY,font=MF,cy=420000)
    rect(s,x+pad_s,y0+1260000,cw3-2*pad_s,14000,fill=EDGE)
    tb(s,x+pad_s,y0+1420000,cw3-2*pad_s,'CAPABILITY DISCOVERY',size=10.5,color=FOOT,font=MF,bold=True,ls=0.8,cy=260000)
    rect(s,x+pad_s,y0+1760000,cw3-2*pad_s,520000,fill=CODEBG,line=EDGE,lw=12700,preset='roundRect',text=api,size=12.5,bold=True,color=CODE,font=MF)
rect(s,M,4740000,CW,1240000,fill=CODEBG,line=EDGE,lw=12700)
rect(s,M+95000,4740000+140000,46000,1240000-280000,fill=CY,preset='roundRect')
tb(s,M+280000,4740000,CW-560000,None,paras=[
    {'text':'Write one ContainerRequest in code — the SDK maps it to the native OS sandbox across Windows, macOS & Linux.','sizePt':16.5,'bold':True,'color':INK,'spaceAfterPt':6},
    {'text':'Requirements: Windows 11 24H2+ (physical or Windows 365), macOS, or Linux  ·  .NET 8+ / Node.js 18+ / Rust.','sizePt':14,'color':MUTED},
],cy=1240000)
footer(s,'',22)
note(s,'聚焦到刚才架构图中间那块 MXC SDK，工程接入其实非常轻。官方在 NuGet、npm 和 crates.io 上同步维护了 .NET、TypeScript 和 Rust 三套 SDK。不管你用哪门语言，写法都是一致的：在代码里填好统一的 ContainerRequest，声明清楚文件、网络、命令和超时，SDK 会自动帮你序列化成稳定的 1.0.0 标准契约，不用你操心底层系统的差异。\n接入时只要养成一个好习惯：启动任务前先调 getPlatformSupport 确认当前机器的支撑情况；如果在 Windows 上跑，还可以顺手调一下 probe，提前确认你配的这份策略能不能拿到预期的原生隔离级别。')

# ---------- 22 FIVE STEPS + CODE
s=22
header(s,'05 · INTEGRATING MXC','Integrate in five steps')
st=[('Discover','Call getPlatformSupport() and probe() before launching.'),('Describe','Build a ContainerRequest: files, network, UI, command & timeout.'),('Learn','Run in Learning Mode first; review the blocked-access report.'),('Enforce','Switch to Enforcement; verify both allowed and denied cases.'),('Operate','Handle timeouts cleanly; surface clear denial reasons to users.')]
lw0=5350000; y=1560000; rh=780000
for i,(t,d) in enumerate(st):
    rect(s,M,y,lw0,rh,fill=CARD,line=EDGE,lw=12700)
    medallion(s,M+360000,y+rh//2,diam=380000,glyph=str(i+1),fill=CY,size=14)
    tb(s,M+680000,y,lw0-800000,None,paras=[
        {'text':t,'sizePt':15.8,'bold':True,'color':INK,'spaceAfterPt':2},
        {'text':d,'sizePt':12.6,'color':MUTED},
    ],cy=rh)
    y+=rh+110000
cx0=M+lw0+260000; cw_=W-M-cx0; cy0=1560000; chh=4340000
rect(s,cx0,cy0,cw_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,cx0+240000,cy0+120000,cw_-480000,'host.ts  ·  @microsoft/mxc-sdk',size=11.5,color=FOOT,font=MF,cy=300000)
rect(s,cx0+240000,cy0+450000,cw_-480000,14000,fill=EDGE)
code=['const support = await getPlatformSupport();','if (!support.isSupported) throw support.reason;','','const request = {','  command: "npm test",','  workingDirectory: repo,','  filesystem: { readwritePaths: [repo] },','  network: { egress: { default: "deny" } },','  timeoutMs: 120_000,','};','const result = await runContained(request);','validate(result); // never trust blindly']
tb(s,cx0+240000,cy0+540000,cw_-400000,None,paras=[{'text':l if l else ' '} for l in code],size=12.8,color=CODE,font=MF,cy=3660000)
footer(s,'',23)
note(s,'具体到研发团队写代码，接进项目里基本就是右边这十几行的五步闭环：先调 getPlatformSupport 和 probe 摸清当前机器能力，不支持就明确报错，绝不静默降级成裸跑；接着按默认拒绝写好 ContainerRequest，把命令、工作目录、读写边界和超时定死；开发期打开 Learning Mode 收集真实访问需求；人工审完策略后关掉学习开关，把允许路径和越权拦截路径都测通过；最后在运行期把超时、异常和权限被拒的原因清晰返给上层，别让沙箱正常拦住了危险动作，外面的 Agent 却以为出了故障在那儿盲目重试。')

# ---------- 23 HOST JOB LIST
s=23
header(s,'05 · INTEGRATING MXC','Your host\'s job list')
items=['Set an explicit timeout; handle launch error, workload failure & cancellation separately.','Clean up even when a later step throws — always deprovision persistent sessions.','Spin up one narrow sandbox per action rather than sharing one broad policy.','Write outputs to a dedicated staging folder; teardown never undoes host writes.','Explain policy denials clearly to users and agents; never fail silently.','Require a human to approve every grant suggested by Learning Mode.']
g=160000; tw=(CW-g)/2; th=1240000; y0=1600000
for i,t in enumerate(items):
    x=M+(i%2)*(tw+g); y=y0+(i//2)*(th+150000)
    rect(s,x,y,tw,th,fill=CARD,line=EDGE,lw=12700)
    medallion(s,x+440000,y+th//2,diam=440000,glyph='✓',fill=CY,size=16)
    tb(s,x+820000,y+80000,tw-1040000,t,size=15.8,color=INK,cy=th-160000)
footer(s,'',24)
note(s,'进 Demo 之前，我们最后把宿主程序自己该扛的职责做个小结。跑任务一定要设超时，并且把沙箱没拉起来、里面命令自己跑错了、还有用户主动取消这几类情况区分开；不管哪一步出了异常，收尾的资源回收和 deprovision 都得保证执行，别留下一堆悬空的临时账号。\n分配权限时，坚持一次动作配一个极窄的短命沙箱，远比偷懒复用一个大权限长寿容器安全得多。产出物单独落到专用目录，记住销毁沙箱不会撤回已经写进宿主机目录的文件。遇到越权拦截要大方地把原因告诉用户，而 Learning Mode 跑出来的任何放行建议，上线前都必须经过真人点头。')

# ---------- 24 DEMO 1 (PROCESSCONTAINER: FILES + NETWORK)
s=24
header(s,'06 · LIVE DEMO 1','ProcessContainer: files & network',cx=CW-3400000)
cx_=5550000; cy0=1560000; chh=4400000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+240000,cy0+120000,cx_-480000,'policy.json  ·  @microsoft/mxc-sdk/v1',size=11.5,color=FOOT,font=MF,cy=300000)
rect(s,M+240000,cy0+440000,cx_-480000,14000,fill=EDGE)
pc_code='{\n  "containment": { "type": "processcontainer" },\n  "filesystem": {\n    "readonlyPaths": ["demo/input"],\n    "readwritePaths": ["demo/output"],\n    "deniedPaths": ["demo/protected/denied.txt"]\n  },\n  "network": {\n    "egress": { "default": "deny" },\n    "ingress": { "hostLoopback": "allow" }\n  }\n}'
tb(s,M+240000,cy0+500000,cx_-480000,pc_code,size=11.2,color=CODE,font=MF,cy=2840000)
rect(s,M+240000,cy0+3460000,cx_-480000,740000,fill=CARD,line=EDGE,lw=12700)
tb(s,M+400000,cy0+3460000,cx_-800000,None,paras=[
    {'text':'PS> npm --prefix demo run demo:process','sizePt':12,'bold':True,'color':INK,'spaceAfterPt':4},
    {'text':'getPlatformSupport() → probe() → run()','sizePt':11.5,'color':CY},
],font=MF,cy=740000)
rx=M+cx_+320000; rw=W-M-rx; ry0=1560000; rcard_h=1360000; rgap=160000
pc_cards=[
    ('✓  Scoped read-in / write-out','readonlyPaths + readwritePaths','Workload reads demo/input and writes report.txt into its dedicated output folder.',CY),
    ('✕  Protected file denied','deniedPaths: ["protected/denied.txt"]','The same workload attempting to read a secret file is blocked by the OS kernel.',AMB),
    ('⇄  Directional network leash','egress: deny · hostLoopback: allow / deny','Reaches an explicitly permitted loopback endpoint; blocks unapproved outbound traffic.',CY),
]
for idx,(t,sub,desc,col_) in enumerate(pc_cards):
    yy=ry0+idx*(rcard_h+rgap)
    card_box(s,rx,yy,rw,rcard_h,[
        {'text':t,'sizePt':16,'bold':True,'color':INK,'spaceAfterPt':5},
        {'text':sub,'sizePt':11.8,'bold':True,'color':col_,'spaceAfterPt':5},
        {'text':desc,'sizePt':13,'color':MUTED},
    ],fill=CARD,line=EDGE,lw=12700,pad_x=220000,pad_y=95000)
footer(s,'',25)
note(s,'原理和架构聊透了，我们直接来看三个真机实跑的 Demo。\n第一个 Demo 跑的是大家最常用的 ProcessContainer。我们先调 getPlatformSupport 和 probe 验证当前 Windows 内核的原生容器能力；接着在只读 input、只写 output 的最小权限下顺利生成报告；随后保持策略一字不改，让同一个脚本去偷读 protected/denied.txt 里的机密内容，你会看到操作系统内核当场拦截；最后在完全封死外网的前提下，对比开启和关闭本机回环时对本地监听端口的访问结果，亲眼确认文件和网络两道闸门都在内核层生效。')

# ---------- 25 DEMO 2 (SESSION PTY, image right)
s=25
img(s,'hd-blocked.jpg',W-PANEL,0,PANEL,H,'Hand-drawn sketch panel: a sneaky amber tentacle pulls a robot agent toward a door, but a force field blocks it')
cxl=W-PANEL-2*M
header(s,'06 · LIVE DEMO 2','Isolated incident triage',cx=cxl)
beats=[('01 · Stage','Host stages INC-1042 and its read-only approval inside the ephemeral workspace.'),('02 · Open PTY','Provisions a separate local Windows operator account & isolated Session.'),('03 · Triage','Checks Spooler & W32Time; blocks cross-user reads and service-stop attempts.'),('04 · Teardown','Host validates evidence, stops the session, and deprovisions the temp account.')]
railx=M+180000; y=1560000; bh_d2=980000; gap_d2=120000
rect(s,railx-12700,y+bh_d2//2,25400,3*(bh_d2+gap_d2),fill=EDGE)
for i,(t,d) in enumerate(beats):
    col_b=AMB if i==2 else CY
    medallion(s,railx,y+bh_d2//2,diam=320000,glyph=str(i+1),fill=col_b,size=13)
    card_box(s,M+500000,y,cxl-500000,bh_d2,[
        {'text':t,'sizePt':16.5,'bold':True,'color':col_b,'spaceAfterPt':4},
        {'text':d,'sizePt':13.2,'color':MUTED},
    ],fill=CARD,line=EDGE,lw=12700,pad_x=220000,pad_y=90000)
    y+=bh_d2+gap_d2
footer(s,'',26,cx=cxl)
note(s,'第二个 Demo 模拟很真实的运维一线排障，用的是保留状态的 IsolationSession 配合交互式 PTY 终端。假设来了一张打印和时间服务异常的工单 INC-1042，宿主程序先把工单和只读诊断授权放进共享目录，然后现场建一个独立的 Windows 临时操作员账号和专属 Session，拉起 PowerShell 终端。\n大家一会儿看的时候可以重点盯一眼账号对比：终端里的用户名和 Session ID 跟当前登录电脑的员工完全隔离，它查它的 Spooler 和 W32Time 服务、写它的排查证据，需要时真人工程师还能直接接管这个终端敲命令，全程绝对碰不到员工自己的桌面和剪贴板。等宿主核验完生成的排查报告，立刻关闭会话并把临时账号彻底注销删干净。')

# ---------- 26 DEMO 3 (HYPERLIGHT MICROVM)
s=26
header(s,'06 · LIVE DEMO 3','Untrusted code in a MicroVM',cx=CW-3400000)
cx_=5550000; cy0=1560000; chh=4400000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
rect(s,M+cx_-1650000,cy0+120000,1410000,300000,fill=AMB,preset='roundRect',text='Experimental',size=10.5,bold=True,color=BG,font=MF)
tb(s,M+240000,cy0+120000,cx_-1950000,'hyperlight-policy.json',size=11.5,color=FOOT,font=MF,cy=300000)
rect(s,M+240000,cy0+450000,cx_-480000,14000,fill=EDGE)
vm_code='{\n  "version": "1.1.0-alpha",\n  "containment": "hyperlight",\n  "hyperlight": {\n    "runtime": "python"\n  },\n  "process": {\n    "commandLine": "python eval_untrusted.py",\n    "timeout": 30000\n  }\n}'
tb(s,M+240000,cy0+540000,cx_-480000,vm_code,size=12.5,color=CODE,font=MF,cy=2780000)
rect(s,M+240000,cy0+3460000,cx_-480000,740000,fill=CARD,line=EDGE,lw=12700)
tb(s,M+400000,cy0+3460000,cx_-800000,None,paras=[
    {'text':'PS> npm --prefix demo run demo:microvm','sizePt':12,'bold':True,'color':INK,'spaceAfterPt':4},
    {'text':'Warm snapshot restore: ~90 ms · Unikraft 5.15','sizePt':11.5,'color':CY},
],font=MF,cy=740000)
rx=M+cx_+320000; rw=W-M-rx; ry0=1560000; rcard_h=1360000; rgap=160000
vm_cards=[
    ('✓  Hardware VM kernel','Unikraft 5.15 on Hypervisor','Runs inside a hardware-virtualized partition, never sharing the Windows NT kernel.',CY),
    ('✕  Zero host surface','No host FS, env, or sockets','Host C:\\, environment variables, and TCP sockets (Errno 13) remain unreachable.',AMB),
    ('↺  Instant snapshot wipe','<100 ms restore & reset','Every call restores a warmed memory snapshot; /tmp writes vanish on exit.',CY),
]
for idx,(t,sub,desc,col_) in enumerate(vm_cards):
    yy=ry0+idx*(rcard_h+rgap)
    card_box(s,rx,yy,rw,rcard_h,[
        {'text':t,'sizePt':16,'bold':True,'color':INK,'spaceAfterPt':5},
        {'text':sub,'sizePt':11.8,'bold':True,'color':col_,'spaceAfterPt':5},
        {'text':desc,'sizePt':13,'color':MUTED},
    ],fill=CARD,line=EDGE,lw=12700,pad_x=220000,pad_y=95000)
footer(s,'',27)
note(s,'第三个 Demo 看隔离级别拉满的实验性后端：Hyperlight MicroVM。什么时候非用它不可？比如你的系统得直接运行大模型现写的复杂 Python 逻辑，或者来路不明的第三方代码，你连宿主机的 Windows 内核都不放心跟它共用；可传统虚拟机一开机就要好几秒，在线业务根本等不起。\nHyperlight 走的是另一条路：它在 Windows Hypervisor 硬件虚拟化之上跑极简的 Unikraft 微内核，每次直接从预热好的内存快照恢复，90 毫秒左右就能跑完返回。演示里我们会顺着走完整个闭环：先检查预热快照，确认里面跑的是独立的 Unikraft 5.15 内核而不是 Windows；接着在百毫秒内算完一段不可信的 Python 对账脚本；再故意跑探针代码去摸宿主机 C 盘、环境变量和网络套接字，看它们被硬件边界全部挡死；最后验证上一次运行留在 /tmp 里的文件，在下一次快照恢复时瞬间清零。')

# ---------- 27 CLOSER
s=27
img(s,'hd-closer.jpg',0,0,W,H,'Hand-drawn sketch illustration: three robot agents in bubbles and an engineer giving a thumbs up at sunrise over a city')
x=M+150000; cx=5850000
rect(s,x,480000,520000,38000,fill=CY,preset='roundRect')
tb(s,x,580000,cx,'KEY TAKEAWAYS',size=12,color=CY,font=HF,bold=True,ls=0.9,cy=260000)
tb(s,x,900000,cx,'Box the agent,\nnot the ambition',size=38,color=INK,font=HF,bold=True,cy=1240000)
tk=['1   Push guardrails down to the OS. Prompts alone can\'t enforce.','2   MXC: one policy, the right sandbox on Windows, macOS & Linux.','3   Learn first, then enforce least privilege.','4   Containment isn\'t trust. Validate every result.']
tb(s,x,2300000,cx,None,paras=[{'text':t,'spaceAfterPt':10} for t in tk],size=14.8,color=INK,cy=2260000)
card_box(s,x,4720000,cx,820000,[
    {'text':'learn.microsoft.com/windows/apps/develop/mxc','sizePt':12.5,'bold':True,'color':CY,'spaceAfterPt':5},
    {'text':'github.com/microsoft/mxc','sizePt':12.5,'bold':True,'color':CY},
],fill=CODEBG,line=EDGE,lw=12700,pad_x=200000,pad_y=90000,font=MF)
tb(s,x,5700000,cx,'Thank you  ·  Q&A',size=24,color=INK,font=HF,bold=True,cy=520000)
footer(s,'',28)
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
                if nlines(txt,o.get('sizePt',29.5),o['cx'],HF,True)>1:
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

def add_pages_qr(sl, pages_url):
    clean = (pages_url or '').strip()
    if not clean:
        return
    import io, re, qrcode
    display = re.sub(r'^https?://', '', clean).rstrip('/')
    card = sl.shapes.add_shape(PRESETS['roundRect'], Emu(740000), Emu(4190000), Emu(4720000), Emu(1930000))
    card.shadow.inherit = False
    card.fill.solid(); card.fill.fore_color.rgb = rgb(CODEBG)
    card.line.color.rgb = rgb(EDGE); card.line.width = Emu(12700)
    pad = sl.shapes.add_shape(PRESETS['roundRect'], Emu(860000), Emu(4310000), Emu(1690000), Emu(1690000))
    pad.shadow.inherit = False
    pad.fill.solid(); pad.fill.fore_color.rgb = rgb('FFFFFF')
    pad.line.color.rgb = rgb(EDGE); pad.line.width = Emu(9525)
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=12, border=2)
    qr.add_data(clean); qr.make(fit=True)
    img_buf = io.BytesIO()
    qr.make_image(fill_color='#' + INK, back_color='#FFFFFF').save(img_buf, format='PNG')
    img_buf.seek(0)
    sl.shapes.add_picture(img_buf, Emu(875000), Emu(4325000), Emu(1660000), Emu(1660000))
    cap = sl.shapes.add_textbox(Emu(2710000), Emu(4780000), Emu(2640000), Emu(280000))
    write_text(cap.text_frame, {'text': 'SCAN FOR LIVE SLIDES & DEMO', 'size': 8.5, 'bold': True, 'color': CY, 'font': BF})
    url_tb = sl.shapes.add_textbox(Emu(2710000), Emu(5060000), Emu(2640000), Emu(380000))
    write_text(url_tb.text_frame, {'text': display, 'size': 11, 'bold': True, 'color': INK, 'font': MF})
    url_tb.text_frame.paragraphs[0].runs[0].hyperlink.address = clean

def build(out,img_dir,pages_url=None):
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
        if k==0 and pages_url:
            add_pages_qr(sl, pages_url)
        if k==5:
            add_appear_animation(sl,[rendered[start:start+6] for start in (2,8,14)])
        elif k==7:
            add_appear_animation(sl,[rendered[start:start+5] for start in (2,7,12,17,22)])
        if k in notes: sl.notes_slide.notes_text_frame.text=notes[k]
    prs.save(out); print(f'Saved {out}  ({len(prs.slides)} slides)')

if __name__=='__main__':
    here=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    root=os.path.dirname(here)
    ap=argparse.ArgumentParser()
    ap.add_argument('--out',default=os.path.join(root,'dist','pptx','Guardrails with Sandboxing - MXC.pptx'))
    ap.add_argument('--images',default=os.path.join(here,'assets','images'))
    ap.add_argument('--pages-url',default=os.environ.get('PAGES_URL'))
    a=ap.parse_args(); build(a.out,a.images,pages_url=a.pages_url)
