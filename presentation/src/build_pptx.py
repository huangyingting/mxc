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
import argparse, copy, os
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from PIL import Image
try:
    from .i18n import EN_NOTES, ZH_TEXT_MAP, ZH_DEMO_LAUNCHERS
except ImportError:
    from i18n import EN_NOTES, ZH_TEXT_MAP, ZH_DEMO_LAUNCHERS

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
WF={'Inter Tight':0.50,'Segoe UI Variable Display':0.51,'Segoe UI Variable Text':0.50,'Calibri':0.50,'Microsoft YaHei UI':0.50,'Consolas':0.55,'Inconsolata':0.50,'Plus Jakarta Sans':0.55,'Newsreader':0.58,'JetBrains Mono':0.62,'Aptos Display':0.56,'Aptos':0.53,'Cascadia Code':0.62,'Georgia':0.6,'Arial':0.56,'Courier New':0.62}
def _disp_units(s):
    return sum(1.85 if ord(ch) > 0x2E80 else 1.0 for ch in s)
def nlines(text,size,cx,font,bold):
    f=WF[font]*(1.08 if bold else 1)
    cpl=max(1,int((cx-240000)/(size*12700*f)))
    n=0
    for para in text.split('\n'):
        has_cjk=any(ord(ch)>0x2E80 for ch in para)
        if has_cjk:
            u=_disp_units(para)
            n+=max(1,int((u+cpl-1)//cpl)) if u>0 else 1
        else:
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
note(s,'大家好，欢迎来到今天的分享。这两年大家都在做 Agent，你会发现它的角色变了——以前大模型只是陪你聊天，写段代码让你自己复制去跑；现在的 Agent 是直接在你的电脑或服务器上动手干活，自己读文件、跑命令、调接口。\n一旦它开始自己动手，安全问题就完全不一样了：万一它理解偏了，或者读到了别人埋坑的文档，怎么保证它不把机器搞乱、不把密钥发出去？今天我们就来聊微软开源的沙箱执行层 MXC，看看怎么既放手让 Agent 干活，又在操作系统底层把安全边界守死。')

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
note(s,'今天这 45 分钟，我们顺着落地实战的顺序往下讲。先看看 Agent 自己动手执行命令以后会冒出哪些新风险，为什么光写提示词防不住；接着看 MXC 是什么，怎么把一份策略变成系统内核的硬隔离；然后聊文件、网络、密钥怎么控，四种轻重不同的沙箱怎么选；最后结合写代码、桌面操作和敏感文档处理，看看怎么用十几行代码接进项目，并现场跑三个真机 Demo。')

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
note(s,'先看最核心的变化。以前用聊天助手，它只管出文字，中间永远隔着一个人——命令安不安全，你扫一眼再决定跑不跑，顶多也就是答错了。但自主 Agent 把中间这道人工审核拿掉了，它会自己翻代码库、装包、跑脚本、改配置。\n更麻烦的是，它下一步要敲什么命令，往往是边跑边根据上下文现编出来的，根本没法提前审完。如果让它直接顶着你本人的账号裸跑，你能打开的文件、存着的密码、能连上的内网，全都会暴露在它的一次误判或者外部注入之下。')

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
note(s,'具体会出什么事呢？最常见的就是这四类。第一是提示词注入，比如 Agent 去读开源仓库或网页，里面藏了一句恶意指令，骗它顺手跑一段外部脚本。第二和第三是偷凭证和传数据，生成的脚本去翻你本机的 SSH 私钥、云 Token 或浏览器登录态，转头就悄悄发到外部服务器。\n第四种甚至没有坏人，纯粹是模型太想帮你把活干成了——比如你让它修个测试报错，它觉得直接改线上生产配置最省事，结果好心办了坏事。不管是被人下套还是自己抄近道，只要底下没有硬边界拦着，后果都一样。\n\n参考来源：第四个越权案例改写自 Windows Developer Blog（2026 年 10 月 7 日）。')

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
note(s,'所以做 Agent 一定要记住一句话：不能让 Agent 自己管自己的权限。你在系统提示词里写十遍“绝对不准碰密钥、不准连外网”，再加内容过滤和工具白名单，这些当然要做，但它们本质上都是在劝模型守规矩，只能减少出错概率。\n一旦模型被绕晕了，或者调起的子进程不老实，提示词是拦不住系统调用的。真正能兜底的只有操作系统：哪个目录能读、哪个端口能连、能不能碰剪贴板，必须由内核说了算。模型决定它想试什么，操作系统决定它真能做成什么，两层缺一不可。')

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
note(s,'放到企业级 Windows 上，隔离、身份、可管理性这三大支柱，是 Windows 团队专门打造的操作系统底座，并且跟微软整套企业治理体系打通了：本机隔离靠开源的 MXC，管住本地能碰哪些文件和网络；身份靠接通 Entra Agent ID，让本机跑的 Agent 有自己独立的审计身份，别再去借员工账号；管理则把 Agent 365 和 Intune 延伸到终端设备上，统一盘点机器上跑了哪些 Agent、统一往下发沙箱策略。\n排落地计划时可以留意一下节奏：MXC 本身、Windows 365 云电脑支持，还有云端的 Entra Agent ID 和 Agent 365 都已经正式发布了；而给本机 Copilot CLI、Claude Code 自动绑身份，以及用 Intune 直接下发 MXC 策略，目前还在 Windows 预览阶段。\n\n参考来源：Windows Developer Blog — Logan Iyer（2026 年 10 月 7 日）及《Enterprise Deployment of Governed Agents on Windows》（2026 年 11 月 17 日）。')

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
note(s,'那 MXC 到底是什么？简单说，它是微软按 MIT 协议开源的一套轻量级沙箱执行层，今年十月已经正式发布。不管你要跑大模型现写的代码、第三方插件，还是来路不明的脚本，宿主程序只要写一份 JSON 策略说清楚准读哪儿、准写哪儿、准连哪儿，MXC 就会自动调起当前系统自带的原生沙箱去跑。\n注意它虽然叫 Execution Containers，但不用你去装厚重的 Docker——它直接用 Windows、macOS 和 Linux 内核现成的隔离能力，而且官方同步提供了 .NET、TypeScript 和 Rust 三套 SDK。')

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
note(s,'看具体接口之前，先把这五个词对上号就清楚了。最外层的 Host app 是你自己写的宿主程序，也是唯一可信的一方，负责定权限、喂输入、验结果；真正丢进沙箱里跑的那条命令或脚本叫 Workload；允许访问哪些文件、网络和界面，写下来就是 Policy；把 Policy 跟具体的命令、工作目录、超时打包到一起，就是 Container request；最后在系统底层真正立起隔离墙的机制，就是 Backend。\n整套设计最铁的一条规矩就在页面最下面：宿主拍板，策略描述，内核执行。Agent 可以申请权限，但绝不能自己给自己批权限。')

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
note(s,'把这五个角色从上往下排开，就是 MXC 的四层结构：最上面是可信的宿主程序，往下依次是 SDK 契约、MXC 运行时和操作系统隔离后端。注意中间那条醒目的红线——信任边界就画在第一层宿主程序脚下。\n这里有个很常见的误区：不少团队觉得代码既然关进沙箱跑了，拿回来的结果就天然可信。其实不是，沙箱只能保证进程在里面折腾不出界，但它打印的文字、写出来的文件里照样可能塞了脏数据或者二次注入。宿主站在红线上面，拿回来的任何产出都必须先验过再用。')

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
note(s,'再看一次调用是怎么从头走到尾的。宿主先问 SDK 当前机器支持哪些沙箱后端，再按默认拒绝的原则拼好策略和请求发出去；MXC 在中间把它翻译成当前操作系统的原生沙箱调用，跑完之后收齐输出、退出码和拦截记录，顺手把环境清干净。\n这里分享一条很实在的实战经验：不同系统底层沙箱的脾气不太一样，写自动化测试时千万别只测正常路径通不通，一定要写反向用例，亲手验证敏感文件是不是真读不了、没放行的网是不是真连不上。那底层具体有哪几种盒子可选？我们翻页看。')

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
     'AppContainer, Seatbelt or Bubblewrap. Millisecond startup and fine-grained file, network & UI rules — default for coding agents.',
     'Shared kernel · Process ACL','GA · Default',CY,1),
    ('Session container','Windows 11 only',
     'Separate local Windows login and desktop for GUI apps; blocks clipboard/input bleed (does not filter paths or IPs).',
     'Shared kernel · Own desktop','GA · GUI',CY,2),
    ('WSL container','Windows 11',
     'Runs Linux CLI workloads and toolchains on a Windows host; networking is all-isolated or bridged (no per-IP rules).',
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
note(s,'MXC 从轻到重提供了四种隔离盒子，按场景挑就行。日常写代码、跑单测、处理本地文档，首选最左边的 Process container：它在 Windows、macOS、Linux 上分别对应 AppContainer、Seatbelt 和 Bubblewrap，毫秒级启动，文件和网络都能卡得很细。\n要是你的 Agent 得操作传统桌面软件、需要自己的窗口和剪贴板，就选 Windows 专属的 Session container：它会新建一个独立的本地账号和登录桌面，跟员工自己的桌面完全隔开，也不怕乱发模拟键鼠，但它不管细粒度的路径和 IP 过滤。再往右，在 Windows 上跑 Linux 命令用 WSL 容器；要是跑高危的陌生代码、连系统内核都不敢共用，就选最右边实验性的硬件级 MicroVM。')

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
note(s,'做跨平台应用时，代码虽然只写一套，但落在各家系统底层是什么，心里最好有张图。Windows 11 默认走 ProcessContainer，同时支持独立桌面会话、WSL 容器和学习模式；Linux 默认走 Bubblewrap，macOS 走 Seatbelt。如果团队用的是 Windows 365 云电脑，进程隔离跟物理机一样开箱即用，只有上会话或虚拟化后端时才需要开嵌套虚拟化。\n页面最底下提到的 Windows Sandbox、MicroVM 和 Hyperlight 目前还在实验阶段，得显式加 alpha 配置开启。所以写代码最稳的习惯，就是启动前先调接口查一下当前环境支持什么，别把后端名字写死。')

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
lane(1520000,'One-shot','DEFAULT CHOICE',['Create','Run one workload','Tear down'],'Builds, tests & generated commands: spin up a fresh, narrow ProcessContainer for every action.',True)
lane(2960000,'State-aware','PERSISTENT STATE',['Provision','Start','Execute','Stop','Deprovision'],'ProvisionRequest → ContainerId on IsolationSession (session/UI only) or WSLc — always deprovision.',False)
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
note(s,'选定了哪种盒子，下一步就是定它活多久。MXC 分成用完即焚的 One-shot 和保留状态的 State-aware 两种模式。日常八九成的任务，比如跑一次编译、测一遍脚本、出个报告，都推荐走上面的 One-shot：每次动作现起一个极窄权限的干净沙箱，跑完立马销毁，不留后患。\n只有确实需要跨步保留现场时，才走下面的有状态容器，比如 IsolationSession 或 WSLc：先调 Provision 拿回容器 ID，后面逐条发执行请求。这里要留神它们的能力边界——比如 IsolationSession 只隔账号和桌面，不做文件网络细筛；而且无论中间跑成跑败，最后都必须调 deprovision 把建出来的临时账号删干净。\n\n参考来源：Microsoft Learn — MXC policy examples（Reuse a state-aware sandbox）。')

# ---------- 13 FILES
s=13
header(s,'03 · FEATURES & PLATFORMS','Files: read in, write out')
cx_=5850000; cy0=1540000; chh=4460000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+240000,cy0+120000,cx_-480000,'mxc-config.1.0.0.json  ·  repo-scoped tool',size=11.5,color=FOOT,font=MF,cy=300000)
rect(s,M+240000,cy0+440000,cx_-480000,14000,fill=EDGE)
code=('{\n'
      '  "version": "1.0.0", "containment": "process",\n'
      '  "process": {\n'
      '    "commandLine": "C:\\\\Windows\\\\System32\\\\cmd.exe /d /c build.cmd",\n'
      '    "cwd": "C:\\\\src\\\\my-app",\n'
      '    "env": ["SystemRoot=C:\\\\Windows",\n'
      '            "LOCALAPPDATA=C:\\\\src\\\\mxc-output",\n'
      '            "TEMP=C:\\\\src\\\\mxc-output"]\n'
      '  },\n'
      '  "filesystem": {\n'
      '    "readonlyPaths": ["C:\\\\src\\\\my-app"],\n'
      '    "readwritePaths": ["C:\\\\src\\\\mxc-output"],\n'
      '    "deniedPaths": ["C:\\\\src\\\\my-app\\\\.env",\n'
      '                    "C:\\\\src\\\\my-app\\\\secrets"]\n'
      '  }\n'
      '}')
tb(s,M+240000,cy0+490000,cx_-440000,code,size=11.1,color=CODE,font=MF,cy=3840000)
rx=M+cx_+280000; rw=W-M-rx
rules=[
    ('01 · Read in, write out (& redirect TEMP)','Keep source read-only; route outputs, tool caches, TEMP & LOCALAPPDATA to a dedicated mxc-output folder.',CY),
    ('02 · Carve out secrets inside grants','Ungranted paths (~/.ssh) stay blocked by default; use deniedPaths for .env or secrets/ inside a granted repo.',AMB),
    ('03 · Teardown is not undo','Writes to granted host folders survive container exit — validate mxc-output as an untrusted staging zone.',CY),
]
for i,(h_,d,col_) in enumerate(rules):
    yy=1540000+i*1540000
    card_box(s,rx,yy,rw,1380000,[
        {'text':h_,'sizePt':16.2,'bold':True,'color':col_,'spaceAfterPt':7},
        {'text':d,'sizePt':13.4,'color':MUTED},
    ],fill=CARD,line=EDGE,lw=12700,pad_x=220000,pad_y=105000)
footer(s,'',14)
note(s,'容器和生命周期定好之后，我们对照官方示例看看策略具体管哪三块。第一块是文件，核心口诀叫“只读进、专口出”：代码仓库只给只读权限，另外单建一个 mxc-output 目录接产出；在 Windows 下跑编译工具如果自定义了环境变量，记得把 SystemRoot 带上，再把 LOCALAPPDATA 和 TEMP 全指到这个输出目录里，这样工具写缓存也不需要放开源码目录的写权限。\n第二点看 deniedPaths：像 ~/.ssh 或别的目录，只要没加进白名单，默认本来就碰不到；deniedPaths 真正好用的地方，是在已经放开只读的仓库里把 .env 或 secrets 文件夹单独抠出去禁掉。第三点千万记住：销毁沙箱不会撤回已经写进宿主机输出目录的文件，一定要把输出目录当隔离待检区先验再收。\n\n参考来源：Microsoft Learn — MXC policy examples（learn.microsoft.com/en-us/windows/apps/develop/mxc/examples）。')

# ---------- 14 NETWORK
s=14
header(s,'03 · FEATURES & PLATFORMS','Network: deny by default')
bx0=M; by0=1620000; bw=2550000; bh=4180000
rect(s,bx0,by0,bw,bh,fill=CARD2,line=CY,lw=19050)
rect(s,bx0+200000,by0+260000,bw-400000,290000,fill=CY,preset='roundRect',text='SANDBOXED WORKLOAD',size=9.8,bold=True,color=BG,font=MF)
tb(s,bx0+200000,by0+760000,bw-400000,'Agent process',size=20,bold=True,color=INK,align='center',cy=480000)
rect(s,bx0+200000,by0+1420000,bw-400000,520000,fill=CARD,line=EDGE,lw=12700,preset='roundRect',text='egress: "deny"',size=12,bold=True,color=CY,font=MF)
rect(s,bx0+200000,by0+2080000,bw-400000,520000,fill=CARD,line=EDGE,lw=12700,preset='roundRect',text='hostLoopback: "deny"',size=11.5,bold=True,color=CY,font=MF)
tb(s,bx0+180000,by0+2840000,bw-360000,'Only explicit CIDR or proxy rules open a path.',size=12.8,color=MUTED,align='center',cy=1000000)

conn_w=880000; dx=bx0+bw+conn_w; dw=3100000; dh=1160000; dgap=350000
dests=[('192.0.2.0/24 : tcp 443','Direct egress.allow CIDR + port (no proxy)',CY,'✓'),('Any other IP / loopback','Blocked at OS firewall (default: deny)',AMB,'✕'),('http://127.0.0.1:8080','App-started proxy via networkProxy',CY,'→')]
for i,(t,d,c,g) in enumerate(dests):
    yy=by0+i*(dh+dgap)
    cy_mid=yy+dh//2
    rect(s,bx0+bw,cy_mid-14000,conn_w,28000,fill=c)
    medallion(s,bx0+bw+conn_w//2,cy_mid,diam=300000,glyph=g,fill=c,size=13)
    rect(s,dx,yy,dw,dh,fill=CARD,line=EDGE,lw=12700)
    tb(s,dx+200000,yy+110000,dw-360000,None,paras=[
        {'text':t,'sizePt':14.8,'bold':True,'color':INK,'spaceAfterPt':5},
        {'text':d,'sizePt':12.4,'bold':True,'color':c},
    ],cy=dh-220000)
fx=dx+dw+260000; fw=W-M-fx
card_box(s,fx,by0,fw,bh,[
    {'text':'NETWORK POLICY PATTERNS','sizePt':10.5,'bold':True,'color':CY,'spaceAfterPt':9},
    {'text':'• Direct CIDR: egress.allow[{ to:[{cidr}], ports }] opens TCP port, not one hostname — don’t mix with networkProxy.','sizePt':12.8,'color':INK,'spaceAfterPt':9},
    {'text':'• Hostname rules: start your host proxy first and point runtimeConfig.networkProxy to it.','sizePt':12.8,'color':INK,'spaceAfterPt':9},
    {'text':'• Windows proxy peer: allowedProxyPeer binds a package family or AppContainer profile name (not a PID).','sizePt':12.8,'bold':True,'color':CY,'spaceAfterPt':9},
    {'text':'• Windows mapping: proxy path needs ingress.default: "allow" while hostLoopback: "deny" blocks general loopback.','sizePt':12.6,'color':MUTED},
],fill=CARD,line=EDGE,lw=12700,pad_x=210000,pad_y=130000)
footer(s,'',15)
note(s,'光管住文件还不够，网络要是不封，读出来的东西顺着网线就溜了。MXC 网络默认把出站、入站和本机回环全设成 deny。真要放行，通常就两条路，而且两条不要混着配：\n第一条是直连固定网段，在 egress.allow 里写清楚 IP 段和端口——注意系统防火墙只认 IP 和端口，不认域名；第二条是按域名白名单放行，做法是在宿主外面先起一个受控代理，再用 networkProxy 指过去。在 Windows 的 ProcessContainer 下走代理有两个容易踩坑的细节：一是 allowedProxyPeer 填的是代理的包家族名或 AppContainer 配置名，不是进程 PID；二是得把 ingress.default 设成 allow，同时靠 hostLoopback 设为 deny 锁死普通本地端口，只让指定的代理连通。\n\n参考来源：Microsoft Learn — MXC policy examples（Allow direct access to an IP range & Use an application-provided proxy）。')

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
card_box(s,ex,by0,ew,1220000,[
    {'text':'Credentials stay on the host','sizePt':15.8,'bold':True,'color':CY,'spaceAfterPt':5},
    {'text':'Keep raw keys out of sandbox env & disk. Let your trusted host or proxy attach credentials outside.',  'sizePt':12.8,'color':MUTED},
],fill=CARD,line=EDGE,lw=12700,pad_x=200000,pad_y=90000)
card_box(s,ex,by0+1340000,ew,1220000,[
    {'text':'Visible UI vs. full GUI desktop','sizePt':15.8,'bold':True,'color':CY,'spaceAfterPt':5},
    {'text':'Use SessionContainer + bot login for full desktop apps, or scope ProcessContainer UI explicitly:',  'sizePt':12.8,'color':MUTED},
],fill=CARD,line=EDGE,lw=12700,pad_x=200000,pad_y=90000)
ui_y=by0+2680000; ui_h=bh-2680000
rect(s,ex,ui_y,ew,ui_h,fill=CODEBG,line=EDGE,lw=12700)
tb(s,ex+200000,ui_y+80000,ew-400000,'ui policy  ·  1.0.0 JSON / V1 SDK',size=10.5,color=FOOT,font=MF,cy=240000)
rect(s,ex+200000,ui_y+340000,ew-400000,14000,fill=EDGE)
ui_code=('"ui": {\n'
         '  "disable": true,     // false = visible tool\n'
         '  "clipboard": "none", // or "read" only\n'
         '  "injection": false   // allowInputInjection\n'
         '}')
tb(s,ex+200000,ui_y+380000,ew-360000,ui_code,size=11.4,color=CODE,font=MF,cy=ui_h-440000)
footer(s,'',16)
note(s,'除了文件和网络，最敏感的就是密钥和员工桌面了。防泄密最省心的一条铁律是：永远别把 API Key、SSH 私钥塞进沙箱的环境变量或磁盘里；真要调外部受保接口，让沙箱外面的宿主或代理验完请求后再替它挂凭证。\n桌面和交互则在 ui 这一段里给了三个开关：默认把 disable 设为 true、剪贴板设为 none、按键注入设为 false，把窗口、剪贴板和模拟键鼠全封死；哪怕个别工具确实得弹窗并读一下剪贴板，也只需要把剪贴板开成单向只读，照样严禁写剪贴板和乱发按键。要是得跑完整的传统桌面软件，那就直接丢进配好专用账号的 Session Container。\n\n参考来源：Microsoft Learn — MXC policy examples（Allow a visible tool without broad UI access）。')

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
note(s,'听到这儿大家肯定会犯愁：跑个编译器或复杂脚本动不动要摸几十个系统路径，我哪猜得到该放开哪些目录？纯靠手写肯定报错。所以 MXC 专门提供了 Learning Mode 帮我们探路。\n线上生产开的是严格拦截的 Enforcement，临时排障可以开只记日志的 Permissive；而在调策略阶段最顺手的就是中间这个 Learning Mode——没授权的访问它照样拦住，但会把每次被拒的明细记进活动报告。你先拿真实任务跑几遍，拿着这份报告由工程师逐条过目，看清哪些真是工具必需的依赖、哪些是越界乱摸，人工拍板写进策略后再切回强制模式上线。')

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
note(s,'把这几招串起来，最典型的落地场景就是写代码的 Coding Agent。你想让它改代码、跑测试、提分支，算下来它真正需要的权限其实就这么一点：当前这一个仓库目录可读写，node、python、git 工具链能调，再通过本地代理拉依赖包。至于电脑上别的项目、个人文档、SSH 私钥和外网直连，全给它封死，一点不耽误干活。\n这可不是纸面推演——大家每天在用的 GitHub Copilot CLI，底层就是靠 MXC 的进程隔离来约束模型生成的命令。')

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
note(s,'除了写代码，企业里还有三类场景特别合适。第一是操作老牌客户端或内网软件的桌面 Agent：给它开个 Session Container，配上低权限专用机器人账号——员工在前台干自己的活，Agent 在后台独立桌面里录单据，两边互不打扰。\n第二是审合同、算财报的敏感文档 Agent：直接关进断网的 Process Container，只读一个输入文件夹、只写一个汇总结果，从物理上掐断外传可能。第三是跑在 Windows 365 云电脑上的 Agent，跟云桌面并排运行，进程隔离的体验和本地电脑完全一样。')

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
note(s,'而且 MXC 不是微软自己关起门来用的组件，整个行业接得非常快。像 GitHub Copilot、OpenAI Codex、OpenClaw、Replit、LM Studio 和 Unsloth AI 已经直接用上了；NVIDIA 的 OpenShell 做了集成，OpenAI 和 Meta 也在往上做深；Anthropic Claude Code、Box、Perplexity、Raycast 等一大批厂商也官宣了支持。对企业架构师来说这是个大好事：各家工具底层正在往同一套操作系统安全底座收敛，不用每接一个工具就重搭一套孤立的沙箱。')

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
note(s,'看完场景和生态，如果你们要把 MXC 接进自家的 Agent 平台，整体架构就长这样。当大模型规划器想执行一句 npm test 时，命令绝不能直通操作系统，必须先经过中间那个可信宿主的策略引擎；宿主按任务类型选好现成的最小权限模板，再调 MXC SDK 拉起内核沙箱去跑。\n跑完往回走的那条路同样关键：退出码对不对、输出格式合不合规、中间有没有越权被拒，全由宿主验过关之后，才把结果喂回给大模型。底下再配上替沙箱挂凭证的出站代理、收敛规则用的 Learning Mode，以及管全局身份审计的 Entra Agent ID 和 Agent 365，整套闭环就齐了。')

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
note(s,'落到代码接入上其实非常轻。官方在 NuGet、npm 和 crates.io 上同步维护了 .NET、TypeScript 和 Rust 三套 SDK。不管用哪门语言，套路都一样：在代码里填好统一的 ContainerRequest，写清文件、网络、命令和超时，SDK 会自动转成稳定的 1.0.0 标准配置去调系统原生沙箱。\n写代码时只要养成一个习惯：启动前先调 getPlatformSupport 看一眼当前机器支不支持；在 Windows 上还可以顺手调一下 probe，提前确认这份策略能不能拿到预期的隔离效果。')

# ---------- 22 FIVE STEPS + CODE
s=22
header(s,'05 · INTEGRATING MXC','Integrate in five steps')
st=[('Discover','Call getPlatformSupport() and probe(request) before launch.'),('Describe','Build a V1 ContainerRequest: files, network, ui & timeoutMs.'),('Learn','Run in Learning Mode first; review the blocked-access report.'),('Enforce','Switch to Enforcement; verify allowed AND denied cases.'),('Operate','Handle timeouts & cleanup across success, failure & cancel.')]
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
tb(s,cx0+240000,cy0+120000,cw_-480000,'host.ts  ·  @microsoft/mxc-sdk/v1',size=11.5,color=FOOT,font=MF,cy=300000)
rect(s,cx0+240000,cy0+440000,cw_-480000,14000,fill=EDGE)
code=[
    'const support = getPlatformSupport();',
    'if (!support.isSupported) throw new Error(support.reason);',
    'const request: ContainerRequest = {',
    '  containment: { type: "process" }, command: "npm test",',
    '  workingDirectory: outDir, timeoutMs: 30_000,',
    '  filesystem: { readonlyPaths: [repo], readwritePaths: [outDir] },',
    '  network: { egress: { default: "deny" },',
    '             ingress: { default: "deny", hostLoopback: "deny" } },',
    '  ui: { disable: true, clipboard: "none", allowInputInjection: false },',
    '};',
    'if (probe(request).error) throw new Error("Probe failed");',
    'const result = await run(request); // validate output next!',
]
tb(s,cx0+240000,cy0+490000,cw_-400000,None,paras=[{'text':l} for l in code],size=11.1,color=CODE,font=MF,cy=3740000)
footer(s,'',23)
note(s,'用 V1 SDK 写进项目里，核心就是右边这十几行代码的五步走：第一步先调 getPlatformSupport 和 probe 探路，不支持就直接报错停下，绝不悄悄降级成裸跑；第二步按默认拒绝填好 ContainerRequest，把只读源码、可写输出目录、网络禁令、剪贴板禁令和超时一次定死；第三步开发期先开 Learning Mode 跑几遍，人工审完被拒记录；第四步切回强制模式，正反用例都测过；第五步调 run 执行，并把成功、失败、取消三种分支的清理逻辑都兜住。')

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
note(s,'进 Demo 之前，最后嘱咐一下宿主程序自己要扛稳的六件事：每次调用都要设超时，而且把沙箱没拉起来、里面脚本跑挂了、用户手动取消这三件事分清楚；不管哪一步抛异常，finally 里的资源回收和账号注销都必须走到，别留一堆僵尸账号。\n给权限时宁可一次动作起一个短命的小沙箱，也别图省事共用一个大权限容器；输出落到专门目录，记住关掉沙箱不会撤回已经写到宿主机上的文件；碰到拦截大大方方把原因报出来，而 Learning Mode 给出的任何放行建议，上线前都得过真人这一关。')

# ---------- 24 DEMO 1 (PROCESSCONTAINER: FILES + NETWORK)
s=24
header(s,'06 · LIVE DEMO 1','ProcessContainer: files & network',cx=CW-3400000)
cx_=5550000; cy0=1560000; chh=4400000
rect(s,M,cy0,cx_,chh,fill=CODEBG,line=EDGE,lw=12700)
tb(s,M+240000,cy0+120000,cx_-480000,'policy.json  ·  @microsoft/mxc-sdk/v1',size=11.5,color=FOOT,font=MF,cy=300000)
rect(s,M+240000,cy0+440000,cx_-480000,14000,fill=EDGE)
pc_code='{\n  "containment": { "type": "process" },\n  "filesystem": {\n    "readonlyPaths": ["demo/input"],\n    "readwritePaths": ["demo/output"],\n    "deniedPaths": ["demo/protected/denied.txt"]\n  },\n  "network": {\n    "egress": { "default": "deny" },\n    "ingress": { "default": "deny", "hostLoopback": "allow" }\n  },\n  "ui": { "disable": true, "clipboard": "none", "allowInputInjection": false }\n}'
tb(s,M+240000,cy0+480000,cx_-440000,pc_code,size=10.8,color=CODE,font=MF,cy=2900000)
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
note(s,'好，接下来我们点开右上角打开 Demo Studio，跑三组真机实验。第一个 Demo 测日常用得最多的 ProcessContainer，一共五步，正反对照看文件和网络两道闸门怎么生效。\n第一步先探路：加载左边这份策略，调 getPlatformSupport 和 probe，确认当前 Windows 内核的文件 ACL 和网络过滤都已经就位。\n第二步测合规读写：策略只放开了 input 只读、output 可写，沙箱里读请求文件、往输出目录写报告，权限刚好够用，顺顺当当返回 0。\n第三步马上做反向测试：策略一个字不改，让同一个程序去偷读 deniedPaths 里封死的机密文件 denied.txt——Windows 内核当场拦下，直接报拒绝访问。\n第四步和第五步对比测网络回环开关：我们把公网出站锁死成 deny，第四步先把本机回环 hostLoopback 设成 allow，沙箱里的 curl 能顺利连上本机端口拿到 200；第五步只把这一个开关改成 deny，再跑一模一样的 curl，系统网络栈当场拒绝连接。放行什么、拦住什么，全由系统内核说了算。')

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
note(s,'第二个 Demo 模拟一线运维排障，用的是带状态的 IsolationSession 配交互式 PowerShell 终端，分六步走完从建号到销号的全流程。\n第一步起会话：宿主先把故障工单 INC-1042 和只读诊断授权放进交接目录，然后让 MXC 现场新建一个 Windows 临时操作员账号和独立桌面会话，拉起持久化 PowerShell 终端，跟员工自己的桌面和剪贴板彻底隔开。\n第二步验身份隔离：看右边对比栏就能发现，终端里的临时用户名和 Session ID 跟宿主机员工完全不同；这会儿让它去摸员工的个人目录，Windows 用户权限当场拒绝。\n第三步终端不断开，接着读交接目录里的工单文件，确认这次任务只准做只读诊断。\n第四步测权限边界：用临时账号跑 Get-Service 查打印服务和时间服务的状态，正常放行；但紧接着让它跑 Stop-Service 想停掉系统服务，因为只是普通非管理员账号，立马被系统拦下。\n第五步落盘验真：Agent 把排查结果写进交接目录，宿主程序在外面独立核对工单号、操作员和服务状态无误后才算过关；如果大家在本机跑 Live 模式，底下输入框还能直接敲任意 PowerShell 命令现场试。\n第六步收尾销户：退出终端、关掉会话，调 deprovision 把临时账号从系统里彻底删干净。')

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
note(s,'第三个 Demo 看隔离级别拉满的实验性后端 Hyperlight MicroVM。什么时候会用到它？假设你的业务要直接跑大模型现写的复杂 Python 脚本或陌生插件，连跟宿主机共用 Windows 内核你都不放心；可传统虚拟机开机动不动好几秒，在线请求根本等不起。Hyperlight 的办法是在硬件虚拟化之上跑极简的 Unikraft 微内核，每次直接从预热好的内存快照恢复，不到一百毫秒就能跑完。\n我们分五步验证一遍：第一步检查 alpha 策略和预热好的 bash、Python 内存快照；第二步拉起微虚拟机跑 uname，90 毫秒左右就拿回结果，跑的是独立的 Unikraft 5.15 内核，不是宿主机的 Windows；第三步把一段 AI 写的发票对账 Python 脚本丢进去算完，通过标准输出把 JSON 结果交回宿主；第四步跑一段恶意逃逸探针，同时去摸宿主机 C 盘、读环境变量、连外部 TCP 端口——因为根本没挂宿主磁盘、没传环境变量、也没开网络栈，三路试探全撞在硬件墙上报错；第五步测快照清零：上一次调用刚往 /tmp 写了个残留文件，下一次调用再去读就已经找不到了，每次恢复都是干干净净的初始内存，一点脏状态都留不下。')

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
note(s,'最后用四句话收个尾：第一，把护栏下沉到操作系统去硬拦，别让提示词自己扛雷；第二，用好 MXC 这一份通用策略，在 Windows、macOS 和 Linux 上自动对接各自的原生沙箱；第三，落地时先开 Learning Mode 摸清真实访问需求，人工审过再锁死最小权限；第四，关进沙箱不代表结果可信，拿回来的产出一定要再验一道。\n大家回去如果想挑第一个场景试水，最推荐从写代码的 Agent 或者断网处理敏感文档开始，花半天把读写目录和拒绝列表配好、正反路径一测，马上就能见效。今天的内容就分享到这里，谢谢大家，看看各位有什么问题！')


# ===================== LOCALIZATION & LAYOUT VALIDATOR =====================
def get_localized_slides(lang='en'):
    if lang == 'en':
        return slides
    if lang != 'zh':
        raise ValueError(f'Unsupported language: {lang}')
    out = {}
    for k, ops in slides.items():
        new_ops = []
        for o in ops:
            no = copy.deepcopy(o)
            if 'text' in no and no['text'] in ZH_TEXT_MAP:
                no['text'] = ZH_TEXT_MAP[no['text']]
            if 'paragraphs' in no:
                for p in no['paragraphs']:
                    if p.get('text') in ZH_TEXT_MAP:
                        p['text'] = ZH_TEXT_MAP[p['text']]
            new_ops.append(no)
        out[k] = new_ops
    return out

def get_localized_notes(lang='en'):
    if lang == 'zh':
        return notes
    if lang == 'en':
        return {k: EN_NOTES.get(k, v) for k, v in notes.items()}
    raise ValueError(f'Unsupported language: {lang}')

def validate_layout(slide_map=None):
    target = slides if slide_map is None else slide_map
    errors=[]
    for k in sorted(target):
        ops=target[k]
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

def write_text(tf,o,lang='en'):
    tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    tf.margin_left=Emu(76200); tf.margin_right=Emu(76200)
    tf.margin_top=Emu(57150); tf.margin_bottom=Emu(57150)
    paras=o.get('paragraphs') or [{'text':o['text']}]
    raw_font=o.get('fontFamily',BF)
    eff_font='Microsoft YaHei UI' if (lang=='zh' and raw_font!=MF) else raw_font
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
            f.name=eff_font; f.size=Pt(sz)
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

def add_pages_qr(sl, pages_url, lang='en'):
    clean = (pages_url or '').strip()
    if not clean:
        return
    import io, re, qrcode
    display = re.sub(r'^https?://', '', clean).rstrip('/')
    kicker = ZH_TEXT_MAP.get('SCAN FOR LIVE SLIDES & DEMO', 'SCAN FOR LIVE SLIDES & DEMO') if lang == 'zh' else 'SCAN FOR LIVE SLIDES & DEMO'
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
    write_text(cap.text_frame, {'text': kicker, 'sizePt': 8.5, 'bold': True, 'color': CY, 'fontFamily': BF}, lang=lang)
    url_tb = sl.shapes.add_textbox(Emu(2710000), Emu(5060000), Emu(2640000), Emu(380000))
    write_text(url_tb.text_frame, {'text': display, 'sizePt': 11, 'bold': True, 'color': INK, 'fontFamily': MF}, lang=lang)
    url_tb.text_frame.paragraphs[0].runs[0].hyperlink.address = clean

def build(out,img_dir,pages_url=None,lang='en'):
    loc_slides = get_localized_slides(lang)
    loc_notes = get_localized_notes(lang)
    validate_layout(loc_slides)
    out_dir=os.path.dirname(os.path.abspath(out))
    os.makedirs(out_dir,exist_ok=True)
    prs=Presentation(); prs.slide_width,prs.slide_height=Emu(W),Emu(H)
    blank=prs.slide_layouts[6]
    for k in sorted(loc_slides):
        sl=prs.slides.add_slide(blank)
        sl.background.fill.solid(); sl.background.fill.fore_color.rgb=rgb(BG)
        rendered=[]
        for o in loc_slides[k]:
            op=o['op']
            if op=='add_image': add_picture(sl,o,img_dir)
            elif op=='add_text_box':
                sh=sl.shapes.add_textbox(Emu(o['x']),Emu(o['y']),Emu(o['cx']),Emu(o['cy'])); write_text(sh.text_frame,o,lang=lang)
                rendered.append(sh)
            elif op=='add_shape':
                sh=sl.shapes.add_shape(PRESETS[o['preset']],Emu(o['x']),Emu(o['y']),Emu(o['cx']),Emu(o['cy']))
                sh.shadow.inherit=False
                sh.fill.solid(); sh.fill.fore_color.rgb=rgb(o['fill'])
                if o.get('line'): sh.line.color.rgb=rgb(o['line']); sh.line.width=Emu(o.get('lineWidth',12700))
                else: sh.line.fill.background()
                if o.get('rot'): sh.rotation=o['rot']/60000
                if 'text' in o: write_text(sh.text_frame,o,lang=lang)
                rendered.append(sh)
        if k==0 and pages_url:
            add_pages_qr(sl, pages_url, lang=lang)
        if k==5:
            add_appear_animation(sl,[rendered[start:start+6] for start in (2,8,14)])
        elif k==7:
            add_appear_animation(sl,[rendered[start:start+5] for start in (2,7,12,17,22)])
        if k in loc_notes: sl.notes_slide.notes_text_frame.text=loc_notes[k]
    prs.save(out); print(f'Saved {out}  ({len(prs.slides)} slides, lang={lang})')

if __name__=='__main__':
    here=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    root=os.path.dirname(here)
    ap=argparse.ArgumentParser()
    ap.add_argument('--out',default=os.path.join(root,'dist','pptx','Guardrails with Sandboxing - MXC.pptx'))
    ap.add_argument('--images',default=os.path.join(here,'assets','images'))
    ap.add_argument('--pages-url',default=os.environ.get('PAGES_URL'))
    ap.add_argument('--lang',choices=('en','zh'),default='en')
    a=ap.parse_args(); build(a.out,a.images,pages_url=a.pages_url,lang=a.lang)
