#!/usr/bin/env python3
"""build_html.py - turns the same 28 slides into ONE self-contained HTML file
(images embedded, no external files). Uses the slide content from build_pptx.py.

Usage:  pip install pillow
        python -m presentation.src.build_html [--out dist/html/index.html]
Keys in the browser: <- / -> or Space to navigate, F = fullscreen,
Home/End = first/last, D = open interactive MXC Studio. Print (Ctrl+P) gives one slide per page.
"""
import argparse, base64, html, io, json, os
from PIL import Image
try:
    from . import build_pptx as D
except ImportError:
    import build_pptx as D

PX = 9525  # EMU per CSS px  (slide = 1280 x 720 px)
SVG_POINTS = {
    'rightArrow':'0,24 68,24 68,0 100,50 68,100 68,76 0,76',
    'leftArrow':'32,0 32,24 100,24 100,76 32,76 32,100 0,50',
    'chevron':'0,0 82,0 100,50 82,100 0,100 18,50',
    'homePlate':'0,0 82,0 100,50 82,100 0,100',
}
FALLBACK = {
    'Plus Jakarta Sans':'"Plus Jakarta Sans","Inter","Segoe UI Variable Display","Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif',
    'Newsreader':'"Newsreader","Fraunces",Georgia,"Times New Roman",serif',
    'JetBrains Mono':'"JetBrains Mono","Cascadia Code","Fira Code",Consolas,monospace',
    'Aptos Display':'"Plus Jakarta Sans","Aptos Display","Segoe UI Variable Display","Segoe UI",sans-serif',
    'Aptos':'"Plus Jakarta Sans",Aptos,"Segoe UI Variable Text","Segoe UI",sans-serif',
    'Cascadia Code':'"JetBrains Mono","Cascadia Code","Cascadia Mono",Consolas,monospace',
    'Georgia':'"Newsreader",Georgia,"Times New Roman",serif',
    'Calibri':'"Plus Jakarta Sans",Calibri,"Segoe UI",sans-serif',
    'Courier New':'"JetBrains Mono","Courier New",Consolas,monospace',
    'Arial':'"Plus Jakarta Sans",Arial,Helvetica,sans-serif',
}
e = lambda s: html.escape(s, quote=True)

def data_uri(path, maxw, cache={}):
    if path not in cache:
        im = Image.open(path).convert('RGB'); im.thumbnail((maxw, maxw))
        b = io.BytesIO(); im.save(b, 'JPEG', quality=82, optimize=True)
        cache[path] = 'data:image/jpeg;base64,' + base64.b64encode(b.getvalue()).decode()
    return cache[path]

def box(o):
    return f"left:{o['x']/PX:.1f}px;top:{o['y']/PX:.1f}px;width:{o['cx']/PX:.1f}px;height:{o['cy']/PX:.1f}px;"

def text_html(o):
    paras = o.get('paragraphs') or [{'text': o['text']}]
    raw_font = o.get('fontFamily', D.BF)
    font = FALLBACK.get(raw_font, FALLBACK['Plus Jakarta Sans'])
    is_display = any(p.get('sizePt', o.get('sizePt', 16)) >= 30 for p in paras)
    if is_display and raw_font != D.MF:
        font = '"Newsreader","Plus Jakarta Sans",Georgia,serif'
    ls = f"letter-spacing:{o['letterSpacingPt']*4/3:.1f}px;" if o.get('letterSpacingPt') else ('letter-spacing:-0.022em;' if is_display else '')
    out = []
    for p in paras:
        sz = p.get('sizePt', o.get('sizePt', 16)) * 4/3
        wt = 700 if p.get('bold', o.get('bold')) else (500 if sz >= 22 else 450)
        lh = 1.12 if sz >= 34 else (1.26 if sz >= 22 else 1.38)
        st = (f"font-size:{sz:.1f}px;color:#{p.get('color', o.get('color', D.INK))};"
              f"font-weight:{wt};line-height:{lh};"
              f"text-align:{p.get('align', o.get('align','center'))};")
        out.append(f'<p style="{st}">' + '<br>'.join(e(l) for l in p['text'].split('\n')) + '</p>')
    return f'<div class="t" style="font-family:{e(font)};{ls}">' + ''.join(out) + '</div>'

REVEAL_ITEMS = {
    5: (2, 5, 3),
    7: (2, 4, 5),
}

DEMO_LAUNCHERS = {
    24: ('process', '1', 'Launch Demo 1 · ProcessContainer'),
    25: ('pty', '1', 'Launch Demo 2 · IsolationSession PTY'),
    26: ('microvm', '1', 'Launch Demo 3 · Hyperlight MicroVM'),
}

def reveal_attrs(slide_index, operation_index):
    config = REVEAL_ITEMS.get(slide_index)
    if config is None:
        return 'a', ''
    start, group_size, count = config
    if start <= operation_index < start + group_size * count:
        item = (operation_index - start) // group_size
        return 'a reveal', f' data-reveal="{item}"'
    return 'a', ''

def slide_html(k, ops, img_dir, maxw):
    parts = []
    for operation_index, o in enumerate(ops):
        classes, reveal = reveal_attrs(k, operation_index)
        if o['op'] == 'add_image':
            is_full = o['x'] == 0 and o['y'] == 0 and o['cx'] == D.W and o['cy'] == D.H
            is_side = o['y'] == 0 and o['cy'] == D.H
            extra_cls = ' bleed' if (is_full or is_side) else ' framed-img'
            src = data_uri(os.path.join(img_dir, os.path.basename(o['path'])), maxw)
            parts.append(f'<img class="{classes}{extra_cls}"{reveal} style="{box(o)}" src="{src}" alt="{e(o.get("altText",""))}">')
        elif o['op'] == 'add_text_box':
            parts.append(f'<div class="{classes}"{reveal} style="{box(o)}">{text_html(o)}</div>')
        elif o['op'] == 'add_shape':
            pr = o['preset']
            inner = text_html(o) if 'text' in o else ''
            if pr in SVG_POINTS:
                stroke = f' stroke="#{o["line"]}" stroke-width="{max(1.2,o.get("lineWidth",12700)/PX):.1f}" stroke-linejoin="round"' if o.get('line') else ''
                svg = (
                    '<svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">'
                    f'<polygon points="{SVG_POINTS[pr]}" fill="#{o["fill"]}"{stroke}/></svg>'
                )
                parts.append(f'<div class="{classes} svg-shape"{reveal} style="{box(o)}">{svg}{inner}</div>')
            else:
                fill = o['fill']
                is_card = pr == 'rect' and o['cx'] > 900000 and o['cy'] > 450000
                extra_cls = ' surface-card' if is_card else ''
                if fill == D.CODEBG:
                    extra_cls += ' surface-code'
                st = box(o) + f"background:#{fill};"
                if pr == 'ellipse':
                    st += 'border-radius:50%;'
                elif pr == 'roundRect':
                    st += 'border-radius:999px;'
                elif is_card:
                    st += 'border-radius:14px;'
                if o.get('line'):
                    lw = max(1.2, o.get('lineWidth', 12700) / PX)
                    st += f"box-shadow:inset 0 0 0 {lw:.1f}px #{o['line']}, 0 8px 24px rgba(32,28,22,.05);"
                parts.append(f'<div class="{classes}{extra_cls}"{reveal} style="{st}">{inner}</div>')
    reveal_count = REVEAL_ITEMS[k][2] if k in REVEAL_ITEMS else 0
    action = ''
    if k in DEMO_LAUNCHERS:
        demo_kind, demo_step, label = DEMO_LAUNCHERS[k]
        pos_cls = ' trigger-side' if k == 25 else ' trigger-wide'
        action = (
            f'<button class="live-demo-trigger{pos_cls}" type="button" '
            f'data-demo="{demo_kind}" data-step="{demo_step}">'
            f'<span class="pulse-dot" aria-hidden="true"></span><span>{e(label)}</span>'
            f'<span class="trigger-arrow" aria-hidden="true">↗</span></button>'
        )
    return f'<section class="slide" data-n="{k+1}" data-reveals="{reveal_count}" aria-label="Slide {k+1}">{"".join(parts)}{action}</section>'

LIVE_DEMO = """
<section id="live-demo" class="live-demo" role="dialog" aria-modal="true" aria-labelledby="live-demo-title" hidden>
  <div class="studio-shell">
    <header class="studio-header">
      <div class="studio-brand">
        <div class="studio-eyebrow"><span class="studio-spark"></span><span id="live-demo-eyebrow">INTERACTIVE MXC LIVE LAB</span></div>
        <h2 id="live-demo-title">Session PTY — Stateful Incident Triage</h2>
      </div>
      <div class="studio-header-right">
        <div class="mode-tabs" role="tablist" aria-label="MXC Demo Scenarios">
          <button type="button" class="mode-tab active" data-mode="process" role="tab" aria-selected="true">
            <span class="tab-kicker">Demo 1 · One-Shot Policy</span>
            <strong>ProcessContainer</strong>
          </button>
          <button type="button" class="mode-tab" data-mode="pty" role="tab" aria-selected="false">
            <span class="tab-kicker">Demo 2 · Stateful Session</span>
            <strong>IsolationSession PTY</strong>
          </button>
          <button type="button" class="mode-tab" data-mode="microvm" role="tab" aria-selected="false">
            <span class="tab-kicker">Demo 3 · Hardware VM</span>
            <strong>Hyperlight MicroVM</strong>
          </button>
        </div>
        <div class="studio-status-group">
          <span class="status-badge" id="live-demo-status">READY</span>
          <button id="live-demo-reset" class="ghost-btn" type="button" title="Reset session & clear terminal">Reset</button>
          <button id="live-demo-close" class="close-btn" type="button" aria-label="Close live demo">✕</button>
        </div>
      </div>
    </header>

    <div class="studio-bar">
      <div class="bar-select-group">
        <div class="custom-select" id="cs-kind" data-value="process">
          <button type="button" class="cs-trigger" id="cs-kind-btn" aria-haspopup="listbox" aria-expanded="false" aria-label="Demo">
            <span class="cs-kicker">DEMO</span>
            <span class="cs-label" id="cs-kind-label">Demo 1 · ProcessContainer (5 steps)</span>
            <svg class="cs-chevron" viewBox="0 0 16 16" aria-hidden="true"><path d="M4 6l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
          </button>
          <ul class="cs-menu" id="cs-kind-menu" role="listbox" tabindex="-1" hidden>
            <li role="option" data-val="process" aria-selected="true" class="cs-option selected">
              <span class="cs-opt-main">Demo 1 · ProcessContainer (Files &amp; Network)</span>
              <span class="cs-opt-sub">5 one-shot steps · file ACLs + loopback rules</span>
            </li>
            <li role="option" data-val="pty" aria-selected="false" class="cs-option">
              <span class="cs-opt-main">Demo 2 · IsolationSession (Operator PTY)</span>
              <span class="cs-opt-sub">6 interactive steps · persistent OS session</span>
            </li>
            <li role="option" data-val="microvm" aria-selected="false" class="cs-option">
              <span class="cs-opt-main">Demo 3 · Hyperlight MicroVM (Unikernel VM)</span>
              <span class="cs-opt-sub">5 VM steps · Unikraft kernel + ~90 ms snapshot</span>
            </li>
          </ul>
        </div>

        <div class="custom-select cs-step-select" id="cs-step" data-value="1">
          <button type="button" class="cs-trigger" id="cs-step-btn" aria-haspopup="listbox" aria-expanded="false" aria-label="Step">
            <span class="cs-kicker cs-kicker-accent">STEP</span>
            <span class="cs-label" id="cs-step-label">1 · Provision Session &amp; Open PTY</span>
            <svg class="cs-chevron" viewBox="0 0 16 16" aria-hidden="true"><path d="M4 6l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
          </button>
          <ul class="cs-menu cs-step-menu" id="cs-step-menu" role="listbox" tabindex="-1" hidden></ul>
        </div>

        <button id="live-demo-run" class="primary-run-btn" type="button">
          <span id="live-run-label">Send Step 1</span>
          <span aria-hidden="true">▶</span>
        </button>
        <button id="live-demo-autoplay" class="ghost-btn autoplay-btn" type="button" title="Play all steps sequentially">
          <span id="live-autoplay-label">Auto-Play All</span>
        </button>
      </div>
      <div class="step-pills" id="live-step-pills" aria-label="Step sequence"></div>
      <div class="next-indicator" id="live-demo-next">Ready · Step 1</div>
    </div>

    <div class="studio-grid">
      <section class="terminal-card">
        <div class="terminal-chrome">
          <div class="chrome-dots" aria-hidden="true"><span></span><span></span><span></span></div>
          <div class="chrome-title"><strong id="terminal-mode-title">REAL MXC PTY · PowerShell</strong><span id="terminal-mode-sub">120 × 30 · persistent isolated session</span></div>
          <div class="chrome-pill" id="terminal-conn-pill">loopback bridge · 127.0.0.1:8787</div>
        </div>
        <pre class="studio-terminal" id="live-demo-output" role="log" aria-live="polite"></pre>
        <div class="quick-cmd-bar" id="live-quick-bar">
          <span class="quick-kicker">TRY IN PTY:</span>
          <button type="button" class="quick-chip" data-cmd="whoami">whoami</button>
          <button type="button" class="quick-chip" data-cmd="pwd">pwd</button>
          <button type="button" class="quick-chip" data-cmd="ls">ls</button>
          <button type="button" class="quick-chip" data-cmd="Get-Service Spooler,W32Time">Get-Service Spooler,W32Time</button>
          <button type="button" class="quick-chip" data-cmd="cat triage-evidence.txt">cat triage-evidence.txt</button>
        </div>
        <form class="terminal-composer" id="live-pty-form">
          <span class="composer-prompt" id="composer-prompt">PS session&gt;</span>
          <input id="live-pty-input" type="text" placeholder="After Step 1 provisions the session, type or click a PowerShell command above..." autocomplete="off">
          <button type="submit" id="live-pty-send">Send to PTY</button>
        </form>
      </section>

      <aside class="explainer-rail">
        <article class="narrative-card" id="narrative-card">
          <div class="narrative-top">
            <span class="step-badge" id="explain-step-badge">STEP 1 OF 6</span>
            <span class="boundary-tag" id="explain-boundary-tag">OS Session Boundary</span>
          </div>
          <h3 id="explain-title">Provision Dedicated Operator Session &amp; PTY</h3>
          <p class="explain-summary" id="explain-what">MXC calls Windows IsolationSession APIs to create a fresh local agent user, boot an isolated OS session, stage ticket INC-1042 into the shared workspace, and attach a live PowerShell pseudo-terminal.</p>
          <div class="explain-callout">
            <div class="callout-label" id="explain-why-label">WHY THIS STEP MATTERS</div>
            <p id="explain-why">Unlike one-shot process sandboxes, this shell stays alive across steps 1–6: variables, working directory, and operator context persist without touching the employee desktop or profile.</p>
          </div>
          <div class="explain-code-wrap">
            <span class="code-kicker">UNDER THE HOOD</span>
            <code id="explain-code">provisionContainer({ containment: 'isolation_session' }) → spawnInContainerWithPty()</code>
          </div>
        </article>

        <section class="compare-card">
          <div class="section-head">
            <strong id="compare-heading">Live Boundary Comparison</strong>
            <span id="compare-sub">Host vs. Contained Runtime</span>
          </div>
          <div class="identity-versus">
            <div class="id-box host">
              <div class="id-role">HOST USER SESSION</div>
              <div class="id-name" id="live-host-user">—</div>
              <div class="id-meta">Session ID <strong id="live-host-session">—</strong></div>
            </div>
            <div class="vs-medallion" aria-hidden="true">≠</div>
            <div class="id-box agent">
              <div class="id-role" id="agent-role-label">MXC ISOLATIONSESSION</div>
              <div class="id-name" id="live-agent-user">Not provisioned</div>
              <div class="id-meta"><span id="agent-meta-label">Session ID </span><strong id="live-agent-session">—</strong></div>
            </div>
          </div>
        </section>

        <section class="telemetry-card">
          <div class="section-head">
            <strong>Verified Runtime Evidence</strong>
            <span id="live-demo-chapter">Select Step 1 to begin</span>
          </div>
          <div class="evidence-chips">
            <span class="ev-chip" id="live-spooler">Spooler: pending</span>
            <span class="ev-chip" id="live-w32time">W32Time: pending</span>
            <span class="ev-chip" id="live-extra-chip">Approval: INC-1042 staged</span>
          </div>
          <div class="live-demo-lifecycle" id="live-lifecycle-bar">
            <span data-step="1">1 · PROVISION</span>
            <span data-step="2">2 · IDENTITY</span>
            <span data-step="3">3 · DIAGNOSE</span>
            <span data-step="4">4 · CLEAN UP</span>
          </div>
        </section>
      </aside>
    </div>
  </div>
</section>
"""

CSS = """
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@450;600&family=Newsreader:ital,opsz,wght@0,6..72,600;0,6..72,700;1,6..72,500&family=Plus+Jakarta+Sans:wght@450;500;600;700;800&display=swap');
:root{
  --bg:#%s;--ink:#%s;--accent:#%s;--card:#%s;--codebg:#%s;--edge:#%s;--muted:#%s;--warn:#%s;
  --surface:#fbf8f2;--surface-strong:#efe9dc;--ink-soft:#4a463f;
  --font-sans:'Plus Jakarta Sans','Inter','Segoe UI Variable Text','Segoe UI',system-ui,sans-serif;
  --font-display:'Newsreader','Plus Jakarta Sans',Georgia,serif;
  --font-mono:'JetBrains Mono','Cascadia Code',Consolas,monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:100%%;height:100%%;background:var(--bg);color:var(--ink);font-family:var(--font-sans);overflow:hidden;overscroll-behavior:none;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
#stage{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;overflow:hidden;background:radial-gradient(circle at 18%% 12%%,rgba(255,255,255,.55),transparent 48%%),var(--bg)}
.slide{position:absolute;width:1280px;height:720px;background:linear-gradient(180deg,rgba(255,255,255,.28) 0%%,rgba(255,255,255,0) 38%%),var(--bg);overflow:hidden;display:none;transform-origin:center center}
.slide.on{display:block;animation:slideFade .26s cubic-bezier(.2,.8,.2,1)}
@keyframes slideFade{from{opacity:.86}to{opacity:1}}
.a{position:absolute}
.surface-card{box-shadow:0 10px 28px rgba(38,33,26,.055),inset 0 1px 0 rgba(255,255,255,.55);transition:transform .2s ease}
.surface-code{border-radius:14px!important;box-shadow:0 12px 30px rgba(32,45,58,.07),inset 0 1px 0 rgba(255,255,255,.8)!important}
.framed-img{border-radius:14px;box-shadow:0 14px 34px rgba(35,30,22,.10)}
.reveal{opacity:0;visibility:hidden;transform:translateY(10px);transition:opacity .28s ease,transform .28s cubic-bezier(.2,.8,.2,1),visibility .28s}
.reveal.shown{opacity:1;visibility:visible;transform:translateY(0)}
.svg-shape svg{position:absolute;inset:0;width:100%%;height:100%%;overflow:visible;filter:drop-shadow(0 4px 10px rgba(35,48,62,.08))}
.svg-shape>.t{z-index:1}
img.a{object-fit:cover}
.t{position:absolute;inset:0;padding:6px 12px;display:flex;flex-direction:column;justify-content:center}

#deck-progress{position:fixed;left:0;top:0;height:3px;width:0;z-index:40;background:linear-gradient(90deg,var(--accent),#5d8aa8);transition:width .25s ease;pointer-events:none}

.live-demo-trigger{position:absolute;z-index:12;display:inline-flex;align-items:center;gap:10px;border:1px solid rgba(255,255,255,.28);border-radius:999px;padding:11px 20px;background:linear-gradient(135deg,var(--accent) 0%%,#2c4963 100%%);color:#fbf8f2;font:600 14px/1 var(--font-sans);letter-spacing:.01em;cursor:pointer;box-shadow:0 10px 26px rgba(44,73,99,.28);transition:transform .18s ease,box-shadow .18s ease,filter .18s ease}
.live-demo-trigger.trigger-side{left:48px;top:594px}
.live-demo-trigger.trigger-wide{right:48px;top:52px}
.live-demo-trigger:hover{transform:translateY(-2px);box-shadow:0 14px 32px rgba(44,73,99,.36);filter:brightness(1.05)}
.pulse-dot{width:8px;height:8px;border-radius:50%%;background:#7ee2b8;box-shadow:0 0 0 4px rgba(126,226,184,.24);animation:pulseDot 1.8s infinite}
@keyframes pulseDot{0%%,100%%{box-shadow:0 0 0 3px rgba(126,226,184,.28)}50%%{box-shadow:0 0 0 7px rgba(126,226,184,.06)}}
.trigger-arrow{font-size:15px;opacity:.85}

.live-demo{position:fixed;inset:0;z-index:120;background:radial-gradient(circle at 12%% 0%%,#faf7f0 0%%,var(--bg) 58%%);color:var(--ink);font-family:var(--font-sans);padding:clamp(10px,1.3vw,20px)}
.live-demo[hidden]{display:none!important}
.studio-shell{height:100%%;max-width:1680px;margin:0 auto;display:grid;grid-template-rows:auto auto minmax(0,1fr);gap:9px}
.studio-header{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
.studio-eyebrow{display:inline-flex;align-items:center;gap:7px;color:var(--accent);font:700 10.5px/1 var(--font-sans);letter-spacing:.15em;text-transform:uppercase}
.studio-spark{width:7px;height:7px;border-radius:2px;background:var(--accent);transform:rotate(45deg)}
.studio-brand h2{font:700 clamp(20px,2vw,30px)/1.08 var(--font-display);margin-top:3px;letter-spacing:-.02em;color:var(--ink)}
.studio-header-right{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.mode-tabs{display:flex;gap:4px;background:var(--card);padding:4px;border-radius:10px;border:1px solid var(--edge)}
.mode-tab{border:0;background:transparent;color:var(--muted);padding:5px 11px;border-radius:7px;text-align:left;cursor:pointer;transition:all .16s ease;font-family:var(--font-sans)}
.mode-tab .tab-kicker{display:block;font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;opacity:.8}
.mode-tab strong{display:block;font-size:12px;font-weight:700;margin-top:1px;color:var(--ink)}
.mode-tab.active{background:var(--surface);box-shadow:0 2px 8px rgba(32,28,22,.08);color:var(--accent)}
.mode-tab.active strong{color:var(--accent)}
.studio-status-group{display:flex;align-items:center;gap:7px}
.status-badge{border-radius:999px;padding:6px 11px;font:700 10.5px/1 var(--font-mono);letter-spacing:.07em;background:var(--accent);color:#fbf8f2}
.status-badge[data-tone="running"]{background:#b86b23;color:#fff}
.status-badge[data-tone="ok"]{background:#286b4b;color:#fff}
.status-badge[data-tone="err"]{background:var(--warn);color:#fff}
.ghost-btn{border:1px solid var(--edge);background:var(--surface);color:var(--ink);border-radius:8px;padding:6px 11px;font:600 11.5px/1 var(--font-sans);cursor:pointer}
.ghost-btn:hover{background:var(--card)}
.close-btn{width:30px;height:30px;border-radius:8px;border:1px solid var(--edge);background:var(--surface);color:var(--ink);font-size:14px;cursor:pointer;display:grid;place-items:center}
.close-btn:hover{background:var(--warn);color:#fff;border-color:var(--warn)}

.studio-bar{display:flex;align-items:center;justify-content:space-between;gap:9px;background:var(--surface);border:1px solid var(--edge);border-radius:12px;padding:7px 11px;box-shadow:0 4px 16px rgba(35,30,22,.04);flex-wrap:wrap;position:relative;z-index:30}
.bar-select-group{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0,0,0,0)}
.custom-select{position:relative;min-width:200px;max-width:300px}
.cs-step-select{min-width:220px;max-width:290px}
.cs-trigger{width:100%%;display:flex;align-items:center;gap:8px;border:1px solid var(--edge);border-radius:9px;background:#fff;color:var(--ink);padding:5px 10px 5px 6px;font:600 12px/1.2 var(--font-sans);cursor:pointer;box-shadow:0 2px 5px rgba(32,28,22,.04);transition:border-color .16s ease,box-shadow .16s ease}
.cs-step-select .cs-trigger{border-color:rgba(63,96,126,.48)}
.cs-trigger:hover:not(:disabled){border-color:var(--accent);box-shadow:0 4px 12px rgba(63,96,126,.12)}
.cs-trigger[aria-expanded="true"]{border-color:var(--accent);box-shadow:0 0 0 3px rgba(63,96,126,.15)}
.cs-trigger:disabled{opacity:.52;cursor:wait}
.cs-kicker{background:var(--card);color:var(--muted);border-radius:5px;padding:4px 6px;font:800 9px/1 var(--font-mono);letter-spacing:.09em;flex-shrink:0}
.cs-kicker-accent{background:rgba(63,96,126,.13);color:var(--accent)}
.cs-label{flex:1;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cs-chevron{width:14px;height:14px;color:var(--muted);flex-shrink:0;transition:transform .18s ease}
.cs-trigger[aria-expanded="true"] .cs-chevron{transform:rotate(180deg);color:var(--accent)}
.cs-menu{position:absolute;left:0;top:calc(100%% + 6px);min-width:100%%;width:max-content;max-width:390px;max-height:310px;overflow-y:auto;list-style:none;background:#fffdf9;border:1px solid var(--edge);border-radius:11px;padding:5px;box-shadow:0 18px 40px rgba(24,32,40,.16);z-index:80;animation:menuPop .15s cubic-bezier(.2,.8,.2,1)}
.cs-menu[hidden]{display:none!important}
@keyframes menuPop{from{opacity:0;transform:translateY(-5px) scale(.985)}to{opacity:1;transform:translateY(0) scale(1)}}
.cs-option{display:flex;flex-direction:column;gap:2px;padding:7px 10px;border-radius:8px;cursor:pointer;transition:background .12s ease}
.cs-option:hover,.cs-option.focused{background:var(--bg)}
.cs-option.selected{background:rgba(63,96,126,.11)}
.cs-opt-row{display:flex;align-items:center;justify-content:space-between;gap:10px}
.cs-opt-main{font:600 12.5px/1.25 var(--font-sans);color:var(--ink)}
.cs-option.selected .cs-opt-main{color:var(--accent);font-weight:700}
.cs-opt-sub{font:500 10.5px/1.25 var(--font-mono);color:var(--muted)}
.cs-opt-check{font:700 10.5px/1 var(--font-mono);color:#286b4b}
.primary-run-btn{display:inline-flex;align-items:center;gap:7px;border:0;border-radius:8px;background:linear-gradient(135deg,var(--accent) 0%%,#2a455d 100%%);color:#fff;padding:8px 15px;font:700 12px/1 var(--font-sans);cursor:pointer;box-shadow:0 5px 14px rgba(44,73,99,.24);transition:transform .14s ease,opacity .14s}
.primary-run-btn:hover:not(:disabled){transform:translateY(-1px)}
.primary-run-btn:disabled{opacity:.52;cursor:wait}
.autoplay-btn{padding:7px 12px;border-radius:8px}
.autoplay-btn.active{background:rgba(184,107,35,.14);border-color:#b86b23;color:#8f4e12}
.step-pills{display:flex;align-items:center;gap:5px}
.step-pill{width:24px;height:24px;border-radius:50%%;border:1px solid var(--edge);background:var(--bg);color:var(--muted);font:700 11px/1 var(--font-mono);display:grid;place-items:center;cursor:pointer;transition:all .16s ease}
.step-pill.current{border-color:var(--accent);background:var(--accent);color:#fff;transform:scale(1.08)}
.step-pill.done{border-color:#286b4b;background:rgba(40,107,75,.14);color:#286b4b}
.next-indicator{font:600 11px var(--font-mono);color:var(--accent);margin-left:auto}

.studio-grid{display:grid;grid-template-columns:minmax(0,1.22fr) minmax(0,.98fr);gap:12px;min-height:0}
.terminal-card{display:flex;flex-direction:column;min-height:0;min-width:0;background:#0d151d;border-radius:14px;border:1px solid #243442;overflow:hidden;box-shadow:0 16px 38px rgba(16,24,32,.22)}
.terminal-chrome{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:8px 12px;background:#15212c;border-bottom:1px solid #233342;color:#b8cad8;font:500 11px var(--font-mono);min-width:0;flex-shrink:0}
.chrome-dots{display:flex;gap:5px;flex-shrink:0}.chrome-dots span{width:9px;height:9px;border-radius:50%%;background:#3a4b59}.chrome-dots span:first-child{background:#d46a5e}.chrome-dots span:nth-child(2){background:#deb05b}.chrome-dots span:nth-child(3){background:#5fb384}
.chrome-title{display:flex;align-items:center;gap:8px;min-width:0;flex:1;overflow:hidden}.chrome-title strong{color:#f2f7fb;font-weight:600;white-space:nowrap}.chrome-title span{color:#8fa4b5;font-size:10.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.chrome-pill{border:1px solid #2d4254;border-radius:999px;padding:3px 8px;font-size:10px;color:#9ec5e2;background:#101a23;white-space:nowrap;flex-shrink:0}
.chrome-pill[data-mode="replay"]{border-color:#3a5b4c;color:#9ce0be;background:#10211b}
.studio-terminal{flex:1;min-height:0;overflow:auto;margin:0;padding:13px 15px;color:#e5eff7;font:450 12px/1.46 var(--font-mono);white-space:pre-wrap;word-break:break-word;scrollbar-color:#445a6c #0d151d}
.quick-cmd-bar{display:flex;align-items:center;gap:6px;flex-wrap:wrap;padding:6px 10px;background:#101a23;border-top:1px solid #1e2e3c;flex-shrink:0}
.quick-kicker{color:#7d95a8;font:700 9.5px/1 var(--font-mono);letter-spacing:.08em;margin-right:2px}
.quick-chip{border:1px solid #294053;border-radius:999px;background:#152330;color:#b9d7ee;padding:3px 9px;font:500 10.5px/1.2 var(--font-mono);cursor:pointer;transition:background .14s,border-color .14s}
.quick-chip:hover:not(:disabled){background:#1f3548;border-color:#5da9dd;color:#fff}
.quick-chip:disabled{opacity:.4;cursor:not-allowed}
.terminal-composer{display:flex;align-items:center;gap:7px;padding:7px 10px;background:#121d27;border-top:1px solid #223241;flex-shrink:0}
.composer-prompt{color:#7cc4f8;font:600 11.5px var(--font-mono);white-space:nowrap}
.terminal-composer input{flex:1;min-width:0;border:1px solid #2a3d4f;border-radius:6px;background:#091017;color:#f3f8fc;padding:6px 9px;font:450 11.5px var(--font-mono);outline:none}
.terminal-composer input:focus{border-color:#5da9dd}
.terminal-composer input:disabled{opacity:.48}
.terminal-composer button{border:0;border-radius:6px;background:#2e5b7d;color:#fff;padding:6px 11px;font:600 11.5px var(--font-sans);cursor:pointer}
.terminal-composer button:disabled{opacity:.42;cursor:not-allowed}

.explainer-rail{display:grid;grid-template-rows:auto auto minmax(0,1fr);grid-template-columns:minmax(0,1fr);gap:9px;min-height:0;min-width:0;overflow-y:auto;overflow-x:hidden;padding-right:2px}
.narrative-card,.compare-card,.telemetry-card{background:var(--surface);border:1px solid var(--edge);border-radius:13px;padding:11px 13px;box-shadow:0 6px 18px rgba(35,30,22,.04);min-width:0;max-width:100%%}
.narrative-card{border-top:3px solid var(--accent)}
.narrative-top{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:5px;min-width:0}
.step-badge{background:rgba(63,96,126,.12);color:var(--accent);border-radius:999px;padding:3px 9px;font:800 9.5px/1 var(--font-mono);letter-spacing:.08em;flex-shrink:0}
.boundary-tag{color:var(--muted);font:600 10.5px/1 var(--font-mono);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.narrative-card h3{font:700 17px/1.18 var(--font-display);color:var(--ink);margin-bottom:5px}
.explain-summary{color:var(--ink-soft);font-size:12px;line-height:1.42;margin-bottom:8px}
.explain-callout{background:var(--card);border-left:3px solid var(--accent);border-radius:7px;padding:7px 10px;margin-bottom:8px}
.callout-label{color:var(--accent);font:800 9.5px/1 var(--font-mono);letter-spacing:.1em;margin-bottom:3px}
.explain-callout p{color:var(--ink);font-size:11.5px;line-height:1.38}
.explain-code-wrap{display:flex;align-items:center;gap:7px;background:var(--codebg);border:1px solid var(--edge);border-radius:7px;padding:5px 9px;overflow:hidden;min-width:0}
.code-kicker{color:var(--muted);font:700 9px/1 var(--font-mono);letter-spacing:.08em;white-space:nowrap;flex-shrink:0}
.explain-code-wrap code{font:600 10.5px/1.25 var(--font-mono);color:var(--accent);white-space:normal;word-break:break-word;min-width:0}

.section-head{display:flex;align-items:baseline;justify-content:space-between;gap:8px;margin-bottom:7px;min-width:0}
.section-head strong{font:700 13.5px var(--font-sans);color:var(--ink);flex-shrink:0}
.section-head span{font:500 10.5px var(--font-mono);color:var(--muted);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0}
.identity-versus{display:grid;grid-template-columns:1fr 30px 1fr;gap:7px;align-items:stretch}
.id-box{background:var(--card);border-radius:10px;padding:8px 10px;min-width:0;border-top:3px solid var(--ink)}
.id-box.agent{border-top-color:var(--accent);background:linear-gradient(180deg,#f2f6f9 0%%,var(--card) 100%%)}
.id-role{color:var(--muted);font:700 9px/1 var(--font-mono);letter-spacing:.08em}
.id-name{font:800 16px/1.15 var(--font-sans);color:var(--ink);margin:4px 0 2px;overflow-wrap:anywhere}
.id-box.agent .id-name{color:var(--accent)}
.id-meta{color:var(--muted);font-size:11px}
.vs-medallion{align-self:center;justify-self:center;width:28px;height:28px;border-radius:50%%;background:var(--accent);color:#fff;font:700 16px/1 var(--font-sans);display:grid;place-items:center}

.telemetry-card{display:flex;flex-direction:column;justify-content:space-between;gap:8px}
.evidence-chips{display:flex;gap:6px;flex-wrap:wrap}
.ev-chip{border:1px solid var(--edge);border-radius:999px;background:var(--bg);padding:4px 9px;color:var(--muted);font:600 10.5px/1 var(--font-mono)}
.ev-chip.on{background:#e7f4ed;border-color:#286b4b;color:#1f583d}
.ev-chip.denied{background:#faece8;border-color:var(--warn);color:var(--warn)}
.live-demo-lifecycle{display:grid;grid-template-columns:repeat(4,1fr);gap:5px}
.live-demo-lifecycle span{border-top:3px solid var(--edge);padding-top:4px;color:var(--muted);font:700 9.5px/1.2 var(--font-mono);text-align:center;letter-spacing:.03em}
.live-demo-lifecycle span.done{border-color:#286b4b;color:#1f583d}

@media(max-width:860px){
  .studio-grid{grid-template-columns:1fr;overflow:auto}
  .terminal-card{min-height:280px}
  .custom-select,.cs-step-select{min-width:190px}
  .step-pills{display:none}
}
@media print{@page{size:13.333in 7.5in;margin:0}
 html,body{overflow:visible;background:#fff}#deck-progress,.live-demo,.live-demo-trigger{display:none!important}
 #stage{position:static;display:block}.slide{position:relative;display:block!important;transform:none!important;box-shadow:none;page-break-after:always;break-after:page}.reveal{opacity:1!important;visibility:visible!important;transform:none!important}}
"""
JS = """
(function(){
var S=[].slice.call(document.querySelectorAll('.slide')),i=0,steps=S.map(function(){return 0;}),liveOpen=false,bar=document.getElementById('deck-progress');
function fit(){var v=window.visualViewport,w=v?v.width:innerWidth,h=v?v.height:innerHeight,s=Math.max(w/1280,h/720);S.forEach(function(x){x.style.transform='scale('+s+')';});}
function toggleFullscreen(){if(document.fullscreenElement){document.exitFullscreen();}else if(document.documentElement.requestFullscreen){document.documentElement.requestFullscreen();}}
function renderSteps(){var shown=steps[i];S[i].querySelectorAll('.reveal').forEach(function(x){x.classList.toggle('shown',Number(x.dataset.reveal)<shown);});if(bar)bar.style.width=((i+1)/S.length*100).toFixed(2)+'%';}
function go(n){i=Math.max(0,Math.min(S.length-1,n));S.forEach(function(x,j){x.classList.toggle('on',j===i);});renderSteps();if(history.replaceState)history.replaceState(null,'','#'+(i+1));}
function next(){var total=Number(S[i].dataset.reveals||0);if(steps[i]<total){steps[i]++;renderSteps();}else{go(i+1);}}
function previous(){if(steps[i]>0){steps[i]--;renderSteps();}else{go(i-1);}}
document.getElementById('stage').addEventListener('click',function(ev){if(liveOpen||ev.target.closest('button'))return;ev.clientX<innerWidth/3?previous():next();});
addEventListener('keydown',function(ev){var k=ev.key;
 if(liveOpen){
  if(k==='Escape'){
   if(document.querySelector('.cs-menu:not([hidden])')){closeAllSelects();return;}
   closeLiveDemo();
  }
  return;
 }
 if(k==='ArrowRight'||k==='PageDown'||k===' ')next();else if(k==='ArrowLeft'||k==='PageUp')previous();
 else if(k==='Home')go(0);else if(k==='End')go(S.length-1);
 else if(k==='d'||k==='D'){ev.preventDefault();openLiveDemo(i===25?'pty':(i===26?'microvm':'process'),1);}
 else if(k==='f'||k==='F'){ev.preventDefault();toggleFullscreen();}});
addEventListener('resize',fit);document.addEventListener('fullscreenchange',fit);if(window.visualViewport)window.visualViewport.addEventListener('resize',fit);
addEventListener('hashchange',function(){var n=(parseInt(location.hash.slice(1),10)||1)-1;if(n!==i)go(n);});
fit();go((parseInt(location.hash.slice(1),10)||1)-1);

function byId(id){return document.getElementById(id);}
function cleanLive(value){
 return value
  .replace(/\\x1b\\[2J[\\s\\S]*?\\x1b\\[H/g,'')
  .replace(/\\x1b\\[\\d+;1H/g,'\\n')
  .replace(/\\x1b\\[(\\d*)C/g,function(_,n){return ' '.repeat(Math.min(8,Number(n||1)));})
  .replace(/\\x1b\\][^\\x07]*(?:\\x07|\\x1b\\\\)/g,'')
  .replace(/\\x1b\\[[0-?]*[ -\\/]*[@-~]/g,'')
  .replace(/\\r/g,'')
  .replace(/\\n{3,}/g,'\\n');
}
function setStatus(label,tone){var el=byId('live-demo-status');el.textContent=label;if(tone)el.dataset.tone=tone;else delete el.dataset.tone;}
function markLiveStep(n){var node=byId('live-demo').querySelector('[data-step="'+n+'"]');if(node)node.classList.add('done');}

var STEP_GUIDES={
 pty:[
  {label:'1 · Provision Session & Open PTY',boundary:'OS Session Boundary',title:'Step 1 · Provision Isolated Operator Session & PTY',
   what:'Host calls provisionContainer() + startContainer() to create a temporary Windows agent account, stages INC-1042 & read-only approval into the shared workspace, and opens an interactive 120×30 PowerShell PTY.',
   why:'Proves real OS session creation: the operator shell starts in its own user profile and session while keeping a bidirectional pseudo-terminal open for all subsequent steps.',
   code:"provisionContainer({ containment: 'isolation_session' }) → spawnInContainerWithPty()"},
  {label:'2 · Verify Distinct User & Session ID',boundary:'Identity & Session Separation',title:'Step 2 · Prove Agent Identity ≠ Host Employee',
   what:'Sends a live PowerShell query over the open PTY stdin to inspect $env:USERNAME and the current process Windows SessionId, comparing them side-by-side against your host login.',
   why:'Demonstrates why Session Isolation exists: the agent cannot touch the employee desktop, clipboard, registry hive, or shell history because it runs in a separate Windows logon session.',
   code:'Write-Output "$env:USERNAME | SessionId=$((Get-Process -Id $PID).SessionId)"'},
  {label:'3 · Validate Staged Ticket & Scope',boundary:'Shared Workspace Handoff',title:'Step 3 · Inspect INC-1042 & Enforce Read-Only Scope',
   what:'Inside the same running PowerShell PTY, reads incident-request.json and incident-approval.json from ephemeralWorkspacePath and verifies ticket matching + read-only-diagnostics scope.',
   why:'Shows controlled host-to-session staging: inputs enter only through the explicit shared workspace, and context variables stay in memory for the next command.',
   code:'Get-Content .\\incident-request.json, .\\incident-approval.json | ConvertFrom-Json'},
  {label:'4 · Diagnose Live Windows Services',boundary:'Stateful Interactive Triage',title:'Step 4 · Query Real Spooler & W32Time Status',
   what:'Queries live Windows service health (Spooler and W32Time) inside the isolated session without restarting services or mutating machine configuration.',
   why:'Illustrates real L1/L2 troubleshooting: multi-step diagnostics share one live terminal session, and a human operator can also type custom commands in the prompt below at any time.',
   code:'Get-Service Spooler, W32Time | Select-Object Name, Status'},
  {label:'5 · Write Host-Verified Audit Evidence',boundary:'Untrusted-to-Trusted Verification',title:'Step 5 · Persist Triage Evidence & Validate on Host',
   what:'Writes structured triage-evidence.txt inside the session workspace, then the trusted Host independently reads and validates ticket, operator identity, session ID, and service states.',
   why:'Containment is not blind trust: the Host validates structured output from the workspace before accepting the result.',
   code:'Set-Content .\\triage-evidence.txt → Host verifies ticket=INC-1042'},
  {label:'6 · Exit Shell & Deprovision Account',boundary:'Deterministic Teardown',title:'Step 6 · Close PTY, Stop Session & Deprovision User',
   what:'Sends exit to PowerShell, stops the active Windows session, and deprovisions the temporary OS agent account.',
   why:'Leaves zero lingering operator state or orphan accounts once the incident workflow finishes.',
   code:'terminal.input.end("exit") → stopContainer() → deprovisionContainer()'}
 ],
 process:[
  {label:'1 · Discover Backends & Probe Tier',boundary:'Capability Discovery',title:'Step 1 · Runtime Discovery & ProcessContainer Probe',
   what:'Calls getPlatformSupport() and Windows probe(request) before launching any workload to confirm native BaseContainer / ProcessContainer availability.',
   why:'Host apps never assume sandbox capabilities blindly—preflight probing verifies that path deny rules and loopback controls are enforceable on this machine.',
   code:'getPlatformSupport() + probe(request) → tier: base-container'},
  {label:'2 · Allowed Read-In / Write-Out',boundary:'Filesystem Least Privilege',title:'Step 2 · Read Approved Input & Write Dedicated Report',
   what:'Runs a one-shot ProcessContainer granted read-only access to demo/input and read-write access only to demo/output, producing report.txt.',
   why:'Validates the positive path: narrow, task-scoped file grants let legitimate agent work succeed cleanly.',
   code:'filesystem: { readonlyPaths: ["input"], readwritePaths: ["output"] }'},
  {label:'3 · Block Protected Secret File',boundary:'OS-Enforced Deny Rule',title:'Step 3 · Attempt to Read Protected File → Denied',
   what:'Runs the workload under the exact same policy while attempting to read demo/protected/denied.txt; the Windows kernel blocks access.',
   why:'Prompts persuade, the OS enforces: even if prompt injection tricks the model into reading secrets, the file ACL/container boundary denies it.',
   code:'deniedPaths: ["protected/denied.txt"] → Access is denied'},
  {label:'4 · Reach Approved Loopback Endpoint',boundary:'Directional Network Control',title:'Step 4 · Connect to Explicitly Allowed Endpoint',
   what:'Starts a local HTTP verification listener and launches a ProcessContainer with hostLoopback: "allow" while internet egress stays default-deny.',
   why:'Shows how agents reach a trusted local proxy or broker without opening raw internet access.',
   code:'network: { egress: { default: "deny" }, ingress: { hostLoopback: "allow" } }'},
  {label:'5 · Block Unapproved Network Access',boundary:'Default-Deny Network Leash',title:'Step 5 · Same Workload with Loopback Denied → Blocked',
   what:'Runs against the exact same local HTTP listener with hostLoopback: "deny" and egress: "deny"; curl fails immediately.',
   why:'Proves data exfiltration is cut off at the OS network stack whenever a destination is not explicitly permitted.',
   code:'network: { egress: { default: "deny" }, ingress: { hostLoopback: "deny" } }'}
 ],
 microvm:[
  {label:'1 · Discover Hyperlight & Warm Snapshot',boundary:'Experimental Opt-In',title:'Step 1 · Verify Hyperlight Backend & Warmed VM Snapshots',
   what:'Calls getPlatformSupport() to confirm native hyperlight support, checks pre-warmed Unikraft initrd memory snapshots (bash + python), and dry-runs the 1.1.0-alpha contract.',
   why:'MicroVMs require explicit experimental opt-in and pre-warmed snapshots so subsequent untrusted code launches skip cold OS boot.',
   code:'getPlatformSupport() + wxc-exec --experimental --dry-run (containment: "hyperlight")'},
  {label:'2 · Boot Unikraft Kernel (<100 ms)',boundary:'Hardware Hypervisor Boundary',title:'Step 2 · Prove Separate Unikraft Kernel Inside VM Partition',
   what:'Restores a hardware-virtualized Unikraft 5.15 x86_64 guest on Windows Hypervisor Platform and inspects uname + rootfs in ~90 ms.',
   why:'Unlike process containers that share the host OS kernel, a kernel bug or syscall exploit inside Hyperlight hits a dedicated unikernel VM boundary.',
   code:'containment: "hyperlight" → Unikraft 5.15.148-Ijiraq (Windows Hypervisor Platform)'},
  {label:'3 · Evaluate Untrusted AI Python Code',boundary:'Sub-100ms Code Sandbox',title:'Step 3 · Run AI-Generated Python 3.12 Analytics Safely',
   what:'Executes untrusted model-authored Python reconciliation logic inside the MicroVM and returns structured JSON anomalies to the host.',
   why:'Gives AI coding and data agents near-process responsiveness (~90 ms) with hardware-level virtual machine isolation.',
   code:'hyperlight: { runtime: "python" } → Python 3.12.15 JSON stdout'},
  {label:'4 · Block Host FS, Env & Socket Escape',boundary:'Zero-Surface Air Gap',title:'Step 4 · Adversarial Probe: Host C:\\, Env & TCP Blocked',
   what:'Runs an adversarial script attempting to read C:\\Windows, inspect host environment variables, and open a loopback TCP socket.',
   why:'Zero host mounts, zero inherited host env vars, and zero guest NIC (Errno 13 Permission denied) eliminate exfiltration paths by construction.',
   code:'HOST_FS=UNREACHABLE · HOST_ENV=ISOLATED · GUEST_NET=BLOCKED (Errno 13)'},
  {label:'5 · Verify Stateless Snapshot Reset',boundary:'Deterministic Memory Wipe',title:'Step 5 · Write /tmp Marker in Run A → Gone in Run B',
   what:'Writes /tmp/agent-residue.txt in MicroVM Run A, then immediately launches Run B from the warmed snapshot and checks if the file exists.',
   why:'Every invocation starts from an identical, clean memory image—preventing cross-tenant or cross-prompt poisoning.',
   code:'Run A: /tmp/agent-residue=True → Run B: /tmp/agent-residue=False'}
 ]
};

var liveNext={pty:1,process:1,microvm:1},doneSet={pty:{},process:{},microvm:{}},liveSource=null,ptyReady=false;
var KIND_LABELS={
 process:'Demo 1 · ProcessContainer (5 steps)',
 pty:'Demo 2 · IsolationSession PTY (6 steps)',
 microvm:'Demo 3 · Hyperlight MicroVM (5 steps)'
};

function getKindVal(){return byId('cs-kind').dataset.value||'process';}
function getStepVal(){return Number(byId('cs-step').dataset.value)||1;}
function setControlsDisabled(flag){
 byId('live-demo-run').disabled=flag;
 byId('cs-kind-btn').disabled=flag;
 byId('cs-step-btn').disabled=flag;
 if(flag)closeAllSelects();
}
function closeAllSelects(){
 ['cs-kind','cs-step'].forEach(function(id){
  var root=byId(id);if(!root)return;
  var btn=root.querySelector('.cs-trigger'),menu=root.querySelector('.cs-menu');
  btn.setAttribute('aria-expanded','false');
  menu.hidden=true;
 });
}
function setupCustomSelect(rootId,onPick){
 var root=byId(rootId),btn=root.querySelector('.cs-trigger'),menu=root.querySelector('.cs-menu');
 btn.addEventListener('click',function(ev){
  ev.stopPropagation();
  if(btn.disabled)return;
  var willOpen=menu.hidden;
  closeAllSelects();
  if(willOpen){
   menu.hidden=false;
   btn.setAttribute('aria-expanded','true');
  }
 });
 btn.addEventListener('keydown',function(ev){
  if(btn.disabled)return;
  if(ev.key==='ArrowDown'||ev.key==='ArrowUp'||ev.key==='Enter'||ev.key===' '){
   ev.preventDefault();
   if(menu.hidden){closeAllSelects();menu.hidden=false;btn.setAttribute('aria-expanded','true');}
   var opts=[].slice.call(menu.querySelectorAll('.cs-option'));
   if(!opts.length)return;
   var cur=opts.findIndex(function(o){return o.classList.contains('focused')||o.classList.contains('selected');});
   if(ev.key==='ArrowDown')cur=(cur+1)%opts.length;
   else if(ev.key==='ArrowUp')cur=(cur-1+opts.length)%opts.length;
   else if((ev.key==='Enter'||ev.key===' ')&&cur>=0&&!menu.hidden){onPick(opts[cur].dataset.val);closeAllSelects();return;}
   opts.forEach(function(o,idx){o.classList.toggle('focused',idx===cur);});
   if(opts[cur])opts[cur].scrollIntoView({block:'nearest'});
  }
 });
 menu.addEventListener('click',function(ev){
  ev.stopPropagation();
  var li=ev.target.closest('.cs-option');
  if(!li)return;
  onPick(li.dataset.val);
  closeAllSelects();
  btn.focus();
 });
}
document.addEventListener('click',function(){if(liveOpen)closeAllSelects();});

function setKindVal(val){
 var root=byId('cs-kind');
 root.dataset.value=val;
 byId('cs-kind-label').textContent=KIND_LABELS[val]||val;
 root.querySelectorAll('.cs-option').forEach(function(li){
  var sel=li.dataset.val===val;
  li.classList.toggle('selected',sel);
  li.setAttribute('aria-selected',sel?'true':'false');
 });
}

function setStepVal(val){
 var kind=getKindVal(),list=STEP_GUIDES[kind],n=Math.max(1,Math.min(list.length,Number(val)||1));
 var root=byId('cs-step'),menu=byId('cs-step-menu');
 root.dataset.value=String(n);
 byId('cs-step-label').textContent=list[n-1].label;
 menu.innerHTML=list.map(function(item,idx){
  var stepNum=idx+1,sel=stepNum===n,done=!!doneSet[kind][stepNum];
  return '<li role="option" data-val="'+stepNum+'" aria-selected="'+(sel?'true':'false')+'" class="cs-option'+(sel?' selected':'')+'">'
   +'<div class="cs-opt-row"><span class="cs-opt-main">'+item.label+'</span>'+(done?'<span class="cs-opt-check">✓ DONE</span>':'')+'</div>'
   +'<span class="cs-opt-sub">'+item.boundary+'</span></li>';
 }).join('');
}

function renderExplain(kind,step){
 var list=STEP_GUIDES[kind],idx=Math.max(0,Math.min(list.length-1,(Number(step)||1)-1)),g=list[idx];
 byId('explain-step-badge').textContent='STEP '+(idx+1)+' OF '+list.length;
 byId('explain-boundary-tag').textContent=g.boundary;
 byId('explain-title').textContent=g.title;
 byId('explain-what').textContent=g.what;
 byId('explain-why').textContent=g.why;
 byId('explain-code').textContent=g.code;
 var pills=byId('live-step-pills');
 pills.innerHTML=list.map(function(_,k){
  var n=k+1,cls='step-pill'+(n===idx+1?' current':'')+(doneSet[kind][n]?' done':'');
  return '<button type="button" class="'+cls+'" data-pick="'+n+'" title="'+list[k].label+'">'+n+'</button>';
 }).join('');
 pills.querySelectorAll('button').forEach(function(b){
  b.addEventListener('click',function(){
   if(byId('cs-step-btn').disabled)return;
   setStepVal(b.dataset.pick);
   syncStepSelection();
  });
 });
 byId('live-run-label').textContent='Send Step '+(idx+1);
}

function syncComposerState(){
 var kind=getKindVal(),inp=byId('live-pty-input'),btn=byId('live-pty-send'),form=byId('live-pty-form'),qbar=byId('live-quick-bar');
 if(kind!=='pty'){
  form.style.display='none';
  if(qbar)qbar.style.display='none';
 }else{
  form.style.display='flex';
  if(qbar){
   qbar.style.display='flex';
   qbar.querySelectorAll('.quick-chip').forEach(function(c){c.disabled=!ptyReady||stepRunning;});
  }
  inp.disabled=!ptyReady||stepRunning;btn.disabled=!ptyReady||stepRunning;
  inp.placeholder=ptyReady
   ?(replayMode?'Session PTY ready — click a preset chip above or type whoami, pwd, ls...':'Session PTY is live — type any PowerShell command (e.g. whoami, pwd, ls) and press Enter')
   :'Run Step 1 first to provision the IsolationSession PowerShell PTY...';
 }
}

function syncStepSelection(){
 var kind=getKindVal(),step=getStepVal();
 setStepVal(step);
 byId('live-demo-next').textContent=(step===liveNext[kind]?'Recommended · Step ':'Selected · Step ')+step;
 renderExplain(kind,step);
}

function updateLiveMode(){
 var kind=getKindVal();
 setKindVal(kind);
 document.querySelectorAll('.mode-tab').forEach(function(t){
  var active=t.dataset.mode===kind;
  t.classList.toggle('active',active);
  t.setAttribute('aria-selected',active?'true':'false');
 });
 setStepVal(liveNext[kind]);
 var metaLbl=byId('agent-meta-label');
 if(kind==='pty'){
  byId('live-demo-eyebrow').textContent='06 · LIVE DEMO 2 · STATEFUL SESSION ISOLATION';
  byId('live-demo-title').textContent='Demo 2 · Session PTY — Incident Triage';
  byId('terminal-mode-title').textContent='REAL MXC PTY · PowerShell';
  byId('terminal-mode-sub').textContent='120 × 30 · persistent isolated operator session';
  byId('agent-role-label').textContent='MXC ISOLATIONSESSION';
  if(metaLbl)metaLbl.textContent='Session ID ';
  byId('compare-heading').textContent='Host Login vs. Isolated Agent Session';
  byId('compare-sub').textContent='Separate OS account & session';
  byId('live-agent-user').textContent='Awaiting Step 1';byId('live-agent-session').textContent='—';
  byId('live-spooler').textContent='Spooler: pending';byId('live-w32time').textContent='W32Time: pending';
  byId('live-extra-chip').textContent='Approval: INC-1042 staged';
  byId('live-demo').querySelectorAll('.live-demo-lifecycle span').forEach(function(x,idx){x.textContent=['1 · PROVISION','2 · IDENTITY','3–5 · TRIAGE','6 · CLEAN UP'][idx];});
 }else if(kind==='process'){
  byId('live-demo-eyebrow').textContent='06 · LIVE DEMO 1 · PROCESSCONTAINER GUARDRAILS';
  byId('live-demo-title').textContent='Demo 1 · ProcessContainer — Files & Network';
  byId('terminal-mode-title').textContent='MXC PROCESSCONTAINER · One-Shot Sandboxes';
  byId('terminal-mode-sub').textContent='base-container · file allow/deny + loopback allow/deny';
  byId('agent-role-label').textContent='PROCESSCONTAINER SANDBOX';
  if(metaLbl)metaLbl.textContent='Session ID ';
  byId('compare-heading').textContent='Trusted Host vs. One-Shot Sandbox';
  byId('compare-sub').textContent='Deny-by-default policy';
  byId('live-agent-user').textContent='base-container';byId('live-agent-session').textContent='ephemeral';
  byId('live-spooler').textContent='Files: pending';byId('live-w32time').textContent='Network: pending';
  byId('live-extra-chip').textContent='Protected secret: locked';
  byId('live-demo').querySelectorAll('.live-demo-lifecycle span').forEach(function(x,idx){x.textContent=['1 · PROBE','2–3 · FILES','4–5 · NETWORK','VERIFIED'][idx];});
 }else{
  byId('live-demo-eyebrow').textContent='06 · LIVE DEMO 3 · HARDWARE MICROVM (EXPERIMENTAL)';
  byId('live-demo-title').textContent='Demo 3 · Hyperlight MicroVM — Untrusted Code';
  byId('terminal-mode-title').textContent='MXC HYPERLIGHT · Unikraft MicroVM';
  byId('terminal-mode-sub').textContent='hardware hypervisor partition · bash & python 3.12 snapshots';
  byId('agent-role-label').textContent='HYPERLIGHT MICROVM';
  if(metaLbl)metaLbl.textContent='Warm boot ';
  byId('compare-heading').textContent='Windows Host vs. Hardware MicroVM';
  byId('compare-sub').textContent='Separate Unikraft kernel';
  byId('live-agent-user').textContent='Unikraft VM';byId('live-agent-session').textContent='standby';
  byId('live-spooler').textContent='VM Kernel: pending';byId('live-w32time').textContent='Host/Net escape: pending';
  byId('live-extra-chip').textContent='Snapshot reset: pending';
  byId('live-demo').querySelectorAll('.live-demo-lifecycle span').forEach(function(x,idx){x.textContent=['1 · SNAPSHOT','2 · VM KERNEL','3–4 · ISOLATE','5 · RESET'][idx];});
 }
 syncComposerState();
 syncStepSelection();
}

function isLocalDemoHost(){return ['127.0.0.1','localhost'].includes(location.hostname)&&location.port==='8787';}
var replayMode=!isLocalDemoHost(),replayTimers=[],autoPlaying=false,stepRunning=false,stepDoneWaiters=[],replayBannerShown=false;

function clearReplayTimers(){
 while(replayTimers.length)clearTimeout(replayTimers.pop());
 stepRunning=false;
}
function syncAutoPlayBtn(){
 var btn=byId('live-demo-autoplay'),lbl=byId('live-autoplay-label');
 if(!btn||!lbl)return;
 btn.classList.toggle('active',autoPlaying);
 lbl.textContent=autoPlaying?'Stop Auto-Play':'Auto-Play All';
}
function stopAutoPlay(){
 autoPlaying=false;
 syncAutoPlayBtn();
}
function resolveStepWaiters(ok){
 while(stepDoneWaiters.length){var fn=stepDoneWaiters.shift();fn(ok);}
}

function resetLive(){
 stopAutoPlay();
 clearReplayTimers();
 resolveStepWaiters(false);
 ptyReady=false;
 replayBannerShown=false;
 doneSet={pty:{},process:{},microvm:{}};
 liveNext={pty:1,process:1,microvm:1};
 byId('live-demo-output').textContent='';
 setControlsDisabled(false);
 setStatus(replayMode?'REPLAY READY':'READY');
 byId('live-demo').querySelectorAll('.ev-chip').forEach(function(x){x.classList.remove('on','denied');});
 byId('live-demo').querySelectorAll('.live-demo-lifecycle span.done').forEach(function(x){x.classList.remove('done');});
 updateLiveMode();
}

function inspectLive(value){
 var kind=getKindVal();
 var webId=value.match(/__WEB_IDENTITY_VAL__=([^|\\r\\n]+)\\|(\\d+)/);
 if(webId){byId('live-agent-user').textContent=webId[1].trim();byId('live-agent-session').textContent=webId[2];}
 var agent=value.match(/^AGENT_USER=([^\\n\\r\\\"]+)/m),session=value.match(/^AGENT_SESSION_ID=(\\d+)/m);
 if(agent)byId('live-agent-user').textContent=agent[1].trim();if(session)byId('live-agent-session').textContent=session[1];
 if(kind==='pty'){
  var svc=value.match(/__WEB_SERVICES_VAL__=([^|\\r\\n]+)\\|([^\\r\\n]+)/);
  if(svc){
   byId('live-spooler').textContent='Spooler: '+svc[1].trim();byId('live-spooler').classList.add('on');
   byId('live-w32time').textContent='W32Time: '+svc[2].trim();byId('live-w32time').classList.add('on');
  }
  if(/__WEB_APPROVAL_VAL__=/.test(value)){byId('live-extra-chip').textContent='Approval: verified read-only';byId('live-extra-chip').classList.add('on');}
  if(/__WEB_EVIDENCE_VAL__=written/.test(value)){byId('live-extra-chip').textContent='Evidence: host-verified ✓';byId('live-extra-chip').classList.add('on');}
 }else if(kind==='process'){
  if(value.includes('STEP 2 COMPLETE')){byId('live-spooler').textContent='Write output: ALLOWED ✓';byId('live-spooler').classList.add('on');}
  if(value.includes('STEP 3 COMPLETE')){byId('live-extra-chip').textContent='Secret read: DENIED ✕';byId('live-extra-chip').classList.add('denied');}
  if(value.includes('STEP 4 COMPLETE')){byId('live-w32time').textContent='Loopback allow: REACHED ✓';byId('live-w32time').classList.add('on');}
  if(value.includes('STEP 5 COMPLETE')){byId('live-w32time').textContent='Loopback deny: BLOCKED ✕';byId('live-w32time').classList.remove('on');byId('live-w32time').classList.add('denied');}
 }else if(kind==='microvm'){
  var gk=value.match(/GUEST_KERNEL=([^\\r\\n]+)/);
  if(gk){
   var kshort=gk[1].trim().split('-')[0];
   byId('live-agent-user').textContent=kshort;
   byId('live-spooler').textContent='Kernel: '+kshort+' ✓';byId('live-spooler').classList.add('on');
  }
  var bms=value.match(/MICROVM_BOOT_MS=(\\d+)/);
  if(bms)byId('live-agent-session').textContent=bms[1]+' ms';
  if(value.includes('MICROVM STEP 3 COMPLETE')&&!gk){
   byId('live-spooler').textContent='Python 3.12 eval: OK ✓';byId('live-spooler').classList.add('on');
  }
  if(value.includes('MICROVM STEP 4 COMPLETE')){
   byId('live-w32time').textContent='FS/Env/Net: BLOCKED ✕';byId('live-w32time').classList.add('denied');
  }
  if(value.includes('MICROVM STEP 5 COMPLETE')){
   byId('live-extra-chip').textContent='Snapshot reset: CLEAN ✓';byId('live-extra-chip').classList.add('on');
  }
 }
}

function applyDemoEvent(data){
 var out=byId('live-demo-output');
 if(replayBannerShown&&data.kind==='step-start'){out.textContent='';replayBannerShown=false;}
 if(data.kind==='phase'&&data.text){
  out.textContent+=cleanLive(data.text);
  if(data.user)byId('live-agent-user').textContent=data.user;
  ptyReady=true;syncComposerState();
 }else if(data.kind==='output'){
  out.textContent+=cleanLive(data.text);out.scrollTop=out.scrollHeight;inspectLive(out.textContent);
 }else if(data.kind==='identity'){
  byId('live-agent-user').textContent=data.user;byId('live-agent-session').textContent=data.sessionId;
 }else if(data.kind==='approval'){
  byId('live-extra-chip').textContent='Approval: '+data.ticket+' ('+data.scope+')';byId('live-extra-chip').classList.add('on');
 }else if(data.kind==='services'){
  byId('live-spooler').textContent='Spooler: '+data.spooler;byId('live-spooler').classList.add('on');
  byId('live-w32time').textContent='W32Time: '+data.w32time;byId('live-w32time').classList.add('on');
 }else if(data.kind==='evidence'){
  byId('live-extra-chip').textContent='Evidence: host-verified ✓';byId('live-extra-chip').classList.add('on');
 }else if(data.kind==='step-start'){
  stepRunning=true;
  setStatus((replayMode?'REPLAYING STEP ':'RUNNING STEP ')+data.step,'running');
  byId('live-demo-chapter').textContent=(replayMode?'Replaying captured Step ':'Executing Step ')+data.step+'…';
  setControlsDisabled(true);syncComposerState();
  out.textContent+=(out.textContent?'\\n':'')+'──▶ [Step '+data.step+'] '+STEP_GUIDES[data.demo][data.step-1].label+'\\n';
  out.scrollTop=out.scrollHeight;
 }else if(data.kind==='step-done'){
  stepRunning=false;
  var demo=data.demo,step=Number(data.step),defs=STEP_GUIDES[demo];
  doneSet[demo][step]=true;
  if(demo==='process'){if(step===1)markLiveStep(1);if(step===2||step===3)markLiveStep(2);if(step>=4)markLiveStep(3);if(step===5)markLiveStep(4);}
  else if(demo==='microvm'){if(step===1)markLiveStep(1);if(step===2)markLiveStep(2);if(step===3||step===4)markLiveStep(3);if(step===5)markLiveStep(4);}
  else{if(step===1){markLiveStep(1);ptyReady=true;}if(step===2)markLiveStep(2);if(step>=3&&step<=5)markLiveStep(3);if(step===6){markLiveStep(4);ptyReady=false;}}
  liveNext[demo]=step>=defs.length?1:step+1;
  setStepVal(liveNext[demo]);
  setStatus(step===defs.length?'COMPLETE':'STEP '+step+' OK',step===defs.length?'ok':undefined);
  byId('live-demo-chapter').textContent=step===defs.length?'All steps verified — ready to replay':'Step '+step+' verified · Next up: Step '+liveNext[demo];
  setControlsDisabled(false);
  syncComposerState();syncStepSelection();
  resolveStepWaiters(true);
 }else if(data.kind==='step-error'){
  stepRunning=false;
  setStatus('STEP ERROR','err');
  byId('live-demo-chapter').textContent=data.error;
  setControlsDisabled(false);
  syncComposerState();
  resolveStepWaiters(false);
 }
}

function initReplayMode(){
 replayMode=true;
 var pill=byId('terminal-conn-pill');
 pill.textContent='recorded MXC capture · Windows 11';
 pill.dataset.mode='replay';
 if(window.RECORDED_DEMO&&window.RECORDED_DEMO.hostMeta){
  byId('live-host-user').textContent=window.RECORDED_DEMO.hostMeta.username;
  byId('live-host-session').textContent=window.RECORDED_DEMO.hostMeta.sessionId;
 }
 setStatus('REPLAY READY');
 byId('live-demo-chapter').textContent='Captured on Windows 11 · Click Send Step or Auto-Play All';
 var out=byId('live-demo-output');
 if(!out.textContent){
  replayBannerShown=true;
  out.textContent=[
   '# Recorded @microsoft/mxc-sdk@1.1.0 execution trace (Windows 11 x64)',
   '# Choose any step above and click "Send Step" (or "Auto-Play All") to replay real sandbox output.',
   '# Local live bridge optional: npm --prefix demo run demo:web → http://127.0.0.1:8787'
  ].join('\\n');
 }
 syncComposerState();
}

function expandReplayItems(items){
 var expanded=[];
 (items||[]).forEach(function(item){
  var ev=item.event;
  if(ev.kind==='output'&&typeof ev.text==='string'){
   var nl=ev.text.indexOf('\\n');
   if(nl>0&&nl<ev.text.length-1){
    expanded.push({delayMs:Math.max(95,item.delayMs||0),event:{kind:'output',text:ev.text.slice(0,nl+1)}});
    expanded.push({delayMs:190,event:{kind:'output',text:ev.text.slice(nl+1)}});
    return;
   }
  }
  var d=ev.kind==='step-start'?10:(ev.kind==='step-done'?110:Math.max(105,item.delayMs||0));
  expanded.push({delayMs:d,event:ev});
 });
 return expanded;
}

function hydratePriorPtyState(targetStep){
 if(!window.RECORDED_DEMO||!window.RECORDED_DEMO.steps||!window.RECORDED_DEMO.steps.pty)return;
 for(var s=1;s<targetStep;s++){
  if(doneSet.pty[s])continue;
  var list=window.RECORDED_DEMO.steps.pty[String(s)]||[];
  list.forEach(function(item){
   var ev=item.event;
   if(ev.kind==='phase'){if(ev.user)byId('live-agent-user').textContent=ev.user;ptyReady=true;}
   else if(ev.kind==='identity'){byId('live-agent-user').textContent=ev.user;byId('live-agent-session').textContent=ev.sessionId;}
   else if(ev.kind==='approval'){byId('live-extra-chip').textContent='Approval: '+ev.ticket+' ('+ev.scope+')';byId('live-extra-chip').classList.add('on');}
   else if(ev.kind==='services'){
    byId('live-spooler').textContent='Spooler: '+ev.spooler;byId('live-spooler').classList.add('on');
    byId('live-w32time').textContent='W32Time: '+ev.w32time;byId('live-w32time').classList.add('on');
   }
  });
 }
}

function playRecordedStep(kind,step){
 if(!window.RECORDED_DEMO||!window.RECORDED_DEMO.steps||!window.RECORDED_DEMO.steps[kind]){
  setStatus('NO RECORDING','err');return Promise.resolve(false);
 }
 if(kind==='pty'&&step===1&&doneSet.pty[6]){
  doneSet.pty={};
  byId('live-demo').querySelectorAll('.ev-chip').forEach(function(x){x.classList.remove('on','denied');});
  byId('live-demo').querySelectorAll('.live-demo-lifecycle span.done').forEach(function(x){x.classList.remove('done');});
 }
 if(kind==='pty'&&step>1)hydratePriorPtyState(step);
 var raw=window.RECORDED_DEMO.steps[kind][String(step)];
 if(!raw||!raw.length){setStatus('MISSING STEP','err');return Promise.resolve(false);}
 var seq=expandReplayItems(raw),elapsed=0;
 return new Promise(function(resolve){
  stepDoneWaiters.push(resolve);
  seq.forEach(function(item){
   elapsed+=item.delayMs;
   var tid=setTimeout(function(){applyDemoEvent(item.event);},elapsed);
   replayTimers.push(tid);
  });
 });
}

function closeLiveDemo(){
 closeAllSelects();
 stopAutoPlay();
 clearReplayTimers();
 resolveStepWaiters(false);
 if(liveOpen&&!replayMode)fetch('/api/cancel',{method:'POST',headers:{'X-MXC-Demo':'1'}}).catch(function(){});
 if(liveSource){liveSource.close();liveSource=null;}
 ptyReady=false;
 byId('live-demo').hidden=true;liveOpen=false;
}

async function connectLive(){
 if(!isLocalDemoHost()){initReplayMode();return;}
 var pill=byId('terminal-conn-pill');
 pill.textContent='live bridge · 127.0.0.1:8787';
 delete pill.dataset.mode;
 try{
  var meta=await fetch('/api/meta').then(function(r){if(!r.ok)throw new Error('Local demo metadata unavailable');return r.json();});
  replayMode=false;
  byId('live-host-user').textContent=meta.username;byId('live-host-session').textContent=meta.sessionId;
  if(liveSource)liveSource.close();
  liveSource=new EventSource('/api/events');
  liveSource.onmessage=function(event){applyDemoEvent(JSON.parse(event.data));};
 }catch(_error){
  initReplayMode();
 }
}

async function openLiveDemo(preferredKind,preferredStep){
 byId('live-demo').hidden=false;liveOpen=true;
 replayMode=!isLocalDemoHost();
 if(preferredKind)setKindVal(preferredKind);
 resetLive();
 if(preferredKind){setKindVal(preferredKind);updateLiveMode();}
 if(preferredStep){setStepVal(preferredStep);syncStepSelection();}
 if(replayMode){initReplayMode();return;}
 byId('live-demo-chapter').textContent='Choose Step 1 and click Send Step to execute live';
 await connectLive();
}

async function runSingleStep(kind,step){
 setStepVal(step);syncStepSelection();
 if(replayMode){return await playRecordedStep(kind,step);}
 try{
  var waitDone=new Promise(function(resolve){stepDoneWaiters.push(resolve);});
  var response=await fetch('/api/run-step?demo='+encodeURIComponent(kind)+'&step='+step,{method:'POST',headers:{'X-MXC-Demo':'1'}});
  if(!response.ok){
   var payload=await response.json().catch(function(){return {error:'Step request rejected'};});
   resolveStepWaiters(false);
   throw new Error(payload.error||'Step request failed');
  }
  return await waitDone;
 }catch(error){
  setStatus('BLOCKED','err');byId('live-demo-chapter').textContent=error.message||String(error);
  return false;
 }
}

async function runSelectedLiveStep(){
 if(stepRunning)return;
 stopAutoPlay();
 await runSingleStep(getKindVal(),getStepVal());
}

async function toggleAutoPlay(){
 if(autoPlaying){stopAutoPlay();return;}
 if(stepRunning)return;
 var kind=getKindVal(),total=STEP_GUIDES[kind].length;
 var allDone=true;
 for(var k=1;k<=total;k++){if(!doneSet[kind][k]){allDone=false;break;}}
 if(allDone||(!replayMode&&kind==='pty'&&!ptyReady&&getStepVal()>1)){
  if(!replayMode)await fetch('/api/cancel',{method:'POST',headers:{'X-MXC-Demo':'1'}}).catch(function(){});
  resetLive();setKindVal(kind);updateLiveMode();
  if(replayMode)initReplayMode();
 }
 autoPlaying=true;syncAutoPlayBtn();
 var start=getStepVal();
 if(kind==='pty'&&!ptyReady)start=1;
 for(var s=start;s<=total;s++){
  if(!autoPlaying)break;
  var ok=await runSingleStep(kind,s);
  if(!ok||!autoPlaying)break;
  if(s<total){
   await new Promise(function(r){var t=setTimeout(r,680);replayTimers.push(t);});
  }
 }
 stopAutoPlay();
}

async function sendPtyCommand(rawCmd){
 var cmd=(rawCmd||'').trim();
 if(!cmd||!ptyReady||stepRunning)return;
 if(!replayMode){
  await fetch('/api/pty-input',{method:'POST',headers:{'X-MXC-Demo':'1','Content-Type':'text/plain;charset=utf-8'},body:cmd}).catch(function(){});
  return;
 }
 var norm=cmd.toLowerCase().replace(/\\s+/g,' ');
 var agentUser=byId('live-agent-user').textContent||'R4-C2';
 var promptStr='PS C:\\\\Users\\\\'+agentUser+'\\\\Shared> ';
 if(norm.indexOf('triage-evidence.txt')!==-1&&!doneSet.pty[5]){
  applyDemoEvent({kind:'output',text:cmd+'\\n'});
  var t1=setTimeout(function(){
   applyDemoEvent({kind:'output',text:"Get-Content : Cannot find path 'C:\\\\Users\\\\"+agentUser+"\\\\Shared\\\\triage-evidence.txt' because it does not exist. (Run Step 5 first.)\\n"+promptStr});
  },140);
  replayTimers.push(t1);
  return;
 }
 var presets=(window.RECORDED_DEMO&&window.RECORDED_DEMO.ptyCommands)||[];
 var hit=presets.find(function(p){
  return p.label.toLowerCase()===norm||p.command.toLowerCase()===norm||(p.aliases||[]).indexOf(norm)!==-1;
 });
 if(hit){
  var seq=expandReplayItems(hit.events),wait=0;
  seq.forEach(function(item){
   wait+=item.delayMs;
   var tid=setTimeout(function(){applyDemoEvent(item.event);},wait);
   replayTimers.push(tid);
  });
 }else{
  applyDemoEvent({kind:'output',text:cmd+'\\n'});
  var t2=setTimeout(function(){
   applyDemoEvent({kind:'output',text:'[Recorded Replay] Offline presets: whoami, pwd, ls, Get-Service Spooler,W32Time, cat triage-evidence.txt\\n(Start "npm --prefix demo run demo:web" on Windows for arbitrary live commands.)\\n'+promptStr});
  },140);
  replayTimers.push(t2);
 }
}

async function switchScenario(newKind){
 if(getKindVal()===newKind)return;
 stopAutoPlay();clearReplayTimers();resolveStepWaiters(false);
 setKindVal(newKind);
 if(liveOpen&&!replayMode)await fetch('/api/cancel',{method:'POST',headers:{'X-MXC-Demo':'1'}}).catch(function(){});
 resetLive();
 setKindVal(newKind);
 updateLiveMode();
 if(!replayMode){
  var meta=await fetch('/api/meta').then(function(r){return r.json();}).catch(function(){return null;});
  if(meta){byId('live-host-user').textContent=meta.username;byId('live-host-session').textContent=meta.sessionId;}
 }else{initReplayMode();}
}

setupCustomSelect('cs-kind',function(val){
 if(!byId('cs-kind-btn').disabled)switchScenario(val);
});
setupCustomSelect('cs-step',function(val){
 if(!byId('cs-step-btn').disabled){stopAutoPlay();setStepVal(val);syncStepSelection();}
});
document.querySelectorAll('.mode-tab').forEach(function(tab){
 tab.addEventListener('click',function(){if(!byId('cs-kind-btn').disabled)switchScenario(tab.dataset.mode);});
});
byId('live-demo-run').addEventListener('click',runSelectedLiveStep);
byId('live-demo-autoplay').addEventListener('click',toggleAutoPlay);
byId('live-demo-reset').addEventListener('click',function(){
 var chosen=getKindVal();
 if(!replayMode)fetch('/api/cancel',{method:'POST',headers:{'X-MXC-Demo':'1'}}).catch(function(){});
 resetLive();setKindVal(chosen);updateLiveMode();
 if(!replayMode){fetch('/api/meta').then(function(r){return r.json();}).then(function(m){byId('live-host-user').textContent=m.username;byId('live-host-session').textContent=m.sessionId;}).catch(function(){});byId('live-demo-chapter').textContent='Session reset · Ready for Step 1';}
 else{initReplayMode();}
});
document.querySelectorAll('.quick-chip').forEach(function(chip){
 chip.addEventListener('click',function(){sendPtyCommand(chip.dataset.cmd);});
});
byId('live-pty-form').addEventListener('submit',async function(ev){
 ev.preventDefault();
 var inp=byId('live-pty-input'),cmd=inp.value.trim();
 if(!cmd||!ptyReady||stepRunning)return;
 inp.value='';
 await sendPtyCommand(cmd);
});
document.querySelectorAll('.live-demo-trigger').forEach(function(button){
 button.addEventListener('click',function(event){
  event.stopPropagation();
  openLiveDemo(button.dataset.demo||'pty',Number(button.dataset.step||1));
 });
});
byId('live-demo-close').addEventListener('click',closeLiveDemo);
})();
"""

def load_recorded_demo(img_dir):
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(img_dir)), 'demo-recordings.json'),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'assets', 'demo-recordings.json'),
    ]
    for path in candidates:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    return 'null'

def build(out, img_dir, maxw):
    out_dir = os.path.dirname(os.path.abspath(out))
    os.makedirs(out_dir, exist_ok=True)
    body = ''.join(slide_html(k, D.slides[k], img_dir, maxw) for k in sorted(D.slides))
    recorded_json = load_recorded_demo(img_dir)
    doc = f"""<!DOCTYPE html>
<html lang="en" dir="ltr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light">
<title>Guardrails with Sandboxing - Microsoft Execution Containers (MXC)</title>
<style>{CSS % (D.BG, D.INK, D.CY, D.CARD, D.CODEBG, D.EDGE, D.MUTED, D.AMB)}</style></head>
<body><div id="deck-progress" aria-hidden="true"></div><main id="stage">{body}</main>
{LIVE_DEMO}
<script>window.RECORDED_DEMO={recorded_json};</script>
<script>{JS}</script></body></html>"""
    open(out, 'w', encoding='utf-8').write(doc)
    print(f'Saved {out}  ({len(D.slides)} slides, {os.path.getsize(out)/1e6:.2f} MB)')

if __name__ == '__main__':
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    root = os.path.dirname(here)
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=os.path.join(root, 'dist', 'html', 'index.html'))
    ap.add_argument('--images', default=os.path.join(here, 'assets', 'images'))
    ap.add_argument('--maxw', type=int, default=1100, help='max image edge in px (smaller = smaller file)')
    a = ap.parse_args(); build(a.out, a.images, a.maxw)
