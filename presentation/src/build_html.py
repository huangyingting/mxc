#!/usr/bin/env python3
"""build_html.py - turns the same 28 slides into ONE self-contained HTML file
(images embedded, no external files). Uses the slide content from build_pptx.py.

Usage:  pip install pillow
        python -m presentation.src.build_html [--out dist/html/index.html]
Keys in the browser: <- / -> or Space to navigate, F = fullscreen,
Home/End = first/last, N = speaker notes, D = open interactive MXC Studio.
Print (Ctrl+P) gives one slide per page.
"""
import argparse, base64, html, io, json, os, re
from PIL import Image
import qrcode
try:
    from . import build_pptx as D
    from .i18n import ZH_DEMO_LAUNCHERS, ZH_TEXT_MAP
except ImportError:
    import build_pptx as D
    from i18n import ZH_DEMO_LAUNCHERS, ZH_TEXT_MAP

PX = 9525  # EMU per CSS px  (slide = 1280 x 720 px)
SVG_POINTS = {
    'rightArrow':'0,24 68,24 68,0 100,50 68,100 68,76 0,76',
    'leftArrow':'32,0 32,24 100,24 100,76 32,76 32,100 0,50',
    'downArrow':'24,0 76,0 76,62 100,62 50,100 0,62 24,62',
    'chevron':'0,0 82,0 100,50 82,100 0,100 18,50',
    'homePlate':'0,0 82,0 100,50 82,100 0,100',
}
FALLBACK = {
    'Inter Tight':'"Inter Tight","Segoe UI Variable Text","PingFang SC","Microsoft YaHei UI",-apple-system,BlinkMacSystemFont,sans-serif',
    'Segoe UI Variable Display':'"Inter Tight","Segoe UI Variable Display","PingFang SC","Microsoft YaHei UI",sans-serif',
    'Segoe UI Variable Text':'"Inter Tight","Segoe UI Variable Text","PingFang SC","Microsoft YaHei UI",sans-serif',
    'Calibri':'"Inter Tight","Segoe UI Variable Text",Calibri,"PingFang SC","Microsoft YaHei UI",sans-serif',
    'Consolas':'"Inconsolata",Consolas,"Cascadia Mono",monospace',
    'Inconsolata':'"Inconsolata",Consolas,"Cascadia Mono",monospace',
    'Plus Jakarta Sans':'"Inter Tight","Segoe UI Variable Text","Segoe UI",sans-serif',
    'Newsreader':'"Inter Tight","Segoe UI Variable Display",sans-serif',
    'JetBrains Mono':'"Inconsolata",Consolas,"Cascadia Mono",monospace',
    'Aptos Display':'"Inter Tight","Segoe UI Variable Display",sans-serif',
    'Aptos':'"Inter Tight","Segoe UI Variable Text",sans-serif',
    'Cascadia Code':'"Inconsolata",Consolas,"Cascadia Mono",monospace',
    'Georgia':'"Inter Tight","Segoe UI Variable Display",sans-serif',
    'Courier New':'"Inconsolata",Consolas,monospace',
    'Arial':'"Inter Tight",Arial,Helvetica,sans-serif',
}
e = lambda s: html.escape(s, quote=True)

def data_uri(path, maxw, cache={}):
    key = (path, maxw)
    if key not in cache:
        if path.lower().endswith('.svg'):
            with open(path, 'rb') as f:
                raw = f.read()
            cache[key] = 'data:image/svg+xml;base64,' + base64.b64encode(raw).decode()
        else:
            im = Image.open(path).convert('RGB')
            im.thumbnail((maxw, maxw), Image.Resampling.LANCZOS)
            b = io.BytesIO()
            im.save(b, 'JPEG', quality=88, optimize=True)
            cache[key] = 'data:image/jpeg;base64,' + base64.b64encode(b.getvalue()).decode()
    return cache[key]

def box(o):
    return f"left:{o['x']/PX:.1f}px;top:{o['y']/PX:.1f}px;width:{o['cx']/PX:.1f}px;height:{o['cy']/PX:.1f}px;"

def text_html(o):
    paras = o.get('paragraphs') or [{'text': o['text']}]
    raw_font = o.get('fontFamily', D.BF)
    font = FALLBACK.get(raw_font, FALLBACK['Inter Tight'])
    is_display = any(p.get('sizePt', o.get('sizePt', 16)) >= 29 for p in paras)
    if is_display and raw_font != D.MF:
        font = FALLBACK['Segoe UI Variable Display']
    if o.get('letterSpacingPt'):
        ls = f"letter-spacing:{o['letterSpacingPt']*4/3:.1f}px;"
    elif raw_font == D.MF:
        ls = "font-variation-settings:'wdth' 90;"
    elif is_display:
        ls = 'letter-spacing:-0.028em;'
    else:
        ls = 'letter-spacing:-0.008em;'
    ws = 'white-space:pre-wrap;' if raw_font == D.MF else ''
    default_gap_pt = o.get('spaceAfterPt', 2 if raw_font == D.MF else 6.5)
    out = []
    for idx, p in enumerate(paras):
        sz = p.get('sizePt', o.get('sizePt', 16)) * 4/3
        wt = 700 if p.get('bold', o.get('bold')) else (500 if sz >= 22 else 450)
        if raw_font == D.MF:
            lh = 1.48
        elif sz >= 34:
            lh = 1.14
        elif sz >= 22:
            lh = 1.26
        else:
            lh = 1.42
        mb = ''
        if idx < len(paras) - 1:
            gap_px = p.get('spaceAfterPt', default_gap_pt) * 4/3
            mb = f"margin-bottom:{gap_px:.1f}px;"
        st = (f"font-size:{sz:.1f}px;color:#{p.get('color', o.get('color', D.INK))};"
              f"font-weight:{wt};line-height:{lh};{ws}{mb}"
              f"text-align:{p.get('align', o.get('align','center'))};")
        inner = '<br>'.join(e(l) for l in p['text'].split('\n'))
        if p.get('href'):
            inner = f'<a class="ref-link" href="{e(p["href"])}" target="_blank" rel="noreferrer">{inner}</a>'
        out.append(f'<p style="{st}">{inner}</p>')
    return f'<div class="t" style="font-family:{e(font)};{ls}">' + ''.join(out) + '</div>'

REVEAL_ITEMS = {
    5: (2, 6, 3),
    7: (2, 5, 5),
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

def qr_badge_html(pages_url, left_px=77.7, top_px=440.0, lang='en'):
    clean_url = pages_url.strip()
    if not clean_url:
        return ''
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=1,
        border=2,
    )
    qr.add_data(clean_url)
    qr.make(fit=True)
    matrix = qr.get_matrix()
    n = len(matrix)
    cmds = ''.join(
        f'M{x},{y}h1v1h-1z'
        for y, row in enumerate(matrix)
        for x, cell in enumerate(row)
        if cell
    )
    display_url = re.sub(r'^https?://', '', clean_url).rstrip('/')
    kicker = ZH_TEXT_MAP.get('SCAN FOR LIVE SLIDES & DEMO', 'SCAN FOR LIVE SLIDES & DEMO') if lang == 'zh' else 'SCAN FOR LIVE SLIDES & DEMO'
    svg = (
        f'<svg class="pages-qr-svg" xmlns="http://www.w3.org/2000/svg" '
        f'viewBox="0 0 {n} {n}" shape-rendering="crispEdges" role="img" '
        f'aria-label="QR code for {e(clean_url)}">'
        f'<rect width="{n}" height="{n}" fill="#FFFFFF"/>'
        f'<path d="{cmds}" fill="#1E242B"/></svg>'
    )
    return (
        f'<a class="a pages-qr-card" style="left:{left_px:.1f}px;top:{top_px:.1f}px;" '
        f'href="{e(clean_url)}" target="_blank" rel="noreferrer">'
        f'{svg}'
        f'<span class="pages-qr-meta">'
        f'<span class="pages-qr-kicker">{e(kicker)}</span>'
        f'<strong class="pages-qr-url">{e(display_url)}</strong>'
        f'</span></a>'
    )

def slide_html(k, ops, img_dir, maxw, pages_url=None, lang='en'):
    parts = []
    for operation_index, o in enumerate(ops):
        classes, reveal = reveal_attrs(k, operation_index)
        if o['op'] == 'add_image':
            is_full = o['x'] == 0 and o['y'] == 0 and o['cx'] == D.W and o['cy'] == D.H
            is_side = o['y'] == 0 and o['cy'] == D.H
            fit_mode = o.get('fit', 'cover')
            if fit_mode == 'contain':
                extra_cls = ' contain-img'
            else:
                extra_cls = ' bleed' if (is_full or is_side) else ' framed-img'
            src = data_uri(os.path.join(img_dir, os.path.basename(o['path'])), maxw)
            parts.append(f'<img class="{classes}{extra_cls}"{reveal} style="{box(o)}" src="{src}" alt="{e(o.get("altText",""))}">')
        elif o['op'] == 'add_text_box':
            fn_cls = ' page-num' if o.get('name') == 'footer_num' else ''
            parts.append(f'<div class="{classes}{fn_cls}"{reveal} style="{box(o)}">{text_html(o)}</div>')
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
                if is_card and fill == D.CODEBG:
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
    launchers = ZH_DEMO_LAUNCHERS if lang == 'zh' else DEMO_LAUNCHERS
    if k in launchers:
        demo_kind, demo_step, label = launchers[k]
        action = (
            f'<button class="live-demo-trigger" type="button" '
            f'data-demo="{demo_kind}" data-step="{demo_step}">'
            f'<span class="pulse-dot" aria-hidden="true"></span><span>{e(label)}</span>'
            f'<span class="trigger-arrow" aria-hidden="true">↗</span></button>'
        )
    qr_html = qr_badge_html(pages_url, left_px=77.7, top_px=440.0, lang=lang) if (k == 0 and pages_url) else ''
    slide_lbl = f'第 {k+1} 页' if lang == 'zh' else f'Slide {k+1}'
    return f'<section class="slide" data-n="{k+1}" data-reveals="{reveal_count}" aria-label="{slide_lbl}">{"".join(parts)}{qr_html}{action}</section>'

def presenter_ui_html(lang='en'):
    if lang == 'zh':
        return """
<div id="deck-hud" class="deck-hud" aria-label="讲者工具栏">
  <span class="hud-count" id="hud-count"></span>
  <button type="button" class="hud-btn" id="hud-notes-btn" title="展开/收起中文讲者备注 (N)" aria-controls="notes-drawer" aria-expanded="false">备注 <kbd>N</kbd></button>
  <button type="button" class="hud-btn hud-btn-accent" id="hud-demo-btn" title="打开交互式 MXC 工作台 (D)">工作台 <kbd>D</kbd></button>
</div>

<aside id="notes-drawer" class="notes-drawer" aria-label="讲者备注" hidden>
  <div class="notes-head">
    <strong id="notes-title" lang="zh-CN">讲者备注</strong>
    <button type="button" class="notes-close" id="notes-close" aria-label="关闭讲者备注">✕</button>
  </div>
  <div class="notes-body" id="notes-body" lang="zh-CN"></div>
</aside>
"""
    return """
<div id="deck-hud" class="deck-hud" aria-label="Presenter toolbar">
  <span class="hud-count" id="hud-count"></span>
  <button type="button" class="hud-btn" id="hud-notes-btn" title="Toggle English speaker notes (N)" aria-controls="notes-drawer" aria-expanded="false">Notes <kbd>N</kbd></button>
  <button type="button" class="hud-btn hud-btn-accent" id="hud-demo-btn" title="Open Interactive MXC Studio (D)">Studio <kbd>D</kbd></button>
</div>

<aside id="notes-drawer" class="notes-drawer" aria-label="Speaker notes" hidden>
  <div class="notes-head">
    <strong id="notes-title" lang="en">Speaker Notes</strong>
    <button type="button" class="notes-close" id="notes-close" aria-label="Close speaker notes">✕</button>
  </div>
  <div class="notes-body" id="notes-body" lang="en"></div>
</aside>
"""

LIVE_DEMO = """
<section id="live-demo" class="live-demo" role="dialog" aria-modal="true" aria-label="Interactive MXC Demo" hidden>
  <div class="studio-shell">
    <header class="studio-top">
      <div class="bar-select-group">
        <div class="custom-select" id="cs-kind" data-value="process">
          <button type="button" class="cs-trigger" id="cs-kind-btn" aria-haspopup="listbox" aria-expanded="false" aria-label="Demo">
            <span class="cs-kicker">DEMO</span>
            <span class="cs-label" id="cs-kind-label">Demo 1 · ProcessContainer</span>
            <svg class="cs-chevron" viewBox="0 0 16 16" aria-hidden="true"><path d="M4 6l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
          </button>
          <ul class="cs-menu" id="cs-kind-menu" role="listbox" tabindex="-1" hidden>
            <li role="option" data-val="process" aria-selected="true" class="cs-option selected">
              <span class="cs-opt-main">Demo 1 · ProcessContainer</span>
              <span class="cs-opt-sub">Files &amp; network guardrails · 5 steps</span>
            </li>
            <li role="option" data-val="pty" aria-selected="false" class="cs-option">
              <span class="cs-opt-main">Demo 2 · IsolationSession PTY</span>
              <span class="cs-opt-sub">Stateful operator session · 6 steps</span>
            </li>
            <li role="option" data-val="microvm" aria-selected="false" class="cs-option">
              <span class="cs-opt-main">Demo 3 · Hyperlight MicroVM</span>
              <span class="cs-opt-sub">Hardware unikernel sandbox · 5 steps</span>
            </li>
          </ul>
        </div>

        <div class="custom-select cs-step-select" id="cs-step" data-value="1">
          <button type="button" class="cs-trigger" id="cs-step-btn" aria-haspopup="listbox" aria-expanded="false" aria-label="Step">
            <span class="cs-kicker cs-kicker-accent">STEP</span>
            <span class="cs-label" id="cs-step-label">1 · Discover Backends &amp; Probe Tier</span>
            <svg class="cs-chevron" viewBox="0 0 16 16" aria-hidden="true"><path d="M4 6l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
          </button>
          <ul class="cs-menu cs-step-menu" id="cs-step-menu" role="listbox" tabindex="-1" hidden></ul>
        </div>

        <button id="live-demo-run" class="primary-run-btn" type="button">
          <span id="live-run-label">Run Step 1</span>
          <span aria-hidden="true">▶</span>
        </button>
        <button id="live-demo-autoplay" class="ghost-btn autoplay-btn" type="button" title="Play all steps sequentially">
          <span id="live-autoplay-label">Auto-Play</span>
        </button>
      </div>

      <div class="studio-status-group">
        <span class="status-badge" id="live-demo-status">Replay · Win11</span>
        <button id="live-demo-reset" class="ghost-btn" type="button" title="Reset current demo">Reset</button>
        <button id="live-demo-close" class="close-btn" type="button" aria-label="Close demo">✕</button>
      </div>
    </header>

    <div class="studio-grid">
      <section class="terminal-card">
        <div class="terminal-chrome">
          <div class="chrome-dots" aria-hidden="true"><span></span><span></span><span></span></div>
          <div class="chrome-title"><strong id="terminal-mode-title">ProcessContainer</strong><span id="terminal-mode-sub">base-container · one-shot policy</span></div>
          <div class="term-legend" aria-label="Color legend">
            <span class="lg-item lg-pol">Policy</span>
            <span class="lg-item lg-run">Command</span>
            <span class="lg-item lg-ok">✓ Pass</span>
            <span class="lg-item lg-err">✕ Blocked</span>
            <span class="lg-item lg-log">Log</span>
          </div>
        </div>
        <pre class="studio-terminal" id="live-demo-output" role="log" aria-live="polite"></pre>
        <form class="terminal-composer" id="live-pty-form" hidden>
          <span class="composer-prompt">PS session&gt;</span>
          <input id="live-pty-input" type="text" placeholder="Type a live PowerShell command (e.g. whoami, pwd, ls) and press Enter..." autocomplete="off">
          <button type="submit" id="live-pty-send">Send</button>
        </form>
      </section>

      <aside class="explainer-rail">
        <article class="narrative-card" id="narrative-card">
          <div class="narrative-top">
            <span class="step-badge" id="explain-step-badge">STEP 1 / 5</span>
            <span class="boundary-tag" id="explain-boundary-tag">Capability Discovery</span>
          </div>
          <h3 id="explain-title">Runtime Discovery &amp; Probe</h3>
          <p class="explain-summary" id="explain-what"></p>
          <div class="explain-code-wrap">
            <div class="policy-kicker">ACTIVE MXC POLICY</div>
            <pre id="explain-code"></pre>
          </div>
        </article>

        <section class="compare-card">
          <div class="section-head">
            <strong id="compare-heading">Isolation Boundary</strong>
            <span id="compare-sub">Host vs. Contained</span>
          </div>
          <div class="identity-versus">
            <div class="id-box host">
              <div class="id-role">HOST</div>
              <div class="id-name" id="live-host-user">—</div>
              <div class="id-meta">Session <strong id="live-host-session">—</strong></div>
            </div>
            <div class="vs-medallion" aria-hidden="true">≠</div>
            <div class="id-box agent">
              <div class="id-role" id="agent-role-label">SANDBOX</div>
              <div class="id-name" id="live-agent-user">—</div>
              <div class="id-meta"><span id="agent-meta-label">Session </span><strong id="live-agent-session">—</strong></div>
            </div>
          </div>
          <div class="evidence-chips">
            <span class="ev-chip" id="live-spooler">Files: pending</span>
            <span class="ev-chip" id="live-w32time">Network: pending</span>
            <span class="ev-chip" id="live-extra-chip">Secret: locked</span>
          </div>
        </section>
      </aside>
    </div>
  </div>
</section>
"""

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inconsolata:wdth,wght@50..200,500..700&family=Inter+Tight:ital,wght@0,400..800;1,400..600&display=swap');
:root{
  --bg:#%s;--ink:#%s;--accent:#%s;--card:#%s;--codebg:#%s;--edge:#%s;--muted:#%s;--warn:#%s;
  --surface:#fbf8f2;--surface-strong:#efe9dc;--ink-soft:#4a463f;
  --font-sans:'Inter Tight','Segoe UI Variable Text','PingFang SC','Microsoft YaHei UI',system-ui,sans-serif;
  --font-display:'Inter Tight','Segoe UI Variable Display','PingFang SC','Microsoft YaHei UI',system-ui,sans-serif;
  --font-mono:'Inconsolata','Consolas','Cascadia Mono',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:100%%;height:100%%;background:var(--bg);color:var(--ink);font-family:var(--font-sans);letter-spacing:-0.006em;overflow:hidden;overscroll-behavior:none;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
code,pre,kbd,.pages-qr-kicker,.pages-qr-url{font-variation-settings:"wdth" 90}
#stage{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;overflow:hidden;background:radial-gradient(circle at 18%% 12%%,rgba(255,255,255,.55),transparent 48%%),var(--bg)}
.slide{position:absolute;width:1280px;height:720px;background:linear-gradient(180deg,rgba(255,255,255,.28) 0%%,rgba(255,255,255,0) 38%%),var(--bg);overflow:hidden;display:none;transform-origin:center center}
.slide.on{display:block;animation:slideFade .24s cubic-bezier(.2,.8,.2,1)}
@keyframes slideFade{from{opacity:.88}to{opacity:1}}
.a{position:absolute}
.surface-card{box-shadow:0 10px 28px rgba(38,33,26,.055),inset 0 1px 0 rgba(255,255,255,.55)}
.surface-code{border-radius:14px;box-shadow:0 12px 30px rgba(32,45,58,.06),inset 0 1px 0 rgba(255,255,255,.8)}
.framed-img{border-radius:14px;box-shadow:0 14px 34px rgba(35,30,22,.10);object-fit:cover}
.contain-img{border-radius:10px;object-fit:contain!important}
.reveal{opacity:0;visibility:hidden;transform:translateY(10px);transition:opacity .28s ease,transform .28s cubic-bezier(.2,.8,.2,1),visibility .28s}
.reveal.shown{opacity:1;visibility:visible;transform:translateY(0)}
.svg-shape svg{position:absolute;inset:0;width:100%%;height:100%%;overflow:visible;filter:drop-shadow(0 4px 10px rgba(35,48,62,.08))}
.svg-shape>.t{z-index:1}
img.a{object-fit:cover}
.t{position:absolute;inset:0;padding:10px 14px;display:flex;flex-direction:column;justify-content:center}
.a[style*="border-radius:50%%"]>.t{padding:0;align-items:center;text-align:center}
.page-num{z-index:6}.page-num>.t{padding:2px 8px;align-items:flex-end}.page-num>.t>p{background:rgba(243,238,228,.82);padding:2px 7px;border-radius:999px}
.lang-switch{position:fixed;left:12px;bottom:10px;z-index:46;display:inline-flex;align-items:center;gap:2px;padding:3px;border-radius:999px;background:rgba(251,248,242,.86);backdrop-filter:blur(8px);border:1px solid rgba(201,192,174,.78);box-shadow:0 4px 14px rgba(35,30,22,.07);opacity:.55;transition:opacity .18s ease}
.lang-switch:hover,.lang-switch:focus-within{opacity:1}
.lang-pill{border:0;background:transparent;color:var(--muted);text-decoration:none;border-radius:999px;padding:4px 9px;font:700 10.5px/1 var(--font-mono);letter-spacing:.03em;cursor:pointer;transition:background .14s,color .14s}
.lang-pill:hover{color:var(--ink)}
.lang-pill.active{background:var(--accent);color:#fff;pointer-events:none}
.ref-link{color:inherit;text-decoration:underline;text-decoration-color:rgba(63,96,126,.42);text-underline-offset:3px;transition:color .14s ease,text-decoration-color .14s ease}
.ref-link:hover{color:#1f384e;text-decoration-color:currentColor}
.pages-qr-card{display:inline-flex;align-items:center;gap:18px;padding:12px 22px 12px 12px;border-radius:16px;background:rgba(251,248,242,.95);border:1.5px solid var(--edge);box-shadow:0 12px 32px rgba(35,30,22,.09);text-decoration:none;color:var(--ink);transition:transform .16s ease,border-color .16s ease,box-shadow .16s ease;z-index:5}
.pages-qr-card:hover{transform:translateY(-1px);border-color:var(--accent);box-shadow:0 16px 36px rgba(35,30,22,.13)}
.pages-qr-svg{width:176px;height:176px;border-radius:10px;border:1px solid rgba(201,192,174,.75);flex-shrink:0;display:block;background:#fff}
.pages-qr-meta{display:flex;flex-direction:column;gap:7px;min-width:0;max-width:280px}
.pages-qr-kicker{font:700 11.5px/1.2 var(--font-mono);letter-spacing:.07em;color:var(--accent)}
.pages-qr-url{font:700 15.5px/1.3 var(--font-mono);color:var(--ink);word-break:break-all}

#deck-progress{position:fixed;left:0;top:0;height:3px;width:0;z-index:40;background:linear-gradient(90deg,var(--accent),#5d8aa8);transition:width .25s ease;pointer-events:none}

.deck-hud{position:fixed;left:50%%;bottom:8px;transform:translateX(-50%%);z-index:45;display:inline-flex;align-items:center;gap:6px;padding:3px 7px;border-radius:999px;background:rgba(251,248,242,.82);backdrop-filter:blur(8px);border:1px solid rgba(201,192,174,.75);box-shadow:0 4px 14px rgba(35,30,22,.07);opacity:.38;transition:opacity .18s ease}
.deck-hud:hover,.deck-hud:focus-within,.deck-hud.notes-open{opacity:1}
.hud-count{font:700 10.5px/1 var(--font-mono);color:var(--muted);padding:0 5px}
.hud-btn{border:1px solid var(--edge);background:#fff;color:var(--ink);border-radius:999px;padding:4px 9px;font:600 11px/1 var(--font-sans);cursor:pointer;display:inline-flex;align-items:center;gap:4px}
.hud-btn:hover{border-color:var(--accent);color:var(--accent)}
.hud-btn-accent{background:var(--accent);border-color:var(--accent);color:#fff}
.hud-btn-accent:hover{color:#fff;filter:brightness(1.08)}
.hud-btn kbd{font:700 9.5px/1 var(--font-mono);opacity:.8}

.notes-drawer{position:fixed;left:50%%;bottom:42px;transform:translateX(-50%%);width:min(640px,calc(100vw - 28px));max-height:42vh;z-index:55;background:rgba(251,248,242,.97);backdrop-filter:blur(12px);border:1px solid var(--edge);border-top:3px solid var(--accent);border-radius:14px;box-shadow:0 18px 42px rgba(28,24,18,.16);display:flex;flex-direction:column;overflow:hidden}
.notes-drawer[hidden]{display:none!important}
.notes-head{display:flex;align-items:center;justify-content:space-between;padding:9px 14px;border-bottom:1px solid var(--edge);font:700 12.5px/1 var(--font-sans);color:var(--accent)}
.notes-close{border:0;background:transparent;color:var(--muted);font-size:14px;cursor:pointer;padding:2px 6px;border-radius:6px}
.notes-close:hover{background:var(--card);color:var(--ink)}
.notes-body{padding:12px 15px;overflow-y:auto;font:450 13.5px/1.65 var(--font-sans);color:var(--ink);white-space:pre-wrap}

.live-demo-trigger{position:absolute;right:48px;top:48px;z-index:12;display:inline-flex;align-items:center;gap:9px;border:1px solid rgba(255,255,255,.28);border-radius:999px;padding:10px 18px;background:linear-gradient(135deg,var(--accent) 0%%,#2c4963 100%%);color:#fbf8f2;font:600 13px/1 var(--font-sans);letter-spacing:.01em;cursor:pointer;box-shadow:0 10px 26px rgba(44,73,99,.26);transition:transform .18s ease,box-shadow .18s ease,filter .18s ease}
.live-demo-trigger:hover{transform:translateY(-2px);box-shadow:0 14px 32px rgba(44,73,99,.34);filter:brightness(1.05)}
.pulse-dot{width:8px;height:8px;border-radius:50%%;background:#7ee2b8;box-shadow:0 0 0 4px rgba(126,226,184,.24);animation:pulseDot 1.8s infinite}
@keyframes pulseDot{0%%,100%%{box-shadow:0 0 0 3px rgba(126,226,184,.28)}50%%{box-shadow:0 0 0 7px rgba(126,226,184,.06)}}
.trigger-arrow{font-size:14px;opacity:.88}

.live-demo{position:fixed;inset:0;z-index:120;background:radial-gradient(circle at 12%% 0%%,#faf7f0 0%%,var(--bg) 58%%);color:var(--ink);font-family:var(--font-sans);padding:clamp(14px,1.8vw,24px)}
.live-demo[hidden]{display:none!important}
.studio-shell{height:100%%;max-width:1540px;margin:0 auto;display:grid;grid-template-rows:auto minmax(0,1fr);gap:14px}
.studio-top{display:flex;align-items:center;justify-content:space-between;gap:8px;background:var(--surface);border:1px solid var(--edge);border-radius:14px;padding:8px 11px;box-shadow:0 6px 20px rgba(35,30,22,.05);flex-wrap:nowrap;position:relative;z-index:30}
.bar-select-group{display:flex;align-items:center;gap:7px;min-width:0;flex:1}
.studio-status-group{display:flex;align-items:center;gap:7px;flex-shrink:0}
.status-badge{border-radius:999px;padding:6px 11px;font:600 11px/1 var(--font-mono);background:var(--card);color:var(--muted);border:1px solid var(--edge)}
.status-badge[data-tone="running"]{background:#fef4ea;border-color:#d48a42;color:#9a5313}
.status-badge[data-tone="ok"]{background:#e7f4ed;border-color:#286b4b;color:#1f583d}
.status-badge[data-tone="err"]{background:#faece8;border-color:var(--warn);color:var(--warn)}
.ghost-btn{border:1px solid var(--edge);background:#fff;color:var(--ink);border-radius:9px;padding:8px 13px;font:600 12px/1 var(--font-sans);cursor:pointer;transition:background .14s,border-color .14s}
.ghost-btn:hover{background:var(--card);border-color:var(--accent)}
.close-btn{width:32px;height:32px;border-radius:9px;border:1px solid var(--edge);background:#fff;color:var(--ink);font-size:14px;cursor:pointer;display:grid;place-items:center;transition:all .14s}
.close-btn:hover{background:var(--warn);color:#fff;border-color:var(--warn)}

.custom-select{position:relative;min-width:185px;max-width:245px;flex:1 1 215px}
.cs-step-select{min-width:210px;max-width:290px;flex:1.2 1 260px}
.cs-trigger{width:100%%;display:flex;align-items:center;gap:8px;border:1px solid var(--edge);border-radius:9px;background:#fff;color:var(--ink);padding:6px 10px 6px 7px;font:600 12.5px/1.2 var(--font-sans);cursor:pointer;transition:border-color .16s ease,box-shadow .16s ease}
.cs-step-select .cs-trigger{border-color:rgba(63,96,126,.45)}
.cs-trigger:hover:not(:disabled){border-color:var(--accent);box-shadow:0 3px 10px rgba(63,96,126,.10)}
.cs-trigger[aria-expanded="true"]{border-color:var(--accent);box-shadow:0 0 0 3px rgba(63,96,126,.14)}
.cs-trigger:disabled{opacity:.52;cursor:wait}
.cs-kicker{background:var(--card);color:var(--muted);border-radius:5px;padding:4px 6px;font:800 9px/1 var(--font-mono);letter-spacing:.08em;flex-shrink:0}
.cs-kicker-accent{background:rgba(63,96,126,.13);color:var(--accent)}
.cs-label{flex:1;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cs-chevron{width:14px;height:14px;color:var(--muted);flex-shrink:0;transition:transform .18s ease}
.cs-trigger[aria-expanded="true"] .cs-chevron{transform:rotate(180deg);color:var(--accent)}
.cs-menu{position:absolute;left:0;top:calc(100%% + 6px);min-width:100%%;width:max-content;max-width:380px;max-height:320px;overflow-y:auto;list-style:none;background:#fffdf9;border:1px solid var(--edge);border-radius:11px;padding:5px;box-shadow:0 18px 40px rgba(24,32,40,.16);z-index:80;animation:menuPop .14s cubic-bezier(.2,.8,.2,1)}
.cs-menu[hidden]{display:none!important}
@keyframes menuPop{from{opacity:0;transform:translateY(-4px) scale(.985)}to{opacity:1;transform:translateY(0) scale(1)}}
.cs-option{display:flex;flex-direction:column;gap:2px;padding:8px 10px;border-radius:8px;cursor:pointer;transition:background .12s ease}
.cs-option:hover,.cs-option.focused{background:var(--bg)}
.cs-option.selected{background:rgba(63,96,126,.11)}
.cs-opt-row{display:flex;align-items:center;justify-content:space-between;gap:10px}
.cs-opt-main{font:600 12.5px/1.25 var(--font-sans);color:var(--ink)}
.cs-option.selected .cs-opt-main{color:var(--accent);font-weight:700}
.cs-opt-sub{font:500 10.5px/1.25 var(--font-mono);color:var(--muted)}
.cs-opt-check{font:700 10.5px/1 var(--font-mono);color:#286b4b}
.primary-run-btn{display:inline-flex;align-items:center;gap:7px;border:0;border-radius:9px;background:linear-gradient(135deg,var(--accent) 0%%,#2a455d 100%%);color:#fff;padding:8px 16px;font:700 12.5px/1 var(--font-sans);cursor:pointer;box-shadow:0 5px 14px rgba(44,73,99,.22);transition:transform .14s ease,opacity .14s}
.primary-run-btn:hover:not(:disabled){transform:translateY(-1px)}
.primary-run-btn:disabled{opacity:.52;cursor:wait}
.autoplay-btn.active{background:#fef4ea;border-color:#d48a42;color:#9a5313}

.studio-grid{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(0,.9fr);gap:14px;min-height:0}
.terminal-card{display:flex;flex-direction:column;min-height:0;min-width:0;background:#0d151d;border-radius:14px;border:1px solid #243442;overflow:hidden;box-shadow:0 16px 38px rgba(16,24,32,.20)}
.terminal-chrome{display:flex;align-items:center;gap:10px;padding:10px 14px;background:#15212c;border-bottom:1px solid #233342;color:#b8cad8;font:500 11.5px var(--font-mono);min-width:0;flex-shrink:0}
.chrome-dots{display:flex;gap:6px;flex-shrink:0}.chrome-dots span{width:9px;height:9px;border-radius:50%%;background:#3a4b59}.chrome-dots span:first-child{background:#d46a5e}.chrome-dots span:nth-child(2){background:#deb05b}.chrome-dots span:nth-child(3){background:#5fb384}
.chrome-title{display:flex;align-items:baseline;gap:10px;min-width:0;flex:1;overflow:hidden}.chrome-title strong{color:#f2f7fb;font-weight:600;white-space:nowrap}.chrome-title span{color:#8fa4b5;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.term-legend{display:flex;align-items:center;gap:6px;flex-shrink:0;flex-wrap:wrap}
.lg-item{font:700 10px/1 var(--font-mono);padding:3px 7px;border-radius:999px;border:1px solid transparent}
.lg-pol{color:#93c5fd;background:rgba(147,197,253,.12);border-color:rgba(147,197,253,.28)}
.lg-run{color:#fcd34d;background:rgba(252,211,77,.12);border-color:rgba(252,211,77,.28)}
.lg-ok{color:#4ade80;background:rgba(74,222,128,.13);border-color:rgba(74,222,128,.32)}
.lg-err{color:#f87171;background:rgba(248,113,113,.14);border-color:rgba(248,113,113,.34)}
.lg-log{color:#cbd5e1;background:rgba(203,213,225,.10);border-color:rgba(203,213,225,.22)}
.studio-terminal{flex:1;min-height:0;overflow:auto;margin:0;padding:14px 16px;color:#cbd5e1;font:450 12.2px/1.55 var(--font-mono);white-space:pre-wrap;word-break:break-word;scrollbar-color:#445a6c #0d151d}
.tl{display:block;padding:1px 6px;border-radius:4px}
.tl-hint{color:#94a3b8;font-style:italic}
.tl-step{color:#7cc4f8;font-weight:700;margin-top:10px;padding-top:7px;border-top:1px dashed rgba(124,196,248,.24)}
.tl-step:first-child{margin-top:0;padding-top:1px;border-top:0}
.tl-policy{color:#93c5fd;font-weight:700;margin-top:3px}
.tl-policy-sub{color:#bfdbfe}
.tl-run{color:#fcd34d;font-weight:700;margin-top:3px}
.tl-pass{color:#4ade80;font-weight:700;background:rgba(34,197,94,.13);border-left:3px solid #22c55e;padding:4px 9px;margin:4px 0 2px;border-radius:0 6px 6px 0}
.tl-fail{color:#f87171;font-weight:700;background:rgba(239,68,68,.15);border-left:3px solid #ef4444;padding:4px 9px;margin:4px 0 2px;border-radius:0 6px 6px 0}
.tl-err{color:#fca5a5}
.tl-why{color:#fde68a;padding-left:9px;margin-bottom:2px}
.tl-log{color:#cbd5e1}
.terminal-composer{display:flex;align-items:center;gap:8px;padding:8px 12px;background:#121d27;border-top:1px solid #223241;flex-shrink:0}
.terminal-composer[hidden]{display:none!important}
.composer-prompt{color:#7cc4f8;font:600 11.5px var(--font-mono);white-space:nowrap}
.terminal-composer input{flex:1;min-width:0;border:1px solid #2a3d4f;border-radius:7px;background:#091017;color:#f3f8fc;padding:7px 10px;font:450 12px var(--font-mono);outline:none}
.terminal-composer input:focus{border-color:#5da9dd}
.terminal-composer input:disabled{opacity:.48}
.terminal-composer button{border:0;border-radius:7px;background:#2e5b7d;color:#fff;padding:7px 13px;font:600 12px var(--font-sans);cursor:pointer}
.terminal-composer button:disabled{opacity:.42;cursor:not-allowed}

.explainer-rail{display:flex;flex-direction:column;gap:14px;min-height:0;min-width:0;overflow-y:auto;overflow-x:hidden}
.narrative-card,.compare-card{background:var(--surface);border:1px solid var(--edge);border-radius:14px;padding:16px 18px;box-shadow:0 6px 20px rgba(35,30,22,.04);min-width:0}
.narrative-card{border-top:3px solid var(--accent);display:flex;flex-direction:column;gap:9px}
.narrative-top{display:flex;align-items:center;justify-content:space-between;gap:8px;min-width:0}
.step-badge{background:rgba(63,96,126,.12);color:var(--accent);border-radius:999px;padding:4px 10px;font:800 10px/1 var(--font-mono);letter-spacing:.08em;flex-shrink:0}
.boundary-tag{color:var(--muted);font:600 11px/1 var(--font-mono);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.narrative-card h3{font:700 19px/1.2 var(--font-display);color:var(--ink)}
.explain-summary{color:var(--ink-soft);font-size:13.5px;line-height:1.5}
.explain-code-wrap{background:var(--codebg);border:1px solid var(--edge);border-radius:9px;padding:9px 12px;margin-top:2px;min-width:0}
.policy-kicker{font:800 9.5px/1 var(--font-mono);letter-spacing:.08em;color:var(--muted);margin-bottom:6px}
.explain-code-wrap pre{margin:0;font:600 11.2px/1.42 var(--font-mono);color:var(--accent);white-space:pre-wrap;word-break:break-word}

.compare-card{display:flex;flex-direction:column;gap:14px}
.section-head{display:flex;align-items:baseline;justify-content:space-between;gap:8px;min-width:0}
.section-head strong{font:700 14px var(--font-sans);color:var(--ink)}
.section-head span{font:500 11px var(--font-mono);color:var(--muted)}
.identity-versus{display:grid;grid-template-columns:1fr 32px 1fr;gap:10px;align-items:stretch}
.id-box{background:var(--card);border-radius:11px;padding:12px 14px;min-width:0;border-top:3px solid var(--ink)}
.id-box.agent{border-top-color:var(--accent);background:linear-gradient(180deg,#f2f6f9 0%%,var(--card) 100%%)}
.id-role{color:var(--muted);font:700 9.5px/1 var(--font-mono);letter-spacing:.08em}
.id-name{font:800 17px/1.2 var(--font-sans);color:var(--ink);margin:6px 0 3px;overflow-wrap:anywhere}
.id-box.agent .id-name{color:var(--accent)}
.id-meta{color:var(--muted);font-size:12px}
.vs-medallion{align-self:center;justify-self:center;width:30px;height:30px;border-radius:50%%;background:var(--accent);color:#fff;font:700 16px/1 var(--font-sans);display:grid;place-items:center}

.evidence-chips{display:flex;gap:7px;flex-wrap:wrap;padding-top:4px;border-top:1px solid var(--edge)}
.ev-chip{border:1px solid var(--edge);border-radius:999px;background:var(--bg);padding:5px 11px;color:var(--muted);font:600 11px/1 var(--font-mono)}
.ev-chip.on{background:#e7f4ed;border-color:#286b4b;color:#1f583d}
.ev-chip.denied{background:#faece8;border-color:var(--warn);color:var(--warn)}

@media(max-width:960px){
  .studio-top,.bar-select-group{flex-wrap:wrap}
  .studio-grid{grid-template-columns:1fr;overflow:auto}
  .terminal-card{min-height:300px}
}
@media print{@page{size:13.333in 7.5in;margin:0}
 html,body{overflow:visible;background:#fff}#deck-progress,.deck-hud,.notes-drawer,.live-demo,.live-demo-trigger,.lang-switch{display:none!important}
 #stage{position:static;display:block}.slide{position:relative;display:block!important;transform:none!important;box-shadow:none;page-break-after:always;break-after:page;print-color-adjust:exact;-webkit-print-color-adjust:exact}.reveal{opacity:1!important;visibility:visible!important;transform:none!important}}
"""
DECK_JS = """
(function(){
var S=[].slice.call(document.querySelectorAll('.slide')),i=0,steps=S.map(function(){return 0;});
var bar=document.getElementById('deck-progress');
var deck={S:S,getIndex:function(){return i;},isOverlayOpen:function(){return false;},onRender:null};
window.__DECK__=deck;
function fit(){var v=window.visualViewport,w=v?v.width:innerWidth,h=v?v.height:innerHeight,s=Math.min(w/1280,h/720);S.forEach(function(x){x.style.transform='scale('+s+')';});}
function toggleFullscreen(){if(document.fullscreenElement){document.exitFullscreen();}else if(document.documentElement.requestFullscreen){document.documentElement.requestFullscreen();}}
function renderSteps(){var shown=steps[i];S[i].querySelectorAll('.reveal').forEach(function(x){x.classList.toggle('shown',Number(x.dataset.reveal)<shown);});if(bar)bar.style.width=((i+1)/S.length*100).toFixed(2)+'%';if(deck.onRender)deck.onRender(i,S.length);}
function parseHash(){var h=location.hash.slice(1),m=/^(?:slide-)?(\\d+)$/.exec(h);return m?Number(m[1])-1:0;}
function syncLangLinks(){
 var h='#'+(i+1);
 document.querySelectorAll('a.lang-pill[data-base-href]').forEach(function(a){a.href=a.dataset.baseHref+h;});
}
function go(n){i=Math.max(0,Math.min(S.length-1,n));S.forEach(function(x,j){x.classList.toggle('on',j===i);});renderSteps();if(history.replaceState)history.replaceState(null,'','#'+(i+1));syncLangLinks();}
function next(){var total=Number(S[i].dataset.reveals||0);if(steps[i]<total){steps[i]++;renderSteps();}else{go(i+1);}}
function previous(){if(steps[i]>0){steps[i]--;renderSteps();}else{go(i-1);}}
document.getElementById('stage').addEventListener('click',function(ev){if(deck.isOverlayOpen()||ev.target.closest('button,a'))return;ev.clientX<innerWidth/3?previous():next();});
addEventListener('keydown',function(ev){
 if(ev.defaultPrevented||ev.altKey||ev.ctrlKey||ev.metaKey)return;
 var k=ev.key;
 if(ev.target&&ev.target.closest('input,textarea,select,[contenteditable="true"]')&&k!=='Escape')return;
 if(deck.isOverlayOpen())return;
 if(k===' '&&ev.target&&ev.target.closest('button,a'))return;
 if(k==='ArrowRight'||k==='ArrowDown'||k==='PageDown'||(k===' '&&!ev.shiftKey)){ev.preventDefault();next();}
 else if(k==='ArrowLeft'||k==='ArrowUp'||k==='PageUp'||(k===' '&&ev.shiftKey)){ev.preventDefault();previous();}
 else if(k==='Home'){ev.preventDefault();go(0);}
 else if(k==='End'){ev.preventDefault();go(S.length-1);}
 else if(k==='f'||k==='F'){ev.preventDefault();toggleFullscreen();}});
addEventListener('resize',fit);document.addEventListener('fullscreenchange',fit);if(window.visualViewport)window.visualViewport.addEventListener('resize',fit);
addEventListener('hashchange',function(){var n=parseHash();if(n!==i)go(n);});
fit();go(parseHash());
})();
"""

STUDIO_JS = """
(function(){
var deck=window.__DECK__;if(!deck)return;
var liveOpen=false;
function byId(id){return document.getElementById(id);}
deck.isOverlayOpen=function(){return liveOpen;};
addEventListener('keydown',function(ev){
 if(ev.altKey||ev.ctrlKey||ev.metaKey)return;
 if(liveOpen&&ev.key==='Escape'){
  ev.preventDefault();
  if(document.querySelector('.cs-menu:not([hidden])')){closeAllSelects();return;}
  closeLiveDemo();
 }
},true);

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
  {label:'1 · Provision Session & Open PTY',boundary:'OS Session',title:'Provision Isolated Operator Session',
   what:'Creates a fresh local Windows operator account, stages ticket INC-1042 into the shared workspace, and opens a persistent PowerShell PTY.',
   code:'{\\n  "containment": "isolation_session",\\n  "network": { "egress": "allow", "ingress": "allow" },\\n  "ui": "separate session desktop (SendInput blocked)"\\n}'},
  {label:'2 · Verify User & Profile Block',boundary:'SID Isolation',title:'Operator Allowed vs. Employee Profile Blocked',
   what:'Confirms separate Windows username/SessionId (PASS), then tries reading the host employee profile—blocked by Windows NTFS ACLs (BLOCKED).',
   code:'{\\n  "containment": "isolation_session",\\n  "agentUser": "temporary local Windows SID",\\n  "hostProfileAccess": "denied by OS user ACL"\\n}'},
  {label:'3 · Check Ticket & Scope',boundary:'Workspace Handoff',title:'Read Staged INC-1042 Approval',
   what:'Reads staged JSON files from the shared handoff folder and confirms read-only-diagnostics scope while keeping shell state alive.',
   code:'{\\n  "ephemeralWorkspace": ["incident-request.json", "incident-approval.json"],\\n  "scope": "read-only-diagnostics"\\n}'},
  {label:'4 · Diagnose Services (Stop Blocked)',boundary:'Privilege Boundary',title:'Read Service Health (PASS) vs. Stop-Service (BLOCKED)',
   what:'Reads live Spooler & W32Time status (PASS), then attempts Stop-Service Spooler—rejected by Windows service permissions (BLOCKED).',
   code:'{\\n  "operatorPrivilege": "standard non-admin user",\\n  "Get-Service": "allowed",\\n  "Stop-Service": "denied by OS"\\n}'},
  {label:'5 · Write Host-Verified Evidence',boundary:'Output Validation',title:'Persist Audit Evidence & Verify on Host',
   what:'Writes triage-evidence.txt in the handoff folder; the trusted Host independently validates ticket, operator, and service fields.',
   code:'{\\n  "handoffFile": "triage-evidence.txt",\\n  "hostValidation": ["ticket=INC-1042", "operator", "services"]\\n}'},
  {label:'6 · Exit & Deprovision Account',boundary:'Clean Teardown',title:'Close PTY & Delete Temporary User',
   what:'Exits PowerShell, stops the isolated Windows session, and deletes the temporary OS operator account.',
   code:'stopContainer(containerId)\\n→ deprovisionContainer(containerId)'}
 ],
 process:[
  {label:'1 · Discover & Probe Policy',boundary:'Capability Probe',title:'Verify Native ProcessContainer Policy',
   what:'Loads policy.json and runs getPlatformSupport() + probe(request) to confirm Windows kernel ACL and network filters are active.',
   code:'{\\n  "containment": "processcontainer",\\n  "filesystem": {\\n    "readonlyPaths": ["input"],\\n    "readwritePaths": ["output"],\\n    "deniedPaths": ["protected/denied.txt"]\\n  }\\n}'},
  {label:'2 · Allowed File Read / Write',boundary:'Least Privilege (PASS)',title:'Read input/ & Write output/report.txt → PASS',
   what:'Runs cmd /c type input\\\\request.txt > output\\\\report.txt. Because both folders are granted in policy, the command succeeds (exit 0).',
   code:'{\\n  "readonlyPaths": ["input"],\\n  "readwritePaths": ["output"],\\n  "deniedPaths": ["protected/denied.txt"]\\n}'},
  {label:'3 · Block Protected Secret',boundary:'Kernel ACL (BLOCKED)',title:'Read protected/denied.txt → BLOCKED (exit 1)',
   what:'Runs cmd /c type protected\\\\denied.txt under the exact same policy. Windows kernel ACL blocks the read with Access Denied (exit 1).',
   code:'{\\n  "readonlyPaths": ["input"],\\n  "readwritePaths": ["output"],\\n  "deniedPaths": ["protected/denied.txt"]\\n}'},
  {label:'4 · Allow Loopback Endpoint',boundary:'Loopback Allow (PASS)',title:'curl Local Listener with hostLoopback: "allow" → PASS',
   what:'Sets hostLoopback: "allow" (while internet egress stays "deny"). curl reaches 127.0.0.1 and receives HTTP 200 (exit 0).',
   code:'{\\n  "network": {\\n    "egress": { "default": "deny" },\\n    "ingress": { "hostLoopback": "allow" }\\n  }\\n}'},
  {label:'5 · Block Unapproved Network',boundary:'Loopback Deny (BLOCKED)',title:'Same curl with hostLoopback: "deny" → BLOCKED (exit 7)',
   what:'Flips only hostLoopback to "deny" and runs the exact same curl command—blocked immediately by the OS network stack (exit 7).',
   code:'{\\n  "network": {\\n    "egress": { "default": "deny" },\\n    "ingress": { "hostLoopback": "deny" }\\n  }\\n}'}
 ],
 microvm:[
  {label:'1 · Verify Warm Snapshots',boundary:'Experimental Opt-In',title:'Validate 1.1.0-alpha Hyperlight Policy',
   what:'Checks native hyperlight support and pre-warmed Unikraft unikernel memory snapshots (bash + python) for <100 ms restore.',
   code:'{\\n  "version": "1.1.0-alpha",\\n  "containment": "hyperlight",\\n  "hyperlight": { "runtime": "bash | python" }\\n}'},
  {label:'2 · Boot Unikraft Kernel',boundary:'Hardware Hypervisor',title:'Run uname Inside Unikraft VM → PASS',
   what:'Restores a hardware-virtualized Unikraft 5.15 x86_64 guest on Windows Hypervisor Platform—never sharing the Windows NT kernel.',
   code:'{\\n  "containment": "hyperlight",\\n  "hyperlight": { "runtime": "bash" },\\n  "process": { "commandLine": "uname -sr && ls /" }\\n}'},
  {label:'3 · Run Untrusted Python',boundary:'Fast VM Compute',title:'Evaluate AI-Generated Python 3.12 → PASS',
   what:'Executes untrusted invoice-reconciliation Python inside the isolated VM and returns structured JSON stdout to the host.',
   code:'{\\n  "containment": "hyperlight",\\n  "hyperlight": { "runtime": "python" },\\n  "process": { "timeout": 30000 }\\n}'},
  {label:'4 · Block Host & Socket Escape',boundary:'Zero-Surface Air Gap',title:'Probe Host C:\\\\, Env & TCP → ALL BLOCKED',
   what:'Runs adversarial Python attempting to read C:\\\\Windows, USERPROFILE, and a TCP socket—all 3 fail with OS/Python exceptions.',
   code:'{\\n  "containment": "hyperlight",\\n  "hostMounts": "none",\\n  "hostEnv": "isolated",\\n  "networkStack": "disabled (Errno 13)"\\n}'},
  {label:'5 · Verify Snapshot Reset',boundary:'Memory Wipe',title:'Write /tmp in Run A → Gone in Run B',
   what:'Run A writes /tmp/agent-residue.txt (PASS), then Run B tries to read it—FileNotFoundError (BLOCKED), proving instant memory reset.',
   code:'{\\n  "containment": "hyperlight",\\n  "snapshotRestore": "clean memory image on every call"\\n}'}
 ]
};

var liveNext={pty:1,process:1,microvm:1},doneSet={pty:{},process:{},microvm:{}},liveSource=null,ptyReady=false;
var KIND_LABELS={
 process:'Demo 1 · ProcessContainer',
 pty:'Demo 2 · IsolationSession PTY',
 microvm:'Demo 3 · Hyperlight MicroVM'
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

function renderExplain(kind,step,keepExplicitRun){
 var list=STEP_GUIDES[kind],idx=Math.max(0,Math.min(list.length-1,(Number(step)||1)-1)),cur=idx+1,g=list[idx];
 byId('explain-step-badge').textContent='STEP '+cur+' / '+list.length;
 byId('explain-boundary-tag').textContent=g.boundary;
 byId('explain-title').textContent=g.title;
 byId('explain-what').textContent=g.what;
 byId('explain-code').textContent=g.code;
 var nextTarget=(!keepExplicitRun&&doneSet[kind][cur]&&cur<list.length)?cur+1:cur;
 byId('live-demo-run').dataset.targetStep=String(nextTarget);
 byId('live-run-label').textContent=(nextTarget!==cur)?('Next: Step '+nextTarget):('Run Step '+cur);
}

function syncComposerState(){
 var kind=getKindVal(),inp=byId('live-pty-input'),btn=byId('live-pty-send'),form=byId('live-pty-form');
 if(kind!=='pty'||replayMode){
  form.hidden=true;
 }else{
  form.hidden=false;
  inp.disabled=!ptyReady||stepRunning;btn.disabled=!ptyReady||stepRunning;
  inp.placeholder=ptyReady
   ?'Live PTY active — type any PowerShell command (e.g. whoami, pwd, ls) and press Enter'
   :'Run Step 1 first to provision the live IsolationSession PowerShell PTY...';
 }
}

function syncStepSelection(explicitPick){
 var kind=getKindVal(),step=getStepVal();
 setStepVal(step);
 renderExplain(kind,step,!!explicitPick);
}

function idleModeLabel(){return replayMode?'Replay · Win11':'Live · 127.0.0.1:8787';}

function updateLiveMode(){
 var kind=getKindVal();
 setKindVal(kind);
 setStepVal(liveNext[kind]);
 var metaLbl=byId('agent-meta-label');
 if(kind==='pty'){
  byId('terminal-mode-title').textContent='IsolationSession · PowerShell PTY';
  byId('terminal-mode-sub').textContent='120 × 30 persistent session';
  byId('agent-role-label').textContent='ISOLATED SESSION';
  if(metaLbl)metaLbl.textContent='Session ';
  byId('compare-heading').textContent='Identity & Session Boundary';
  byId('compare-sub').textContent='Host vs. Operator';
  byId('live-agent-user').textContent='Pending';byId('live-agent-session').textContent='—';
  byId('live-spooler').textContent='Spooler: —';byId('live-w32time').textContent='W32Time: —';
  byId('live-extra-chip').textContent='INC-1042: staged';
 }else if(kind==='process'){
  byId('terminal-mode-title').textContent='ProcessContainer · One-Shot Policy';
  byId('terminal-mode-sub').textContent='base-container · files & loopback';
  byId('agent-role-label').textContent='PROCESS SANDBOX';
  if(metaLbl)metaLbl.textContent='Tier ';
  byId('compare-heading').textContent='Least-Privilege Boundary';
  byId('compare-sub').textContent='Deny by default';
  byId('live-agent-user').textContent='AppContainer';byId('live-agent-session').textContent='base-container';
  byId('live-spooler').textContent='Files: —';byId('live-w32time').textContent='Loopback: —';
  byId('live-extra-chip').textContent='Secret: locked';
 }else{
  byId('terminal-mode-title').textContent='Hyperlight · Unikraft MicroVM';
  byId('terminal-mode-sub').textContent='hardware hypervisor · bash & python 3.12';
  byId('agent-role-label').textContent='MICROVM GUEST';
  if(metaLbl)metaLbl.textContent='Warm boot ';
  byId('compare-heading').textContent='Hardware VM Boundary';
  byId('compare-sub').textContent='Separate kernel';
  byId('live-agent-user').textContent='Unikraft VM';byId('live-agent-session').textContent='~90 ms';
  byId('live-spooler').textContent='Kernel: —';byId('live-w32time').textContent='Escape: —';
  byId('live-extra-chip').textContent='Reset: —';
 }
 syncComposerState();
 syncStepSelection();
}

function isLocalDemoHost(){return ['127.0.0.1','localhost'].includes(location.hostname)&&location.protocol.indexOf('http')===0;}
var replayMode=!isLocalDemoHost(),replayTimers=[],autoPlaying=false,stepRunning=false,stepDoneWaiters=[],replayBannerShown=false;

function escHtml(s){
 return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function renderTerminalHtml(text){
 var lines=String(text||'').split('\\n'),inPolicy=false,html=[];
 for(var idx=0;idx<lines.length;idx++){
  var raw=lines[idx];
  if(/^(?:MICROVM )?STEP \\d+ COMPLETE/.test(raw)||/^MICROVM_BOOT_MS=\\d+$/.test(raw))continue;
  if(idx===lines.length-1&&raw==='')continue;
  var safe=escHtml(raw),cls='tl-log';
  if(/^# /.test(raw)){inPolicy=false;cls='tl-hint';}
  else if(/^──▶/.test(raw)){inPolicy=false;cls='tl-step';}
  else if(/^\\[POLICY\\]/.test(raw)){inPolicy=true;cls='tl-policy';}
  else if(inPolicy&&/^  \\S/.test(raw)){cls='tl-policy-sub';}
  else if(/^\\[RUN\\]/.test(raw)||/^PS [^>]+>/.test(raw)){inPolicy=false;cls='tl-run';}
  else if(/^✓ PASS/.test(raw)){inPolicy=false;cls='tl-pass';}
  else if(/^✕ BLOCKED/.test(raw)||/^\\[Error\\]/.test(raw)){inPolicy=false;cls='tl-fail';}
  else if(/^  stderr \\(OS\\):/.test(raw)){inPolicy=false;cls='tl-err';}
  else if(/^Why:/.test(raw)){inPolicy=false;cls='tl-why';}
  else{inPolicy=false;}
  html.push('<span class="tl '+cls+'">'+(safe||' ')+'</span>');
 }
 return html.join('');
}

function clearReplayTimers(){
 while(replayTimers.length)clearTimeout(replayTimers.pop());
 stepRunning=false;
}
function syncAutoPlayBtn(){
 var btn=byId('live-demo-autoplay'),lbl=byId('live-autoplay-label');
 if(!btn||!lbl)return;
 btn.classList.toggle('active',autoPlaying);
 lbl.textContent=autoPlaying?'Stop':'Auto-Play';
}
function stopAutoPlay(){
 autoPlaying=false;
 syncAutoPlayBtn();
}
function resolveStepWaiters(ok){
 while(stepDoneWaiters.length){var fn=stepDoneWaiters.shift();fn(ok);}
}

var rawLog='',displayLog='';
function flushTerminal(){
 var out=byId('live-demo-output');
 out.innerHTML=renderTerminalHtml(displayLog);
 out.scrollTop=out.scrollHeight;
}
function appendTerminal(chunk){
 displayLog+=chunk;
 flushTerminal();
}

function resetLive(){
 stopAutoPlay();
 clearReplayTimers();
 resolveStepWaiters(false);
 ptyReady=false;
 replayBannerShown=false;
 rawLog='';displayLog='';
 doneSet={pty:{},process:{},microvm:{}};
 liveNext={pty:1,process:1,microvm:1};
 byId('live-demo-output').innerHTML='';
 setControlsDisabled(false);
 setStatus(idleModeLabel());
 byId('live-demo').querySelectorAll('.ev-chip').forEach(function(x){x.classList.remove('on','denied');});
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
  if(/__WEB_APPROVAL_VAL__=/.test(value)){byId('live-extra-chip').textContent='INC-1042: read-only ✓';byId('live-extra-chip').classList.add('on');}
  if(/__WEB_EVIDENCE_VAL__=written/.test(value)){byId('live-extra-chip').textContent='Evidence: verified ✓';byId('live-extra-chip').classList.add('on');}
 }else if(kind==='process'){
  if(value.includes('STEP 2 COMPLETE')){byId('live-spooler').textContent='Files: ALLOWED ✓';byId('live-spooler').classList.add('on');}
  if(value.includes('STEP 3 COMPLETE')){byId('live-extra-chip').textContent='Secret: DENIED ✕';byId('live-extra-chip').classList.add('denied');}
  if(value.includes('STEP 4 COMPLETE')){byId('live-w32time').textContent='Loopback: ALLOWED ✓';byId('live-w32time').classList.add('on');}
  if(value.includes('STEP 5 COMPLETE')){byId('live-w32time').textContent='Loopback: BLOCKED ✕';byId('live-w32time').classList.remove('on');byId('live-w32time').classList.add('denied');}
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
   byId('live-spooler').textContent='Python 3.12: OK ✓';byId('live-spooler').classList.add('on');
  }
  if(value.includes('MICROVM STEP 4 COMPLETE')){
   byId('live-w32time').textContent='Escape: BLOCKED ✕';byId('live-w32time').classList.add('denied');
  }
  if(value.includes('MICROVM STEP 5 COMPLETE')){
   byId('live-extra-chip').textContent='Snapshot: CLEAN ✓';byId('live-extra-chip').classList.add('on');
  }
 }
}

function formatDisplayChunk(text){
 return text
  .replace(/Write-Output \\("__WEB_IDENTITY_" \\+ "VAL__=[^"]*"\\)/g,'Write-Output "$env:USERNAME | SessionId=$((Get-Process -Id $PID).SessionId)"')
  .replace(/; Write-Output \\("__WEB_APPROVAL_" \\+ "VAL__=[^"]*"\\)/g,'; "$($request.ticket) | $($approval.scope)"')
  .replace(/; Write-Output \\("__WEB_SERVICES_" \\+ "VAL__=[^"]*"\\)/g,'; "Spooler=$($spooler.Status) | W32Time=$($time.Status)"')
  .replace(/; Write-Output \\("__WEB_EVIDENCE_" \\+ "VAL__=written"\\)/g,'')
  .replace(/__WEB_IDENTITY_VAL__=([^|\\n]+)\\|(\\d+)/g,'$1 | SessionId=$2')
  .replace(/__WEB_APPROVAL_VAL__=([^|\\n]+)\\|([^\\n]+)/g,'$1 | $2')
  .replace(/__WEB_SERVICES_VAL__=([^|\\n]+)\\|([^\\n]+)/g,'Spooler=$1 | W32Time=$2')
  .replace(/__WEB_EVIDENCE_VAL__=written/g,'Saved .\\\\triage-evidence.txt (verified by host)');
}

function applyDemoEvent(data){
 if(replayBannerShown&&data.kind==='step-start'){rawLog='';displayLog='';replayBannerShown=false;}
 if(data.kind==='phase'&&data.text){
  var ptxt=cleanLive(data.text);rawLog+=ptxt;
  appendTerminal(formatDisplayChunk(ptxt));
  if(data.user)byId('live-agent-user').textContent=data.user;
  ptyReady=true;syncComposerState();
 }else if(data.kind==='output'){
  var ctxt=cleanLive(data.text);rawLog+=ctxt;
  appendTerminal(formatDisplayChunk(ctxt));inspectLive(rawLog);
 }else if(data.kind==='identity'){
  byId('live-agent-user').textContent=data.user;byId('live-agent-session').textContent=data.sessionId;
 }else if(data.kind==='approval'){
  byId('live-extra-chip').textContent='INC-1042: read-only ✓';byId('live-extra-chip').classList.add('on');
 }else if(data.kind==='services'){
  byId('live-spooler').textContent='Spooler: '+data.spooler;byId('live-spooler').classList.add('on');
  byId('live-w32time').textContent='W32Time: '+data.w32time;byId('live-w32time').classList.add('on');
 }else if(data.kind==='evidence'){
  byId('live-extra-chip').textContent='Evidence: verified ✓';byId('live-extra-chip').classList.add('on');
 }else if(data.kind==='step-start'){
  stepRunning=true;
  setStatus('Running Step '+data.step+'…','running');
  setControlsDisabled(true);syncComposerState();
  appendTerminal((displayLog?'\\n':'')+'──▶ [Step '+data.step+'] '+STEP_GUIDES[data.demo][data.step-1].label+'\\n');
 }else if(data.kind==='step-done'){
  stepRunning=false;
  var demo=data.demo,step=Number(data.step),defs=STEP_GUIDES[demo];
  doneSet[demo][step]=true;
  if(demo==='pty'){if(step===1)ptyReady=true;if(step===6)ptyReady=false;}
  liveNext[demo]=step>=defs.length?1:step+1;
  setStepVal(step);
  setStatus(step===defs.length?'All Verified ✓':'Step '+step+' ✓','ok');
  setControlsDisabled(false);
  syncComposerState();syncStepSelection(false);
  resolveStepWaiters(true);
 }else if(data.kind==='step-error'){
  stepRunning=false;
  setStatus('Error','err');
  appendTerminal('\\n[Error] '+data.error+'\\n');
  setControlsDisabled(false);
  syncComposerState();
  resolveStepWaiters(false);
 }
}

function initReplayMode(){
 replayMode=true;
 if(window.RECORDED_DEMO&&window.RECORDED_DEMO.hostMeta){
  byId('live-host-user').textContent=window.RECORDED_DEMO.hostMeta.username;
  byId('live-host-session').textContent=window.RECORDED_DEMO.hostMeta.sessionId;
 }
 setStatus(idleModeLabel());
 if(!displayLog){
  replayBannerShown=true;
  appendTerminal('# Select any step or click "Run Step 1" / "Auto-Play" — Policy is blue, Pass is green, Blocked is red.');
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
   else if(ev.kind==='approval'){byId('live-extra-chip').textContent='INC-1042: read-only ✓';byId('live-extra-chip').classList.add('on');}
   else if(ev.kind==='services'){
    byId('live-spooler').textContent='Spooler: '+ev.spooler;byId('live-spooler').classList.add('on');
    byId('live-w32time').textContent='W32Time: '+ev.w32time;byId('live-w32time').classList.add('on');
   }
  });
 }
}

function playRecordedStep(kind,step){
 if(!window.RECORDED_DEMO||!window.RECORDED_DEMO.steps||!window.RECORDED_DEMO.steps[kind]){
  setStatus('No Recording','err');return Promise.resolve(false);
 }
 if(kind==='pty'&&step===1&&doneSet.pty[6]){
  doneSet.pty={};
  byId('live-demo').querySelectorAll('.ev-chip').forEach(function(x){x.classList.remove('on','denied');});
 }
 if(kind==='pty'&&step>1)hydratePriorPtyState(step);
 var raw=window.RECORDED_DEMO.steps[kind][String(step)];
 if(!raw||!raw.length){setStatus('Missing Step','err');return Promise.resolve(false);}
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
 try{
  var meta=await fetch('/api/meta').then(function(r){if(!r.ok)throw new Error('Local demo metadata unavailable');return r.json();});
  replayMode=false;
  setStatus(idleModeLabel());
  byId('live-host-user').textContent=meta.username;byId('live-host-session').textContent=meta.sessionId;
  syncComposerState();
  if(liveSource)liveSource.close();
  liveSource=new EventSource('/api/events');
  liveSource.onmessage=function(event){applyDemoEvent(JSON.parse(event.data));};
 }catch(_error){
  initReplayMode();
 }
}

async function openLiveDemo(preferredKind,preferredStep){
 if(deck.closeNotes)deck.closeNotes();
 byId('live-demo').hidden=false;liveOpen=true;
 replayMode=!isLocalDemoHost();
 if(preferredKind)setKindVal(preferredKind);
 resetLive();
 if(preferredKind){setKindVal(preferredKind);updateLiveMode();}
 if(preferredStep){setStepVal(preferredStep);syncStepSelection();}
 if(replayMode){initReplayMode();return;}
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
  setStatus('Blocked','err');
  appendTerminal('\\n[Error] '+(error.message||String(error))+'\\n');
  return false;
 }
}

async function runSelectedLiveStep(){
 if(stepRunning)return;
 stopAutoPlay();
 var target=Number(byId('live-demo-run').dataset.targetStep)||getStepVal();
 await runSingleStep(getKindVal(),target);
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
 if(!byId('cs-step-btn').disabled){stopAutoPlay();setStepVal(val);syncStepSelection(true);}
});
byId('live-demo-run').addEventListener('click',runSelectedLiveStep);
byId('live-demo-autoplay').addEventListener('click',toggleAutoPlay);
byId('live-demo-reset').addEventListener('click',function(){
 var chosen=getKindVal();
 if(!replayMode)fetch('/api/cancel',{method:'POST',headers:{'X-MXC-Demo':'1'}}).catch(function(){});
 resetLive();setKindVal(chosen);updateLiveMode();
 if(!replayMode){fetch('/api/meta').then(function(r){return r.json();}).then(function(m){byId('live-host-user').textContent=m.username;byId('live-host-session').textContent=m.sessionId;}).catch(function(){});}
 else{initReplayMode();}
});
byId('live-pty-form').addEventListener('submit',async function(ev){
 ev.preventDefault();
 var inp=byId('live-pty-input'),cmd=inp.value.trim();
 if(!cmd||!ptyReady||stepRunning||replayMode)return;
 inp.value='';
 await fetch('/api/pty-input',{method:'POST',headers:{'X-MXC-Demo':'1','Content-Type':'text/plain;charset=utf-8'},body:cmd}).catch(function(){});
});
document.querySelectorAll('.live-demo-trigger').forEach(function(button){
 button.addEventListener('click',function(event){
  event.stopPropagation();
  openLiveDemo(button.dataset.demo||'pty',Number(button.dataset.step||1));
 });
});
deck.openStudio=openLiveDemo;
byId('live-demo-close').addEventListener('click',closeLiveDemo);
})();
"""

PRESENTER_JS = """
(function(){
var isLocalOrigin=location.protocol==='file:'||['127.0.0.1','localhost','[::1]','::1'].includes(location.hostname);
if(!isLocalOrigin){
 ['deck-hud','notes-drawer'].forEach(function(id){var el=document.getElementById(id);if(el)el.remove();});
 return;
}
var deck=window.__DECK__;if(!deck)return;
var hudCount=document.getElementById('hud-count'),hud=document.getElementById('deck-hud'),notesDrawer=document.getElementById('notes-drawer');
if(!hud||!notesDrawer)return;
var S=deck.S,notesOpen=false;
function byId(id){return document.getElementById(id);}
function pad2(n){return n<10?'0'+n:String(n);}
function syncNotes(){
 var idx=deck.getIndex(),num=pad2(idx+1),isZh=(document.documentElement.lang||'').indexOf('zh')===0;
 if(hudCount)hudCount.textContent=num+' / '+pad2(S.length);
 byId('notes-title').textContent=(isZh?'讲者备注 · 第 '+num+' 页':'Speaker Notes · Slide '+num);
 var txt=(window.SLIDE_NOTES&&window.SLIDE_NOTES[String(idx)])||(isZh?'（本页无备注）':'(No notes for this slide)');
 byId('notes-body').textContent=txt;
}
function toggleNotes(force){
 notesOpen=typeof force==='boolean'?force:!notesOpen;
 notesDrawer.hidden=!notesOpen;
 hud.classList.toggle('notes-open',notesOpen);
 byId('hud-notes-btn').setAttribute('aria-expanded',String(notesOpen));
 syncNotes();
}
function defaultDemoForSlide(idx){
 return idx===25?'pty':(idx===26?'microvm':'process');
}
deck.closeNotes=function(){if(notesOpen)toggleNotes(false);};
deck.onRender=function(){syncNotes();};
syncNotes();
byId('hud-notes-btn').addEventListener('click',function(ev){ev.stopPropagation();toggleNotes();});
byId('notes-close').addEventListener('click',function(){toggleNotes(false);});
byId('hud-demo-btn').addEventListener('click',function(ev){ev.stopPropagation();if(deck.openStudio)deck.openStudio(defaultDemoForSlide(deck.getIndex()),1);});
addEventListener('keydown',function(ev){
 if(ev.altKey||ev.ctrlKey||ev.metaKey||deck.isOverlayOpen())return;
 var k=ev.key;
 if(ev.target&&ev.target.closest('input,textarea,select,[contenteditable="true"]')&&k!=='Escape')return;
 if(k==='Escape'&&notesOpen){ev.preventDefault();toggleNotes(false);return;}
 if(k==='n'||k==='N'){ev.preventDefault();toggleNotes();}
 else if(k==='d'||k==='D'){ev.preventDefault();if(deck.openStudio)deck.openStudio(defaultDemoForSlide(deck.getIndex()),1);}
},true);
})();
"""

PUBLIC_JS = DECK_JS + "\n" + STUDIO_JS
JS = PUBLIC_JS + "\n" + PRESENTER_JS

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

def lang_switch_html(lang='en', alt_href=None):
    if not alt_href:
        return ''
    if lang == 'zh':
        return (
            '<nav class="lang-switch" aria-label="语言切换">'
            f'<a class="lang-pill" href="{e(alt_href)}" data-base-href="{e(alt_href)}">EN</a>'
            '<span class="lang-pill active" aria-current="page">中文</span>'
            '</nav>'
        )
    return (
        '<nav class="lang-switch" aria-label="Language switch">'
        '<span class="lang-pill active" aria-current="page">EN</span>'
        f'<a class="lang-pill" href="{e(alt_href)}" data-base-href="{e(alt_href)}">中文</a>'
        '</nav>'
    )

def build(out, img_dir, maxw, pages_url=None, presenter_tools=True, lang='en', alt_href=None):
    loc_slides = D.get_localized_slides(lang)
    loc_notes = D.get_localized_notes(lang)
    D.validate_layout(loc_slides)
    out_dir = os.path.dirname(os.path.abspath(out))
    os.makedirs(out_dir, exist_ok=True)
    effective_url = pages_url if pages_url is not None else os.environ.get('PAGES_URL')
    body = ''.join(slide_html(k, loc_slides[k], img_dir, maxw, pages_url=effective_url, lang=lang) for k in sorted(loc_slides))
    recorded_json = load_recorded_demo(img_dir)
    if presenter_tools:
        notes_json = json.dumps({str(k): v for k, v in loc_notes.items()}, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
        presenter_block = f"\n{presenter_ui_html(lang)}"
        data_script = f"<script>window.RECORDED_DEMO={recorded_json};window.SLIDE_NOTES={notes_json};</script>"
        script_body = JS
    else:
        presenter_block = ""
        data_script = f"<script>window.RECORDED_DEMO={recorded_json};</script>"
        script_body = PUBLIC_JS
    html_lang = 'zh-CN' if lang == 'zh' else 'en'
    doc_title = (
        '用操作系统沙箱筑牢 Agent 护栏 — Microsoft Execution Containers (MXC)'
        if lang == 'zh'
        else 'Guardrails with Sandboxing - Microsoft Execution Containers (MXC)'
    )
    lang_nav = lang_switch_html(lang=lang, alt_href=alt_href)
    doc = f"""<!DOCTYPE html>
<html lang="{html_lang}" dir="ltr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light">
<title>{doc_title}</title>
<style>{CSS % (D.BG, D.INK, D.CY, D.CARD, D.CODEBG, D.EDGE, D.MUTED, D.AMB)}</style></head>
<body><div id="deck-progress" aria-hidden="true"></div><main id="stage">{body}</main>{lang_nav}{presenter_block}
{LIVE_DEMO}
{data_script}
<script>{script_body}</script></body></html>"""
    with open(out, 'w', encoding='utf-8', newline='\n') as f:
        f.write(doc)
    print(f'Saved {out}  ({len(loc_slides)} slides, lang={lang}, {os.path.getsize(out)/1e6:.2f} MB)')

if __name__ == '__main__':
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    root = os.path.dirname(here)
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=os.path.join(root, 'dist', 'html', 'index.html'))
    ap.add_argument('--images', default=os.path.join(here, 'assets', 'images'))
    ap.add_argument('--maxw', type=int, default=1600, help='max image edge in px')
    ap.add_argument('--pages-url', default=None, help='GitHub Pages URL for Slide 1 QR code')
    ap.add_argument('--lang', choices=('en', 'zh'), default='en', help='Deck language (en or zh)')
    ap.add_argument('--alt-href', default=None, help='Alternate language deck link')
    ap.add_argument('--public-pages', action='store_true', help='omit Notes, Studio, and N/D shortcuts for public GitHub Pages')
    a = ap.parse_args(); build(a.out, a.images, a.maxw, pages_url=a.pages_url, presenter_tools=not a.public_pages, lang=a.lang, alt_href=a.alt_href)
