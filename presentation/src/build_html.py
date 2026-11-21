#!/usr/bin/env python3
"""build_html.py - turns the same 28 slides into ONE self-contained HTML file
(images embedded, no external files). Uses the slide content from build_pptx.py.

Usage:  pip install pillow
        python -m presentation.src.build_html [--out dist/html/index.html]
Keys in the browser: <- / -> or Space to navigate, F = fullscreen,
Home/End = first/last. Print (Ctrl+P) gives one slide per page.
"""
import argparse, base64, html, io, json, os
from PIL import Image
try:
    from . import build_pptx as D
except ImportError:
    import build_pptx as D

PX = 9525  # EMU per CSS px  (slide = 1280 x 720 px)
SVG_POINTS = {
    'rightArrow':'0,25 70,25 70,0 100,50 70,100 70,75 0,75',
    'leftArrow':'30,0 30,25 100,25 100,75 30,75 30,100 0,50',
    'chevron':'0,0 82,0 100,50 82,100 0,100 18,50',
    'homePlate':'0,0 82,0 100,50 82,100 0,100',
}
FALLBACK = {
    'Aptos Display':'"Aptos Display","Segoe UI Variable Display","Segoe UI",Arial,sans-serif',
    'Aptos':'Aptos,"Segoe UI Variable Text","Segoe UI",Arial,sans-serif',
    'Cascadia Code':'"Cascadia Code","Cascadia Mono",Consolas,monospace',
    'Georgia':'Georgia,"Times New Roman",serif',
    'Calibri':'Calibri,Carlito,"Segoe UI",Arial,sans-serif',
    'Courier New':'"Courier New",Consolas,monospace',
    'Arial':'Arial,Helvetica,sans-serif',
}
e = lambda s: html.escape(s, quote=True)

def data_uri(path, maxw, cache={}):
    if path not in cache:
        im = Image.open(path).convert('RGB'); im.thumbnail((maxw, maxw))
        b = io.BytesIO(); im.save(b, 'JPEG', quality=80, optimize=True)
        cache[path] = 'data:image/jpeg;base64,' + base64.b64encode(b.getvalue()).decode()
    return cache[path]

def box(o):
    return f"left:{o['x']/PX:.1f}px;top:{o['y']/PX:.1f}px;width:{o['cx']/PX:.1f}px;height:{o['cy']/PX:.1f}px;"

def text_html(o):
    paras = o.get('paragraphs') or [{'text': o['text']}]
    font = FALLBACK.get(o.get('fontFamily', D.BF), 'sans-serif')
    ls = f"letter-spacing:{o['letterSpacingPt']*4/3:.1f}px;" if o.get('letterSpacingPt') else ''
    out = []
    for p in paras:
        sz = p.get('sizePt', o.get('sizePt', 16)) * 4/3
        st = (f"font-size:{sz:.1f}px;color:#{p.get('color', o.get('color', D.INK))};"
              f"font-weight:{700 if p.get('bold', o.get('bold')) else 400};"
              f"text-align:{p.get('align', o.get('align','center'))};")
        out.append(f'<p style="{st}">' + '<br>'.join(e(l) for l in p['text'].split('\n')) + '</p>')
    return f'<div class="t" style="font-family:{e(font)};{ls}">' + ''.join(out) + '</div>'

REVEAL_ITEMS = {
    5: (2, 5, 3),
    7: (2, 4, 5),
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
            src = data_uri(os.path.join(img_dir, os.path.basename(o['path'])), maxw)
            parts.append(f'<img class="{classes}"{reveal} style="{box(o)}" src="{src}" alt="{e(o.get("altText",""))}">')
        elif o['op'] == 'add_text_box':
            parts.append(f'<div class="{classes}"{reveal} style="{box(o)}">{text_html(o)}</div>')
        elif o['op'] == 'add_shape':
            pr = o['preset']
            inner = text_html(o) if 'text' in o else ''
            if pr in SVG_POINTS:
                stroke = f' stroke="#{o["line"]}" stroke-width="{max(1,o.get("lineWidth",12700)/PX):.1f}"' if o.get('line') else ''
                svg = (
                    '<svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">'
                    f'<polygon points="{SVG_POINTS[pr]}" fill="#{o["fill"]}"{stroke}/></svg>'
                )
                parts.append(f'<div class="{classes} svg-shape"{reveal} style="{box(o)}">{svg}{inner}</div>')
            else:
                st = box(o) + f"background:#{o['fill']};"
                if pr == 'ellipse': st += 'border-radius:50%;'
                if pr == 'roundRect': st += f"border-radius:{min(o['cx'],o['cy'])/PX/6:.1f}px;"
                if o.get('line'): st += f"box-shadow:inset 0 0 0 {max(1,o.get('lineWidth',12700)/PX):.1f}px #{o['line']};"
                parts.append(f'<div class="{classes}"{reveal} style="{st}">{inner}</div>')
    reveal_count = REVEAL_ITEMS[k][2] if k in REVEAL_ITEMS else 0
    return f'<section class="slide" data-n="{k+1}" data-reveals="{reveal_count}" aria-label="Slide {k+1}">{"".join(parts)}</section>'

CSS = """
:root{--bg:#%s;--ink:#%s;--chrome:#2b2b2b;--chrome-ink:#f3eee4}
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:100%%;height:100%%;background:var(--bg);color:var(--ink);overflow:hidden;overscroll-behavior:none}
#stage{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;overflow:hidden}
.slide{position:absolute;width:1280px;height:720px;background:var(--bg);overflow:hidden;display:none;transform-origin:center center}
.slide.on{display:block}
.a{position:absolute}
.reveal{visibility:hidden}
.reveal.shown{visibility:visible}
.svg-shape svg{position:absolute;inset:0;width:100%%;height:100%%;overflow:visible}
.svg-shape>.t{z-index:1}
img.a{object-fit:cover}
.t{position:absolute;inset:0;padding:4px 10px;display:flex;flex-direction:column;justify-content:center;line-height:1.2}
@media (prefers-color-scheme:dark){:root{--chrome:#1a1a1a}}
@media print{@page{size:13.333in 7.5in;margin:0}
 html,body{overflow:visible;background:#fff}
 #stage{position:static;display:block}.slide{position:relative;display:block!important;transform:none!important;box-shadow:none;page-break-after:always;break-after:page}.reveal{visibility:visible!important}}
"""
JS = """
(function(){
var S=[].slice.call(document.querySelectorAll('.slide')),i=0,steps=S.map(function(){return 0;});
function fit(){var v=window.visualViewport,w=v?v.width:innerWidth,h=v?v.height:innerHeight,s=Math.max(w/1280,h/720);S.forEach(function(x){x.style.transform='scale('+s+')';});}
function toggleFullscreen(){if(document.fullscreenElement){document.exitFullscreen();}else if(document.documentElement.requestFullscreen){document.documentElement.requestFullscreen();}}
function renderSteps(){var shown=steps[i];S[i].querySelectorAll('.reveal').forEach(function(x){x.classList.toggle('shown',Number(x.dataset.reveal)<shown);});}
function go(n){i=Math.max(0,Math.min(S.length-1,n));S.forEach(function(x,j){x.classList.toggle('on',j===i);});
 renderSteps();
 if(history.replaceState)history.replaceState(null,'','#'+(i+1));}
function next(){var total=Number(S[i].dataset.reveals||0);if(steps[i]<total){steps[i]++;renderSteps();}else{go(i+1);}}
function previous(){if(steps[i]>0){steps[i]--;renderSteps();}else{go(i-1);}}
document.getElementById('stage').addEventListener('click',function(ev){ev.clientX<innerWidth/3?previous():next();});
addEventListener('keydown',function(ev){var k=ev.key;
 if(k==='ArrowRight'||k==='PageDown'||k===' ')next();else if(k==='ArrowLeft'||k==='PageUp')previous();
 else if(k==='Home')go(0);else if(k==='End')go(S.length-1);
 else if(k==='f'||k==='F'){ev.preventDefault();toggleFullscreen();}});
addEventListener('resize',fit);document.addEventListener('fullscreenchange',fit);if(window.visualViewport)window.visualViewport.addEventListener('resize',fit);
addEventListener('hashchange',function(){var n=(parseInt(location.hash.slice(1),10)||1)-1;if(n!==i)go(n);});
fit();go((parseInt(location.hash.slice(1),10)||1)-1);
})();
"""

def build(out, img_dir, maxw):
    out_dir = os.path.dirname(os.path.abspath(out))
    os.makedirs(out_dir, exist_ok=True)
    body = ''.join(slide_html(k, D.slides[k], img_dir, maxw) for k in sorted(D.slides))
    doc = f"""<!DOCTYPE html>
<html lang="en" dir="ltr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>Guardrails with Sandboxing - Microsoft Execution Containers (MXC)</title>
<style>{CSS % (D.BG, D.INK)}</style></head>
<body><main id="stage">{body}</main>
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
