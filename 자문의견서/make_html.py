# -*- coding: utf-8 -*-
"""자문 의견서 -> 인쇄용 HTML (Chromium print-to-PDF 입력)."""
from __future__ import annotations
import html as H, re, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import os, importlib
C = importlib.import_module(os.environ.get("CONTENT_MODULE", "content"))

BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def inl(text: str) -> str:
    return BOLD_RE.sub(r"<strong>\1</strong>", H.escape(text))


CSS = """
@page { size: A4; margin: 20mm 18mm 18mm 18mm; }
* { box-sizing: border-box; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body {
  font-family: "NanumGothic","Nanum Gothic","Malgun Gothic",sans-serif;
  font-size: 10pt; line-height: 1.72; color: #1a1a1a; margin: 0;
  word-break: keep-all; overflow-wrap: break-word;
}
.cover { text-align: center; padding: 8mm 0 4mm; border-bottom: 2.4pt solid #0b3d5c; margin-bottom: 7mm; }
.cover .eyebrow { font-size: 8.6pt; letter-spacing: .28em; color: #5c7d92; margin-bottom: 5mm; }
.cover h1 { font-size: 21pt; color: #0b3d5c; margin: 0 0 3mm; line-height: 1.34; font-weight: 800; letter-spacing: -.01em; }
.cover .sub { font-size: 11pt; color: #4d5b64; margin: 0 auto; max-width: 150mm; line-height: 1.6; }

table.meta { width: 100%; border-collapse: collapse; margin: 0 0 8mm; font-size: 9.3pt; }
table.meta th { width: 26%; text-align: left; background: #f4f7f9; color: #0b3d5c;
  font-weight: 700; padding: 2.4mm 3mm; border: .4pt solid #cfdae2; vertical-align: top; }
table.meta td { padding: 2.4mm 3mm; border: .4pt solid #cfdae2; vertical-align: top; }

h2.s { font-size: 14.5pt; color: #0b3d5c; font-weight: 800; margin: 9mm 0 3.2mm;
  padding-bottom: 1.6mm; border-bottom: 1.2pt solid #0b3d5c; break-after: avoid; letter-spacing: -.01em; }
h3.s { font-size: 11.6pt; color: #0b3d5c; font-weight: 700; margin: 6.5mm 0 2.4mm; break-after: avoid; }
h4.s { font-size: 10.4pt; color: #1a1a1a; font-weight: 700; margin: 5mm 0 1.8mm; break-after: avoid; }
p { margin: 0 0 2.6mm; text-align: justify; }

ul.s, ol.s { margin: 0 0 3.2mm; padding-left: 6.5mm; }
ul.s li, ol.s li { margin-bottom: 1.4mm; text-align: justify; }
ul.s li::marker { color: #0b3d5c; }
ol.s li::marker { color: #0b3d5c; font-weight: 700; }

.callout { background: #eef4f8; border-left: 3.4pt solid #0b3d5c; padding: 3.4mm 4.2mm;
  margin: 3.4mm 0 4.4mm; font-size: 10pt; font-weight: 700; color: #0a2e45;
  line-height: 1.68; text-align: justify; break-inside: avoid; }

table.d { width: 100%; border-collapse: collapse; margin: 2.4mm 0 5mm; font-size: 8.9pt; line-height: 1.55; }
table.d th { background: #dce6ee; color: #0b3d5c; font-weight: 700; text-align: left;
  padding: 2.2mm 2.4mm; border: .4pt solid #b8c8d4; vertical-align: middle; font-size: 9.1pt; }
table.d td { padding: 2.2mm 2.4mm; border: .4pt solid #cfdae2; vertical-align: top; text-align: justify; }
table.d tr { break-inside: avoid; }
table.d tbody tr:nth-child(even) td { background: #fafcfd; }
table.d td:first-child { font-weight: 600; color: #123c56; }

pre.tree { font-family: "NanumGothicCoding","Nanum Gothic Coding",monospace; font-size: 9pt;
  background: #f6f8fa; border: .4pt solid #dde5eb; border-radius: 1.2mm;
  padding: 3.4mm 4mm; margin: 2.4mm 0 4mm; line-height: 1.62; white-space: pre-wrap; break-inside: avoid; }

hr.s { border: 0; border-top: .6pt solid #c8c8c8; margin: 6mm 0 4mm; }
.pb { break-before: page; }
h2.s, h3.s, h4.s, .callout, pre.tree { break-inside: avoid; }
"""


def render() -> str:
    o = ['<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">',
         f'<title>{H.escape(C.TITLE)}</title><style>{CSS}</style></head><body>']

    o.append('<div class="cover"><div class="eyebrow">A D V I S O R Y &nbsp; O P I N I O N</div>')
    o.append(f'<h1>{H.escape(C.TITLE)}</h1>')
    o.append(f'<div class="sub">{H.escape(C.SUBTITLE)}</div></div>')

    o.append('<table class="meta">')
    for k, v in C.META:
        o.append(f'<tr><th>{H.escape(k)}</th><td>{H.escape(v)}</td></tr>')
    o.append('</table>')

    pb = False
    for b in C.BLOCKS:
        k = b[0]
        if k == "pagebreak":
            pb = True
            continue
        cls = ' pb' if pb else ''
        pb = False

        if k == "h1":
            o.append(f'<h2 class="s{cls}">{H.escape(b[1])}</h2>')
        elif k == "h2":
            o.append(f'<h3 class="s{cls}">{H.escape(b[1])}</h3>')
        elif k == "h3":
            o.append(f'<h4 class="s{cls}">{H.escape(b[1])}</h4>')
        elif k == "p":
            o.append(f'<p class="{cls.strip()}">{inl(b[1])}</p>')
        elif k == "mono":
            o.append(f'<pre class="tree{cls}">{H.escape(b[1])}</pre>')
        elif k == "ul":
            o.append(f'<ul class="s{cls}">' + "".join(f'<li>{inl(i)}</li>' for i in b[1]) + '</ul>')
        elif k == "ol":
            o.append(f'<ol class="s{cls}">' + "".join(f'<li>{inl(i)}</li>' for i in b[1]) + '</ol>')
        elif k == "callout":
            o.append(f'<div class="callout{cls}">{inl(b[1])}</div>')
        elif k == "table":
            hd, rows = b[1], b[2]
            ws = b[3] if len(b) > 3 else None
            o.append(f'<table class="d{cls}">')
            if ws:
                o.append('<colgroup>' + "".join(f'<col style="width:{w}%">' for w in ws) + '</colgroup>')
            o.append('<thead><tr>' + "".join(f'<th>{inl(x)}</th>' for x in hd) + '</tr></thead><tbody>')
            for r in rows:
                o.append('<tr>' + "".join(f'<td>{inl(c)}</td>' for c in r) + '</tr>')
            o.append('</tbody></table>')
        elif k == "hr":
            o.append(f'<hr class="s{cls}">')

    o.append('</body></html>')
    return "\n".join(o)


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "out.html")
    out.write_text(render(), encoding="utf-8")
    print("html:", out, out.stat().st_size, "bytes")
