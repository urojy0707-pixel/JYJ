# -*- coding: utf-8 -*-
"""자문 의견서 -> HWPX (한글 2010+ 개방형 표준 포맷) 생성기."""
from __future__ import annotations
import re, sys, warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=DeprecationWarning)
sys.path.insert(0, str(Path(__file__).parent))

from hwpx.document import HwpxDocument
import os, importlib
C = importlib.import_module(os.environ.get("CONTENT_MODULE", "content"))

BODY_FONT = "맑은 고딕"
INK      = "#1A1A1A"
ACCENT   = "#0B3D5C"
MUTED    = "#555555"
RULE     = "#C8C8C8"
BOX_BG   = "#EEF4F8"
HEAD_BG  = "#DCE6EE"

BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def emit_inline(para, text, *, size=10.0, color=INK, font=BODY_FONT, base_bold=False):
    """`**bold**` 마크업을 런 단위로 분해해 문단에 기록."""
    pos = 0
    for m in BOLD_RE.finditer(text):
        if m.start() > pos:
            para.add_run(text[pos:m.start()], bold=base_bold, size=size, color=color, font=font)
        para.add_run(m.group(1), bold=True, size=size, color=color, font=font)
        pos = m.end()
    if pos < len(text):
        para.add_run(text[pos:], bold=base_bold, size=size, color=color, font=font)
    if not text:
        para.add_run("", size=size, color=color, font=font)


def plain(text):
    return BOLD_RE.sub(r"\1", text)


def build(out_path: str) -> str:
    doc = HwpxDocument.new()
    doc.page.setup(
        paper_size="A4",
        margins_mm={"left": 20, "right": 20, "top": 20, "bottom": 18},
        header_margin_mm=10, footer_margin_mm=10,
    )

    def para(text="", *, size=10.0, color=INK, bold=False, align="justify",
             before=0, after=3.0, spacing=160, indent_left=0.0, inline=True,
             font=BODY_FONT, page_break=False):
        p = doc.add_paragraph(include_run=False)
        if inline:
            emit_inline(p, text, size=size, color=color, font=font, base_bold=bold)
        else:
            p.add_run(text, bold=bold, size=size, color=color, font=font)
        idx = len(doc.paragraphs) - 1
        doc.set_paragraph_format(
            paragraph_index=idx, alignment=align, line_spacing_percent=spacing,
            spacing_before_pt=before, spacing_after_pt=after,
            indent_left_mm=indent_left, page_break_before=page_break or None,
        )
        return p

    # ── 표지 영역 ──────────────────────────────────────────────
    para("", after=6)
    para(C.TITLE, size=19.0, bold=True, color=ACCENT, align="center",
         after=4, spacing=150, inline=False)
    para(C.SUBTITLE, size=11.0, color=MUTED, align="center", after=10,
         spacing=150, inline=False)

    meta_tbl = doc.add_table(len(C.META), 2)
    meta_tbl.set_column_widths([26, 74])
    for r, (k, v) in enumerate(C.META):
        meta_tbl.set_cell_text(r, 0, k)
        meta_tbl.set_cell_text(r, 1, v)
    style_table(doc, meta_tbl, header=False, key_col=True)
    para("", after=6)

    # ── 본문 ──────────────────────────────────────────────────
    pending_break = False
    for block in C.BLOCKS:
        kind = block[0]

        if kind == "pagebreak":
            pending_break = True
            continue

        pb, pending_break = pending_break, False

        if kind == "h1":
            para(block[1], size=14.5, bold=True, color=ACCENT, align="left",
                 before=10, after=4, spacing=150, inline=False, page_break=pb)
        elif kind == "h2":
            para(block[1], size=12.0, bold=True, color=ACCENT, align="left",
                 before=7, after=3, spacing=150, inline=False, page_break=pb)
        elif kind == "h3":
            para(block[1], size=10.8, bold=True, color=INK, align="left",
                 before=5, after=2, spacing=150, inline=False, page_break=pb)
        elif kind == "p":
            para(block[1], page_break=pb)
        elif kind == "mono":
            for line in block[1].split("\n"):
                para(line, size=9.5, font="나눔고딕코딩", align="left",
                     after=0.5, spacing=140, indent_left=4.0, inline=False)
        elif kind == "ul":
            for item in block[1]:
                p = doc.add_paragraph(include_run=False)
                p.add_run("· ", size=10.0, color=ACCENT, bold=True, font=BODY_FONT)
                emit_inline(p, item, size=10.0)
                doc.set_paragraph_format(
                    paragraph_index=len(doc.paragraphs) - 1, alignment="justify",
                    line_spacing_percent=160, spacing_after_pt=1.5,
                    indent_left_mm=5.0, first_line_indent_mm=-3.5)
        elif kind == "ol":
            for n, item in enumerate(block[1], 1):
                p = doc.add_paragraph(include_run=False)
                p.add_run(f"{n}. ", size=10.0, color=ACCENT, bold=True, font=BODY_FONT)
                emit_inline(p, item, size=10.0)
                doc.set_paragraph_format(
                    paragraph_index=len(doc.paragraphs) - 1, alignment="justify",
                    line_spacing_percent=160, spacing_after_pt=1.5,
                    indent_left_mm=6.0, first_line_indent_mm=-4.5)
        elif kind == "callout":
            para("", after=1.5)
            box = doc.add_table(1, 1)
            box.set_column_widths([100])
            cell = box.cell(0, 0)
            cell.text = ""
            cp = cell.paragraphs[0]
            emit_inline(cp, block[1], size=10.0, color="#0A2E45", base_bold=True)
            box.set_cell_shading(0, 0, BOX_BG)
            box.set_cell_borders(0, 0, color=ACCENT, line_type="SOLID")
            para("", after=3)
        elif kind == "table":
            headers, rows = block[1], block[2]
            widths = block[3] if len(block) > 3 else None
            para("", after=1.0)
            t = doc.add_table(len(rows) + 1, len(headers))
            if widths:
                t.set_column_widths(widths)
            else:
                t.equalize_column_widths()
            for c, h in enumerate(headers):
                t.set_cell_text(0, c, plain(h))
            for r, row in enumerate(rows, 1):
                for c, v in enumerate(row):
                    t.set_cell_text(r, c, plain(v))
            style_table(doc, t, header=True)
            para("", after=3)
        elif kind == "hr":
            p = doc.add_paragraph(include_run=False)
            p.add_run("", size=9.0, font=BODY_FONT)
            doc.set_paragraph_format(paragraph_index=len(doc.paragraphs) - 1,
                                     spacing_before_pt=4, spacing_after_pt=4,
                                     bottom_border=True, border_color=RULE)

    doc.set_header_text("건강기능식품 원료 자문 의견서 — 아르기닌·아연·비타민 B군")
    doc.set_footer_text("- 대외비 자문 문서 -")

    doc.save_to_path(out_path)
    return out_path


def style_table(doc, t, *, header=True, key_col=False):
    """표 셀 서식: 헤더 음영/굵게, 본문 가독성."""
    ncols = t.column_count
    for r, row in enumerate(t.rows):
        for c in range(ncols):
            try:
                cell = t.cell(r, c)
            except Exception:
                continue
            is_head = header and r == 0
            is_key = key_col and c == 0
            for p in cell.paragraphs:
                txt = p.text
                p.clear_text()
                p.add_run(txt, size=9.2 if not is_head else 9.4,
                          bold=is_head or is_key,
                          color=ACCENT if (is_head or is_key) else INK,
                          font=BODY_FONT)
            if is_head:
                t.set_cell_shading(r, c, HEAD_BG)
            elif is_key and key_col:
                t.set_cell_shading(r, c, "#F5F7F9")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "out.hwpx"
    p = build(out)
    d = HwpxDocument.open(p)
    rep = d.validate()
    print("saved:", p)
    print("validate issues:", len(rep.issues))
    for i in rep.issues[:10]:
        print("  !", i)
