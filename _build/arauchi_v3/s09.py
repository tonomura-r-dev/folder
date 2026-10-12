# -*- coding: utf-8 -*-
"""S9 想定セグメント×成果：表の中の文字を、原本の8〜10ptから11ptにする（2026-10-12）。
箱の大きさ・位置は変えず、箱の中の左右の余白を詰めて収める（文言は削らない）。
薄い灰色だった施策方針の文字は、読める濃さ（595959）にする。"""
from _build.arauchi_v3.common import (Inches, ids, move, qn, set_anchor, set_color, set_insets, set_line_spacing,
                                      set_size_all)

PANEL = (34, 35, 37, 38, 48, 49, 51, 52, 54, 55, 57, 58, 60, 61)      # 具体の施策方針のマス
LEFT = (9, 10, 11, 12, 17, 18, 19, 25, 26, 27, 28, 29, 30)            # 流入元・LINE誘導の列


def apply(ctx):
    b = ids(ctx.orig[9])
    for k in PANEL:
        set_size_all(b[k], 11)
        for i in range(len(b[k].text_frame.paragraphs)):
            set_color(b[k], i, "595959")
        set_line_spacing(b[k], 100)
        set_insets(b[k], l=0.03, r=0.03, t=0.03, b=0.03)
        set_anchor(b[k], "ctr")
    for k in LEFT:
        set_size_all(b[k], 11)
        set_insets(b[k], l=0.03, r=0.03)
        set_line_spacing(b[k], 100)
    # 右の数字（20ptのまま）と、その説明（9→11pt）。見出しは12pt
    for k in (41, 42):
        for p in b[k].text_frame._txBody.findall(qn("a:p"))[1:]:
            for rpr in list(p.iter(qn("a:rPr"))) + list(p.iter(qn("a:endParaRPr"))):
                rpr.set("sz", "1100")
        b[k].top = b[k].top - Inches(0.06)
    for p in b[40].text_frame._txBody.findall(qn("a:p")):
        for rpr in list(p.iter(qn("a:rPr"))) + list(p.iter(qn("a:endParaRPr"))):
            rpr.set("sz", "1200")
    # 上の要旨と、表の見出し（12→13pt）
    for k in (68, 5, 6, 7):
        set_size_all(b[k], 13)
