# -*- coding: utf-8 -*-
"""チェングロウスver1.1 S1〜S24（個別パート）の文字サイズを上げる（2026-09-24 殿村さん指示）。
S13・S14（図が詰まった汎用図）は据え置き。
ルール：タイトル（上部16pt）・20pt以上は据え置き／10.5pt以下は+2pt／11〜17ptは+1pt。
SMALL_ONLY_1 に入れたスライド・図形は、あふれ対策で一律+1pt。
  python3 _build/font_up_chengrowth.py <入力> <出力>
"""
import sys
from pptx import Presentation
from pptx.util import Pt, Cm

SRC, OUT = sys.argv[1], sys.argv[2]
LAST = 24
SKIP = {13, 14}  # 図が詰まっていて崩れるため据え置き
SMALL_ONLY_1 = set(a for a in sys.argv[3:])  # "s13" or "s13:図形名"

prs = Presentation(SRC)

def paras(sh):
    if getattr(sh, "has_table", False) and sh.has_table:
        for row in sh.table.rows:
            for cell in row.cells: yield from cell.text_frame.paragraphs
    elif sh.has_text_frame:
        yield from sh.text_frame.paragraphs

def leaves(shapes):
    for sh in shapes:
        if sh.shape_type == 6: yield from leaves(sh.shapes)
        else: yield sh

n = 0
for i, s in enumerate(prs.slides, 1):
    if i > LAST: break
    if i in SKIP: continue
    for sh in leaves(s.shapes):
        is_title = sh.top is not None and sh.top < Cm(1.5)
        only1 = f"s{i:02d}" in SMALL_ONLY_1 or f"s{i:02d}:{sh.name}" in SMALL_ONLY_1
        for p in paras(sh):
            for r in p.runs:
                sz = r.font.size
                if sz is None or not r.text.strip(): continue
                pt = sz.pt
                if is_title or pt >= 20: continue
                add = 2 if (pt <= 10.5 and not only1) else 1
                r.font.size = Pt(pt + add); n += 1
prs.save(OUT)
print("runs changed:", n)
