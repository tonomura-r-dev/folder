# -*- coding: utf-8 -*-
"""サンクスLINE誘導・業界別 ver3.0（殿村さんPC保存版）→ ver3.1（2026-10-07 殿村さん指示）。
5枚目の費用を「別途お見積もり」から「費用目安＋個別見積」の見せ方に変更。他のページは触らない。
  ① サンクスLINE誘導ツール｜費用目安：初期15万円／月額3万円〜
  ② LINE運用コンサル｜費用目安：初期10万円〜／月額20万円〜／6か月〜
  ※施策内容・配信本数・連携範囲等により、費用は個別にお見積もりいたします。
支援内容（①導入→②配信設計→③効果計測→④改善）はそのまま。費用ブロックは目立たせない（塗りなし・細い線）。
  python3 _build/patch_thanks_line_gyokai_v31_fee.py <業界別ver3.0.pptx> [出力]
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_parts import *  # noqa

ROOT = Path(__file__).resolve().parent.parent
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "20261007_サンクスLINE誘導のご提案_業界別ver3.1.pptx"
L, R = 2.2, 31.7
CW = R - L

prs = Presentation(sys.argv[1])
s = prs.slides[4]
assert s.shapes[0].text_frame.text.startswith("DYMのコンサル支援内容"), "5枚目がDYMの支援ページではない"

# 旧の費用ブロック（区切り線＋「費用」＋箱4つ）を外す。y=13.7以下にある図形が対象
for sh in list(s.shapes):
    if sh.top / 360000 >= 13.6:
        sh._element.getparent().remove(sh._element)

TOP = 13.55
hline(s, L, R, TOP, LGRAY, 0.75)
GAP = 0.8
BW = (CW - GAP) / 2
blocks = [
    ("① サンクスLINE誘導ツール｜費用目安", GREEN, [("初期", "15万円"), ("月額", "3万円〜")]),
    ("② LINE運用コンサル｜費用目安", NAVY, [("初期", "10万円〜"), ("月額", "20万円〜"), ("期間", "6か月〜")]),
]
HY = TOP + 0.3          # 見出し
AY = HY + 0.85          # 金額の行
AH = 1.35
for i, (title, mark, items) in enumerate(blocks):
    bx = L + i * (BW + GAP)
    rect(s, bx, HY + 0.12, 0.18, 0.5, mark)
    label(s, bx + 0.35, HY, BW - 0.35, 0.75, [(title, 11.5, True, NAVY, 0)], anchor=MSO_ANCHOR.MIDDLE)
    iw = BW / len(items)
    for j, (k, v) in enumerate(items):
        ix = bx + j * iw
        if j:
            vline(s, ix, AY + 0.2, AY + AH - 0.2, LGRAY, 0.75)
        rich(s, ix, AY, iw, AH, [[(k + "　", 9.5, False, GRAY), (v, 16, True, NAVY)]], align=PP_ALIGN.CENTER)
    if i:
        vline(s, bx - GAP / 2, HY, AY + AH, LGRAY, 0.75)
label(s, L, AY + AH + 0.1, CW, 0.55, [("※施策内容・配信本数・連携範囲等により、費用は個別にお見積もりいたします。", 9, False, GRAY, 0)])

prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
