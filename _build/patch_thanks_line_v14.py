# -*- coding: utf-8 -*-
"""サンクスLINEのご提案 ver1.3 → ver1.4（2026-10-01）
S2：余白が大きいので、文字と図形を大きくしてページを埋める（流れの図は上、強み3つは下）。
  python3 _build/patch_thanks_line_v14.py <ver1.3.pptx>
"""
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20260930_サンクスLINEのご提案ver1.4.pptx"
prs = Presentation(sys.argv[1])
s = prs.slides[1]
by = lambda n: next(x for x in s.shapes if x.name == n)


def place(n, x, y, w, h):
    sh = by(n)
    sh.left, sh.top, sh.width, sh.height = Cm(x), Cm(y), Cm(w), Cm(h)
    return sh


def size(sh, first, rest=None):
    for i, para in enumerate(sh.text_frame.paragraphs):
        for r in para.runs:
            r.font.size = Pt(first if i == 0 else (rest or first))


# 流れの図（1段目）：横いっぱい・高さ3.8cm
Y, H, AW, AH, G = 5.7, 3.8, 0.8, 1.0, 0.2
x = 1.46
for n, w in (("Rounded Rectangle 4", 4.4), ("Right Arrow 5", AW), ("Rounded Rectangle 6", 5.2), ("Right Arrow 7", AW),
             ("Rounded Rectangle 8", 4.0), ("Right Arrow 9", AW)):
    if n.startswith("Right"):
        place(n, x, Y + (H - AH) / 2, w, AH)
    else:
        place(n, x, Y, w, H)
    x += w + G
place("Rounded Rectangle 10", x, Y, 26.06 - x, (H - 0.2) / 2)
place("Rounded Rectangle 11", x, Y + (H - 0.2) / 2 + 0.2, 26.06 - x, (H - 0.2) / 2)
size(by("Rounded Rectangle 4"), 17)
size(by("Rounded Rectangle 6"), 17, 13)
size(by("Rounded Rectangle 8"), 16)
size(by("Rounded Rectangle 10"), 13)
for r in by("Rounded Rectangle 10").text_frame.paragraphs[0].runs:   # 折り返さないよう短く
    r.text = r.text.replace("友だち追加の画面へ", "友だち追加へ")
size(by("Rounded Rectangle 11"), 13)

# 強み3つ（2段目）：高さ5cm・見出し17pt・本文14pt
CY, CH, CW, CG = 10.4, 5.6, (24.6 - 0.4) / 3, 0.2
for i, n in enumerate(("Rounded Rectangle 12", "Rounded Rectangle 13", "Rounded Rectangle 14")):
    sh = place(n, 1.46 + i * (CW + CG), CY, CW, CH)
    size(sh, 16, 14)
    for para in sh.text_frame.paragraphs:
        para.line_spacing = 1.3
prs.save(str(OUT))
print("saved:", OUT.name)
