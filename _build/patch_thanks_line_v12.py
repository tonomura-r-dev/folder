# -*- coding: utf-8 -*-
"""サンクスLINE簡易資料 ver1.1（殿村さんPC修正版）→ ver1.2
費用がS2とS5の2か所にあるので、S5の費用（初期・月額のチップと※注記）を消す。
S2の費用の枠（API併用・追加費用の注記つき）を残す。S5のリード「月3万円〜を足すだけ」はそのまま。

  python3 _build/patch_thanks_line_v12.py <ver1.1.pptx>
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Cm

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20260930_サンクスLINEのご提案（簡易版）ver1.2.pptx"
prs = Presentation(sys.argv[1])
s5 = prs.slides[4]
by = lambda n: next(x for x in s5.shapes if x.name == n)

for n in ("Rounded Rectangle 8", "Rounded Rectangle 9", "TextBox 10"):
    el = by(n)._element
    el.getparent().remove(el)

# 帯を、消した費用の位置まで上げる
band = by("Rounded Rectangle 11")
band.top = Cm(12.26)

# 本文（区切り線〜注記の間）を上下中央へ
DIV_Y, FOOT_Y = 3.86, 17.35
body = [sh for sh in s5.shapes if Cm(DIV_Y + 0.1) <= sh.top < Cm(FOOT_Y - 0.1)]
top = min(sh.top for sh in body)
bottom = max(sh.top + sh.height for sh in body)
dy = int((Cm(DIV_Y + 0.4) + Cm(FOOT_Y - 0.3)) / 2 - (top + bottom) / 2)
for sh in body:
    sh.top = sh.top + dy

prs.save(str(OUT))
print("saved:", OUT.name)
