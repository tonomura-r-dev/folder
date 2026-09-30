# -*- coding: utf-8 -*-
"""サンクスLINE簡易資料 ver1.1（殿村さんPC修正版）→ ver1.2
費用がS2とS5の2か所にあるので、S2の費用の枠を消し、S5に費用を残す（殿村さん指示）。
S2の枠にあった「APIツールと併用できる」は、S5の※注記に移す。

  python3 _build/patch_thanks_line_v12.py <ver1.1.pptx>
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Cm

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20260930_サンクスLINEのご提案（簡易版）ver1.2.pptx"
prs = Presentation(sys.argv[1])
S = list(prs.slides)
DIV_Y, FOOT_Y = 3.86, 17.35


def by(s, n):
    return next(x for x in s.shapes if x.name == n)


def center(s):
    """本文（区切り線〜注記の間）を上下中央へ"""
    body = [sh for sh in s.shapes if Cm(DIV_Y + 0.1) <= sh.top < Cm(FOOT_Y - 0.1)]
    top = min(sh.top for sh in body)
    bottom = max(sh.top + sh.height for sh in body)
    dy = int((Cm(DIV_Y + 0.4) + Cm(FOOT_Y - 0.3)) / 2 - (top + bottom) / 2)
    for sh in body:
        sh.top = sh.top + dy


# S2：費用の枠を消す
s2 = S[1]
for n in ("Rounded Rectangle 15", "Rounded Rectangle 16", "TextBox 17", "Connector 18", "TextBox 19"):
    el = by(s2, n)._element
    el.getparent().remove(el)
center(s2)

# S5：※注記にAPI併用を足す
note = by(S[4], "TextBox 10")
ts = list(note._element.iter(qn("a:t")))
full = "".join(t.text or "" for t in ts)
ts[0].text = "※APIツール（Lステップ等）と併用可／完了画面を増やす場合は追加費用"
for t in ts[1:]:
    t.text = ""

prs.save(str(OUT))
print("saved:", OUT.name)
