# -*- coding: utf-8 -*-
"""横浜シミズ資料（18枚）：左上のタイトルの位置をそろえる（2026-10-29 殿村さん「左上のアジェンダの縦横関係整えて」）。
ページによって、タイトルの枠の左余白（0／0.03／0.1インチ）と位置が違い、文字の出だしがずれていた。
いちばん多い形（2・3ページなど：x 0.598・y 0.15・高さ0.37・左右余白0.03・上下余白0.01インチ・上下中央）にそろえる。
色も、8・9・14ページだけ濃い紺（002060）だったので、ほかと同じ 323366 にそろえる。文字・大きさは変えない。表紙（1ページ）と裏表紙は触らない。出力を省くと上書き。

    python -m _build.patch_yokohama_shimizu_title_align <資料.pptx> [出力.pptx]
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

X, Y, H = Emu(546811), Emu(137160), Emu(338328)
INS = dict(lIns="28800", rIns="28800", tIns="7200", bIns="7200", anchor="ctr")


def main(src, out):
    prs = Presentation(src)
    ref = prs.slides[1]
    t0 = next(sh for sh in ref.shapes if sh.shape_id == 2)
    x, y, h = t0.left, t0.top, t0.height
    n = 0
    for s in list(prs.slides)[1:]:
        cands = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and sh.top < Emu(548640)]
        if not cands:
            continue
        sh = cands[0]
        sh.left, sh.top, sh.height = x, y, h
        bp = sh.text_frame._txBody.find(qn("a:bodyPr"))
        for k, v in INS.items():
            bp.set(k, v)
        for c in sh.text_frame._txBody.iter(qn("a:srgbClr")):
            if c.get("val") == "002060":
                c.set("val", "323366")
        n += 1
    prs.save(out)
    print("saved:", out, n)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1])
