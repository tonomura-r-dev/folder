# -*- coding: utf-8 -*-
"""指定したページだけ、色の割り当て（clrMap）を提案FMTと同じにする（2026-10-29 エムアンドエムサービス P3・P28）。

    python -m _build.fix_clrmap_pages <入力.pptx> <出力.pptx> 3 28

  原因：この資料のスライドマスターは「文字2（tx2）＝薄いベージュ（lt2）」の割り当てで、FMTのページ（tx2＝濃い青 dk2）を
  持ってくると、見出しの帯・札・「DYM→」がベージュや白っぽい文字になる。ページの側で FMT と同じ割り当てを上書きして直す。
  図形・文字・位置は触らない。
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

FMT_MAP = dict(bg1="lt1", tx1="dk1", bg2="lt2", tx2="dk2", accent1="accent1", accent2="accent2", accent3="accent3",
               accent4="accent4", accent5="accent5", accent6="accent6", hlink="hlink", folHlink="folHlink")


def main(src, out, pages):
    prs = Presentation(src)
    for n in pages:
        sld = prs.slides[n - 1]._element
        ovr = sld.find(qn("p:clrMapOvr"))
        if ovr is None:
            ovr = sld.makeelement(qn("p:clrMapOvr"), {})
            sld.append(ovr)
        for c in list(ovr):
            ovr.remove(c)
        ovr.append(ovr.makeelement(qn("a:overrideClrMapping"), FMT_MAP))
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], [int(x) for x in sys.argv[3:]])
