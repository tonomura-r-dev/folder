# -*- coding: utf-8 -*-
"""横浜シミズ：料金プラン（P15）の後ろに、成果報酬の料金プランのページを1枚足す（2026-10-27 殿村さん・佐村さんFB）。

    python -m _build.patch_yokohama_shimizu_seika_plan <殿村さんのPC保存版.pptx> <出力.pptx>

  コンサル版＝既にあるP15（3つのプラン）。成果報酬版＝P15を写して、プラン②・③の月額運用費を成果報酬に（プラン①はコンサル版と同じ）。
  ・初期費用はコンサル版と同じ 295,000円〜
  ・月額運用費の代わりに「成果報酬（応募1件）」：プラン② 6,000円／プラン③ 10,000円
  ・月額ツール・LINE関連費 65,000円〜（SIM②_成果報酬の45行と同じ）
  運用内容のページ（P16→P17）はそのまま両方に効く。ほかのページは変えない。
"""
import copy
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

from _build.patch_tanaka_v4_thanks_option import cell, set_lines, sid
from _build.patch_tanaka_v5 import copy_slide


def drop_col(tbl_shape, idx):
    tbl = tbl_shape.table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    grid.remove(grid.findall(qn("a:gridCol"))[idx])
    for tr in tbl.findall(qn("a:tr")):
        tr.remove(tr.findall(qn("a:tc"))[idx])


def main(src, out):
    prs = Presentation(src)
    assert len(prs.slides) == 18
    base = prs.slides[14]
    assert sid(base, 635).text_frame.text == "料金プラン：3つのプランの費用"
    s = copy_slide(prs, base)
    lst = prs.slides._sldIdLst
    new = list(lst)[-1]
    lst.remove(new)
    list(lst)[14].addnext(new)

    set_lines(sid(s, 635).text_frame, ["料金プラン（成果報酬）：月額を抑え、応募1件ごとにお支払い"])
    set_lines(sid(s, 636).text_frame, ["成果報酬プラン"])
    set_lines(sid(s, 637).text_frame, ["初期費用は同じ。月額運用費の代わりに、応募1件ごとの成果報酬"])
    tshape = next(sh for sh in s.shapes if sh.has_table)
    tb = tshape.table
    assert tb.cell(2, 0).text == "月額運用費"
    # プラン①（初期設計・構築のみ）の列もコンサル版と同じ内容で残す（2026-10-27 殿村さん「コンサルと全く同じ内容でいい」）
    cell(tb, 2, 0, "成果報酬\n（応募1件）")
    cell(tb, 2, 2, "20,000円")
    cell(tb, 2, 3, "30,000円")
    cell(tb, 4, 2, "65,000円〜\n＋成果報酬")
    cell(tb, 4, 3, "65,000円〜\n＋成果報酬")
    assert [tb.cell(1, c).text for c in (1, 2, 3)] == ["295,000円〜"] * 3
    assert [tb.cell(3, c).text for c in (1, 2, 3)] == ["65,000円〜"] * 3
    set_lines(sid(s, 642).text_frame, ["※成果報酬の対象＝LINEからの新規応募・再応募（再応募は、勤務希望をLINEで受け付ける場合）。契約期間は6か月〜"])
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

# 2026-10-27 追記：P18（今後の展開）のまとめの最後の1行は、殿村さん指示で
# 「月額215,000円〜（ツール費込み）で始め、成果を見て調整します。」→「コンサル・成果報酬のどちらのプランも、成果を見て調整します。」に直した（PC保存版に直接反映）。
