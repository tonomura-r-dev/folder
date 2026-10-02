# -*- coding: utf-8 -*-
"""ver4.0 仕上げ：6枚目の本文を14pt、17・18枚目の表を12pt、注記を短くして下に収める（2026-10-02）
  python3 _build/patch_chengrowth_v40_finish.py <入力.pptx> <出力.pptx>
"""
import sys
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

E = 360000
prs = Presentation(sys.argv[1])

# 6枚目：図形内の本文を13pt（見出し枠は据え置き）
for sh in prs.slides[5].shapes:
    if sh.shape_type == 1 and not (sh.height < 1.3 * E and sh.width < 12 * E):
        for r in sh._element.iter(qn("a:rPr")):
            if r.get("sz") and int(r.get("sz")) <= 1200:
                r.set("sz", "1300")

# 17・18枚目：表を12pt、行の高さを少し広げる
for n in (17, 18):
    s = prs.slides[n - 1]
    for sh in s.shapes:
        if getattr(sh, "has_table", False) and sh.has_table:
            for tag in ("a:rPr", "a:endParaRPr"):
                for r in sh._element.iter(qn(tag)):
                    r.set("sz", "1200" if n == 17 else "1100")
            if n == 17:
                for row in sh.table.rows:
                    row.height = int(row.height * 1.08)
            sh.top = Cm(4.6 if n == 17 else 4.4)
        if sh.name == "TextBox 8":
            tf = sh.text_frame
            p = tf.paragraphs[0]
            p.runs[0].text = ("※クリック数は、弊社シミュレーションの9月の値を、動線ごとの友だち数と配信ごとのクリック比率で振り分けたもの。"
                              "反応率の補正（1.46倍）と複数回のクリック（2.05倍）を含むため、単純な掛け算とは一致しない／画面はイメージ")
            for r in p.runs[1:]:
                r._r.getparent().remove(r._r)
            sh.top = Cm(16.6 if n == 17 else 17.0)

# 25枚目（スケジュール）：1行に収める
for sh in prs.slides[24].shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("LINEを追加してもらう期間"):
        pp = sh.text_frame.paragraphs[1]
        pp.runs[0].text = "離脱防止ポップアップなどで、LINEを追加してもらう（数字はシミュレーションに含めない）"
        for r in pp.runs[1:]:
            r._r.getparent().remove(r._r)
prs.save(sys.argv[2])
print("saved", sys.argv[2])
