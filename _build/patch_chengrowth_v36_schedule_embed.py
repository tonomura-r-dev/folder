# -*- coding: utf-8 -*-
"""チェングロウス ver3.5 → ver3.6：23枚目に「リニューアルのサイト制作への組み込み」の行を追加（2026-10-02）
サンクスLINE誘導・Profile+の実装は、リニューアルのサイト制作に組み込む。時期はサイト制作の進行に応じて決める。
  python3 _build/patch_chengrowth_v36_schedule_embed.py <ver3.5.pptx>
"""
import copy
import sys
from pathlib import Path
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20260929_チェングロウス_LINE公式アカウント運用のご提案ver3.6.pptx"
prs = Presentation(sys.argv[1])
s = prs.slides[22]
by = {sh.name: sh for sh in s.shapes}


def set_par(sh, k, text):
    p = sh.text_frame.paragraphs[k]
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


# 4月の行
set_par(by["Rounded Rectangle 11"], 0, "リニューアル公開と同時に運用を本格化（月額の固定費）")
set_par(by["Rounded Rectangle 11"], 1, "「LINEで登録」・サンクスLINE誘導が稼働し、計測と配信を本格化")
set_par(by["TextBox 14"], 0, "※LINE Profile+の審査期間は、申請内容により変わります／組み込みの時期は、サイト制作の進行に応じて決定／求人ボックス・Indeed経由の応募者への導線は、着地先を確認のうえ決定")

# 新しい行（2行目の複製）を、初期構築の次に入れる
lab = copy.deepcopy(by["Rounded Rectangle 8"]._element)
box = copy.deepcopy(by["Rounded Rectangle 9"]._element)
by["Rounded Rectangle 9"]._element.addnext(box)
by["Rounded Rectangle 9"]._element.addnext(lab)
new_lab = s.shapes[[sh._element for sh in s.shapes].index(lab)]
new_box = s.shapes[[sh._element for sh in s.shapes].index(box)]
set_par(new_lab, 0, "〜3月")
set_par(new_box, 0, "リニューアルのサイト制作への組み込み")
set_par(new_box, 1, "サンクスLINE誘導・Profile+の実装を、リニューアルのサイト制作に組み込む")
# 並び：申請→初期構築→組み込み→LINE追加→4月→翌3月
pairs = [("Rounded Rectangle 4", "Rounded Rectangle 5"), ("Rounded Rectangle 6", "Rounded Rectangle 7"),
         (new_lab, new_box), ("Rounded Rectangle 8", "Rounded Rectangle 9"),
         ("Rounded Rectangle 10", "Rounded Rectangle 11"), ("Rounded Rectangle 12", "Rounded Rectangle 13")]
pairs = [(by[a] if isinstance(a, str) else a, by[b] if isinstance(b, str) else b) for a, b in pairs]
top0 = min(by["Rounded Rectangle 4"].top, by["Rounded Rectangle 5"].top)
bottom = by["Rounded Rectangle 13"].top + by["Rounded Rectangle 13"].height
gap = int(0.25 * 360000)
h = int((bottom - top0 - 5 * gap) / 6)
for i, (a, b) in enumerate(pairs):
    for sh in (a, b):
        sh.top = top0 + i * (h + gap)
        sh.height = h
for n, c in enumerate(s.shapes._spTree.iter(__import__("pptx").oxml.ns.qn("p:cNvPr")), start=2):
    c.set("id", str(n))
prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
