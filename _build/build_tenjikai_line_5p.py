# -*- coding: utf-8 -*-
"""展示会 LINE運用方針（テキストベース5枚・4:3）。文言は _drafts/展示会LINE_スライド要旨（2026-10-09）.md のとおり。
器はDYMの4:3マスターの既存PPTX（サンクスLINE Appendix 等）。ヘッダー（タイトル・リード・区切り線）だけ残して中身を入れる。
  python3 _build/build_tenjikai_line_5p.py <4:3のDYM資料.pptx> [出力.pptx]
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_parts import INK, NAVY, GRAY, find, keep_header_only, rich, set_paras, style_run  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1])
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "20261009_展示会LINE運用方針.pptx"

SLIDES = [
    ("課題｜同じ方に広告が出すぎている",
     ["少ない方に広告が重なり、効率が落ちている"]),
    ("方針｜エリア予算を見直し、浮いた予算をLINEへ",
     ["オーディエンスに対して予算が多いエリアを削り、新規予算10〜30万円をつくる"]),
    ("LINE広告（通常）｜Metaで届いていない方に届ける",
     ["媒体を変え、Metaとは違うユーザー層に展示会を案内する"]),
    ("サンクスLINE｜申し込んだ方を来場まで追う",
     ["申込完了画面からLINEへ案内し、前日・当日のお知らせで来場につなげる"]),
    ("友だち登録後の配信と仮の数値",
     ["今の友だちとサンクスLINEで増えた友だちへ、展示会ごとに配信で案内する"]),
]

X, W = 1.79, 23.97          # 本文の左端・幅（リード行と同じ左端）
Y0 = 4.8                    # 本文の開始位置（区切り線 3.86 の下）
B = 16                      # 本文の文字サイズ


def text(s, y, runs_by_para, h=1.0, x=X, w=W):
    return rich(s, x, y, w, h, runs_by_para, anchor=MSO_ANCHOR.TOP)


def table_rows(s, y, rows, kw=5.4, h=0.95, size=B):
    """左＝項目（紺・太字）、右＝内容。行ごとに同じ高さで並べる"""
    for i, (k, v) in enumerate(rows):
        text(s, y + i * h, [[(k, size, True, NAVY)]], h=h, w=kw)
        text(s, y + i * h, [[(v, size, False, INK)]], h=h, x=X + kw, w=W - kw)
    return y + len(rows) * h


prs = Presentation(str(SRC))
# ヘッダー一式（TextBox 1/2・Connector 3）がそろっているページを5枚だけ残す
lst = prs.slides._sldIdLst
keep = [i for i, s in enumerate(prs.slides)
        if {"TextBox 1", "TextBox 2", "Connector 3"} <= {sh.name for sh in s.shapes}][:5]
assert len(keep) == 5, keep
for i, el in reversed(list(enumerate(list(lst)))):
    if i not in keep:
        prs.part.drop_rel(el.rId)
        lst.remove(el)

for s, (title, lead) in zip(prs.slides, SLIDES):
    keep_header_only(s)
    t1, t2 = find(s, "TextBox 1"), find(s, "TextBox 2")
    t1.left, t1.top, t1.width, t1.height = Cm(1.52), Cm(0.38), Cm(18.1), Cm(0.94)
    t2.left, t2.top, t2.width, t2.height = Cm(1.79), Cm(2.27), Cm(25.31), Cm(0.96 if len(lead) == 1 else 1.5)
    set_paras(t1, [(title, 0)])
    set_paras(t2, [(l, 0) for l in lead])
    for r in t1.text_frame._txBody.iter(qn("a:rPr")):
        r.set("sz", "1600")
    for r in t2.text_frame._txBody.iter(qn("a:rPr")):
        r.set("sz", "1600")

s1, s2, s3, s4, s5 = prs.slides

# 1. 課題（リードは1行にそろえ、数値は本文へ）
y = table_rows(s1, Y0, [("フリークエンシー", "5.24"), ("CPA", "15,000円以上")])
text(s1, y + 0.6, [[("オーディエンスの少ないエリアに予算が多い", B, False, INK)]])
table_rows(s1, y + 1.8, [("三鷹", "10,700〜12,600"), ("高松", "上限25,000"), ("日向", "4,400")])

# 2. 方針
table_rows(s2, Y0, [("例", "銀座店 30万円 → 20万円"), ("使い道", "LINE広告（通常）")])

# 3. LINE広告（通常）
table_rows(s3, Y0, [("予約", "仮に予算20万円 → クリック2,000 → 予約20件"), ("予約1件あたり", "10,000円")])

# 4. サンクスLINE
table_rows(s4, Y0, [("来場率", "仮に 50% → 70%（申込100件あたり 来場50 → 70人）"),
                    ("費用", "初期15万円・月額3万円〜")])

# 5. 配信と仮の数値
y = table_rows(s5, Y0, [("登録直後", "概要・予約"), ("2週間前", "見どころ"), ("3日前", "リマインド・アクセス"),
                        ("当日朝", "開催のお知らせ"), ("終了後", "お礼・次回案内")])
table_rows(s5, y + 0.6, [("予約", "仮に 26件／月"), ("予約1件あたり", "8,846円（現状：15,000円以上）")])

# 字体をメイリオに統一（latin → ea → cs の順）
AFTER = {qn(x) for x in ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")}
for s in prs.slides:
    for r in s.shapes._spTree.iter(qn("a:r")):
        if r.find(qn("a:rPr")) is None:
            r.insert(0, r.makeelement(qn("a:rPr"), {"lang": "ja-JP"}))
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for rpr in s.shapes._spTree.iter(qn(tag)):
            for ft in ("a:latin", "a:ea", "a:cs"):
                for e in rpr.findall(qn(ft)):
                    rpr.remove(e)
            nxt = next((c for c in rpr if c.tag in AFTER), None)
            for ft in ("a:latin", "a:ea", "a:cs"):
                e = rpr.makeelement(qn(ft), {"typeface": "メイリオ"})
                if nxt is not None:
                    nxt.addprevious(e)
                else:
                    rpr.append(e)

prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
