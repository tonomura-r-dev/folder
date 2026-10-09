# -*- coding: utf-8 -*-
"""展示会 LINE運用方針（簡潔な3枚・4:3）。課題 → 方針（店舗ごとにMetaの一部をLINE広告へ）→ サンクスLINE。
5枚版（build_tenjikai_line_5p.py）の論理を3枚にまとめたもの。方針の文は殿村さん確定（2026-10-09）。
  python3 _build/build_tenjikai_line_3p.py <4:3のDYM資料.pptx> [出力.pptx]
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Cm

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_parts import (CARD, GRAY, INK, NAVY, WHITE, arrow_line, find, keep_header_only, rect,  # noqa: E402
                        rich, set_paras, shape)

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1])
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "20261009_展示会LINE運用方針_まとめ.pptx"

POLICY = "同じ方に広告が出すぎている店舗のMeta予算を減らし、減らした分をその店舗のLINE広告に回す方針"
SLIDES = [
    ("課題｜同じ方に広告が出すぎている", "少ない方に広告が重なり、予約1件に15,000円以上かかっている"),
    ("方針｜Meta予算の一部をLINE広告へ", "広告費を増やさずに、Metaで届いていない方にも展示会を案内する"),
    ("サンクスLINE｜申し込んだ方を来場まで追う", "申込完了画面からLINEへ案内し、前日・当日のお知らせで来場につなげる"),
]

X, W = 1.79, 23.97
Y0 = 4.5
S = 13


def text(s, y, runs, h=0.8, x=X, w=W):
    return rich(s, x, y, w, h, runs, anchor=MSO_ANCHOR.TOP)


def head(s, y, txt):
    rect(s, X, y + 0.14, 0.15, 0.5, NAVY)
    text(s, y, [[(txt, 13.5, True, NAVY)]], x=X + 0.35, w=W - 0.35)
    return y + 0.85


def rows(s, y, items, kw=4.6, h=0.72):
    for i, (k, v) in enumerate(items):
        text(s, y + i * h, [[(k, S, True, INK)]], h=h, x=X + 0.35, w=kw)
        text(s, y + i * h, [[(v, S, False, INK)]], h=h, x=X + 0.35 + kw, w=W - 0.35 - kw)
    return y + len(items) * h


def flow(s, y, items, h=1.35, gap=0.75, size=11.5):
    x0, ww = X + 0.35, W - 0.35
    bw = (ww - gap * (len(items) - 1)) / len(items)
    for i, t in enumerate(items):
        last = i == len(items) - 1
        x = x0 + i * (bw + gap)
        paras = [(ln, size, last, WHITE if last else INK, 0) for ln in t.split("\n")]
        shape(s, MSO_SHAPE.RECTANGLE, x, y, bw, h, paras, fill=NAVY if last else WHITE, line=NAVY, lw=1.0,
              margins=(0.1, 0.05, 0.1, 0.05))
        if not last:
            arrow_line(s, x + bw + 0.1, y + h / 2, x + bw + gap - 0.1, y + h / 2, NAVY, w=1.5)
    return y + h


prs = Presentation(str(SRC))
lst = prs.slides._sldIdLst
keep = [i for i, s in enumerate(prs.slides)
        if {"TextBox 1", "TextBox 2", "Connector 3"} <= {sh.name for sh in s.shapes}][:3]
assert len(keep) == 3, keep
for i, el in reversed(list(enumerate(list(lst)))):
    if i not in keep:
        prs.part.drop_rel(el.rId)
        lst.remove(el)

for s, (title, lead) in zip(prs.slides, SLIDES):
    keep_header_only(s)
    t1, t2 = find(s, "TextBox 1"), find(s, "TextBox 2")
    t1.left, t1.top, t1.width, t1.height = Cm(1.52), Cm(0.38), Cm(18.1), Cm(0.94)
    t2.left, t2.top, t2.width, t2.height = Cm(1.79), Cm(2.27), Cm(25.31), Cm(0.96)
    set_paras(t1, [(title, 0)])
    set_paras(t2, [(lead, 0)])
    for t in (t1, t2):
        for r in t.text_frame._txBody.iter(qn("a:rPr")):
            r.set("sz", "1600")

s1, s2, s3 = prs.slides

# 1. 課題
y = head(s1, Y0, "現状の数値")
y = rows(s1, y, [("フリークエンシー", "5.24（同じ方が月に5回以上、同じ広告を見ている）"),
                 ("CPA", "15,000円以上（予約1件あたりの広告費）")])
y = head(s1, y + 0.5, "なぜ効率が落ちるのか")
y = flow(s1, y + 0.1, ["オーディエンスが少ない", "予算を使い切るため、\n同じ方に繰り返し表示", "予約は増えず、\nCPAだけ上がる"])
y = head(s1, y + 0.6, "オーディエンスの少ないエリア（例）")
rows(s1, y, [("三鷹", "10,700〜12,600"), ("高松", "上限25,000"), ("日向", "4,400")])

# 2. 方針
y = head(s2, Y0, "進め方")
rect(s2, X + 0.35, y + 0.05, W - 0.35, 1.5, CARD,
     [(POLICY, 14, True, NAVY, 0)], line=NAVY, lw=1.0, margins=(0.4, 0.1, 0.4, 0.1))
y = head(s2, y + 2.05, "銀座店の例（月）")
y = flow(s2, y + 0.1, ["Meta広告\n30万円", "Meta広告 20万円\n（同じ方への重複表示を削る）", "LINE広告 10万円\n（減らした分を回す）"])
y = head(s2, y + 0.6, "ポイント")
y = rows(s2, y, [("広告費", "店舗ごとの総額は変わらない（銀座店は30万円のまま）"),
                 ("LINE広告", "Metaとは別の媒体。国内の月間利用者数 1億人以上（LINEヤフー公表）"),
                 ("見直さない店舗", "オーディエンスに対して予算がちょうどいい店舗は今のまま")])
text(s2, y + 0.35, [[("仮に銀座店でLINE広告10万円 → クリック1,000（CPC 100円）→ 予約10件（予約率1.0%）", 11, False, GRAY)]], h=0.6)

# 3. サンクスLINE
y = head(s3, Y0, "流れ")
y = flow(s3, y + 0.1, ["申込フォーム", "申込完了画面", "LINE\n友だち追加", "前日・当日の\nお知らせ", "来場"], h=1.5, gap=0.6)
y = head(s3, y + 0.6, "友だちへの配信")
y = rows(s3, y, [("タイミング", "登録直後 → 2週間前 → 3日前 → 当日朝 → 終了後"),
                 ("内容", "概要・予約 → 見どころ → リマインド・アクセス → 開催のお知らせ → お礼・次回案内")])
y = head(s3, y + 0.5, "実績・費用")
rows(s3, y, [("実績", "弊社運用の美容クリニックで、予約後の来院率 40〜50%改善"),
             ("費用", "初期15万円・月額3万円〜（追加でかかる費用はこれだけ）")])

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
