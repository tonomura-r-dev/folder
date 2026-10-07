# -*- coding: utf-8 -*-
"""サンクスLINE誘導・業界別 ver3.5 に「前後検索データを活用したステップ配信設計」を1枚追加（2026-10-07 殿村さん指示）。
2枚目「機能・費用」と3枚目「実例」の間に入れる。既存5枚は触らない。
  左＝① 前後検索データを分析（検索前後のニーズを把握。抽象化した検索推移の棒グラフ）
  中＝② 配信内容・順番を設計（検索データ → LINE配信 への変換）
  右＝③ アクションが起きやすい追加直後に配信（友だち追加→初回配信→数日後→1〜2週間のタイムライン）
  下＝結論「前後検索でユーザーニーズを捉え、友だち追加直後のステップ配信に反映」
特定KW・業界データ・数値（開封率/CVR/Day固定）は入れない。白＋紺、LINE側だけ緑。
  python3 _build/patch_thanks_line_gyokai_add_zengo.py <業界別ver3.5.pptx> [出力]
"""
import copy
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_parts import *  # noqa

ROOT = Path(__file__).resolve().parent.parent
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "20261007_サンクスLINE誘導のご提案_業界別ver3.6.pptx"
PALE, PALE2, OFF = "E4E8F6", "C9D2EC", "F2F2F2"
L, R = 2.2, 31.7

prs = Presentation(sys.argv[1])
S = list(prs.slides)
assert S[1].shapes[0].text_frame.text.startswith("サンクスLINE｜機能") and S[2].shapes[0].text_frame.text.startswith("実例"), "2・3枚目が想定と違う"
ref = S[3]  # ヘッダー（タイトル・リード・区切り線）の雛形
s = prs.slides.add_slide(ref.slide_layout)
for sh in list(s.shapes):
    sh._element.getparent().remove(sh._element)
for nm in ("TextBox 1", "TextBox 2", "Connector 3"):
    s.shapes._spTree.append(copy.deepcopy(next(x for x in ref.shapes if x.name == nm)._element))
set_paras(find(s, "TextBox 1"), [("前後検索データを活用したステップ配信設計", 0)])
set_paras(find(s, "TextBox 2"), [("アクションが起きやすい友だち追加直後に、ユーザーの検討行動に合わせた配信を設計", 0)])
lst = prs.slides._sldIdLst  # 3枚目へ移動
el = lst[-1]
lst.remove(el)
lst.insert(2, el)

# ---- 3ブロックの枠 ----
GAPB = 1.1
BW = [8.6, 8.6, R - L - 2 * GAPB - 17.2]
BX = [L, L + BW[0] + GAPB, L + BW[0] + BW[1] + 2 * GAPB]
TY, BY0, BY1 = 4.5, 6.6, 14.9


def block_head(i, num, step, head, accent):
    x, w = BX[i], BW[i]
    shape(s, MSO_SHAPE.OVAL, x, TY, 0.85, 0.85, [(num, 11, True, WHITE, 0)], fill=accent, margins=(0, 0, 0, 0))
    label(s, x + 1.0, TY - 0.05, w - 1.0, 0.95, [(step, 10.5, False, GRAY, 0)], anchor=MSO_ANCHOR.MIDDLE)
    label(s, x, TY + 0.95, w, 0.9, [(head, 14, True, NAVY, 0)], anchor=MSO_ANCHOR.MIDDLE)
    hline(s, x, x + w, TY + 1.9, accent, 1.5)


block_head(0, "1", "前後検索データを分析", "検索前後のニーズを把握", NAVY)
block_head(1, "2", "配信内容・順番・タイミングを設計", "配信内容・順番を設計", NAVY)
block_head(2, "3", "友だち追加直後からステップ配信", "アクションが起きやすい追加直後に配信", GREEN_TX)
for i in range(2):
    xa = BX[i] + BW[i] + 0.15
    shape(s, MSO_SHAPE.RIGHT_ARROW, xa, 9.6, GAPB - 0.3, 1.3, None, fill=NAVY, adj=0.45)

# ---- ① 前後検索の画面風（横軸＝検索起点からの日数、縦軸＝UU、点＝キーワード。文字は伏せた帯で抽象化）----
import random
x, w = BX[0], BW[0]
gy0, gy1 = BY0 + 0.2, BY0 + 4.7
TEAL = "2EC4B6"
ax0, ax1 = x + 0.7, x + w - 0.1          # プロット域（横）
ay0, ay1 = gy0 + 0.2, gy1 - 0.1          # プロット域（縦）
hline(s, ax0, ax1, ay1, GRAY, 0.75)
vline(s, ax0, ay0, ay1, GRAY, 0.75)
cx = (ax0 + ax1) / 2
vline(s, cx, ay0, ay1, GRAY, 0.75)
label(s, x - 0.1, ay0 + 0.5, 0.8, 0.5, [("UU", 7, False, GRAY, 0)], align=PP_ALIGN.CENTER)
ticks = [("-15日", 0.0), ("-7日", 0.17), ("-3日", 0.33), ("0日(検索起点)", 0.5), ("3日", 0.67), ("7日", 0.83), ("15日", 1.0)]
for tname, tp in ticks:
    tx = ax0 + (ax1 - ax0) * tp
    tb = label(s, tx - 1.1, ay1 + 0.02, 2.2, 0.4, [(tname, 6.5, False, GRAY, 0)], align=PP_ALIGN.CENTER)
    tb.text_frame.word_wrap = False
rng = random.Random(7)
pts = []
for _ in range(70):
    # 検索起点の近くに密集、日数が離れるほどまばら。上に行くほど少ない
    d = rng.gauss(0, 0.32)
    d = max(-1, min(1, d))
    u = rng.random() ** 2.2
    if abs(d) < 0.03:
        u = max(u, rng.random() ** 1.2)
    pts.append((d, u))
pts.sort(key=lambda p: -p[1])
for i, (d, u) in enumerate(pts):
    px = cx + d * (ax1 - ax0) / 2
    py = ay1 - 0.15 - u * (ay1 - ay0 - 0.4)
    shape(s, MSO_SHAPE.OVAL, px - 0.07, py - 0.07, 0.14, 0.14, None, fill=TEAL)
    if i % 4 == 0 and px + 0.9 < ax1:
        lw_ = 0.3 + rng.random() * 0.45
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, px + 0.12, py - 0.07, lw_, 0.14, None, fill="C9CED9", adj=0.5)
label(s, x, gy1 + 0.35, w, 0.5, [("※実際の前後検索データをもとに分析（キーワードは案件ごと）", 8, False, GRAY, 0)], align=PP_ALIGN.CENTER)
label(s, x + 0.3, gy1 + 0.9, w - 0.6, 3.0, [(t, 10.5, False, INK, 4) for t in
                                            ("検討前に何を検索しているか", "検討後に何を検索しているか", "ユーザーの不安・比較軸・関心テーマを把握")])

# ---- ② 検索データ → LINE配信 への変換 ----
x, w = BX[1], BW[1]
cw = (w - 1.6) / 2
lx, rx = x, x + cw + 1.6
rect(s, lx, BY0, cw, 0.75, PALE, [("検索データ", 10.5, True, NAVY, 0)])
rect(s, rx, BY0, cw, 0.75, GREEN_BG, [("LINE配信", 10.5, True, GREEN_TX, 0)])
rows = [("不安", "どの情報を"), ("比較軸", "どの順番で"), ("関心テーマ", "いつ配信するか")]
ry = BY0 + 1.2
for a, b in rows:
    shape(s, MSO_SHAPE.RECTANGLE, lx, ry, cw, 1.1, [(a, 11.5, True, NAVY, 0)], fill=WHITE, line=NAVY, lw=1.0)
    arrow_line(s, lx + cw + 0.2, ry + 0.55, rx - 0.2, ry + 0.55, NAVY, 1.75)
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, rx, ry, cw, 1.1, [(b, 11.5, True, WHITE, 0)], fill=GREEN, adj=0.3)
    ry += 1.1 + 0.55
label(s, x, ry + 0.2, w, 1.6, [("ニーズをもとに、配信内容・順番・タイミングを", 10.5, False, INK, 2), ("業界・商材ごとに設計", 10.5, False, INK, 0)],
      align=PP_ALIGN.CENTER)

# ---- ③ 追加直後からのタイムライン ----
x, w = BX[2], BW[2]
tx = x + 2.9                       # 縦線
vline(s, tx, BY0 + 0.3, BY1 - 0.9, GREEN, 2.0)
tl = [("友だち追加", None, True), ("初回配信", "興味喚起", False), ("数日後", "比較検討情報・不安解消", False), ("1〜2週間", "CVオファー", False)]
ty = BY0 + 0.3
step = (BY1 - 0.9 - BY0 - 0.3) / (len(tl) - 1)
for k, (when, what, start) in enumerate(tl):
    yy = ty + k * step
    shape(s, MSO_SHAPE.OVAL, tx - 0.3, yy - 0.3, 0.6, 0.6, None, fill=GREEN if start else WHITE, line=GREEN, lw=2.0)
    label(s, x, yy - 0.5, 2.5, 1.0, [(when, 11 if not start else 12, True, GREEN_TX if start else INK, 0)], align=PP_ALIGN.RIGHT,
          anchor=MSO_ANCHOR.MIDDLE)
    if what:
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, tx + 0.7, yy - 0.5, x + w - tx - 0.7, 1.0, [(what, 11, True, GREEN_TX, 0)], fill=GREEN_BG,
              adj=0.3, margins=(0.15, 0, 0.15, 0))
    else:
        label(s, tx + 0.7, yy - 0.5, x + w - tx - 0.7, 1.0, [("サンクスLINE誘導などで追加", 10, False, GRAY, 0)], anchor=MSO_ANCHOR.MIDDLE)
label(s, x, BY1 + 0.15, w, 0.6, [("※配信内容・間隔は業界・商材ごとに設計", 9, False, GRAY, 0)])

# ---- 結論（1文）----
label(s, L, 16.0, R - L, 1.4, [("前後検索でユーザーニーズを捉え、友だち追加直後のステップ配信に反映", 17, True, NAVY, 0)],
      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
