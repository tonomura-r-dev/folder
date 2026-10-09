# -*- coding: utf-8 -*-
"""サンクスLINE Appendix（殿村さんPC保存・4:3・7枚）の「実例｜予約後の来院率」1枚だけを
「サンクスLINE｜導入・活用実績」に組み直す（2026-10-08 殿村さん指示）。他のスライドは触らない。
  上 約30%＝クリニックの参考事例（出典区分：LINEヤフーの参考事例／予約完了→LINE友だち追加→予約確認・前日のお知らせ／予約後の来院率 40〜50%改善）
  下 約70%＝人材業界｜サンクスLINE誘導ツール導入実績
      ①導入前後3か月比較 1,532人→2,103人 導入前比137.3%（+571人）／②前年同期間比較 1,650人→2,103人 前年比127.5%（+453人）
      ＋月別比較の小さな棒グラフ（7〜9月・2025年/2026年・前年比107%/151%/127%）＋結論1文
  白＋紺、緑は増加数・前年比の強調だけ。数字は指示どおり（変更しない）。
  python3 _build/patch_thanks_line_appendix_s3_jisseki.py <Appendix.pptx> [出力]
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_parts import *  # noqa

ROOT = Path(__file__).resolve().parent.parent
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "20261008_サンクスLINE誘導のご提案_Appendix_ver1.1.pptx"
PALE, NAVY_LT = "E4E8F6", "8E9BC4"
L, R = 1.79, 25.76
CW = R - L

prs = Presentation(sys.argv[1])
s = next(sl for sl in prs.slides
         if any(sh.has_text_frame and sh.text_frame.text.startswith("実例｜予約後の来院率") for sh in sl.shapes))
keep_header_only(s)
set_paras(find(s, "TextBox 1"), [("サンクスLINE｜導入・活用実績", 0)])
set_paras(find(s, "TextBox 2"), [("LINE活用の参考事例と、サンクスLINE誘導ツールの導入実績", 0)])

# ================= 上段：クリニックの参考事例（約30%）=================
TY, TH = 4.3, 3.85
shape(s, MSO_SHAPE.RECTANGLE, L, TY, CW, TH, None, fill=WHITE, line=LGRAY, lw=1.0)
label(s, L + 0.35, TY + 1.2, 8.7, 0.7, [("クリニック｜LINE活用による来院率改善", 12, True, NAVY, 0)], anchor=MSO_ANCHOR.MIDDLE)
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, L + 0.45, TY + 1.95, 4.3, 0.55, [("弊社運用の実績", 8.5, True, GRAY, 0)],
      fill=WHITE, line=LGRAY, lw=0.75, adj=0.5, margins=(0.05, 0, 0.05, 0))
# 流れ（3ステップ）
fx, bw, gap, bh = 10.7, 2.75, 0.4, 1.15
by = TY + (TH - bh) / 2
steps = [(["予約完了"], PALE, NAVY), (["LINE", "友だち追加"], GREEN, WHITE), (["予約確認・", "前日のお知らせ"], NAVY, WHITE)]
for i, (ls, f, c) in enumerate(steps):
    x = fx + i * (bw + gap)
    shape(s, MSO_SHAPE.RECTANGLE, x, by, bw, bh, [(t, 10, True, c, 0) for t in ls], fill=f, margins=(0.03, 0.02, 0.03, 0.02))
    if i < len(steps) - 1:
        arrow_line(s, x + bw + 0.06, by + bh / 2, x + bw + gap - 0.06, by + bh / 2, NAVY, 1.5)
# 実績
rx, rw = 19.95, 5.65
label(s, rx, TY + 0.85, rw, 0.5, [("予約後の来院率", 10, True, INK, 0)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
tb = rich(s, rx, TY + 1.4, rw, 1.4, [[("40〜50%", 22, True, NAVY), ("改善", 12, True, NAVY)]], align=PP_ALIGN.CENTER)
tb.text_frame.word_wrap = False

# ================= 下段：人材業界の導入実績（約70%）=================
HY = 8.55
label(s, L, HY, 10.5, 0.75, [("人材業界｜サンクスLINE誘導ツール導入実績", 13, True, NAVY, 0)], anchor=MSO_ANCHOR.MIDDLE)
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, L + 10.7, HY + 0.1, 3.2, 0.55, [("DYMの導入実績", 8.5, True, WHITE, 0)], fill=NAVY, adj=0.5,
      margins=(0.05, 0, 0.05, 0))

CY, CH, CWD = 9.45, 5.75, 8.0


def card(x0, title, lab_a, val_a, lab_b, val_b, rate_label, rate, inc):
    shape(s, MSO_SHAPE.RECTANGLE, x0, CY, CWD, CH, None, fill=WHITE, line=LGRAY, lw=1.0)
    rect(s, x0, CY, CWD, 0.75, PALE, [(title, 11.5, True, NAVY, 0)])
    gw = 3.3
    ax, bx = x0 + 0.3, x0 + CWD - 0.3 - gw
    label(s, ax, CY + 1.0, gw, 0.45, [(lab_a, 9.5, False, GRAY, 0)], align=PP_ALIGN.CENTER)
    label(s, ax, CY + 1.45, gw, 0.85, [(val_a, 17, True, INK, 0)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    shape(s, MSO_SHAPE.RIGHT_ARROW, x0 + CWD / 2 - 0.35, CY + 1.6, 0.7, 0.55, None, fill=NAVY, adj=0.4)
    label(s, bx, CY + 1.0, gw, 0.45, [(lab_b, 9.5, False, GRAY, 0)], align=PP_ALIGN.CENTER)
    label(s, bx, CY + 1.45, gw, 0.85, [(val_b, 17, True, NAVY, 0)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    hline(s, x0 + 0.4, x0 + CWD - 0.4, CY + 2.55, LGRAY, 0.75)
    rich(s, x0 + 0.2, CY + 2.7, CWD - 0.4, 1.6, [[(rate_label + " ", 11, True, INK), (rate, 30, True, NAVY)]], align=PP_ALIGN.CENTER)
    rich(s, x0 + 0.2, CY + 4.4, CWD - 0.4, 0.8, [[("増加数 ", 9.5, False, GRAY), (inc, 15, True, GREEN_TX)]], align=PP_ALIGN.CENTER)


card(L, "① 導入前後3か月比較", "導入前", "1,532人", "導入後", "2,103人", "導入前比", "137.3%", "+571人")
card(L + CWD + 0.3, "② 前年同期間比較", "2025年7〜9月", "1,650人", "2026年7〜9月", "2,103人", "前年比", "127.5%", "+453人")

# 月別比較（小さな棒グラフ・図形で作る）
cx0 = L + 2 * CWD + 0.3 + 0.4
cw0 = R - cx0
label(s, cx0, CY + 0.05, cw0, 0.45, [("月別比較（友だち増加数・人）", 9, True, GRAY, 0)])
for i, (name, col) in enumerate((("2025年", NAVY_LT), ("2026年", NAVY))):
    rect(s, cx0 + 0.1 + i * 2.3, CY + 0.62, 0.3, 0.3, col)
    label(s, cx0 + 0.45 + i * 2.3, CY + 0.47, 1.9, 0.6, [(name, 8, False, GRAY, 0)], anchor=MSO_ANCHOR.MIDDLE)
base, hmax, vmax = CY + 4.3, 2.6, 779
months = [("7月", 575, 615, "107%"), ("8月", 516, 779, "151%"), ("9月", 559, 709, "127%")]
bwid, bgap, gstep = 0.8, 0.12, 2.3
c0 = cx0 + 0.95
for i, (m, a, b, r) in enumerate(months):
    c = c0 + i * gstep
    for v, col, xx in ((a, NAVY_LT, c - bgap / 2 - bwid), (b, NAVY, c + bgap / 2)):
        h = hmax * v / vmax
        rect(s, xx, base - h, bwid, h, col)
        lb = label(s, xx - 0.3, base - h - 0.5, bwid + 0.6, 0.45, [(f"{v:,}", 7.5, True, INK, 0)], align=PP_ALIGN.CENTER,
                   anchor=MSO_ANCHOR.BOTTOM)
        lb.text_frame.word_wrap = False
    label(s, c - 1.05, base + 0.05, 2.1, 0.42, [(m, 9, True, INK, 0)], align=PP_ALIGN.CENTER)
    tb = rich(s, c - 1.05, base + 0.45, 2.1, 0.5, [[("前年比 ", 7.5, False, GRAY), (r, 10, True, GREEN_TX)]], align=PP_ALIGN.CENTER)
    tb.text_frame.word_wrap = False
hline(s, cx0 + 0.2, R - 0.1, base, GRAY, 0.75)

# 結論・注記
BY = 15.55
rect(s, L, BY, CW, 0.95, NAVY, [("サンクスLINE誘導ツールの導入後、友だち増加数が導入前比137.3%、前年比127.5%に改善。", 12, True, WHITE, 0)])
label(s, L, 16.65, CW, 0.95,
      [("※上段は弊社運用の実績（美容クリニック・LINE活用による来院率の改善）。サンクスLINE誘導ツールの実績とは別。", 8, False, GRAY, 1),
       ("※下段はサンクスLINE誘導ツール導入前後の友だち増加数（人材業界）。", 8, False, GRAY, 0)])

prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
