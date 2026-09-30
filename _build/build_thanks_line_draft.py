# -*- coding: utf-8 -*-
"""サンクスLINE 簡易資料（5ページ・表紙なし）の下書き。まずP4（広告×LINEの役割分担）だけ図で作って確認する。

チェングロウスの本資料（DYMの書式）をコピーし、不要ページを使い回して中身を作り直す。
部品（書式・図形）は build_chengrowth_v11.py と同じものを使う。

  python3 _build/build_thanks_line_draft.py <出力.pptx>
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_src = (ROOT / "_build/build_chengrowth_v11.py").read_text(encoding="utf-8")
_helpers = _src[_src.index("# ================= helpers"):_src.index("# ============================================================\n# 流用するページの素材")]
_head = _src[_src.index("import shutil"):_src.index("shutil.copyfile(SRC, OUT)")]
_head = _head.replace('SRC = ROOT / "20260922_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver1.2.pptx"',
                      'SRC = ROOT / "20260929_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver2.5.pptx"')
exec(_head)
OUT = Path(sys.argv[1])
shutil.copyfile(SRC, OUT)
prs = Presentation(str(OUT))
S = list(prs.slides)
exec(_helpers)
GRAY, PGRAY = "7F7F7F", "EDEDED"


def node(s, x, y, w, h, lines, fill, col=INK, sz=11, line=None, dash=None):
    sp = box(s, x, y, w, h, fill=fill, line=line, dash=dash, radius=0.10)
    put_text(sp.text_frame, [one(t, sz if i == 0 else sz - 1.5, i == 0, col, align="c", ls=1.2) for i, t in enumerate(lines)],
             anchor="m", ml=0.15, mr=0.15, mt=0.05, mb=0.05)
    return sp


# ============================================================
# P4 広告は「申込みの数」、LINEは「売上」をつくる
# ============================================================
s = S[18]
frame(s, "広告は「申込みの数」、LINEは「売上」をつくる",
      ["広告は申込みを集めるところまで。申込んだ方を来店・成約・リピートまで進めるのがLINEの役割です。",
       "同じ広告費のまま、申込みの「後」を伸ばして成果を増やします。"])

# 役割の見出し
chip(s, 4.1, 4.3, 8.3, 0.75, "広告の役割：申込みを集める", fill=NAVY, sz=11)
chip(s, 13.6, 4.3, 12.7, 0.75, "LINEの役割：申込んだ方を売上につなげる", fill=GREEN, sz=11)

# 行A：広告だけ
yA, h = 5.35, 2.5
node(s, CX0, yA, 2.6, h, ["広告だけ"], PGRAY, GRAY, sz=11)
node(s, 4.1, yA, 3.3, h, ["広告", "月100万円"], PALE, NAVY, sz=12, line=BORDER)
arrow(s, 7.6, yA + 0.8, 0.7, 0.9, fill=GRAY)
node(s, 8.5, yA, 3.9, h, ["申込み", "100件"], PALE, NAVY, sz=12, line=BORDER)
arrow(s, 12.6, yA + 0.8, 0.7, 0.9, fill=GRAY)
node(s, 13.6, yA, 7.0, h, ["フォローはメール・電話だけ", "読まれない・つながらない"], WHITE, GRAY, sz=11, line=GRAY,
     dash=MSO_LINE_DASH_STYLE.DASH)
arrow(s, 20.8, yA + 0.8, 0.7, 0.9, fill=GRAY)
node(s, 21.7, yA, 4.6, h, ["成約 20件", "1件あたり5.0万円"], PGRAY, INK, sz=13)

# 行B：広告＋LINE
yB = 8.25
node(s, CX0, yB, 2.6, h, ["広告", "＋LINE"], GREEN, WHITE, sz=11)
node(s, 4.1, yB, 3.3, h, ["広告", "月100万円"], PALE, NAVY, sz=12, line=BORDER)
arrow(s, 7.6, yB + 0.8, 0.7, 0.9)
node(s, 8.5, yB, 3.9, h, ["申込み", "100件"], PALE, NAVY, sz=12, line=BORDER)
arrow(s, 12.6, yB + 0.8, 0.7, 0.9, fill=GREEN)
sp = box(s, 13.6, yB, 7.0, h, fill=PGREEN, line=GREEN, lw=1.5, radius=0.10)
put_text(sp.text_frame, [one("サンクスLINEで友だちに（月36人）", 11, True, "0B7A3B", align="c", sa=3),
                         one("リマインド・配信・1対1のトーク", 9.5, None, INK, align="c"),
                         one("→ 来店・成約・リピートへ", 9.5, None, INK, align="c")],
         anchor="m", ml=0.15, mr=0.15)
arrow(s, 20.8, yB + 0.8, 0.7, 0.9, fill=GREEN)
node(s, 21.7, yB, 4.6, h, ["成約 24件", "1件あたり約4.3万円"], GREEN, WHITE, sz=13)

# 下：効果3つ
yC, hc, wc = 11.3, 3.2, 8.2
cards = [("成約1件あたりの費用が下がる", "5.0万円 → 約4.3万円", "広告費は同じまま（LINE費用 月3万円込み）"),
         ("取りこぼしていた方を回収", "成約 ＋4件／月", "申込み後に止まっていた方をLINEで後押し"),
         ("友だちは毎月たまる資産", "広告を止めても案内できる", "新商品・キャンペーンを何度でも届けられる")]
for i, (hd, big, sub) in enumerate(cards):
    sp = box(s, CX0 + i * (wc + 0.26), yC, wc, hc, fill=PALE, line=BORDER, radius=0.08)
    put_text(sp.text_frame, [one(hd, 11, True, NAVY, align="c", sa=4),
                             one(big, 16, True, ORANGE, align="c", sa=4),
                             one(sub, 9, None, INK, align="c", ls=1.2)], anchor="m", ml=0.3, mr=0.3)

band(s, 14.9, "弊社実績：美容クリニックで、LINEでのフォローにより予約後の来院率が40〜50%改善", sz=12.5)
foot(s, "※試算例：広告費 月100万円・申込み100件・成約率20%（広告だけ）→24%（広告＋LINE）。友だち数＝申込み100件×完了画面からの遷移80%×友だち追加45%（弊社想定）。"
        "LINE費用はサンクスLINE 月3万円。成約1件あたり＝（広告費＋LINE費用）÷成約件数")

# P4だけ残す
keep = s
lst = prs.slides._sldIdLst
for sl, el in zip(S, list(lst)):
    if sl is not keep:
        prs.part.drop_rel(el.rId); lst.remove(el)
prs.save(str(OUT))
print("saved:", OUT)
