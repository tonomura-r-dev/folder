# -*- coding: utf-8 -*-
"""人材業界（派遣・バイト・単発）：リッチメニューのページ（21枚目）を組み直す（2026-10-21 殿村さん）。

    python -m _build.patch_haken_richmenu <入力.pptx> <出力.pptx>

  入力は、ステップ配信を14日間にした版（`_build.patch_haken_14days` の出力）。
  ボタンは横3×縦2の6つ：仕事情報／マイページ／Web登録・来社予約（上の段＝探す・進む）、
  よくある質問／社員インタビュー／お問い合わせ（下の段＝不安を減らす）。
  「通知を止める」「今は働いている」は外し、「今すぐ探す」「条件に合う新着」は「仕事情報」にまとめた（殿村さん指示）。
  左にスマホのLINE画面の画像（殿村さんが生成。トーク部分を少し詰めた `_images/haken_richmenu_phone.png`）、
  右に段ごとの役割を2つ（ボタン名は画像にあるので書かない）。メニューの段から細い線でつなぐ。
  帯はリードと同じなので外した。ほかのページは変えない。
"""
import sys

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm

from _build.pptx_parts import find, hline, rect, rich, set_para

NAVY = "1E2761"
SUB = "64748B"
LINE_GRAY = "A6A6A6"
IMAGE = "_images/haken_richmenu_phone.png"
LEAD = "求人を探す入口と、応募前の不安を減らすボタンを並べる。"
NOTE = "※ 画面はイメージ。ボタン名は運用しながら調整する"
GROUPS = [("探す・進む", ["条件に合う求人を探し、", "応募・登録はサイトで行う"]),
          ("不安を減らす", ["職場や働き方の疑問に、", "応募の前にお答えする"])]
REMOVE = {"Picture 26", "Rectangle 23", "TextBox 24"} | {f"Rectangle {n}" for n in (5, 8, 11, 14, 17, 20)} \
    | {f"TextBox {n}" for n in range(6, 23)}

# 配置（cm）。画像は殿村さんの生成画像（1024×1536）から、トーク画面の空いた部分（縦200px）を詰めて余白を切った 836×1302px。
# メニューの上の段は y=646〜891px、下の段は 894〜1131px、メニューの右端は x=788px
PW, PH = 836, 1302
ROW_PX = [(646 + 891) / 2, (894 + 1131) / 2]
MENU_RIGHT_PX = 788
IH = 12.7
IW = IH * PW / PH
IX, IY = 4.73, 3.68                          # スマホと右の説明をまとめて中央に寄せる
GX, GW, GH = IX + IW + 2.0, 7.6, 1.85        # 右の説明


def main(src, out):
    prs = Presentation(src)
    s = list(prs.slides)[20]
    assert find(s, "TextBox 1").text_frame.text.endswith("リッチメニュー")
    for sh in list(s.shapes):
        if sh.name in REMOVE:
            sh._element.getparent().remove(sh._element)
    t2 = find(s, "TextBox 2")
    set_para(t2, 0, LEAD)
    while len(t2.text_frame.paragraphs) > 1:
        p = t2.text_frame.paragraphs[1]._p
        p.getparent().remove(p)
    set_para(find(s, "TextBox 25"), 0, NOTE)

    s.shapes.add_picture(IMAGE, Cm(IX), Cm(IY), Cm(IW), Cm(IH))
    for (head, desc), px in zip(GROUPS, ROW_PX):
        cy = IY + IH * px / PH
        hline(s, IX + IW * MENU_RIGHT_PX / PW + 0.15, GX - 0.15, cy, LINE_GRAY, w=1.0)
        rect(s, GX, cy - GH / 2, 0.12, GH, NAVY)
        rich(s, GX + 0.3, cy - GH / 2, GW - 0.3, GH,
             [[(head, 14, True, NAVY)]] + [[(d, 12, False, SUB)] for d in desc],
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
