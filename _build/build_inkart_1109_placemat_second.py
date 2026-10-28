#!/usr/bin/env python3
"""インクアート 11/9 紙製ランチョンマット（2回目）のCR。
10/20の完成CR（_images/inkart_1019_placemat_banner_dym_washi.jpg）の写真・枠・ボタンはそのまま、
文字だけ今回の言葉に差し替える（2026-10-27 殿村さん「それでいいかもね」）。

差し替え：
  主役「紙から おもてなしを」 → 「テーブルを 心地よい雰囲気に」（赤は「心地よい雰囲気に」）
  下の1行「風合い豊かな和紙も選べます」 → 「クリスマス・忘年会・新年会の席に」
  残す：帯（文字は新商品→再告知）／紙製ランチョンマット印刷／写真2枚／裏面印刷も可能！／札3つ／ボタン

使い方: PYTHONPATH=<pylib> python3 _build/build_inkart_1109_placemat_second.py
出力:  _images/inkart_1109_placemat_second_v1_1040.jpg
"""
from PIL import Image, ImageDraw, ImageFont
import numpy as np

SRC = "_images/inkart_1019_placemat_banner_dym_washi.jpg"
OUT = "_images/inkart_1109_placemat_second_v1_1040.jpg"
SERIF = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc"
BROWN = (74, 55, 40)
RED = (200, 48, 46)
WHITE = (255, 255, 255)

im = Image.open(SRC).convert("RGB")
W, H = im.size
a = np.array(im)
cream = tuple(int(v) for v in a[420, 150])          # 下地の色（255,247,228）
d = ImageDraw.Draw(im)

# 主役の行（y 215〜365）と、その下の1行（y 385〜450）を下地で塗りつぶす
d.rectangle([20, 213, 1020, 366], fill=cream)
d.rectangle([150, 366, 890, 456], fill=cream)


def text_w(font, s):
    b = font.getbbox(s)
    return b[2] - b[0]


def draw_outlined(x, y, s, font, fill, stroke=8):
    d.text((x, y), s, font=font, fill=fill, stroke_width=stroke, stroke_fill=WHITE)


# 主役：1行に13字は入らないので2行に。上「テーブルを」（茶・小さめ）、下「心地よい雰囲気に」（赤・大）
f_top = ImageFont.truetype(SERIF, 64, index=2)      # index=2 = JP
f_main = ImageFont.truetype(SERIF, 100, index=2)
s1, s2 = "テーブルを", "心地よい雰囲気に"
draw_outlined((W - text_w(f_top, s1)) // 2, 218, s1, f_top, BROWN, stroke=6)
draw_outlined((W - text_w(f_main, s2)) // 2, 272, s2, f_main, RED)

# 下の1行：茶・細めのフチ
f_sub = ImageFont.truetype(SERIF, 42, index=2)
s3 = "クリスマス・忘年会・新年会の席に"
draw_outlined((W - text_w(f_sub, s3)) // 2, 398, s3, f_sub, BROWN, stroke=5)

# 上の帯「新商品」→「再告知」（殿村さん 10/28）。帯の赤で塗って白文字を載せ直す
ribbon_red = tuple(int(v) for v in a[50, 385])
d.rectangle([395, 14, 645, 97], fill=ribbon_red)
f_rib = ImageFont.truetype(SERIF, 68, index=2)
s4 = "再告知"
bb = d.textbbox((0, 0), s4, font=f_rib)
d.text(((W - (bb[2] - bb[0])) // 2 - bb[0], 56 - (bb[1] + bb[3]) // 2), s4, font=f_rib, fill=WHITE)

im.save(OUT, quality=93)
print(OUT, im.size)
