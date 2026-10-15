"""インクアート10/19 ランチョンマットのバナーの「和風」を「和紙」に直す（1文字だけ）。

DYMの完成バナーの見出し「紙製…」の「紙」を切り出して縮め、
「風合い豊かな和風も選べます」の2つ目の「風」の位置に置き換える。
使い方: python3 _build/fix_inkart_1019_washi.py
"""
from PIL import Image, ImageFilter
import numpy as np

SRC = "_images/inkart_1019_placemat_banner_dym.jpg"
OUT = "_images/inkart_1019_placemat_banner_dym_washi.jpg"

im = Image.open(SRC).convert("RGB")
a = np.array(im).astype(float)
lum = a.mean(axis=2)

# 置き換える「風」（和風の風）の枠と、その左右の文字のすき間
KX0, KX1, KY0, KY1 = 554, 615, 389, 449
GAP_L, GAP_R = (612, 619)  # 風ともの間（背景だけの列）

# 1) 「風」を消す：行ごとに、すき間の列の色で塗る（帯の色の変化も保てる）
for y in range(KY0, KY1):
    bg = a[y, GAP_L:GAP_R].mean(axis=0)
    a[y, KX0:KX1] = bg

# 2) 見出しの「紙」を切り出して不透明度（濃さ）を作る
HX0, HX1, HY0, HY1 = 141, 213, 133, 205
crop = lum[HY0:HY1, HX0:HX1]
bg_l, ink_l = np.percentile(crop, 95), np.percentile(crop, 3)
alpha = np.clip((bg_l - crop) / (bg_l - ink_l), 0, 1)

# 3) 行の文字の大きさに合わせて縮める（文字の送り幅 71px → 53.5px）
scale = 53.5 / 71.0
w = round(alpha.shape[1] * scale); h = round(alpha.shape[0] * scale)
am = Image.fromarray((alpha * 255).astype("uint8")).resize((w, h), Image.LANCZOS)

# 4) 置く位置：元の「風」の中心（584, 419）にそろえる
cx, cy = 584, 419
x0, y0 = cx - w // 2, cy - h // 2

# 文字色＝同じ行の「和」の濃い部分の色
ink = a[394:445, 505:554].reshape(-1, 3)
ink = ink[ink.mean(axis=1) < np.percentile(ink.mean(axis=1), 15)].mean(axis=0)

# 5) 白いふち（元の行の文字と同じ。下の帯の上で見える）→ 文字の順に重ねる
canvas = Image.fromarray(a.clip(0, 255).astype("uint8"))
big = Image.new("L", canvas.size, 0); big.paste(am, (x0, y0))
outline = big.point(lambda v: 255 if v > 60 else 0).filter(ImageFilter.MaxFilter(5)).filter(ImageFilter.GaussianBlur(0.6))
white = Image.new("RGB", canvas.size, (255, 253, 247))
canvas = Image.composite(white, canvas, outline)
inkimg = Image.new("RGB", canvas.size, tuple(int(c) for c in ink))
canvas = Image.composite(inkimg, canvas, big)

canvas.save(OUT, quality=95)
print(OUT, canvas.size)
