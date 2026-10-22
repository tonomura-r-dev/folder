#!/usr/bin/env python3
"""インクアート 11月リッチメニュー（11/1〜）の登録用画像を、生成画像から組み直す。

生成画像（_images/inkart_richmenu_202611_generated_v1.webp・1527×1030）は上段が高さの約56%で、
LINEの6分割テンプレ（上下ちょうど半分）とずれる。6枚の付箋形の枠を切り出し、
1200×810（LINEの許容サイズ）の格子に置き直す。商品一覧の枠のカートは普通のカートに描き直す。

使い方: python3 _build/build_inkart_richmenu_202611.py
出力:   _images/inkart_richmenu_202611_1200x810.png（登録用）
        _images/inkart_richmenu_202611_1200x810_gridcheck.png（格子を重ねた確認用）
"""
from PIL import Image, ImageDraw
import numpy as np

SRC = "_images/inkart_richmenu_202611_generated_v1.webp"
OUT = "_images/inkart_richmenu_202611_1200x810.png"
OUT_GRID = "_images/inkart_richmenu_202611_1200x810_gridcheck.png"

W, H = 1200, 810
CELL_W, CELL_H = W // 3, H // 2          # 400 × 405
GREEN = (2, 145, 95)                     # 生成画像の下地の緑（実測）
ORANGE = (240, 145, 45)                  # #F0912D

# 生成画像の枠の位置（緑の溝から実測）
COLS = [(21, 519), (534, 1003), (1016, 1509)]
ROWS = [(26, 575), (597, 994)]

# 置き先の枠の大きさ（列ごとに幅を変えて、縦横の縮み・伸びを5%以内に収める）
TOP_W = [360, 350, 360]
TOP_H = 393
BOT_W = [360, 350, 360]
BOT_H = 310
TOP_Y = 6
BOT_Y = 440


def redraw_cart(im):
    """上段右（商品一覧／ご注文）のカートと印刷物を消して、普通のカートを描く。"""
    x0, x1 = COLS[2]
    y0, y1 = ROWS[0]
    a = np.array(im).astype(int)
    card = a[y0:y1 + 1, x0:x1 + 1]
    white = card.min(-1) > 235
    ys, xs = np.where(white)
    in_l = x0 + xs.min()                          # 白い内側の左端
    in_r = x1 - 22                                # 右端は枠（橙の縁 約22px）の内側まで
    in_b = y1 - 22                                # 下端も同じ
    border = tuple(int(v) for v in a[y0 + 300, x1 - 8])   # 橙の縁の色（右側で実測）
    d = ImageDraw.Draw(im)
    # 文字の下（y0+268〜）を白で塗り、はみ出していた印刷物の名残は縁ごと塗り直す
    top = y0 + 268
    d.rectangle([in_l, top, x1, y1], fill=border)
    d.rectangle([in_l, top, in_r, in_b], fill=(255, 255, 255))
    # カート
    cx = (in_l + in_r) // 2 + 38      # 取っ手が左に出るぶん右に寄せて、全体の中心を枠の中心に合わせる
    lw = 22
    bx0, bx1 = cx - 110, cx + 120     # かごの上辺
    by0, by1 = top + 60, top + 170    # かごの上下
    # 取っ手
    d.line([(bx0 - 70, by0 - 30), (bx0 - 25, by0 - 30), (bx0, by0)], fill=ORANGE, width=lw, joint="curve")
    # かご（台形）
    pts = [(bx0, by0), (bx1, by0), (bx1 - 30, by1), (bx0 + 22, by1), (bx0, by0)]
    d.line(pts, fill=ORANGE, width=lw, joint="curve")
    # かごの横線2本
    for yy in (by0 + 37, by0 + 74):
        t = (yy - by0) / (by1 - by0)
        lx = bx0 + 22 * t
        rx = bx1 - 30 * t
        d.line([(lx, yy), (rx, yy)], fill=ORANGE, width=lw - 6)
    # 車輪
    r = 20
    for wx in (bx0 + 55, bx1 - 60):
        d.ellipse([wx - r, by1 + 32 - r, wx + r, by1 + 32 + r], fill=ORANGE)
    return im


def main():
    im = Image.open(SRC).convert("RGB")
    im = redraw_cart(im)
    out = Image.new("RGB", (W, H), GREEN)
    for ci, (sx0, sx1) in enumerate(COLS):
        # 上段
        sy0, sy1 = ROWS[0]
        card = im.crop((sx0, sy0, sx1 + 1, sy1 + 1)).resize((TOP_W[ci], TOP_H), Image.LANCZOS)
        x = ci * CELL_W + (CELL_W - TOP_W[ci]) // 2
        out.paste(card, (x, TOP_Y))
        # 下段
        sy0, sy1 = ROWS[1]
        card = im.crop((sx0, sy0, sx1 + 1, sy1 + 1)).resize((BOT_W[ci], BOT_H), Image.LANCZOS)
        out.paste(card, (x, BOT_Y))
    out.save(OUT)
    # 格子の確認用
    g = out.copy()
    d = ImageDraw.Draw(g)
    for i in (1, 2):
        d.line([(i * CELL_W, 0), (i * CELL_W, H)], fill=(255, 0, 0), width=3)
    d.line([(0, CELL_H), (W, CELL_H)], fill=(255, 0, 0), width=3)
    g.save(OUT_GRID)
    print(OUT, out.size)


if __name__ == "__main__":
    main()
