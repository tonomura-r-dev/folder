#!/usr/bin/env python3
"""インクアート 11月リッチメニュー（11/1〜）の登録用画像を組む。

材料
- 生成画像 v4（_images/inkart_richmenu_202611_generated_v4.webp・1526×1030）：上段左（年賀状）と上段中（缶バッジ）の2枠だけ使う
- 10月の実物（_data/printearth/richmenu_202610_actual.jpg・2000×1349・DYMデザイナー版）：上段右（商品一覧／ご注文）と下段3枠をそのまま使う

生成画像は上段が高く、LINEの6分割テンプレ（上下ちょうど半分）とずれるので、枠を切り出して格子に置き直す。
出力は 1200×810（LINEの許容サイズ）。

使い方: python3 _build/build_inkart_richmenu_202611.py
出力:   _images/inkart_richmenu_202611_1200x810.png（登録用）
        _images/inkart_richmenu_202611_1200x810_gridcheck.png（格子を重ねた確認用）
"""
from PIL import Image, ImageDraw
import numpy as np

GEN = "_images/inkart_richmenu_202611_generated_v4.webp"
ACTUAL = "_data/printearth/richmenu_202610_actual.jpg"
OUT = "_images/inkart_richmenu_202611_1200x810.png"
OUT_GRID = "_images/inkart_richmenu_202611_1200x810_gridcheck.png"

W, H = 1200, 810
CW, CH = W // 3, H // 2            # 400 × 405

# 生成画像v4の枠の位置（緑の溝から実測）
GEN_COLS = [(19, 529), (531, 1018)]
GEN_ROW = (20, 559)
GEN_GREEN = np.array([2, 145, 95])


def bbox_nongreen(a, green, tol=60):
    m = (abs(a - green).sum(-1) >= tol)
    ys, xs = np.where(m)
    return xs.min(), ys.min(), xs.max(), ys.max()


def main():
    act = Image.open(ACTUAL).convert("RGB").resize((W, H), Image.LANCZOS)
    aa = np.array(act).astype(int)
    act_green = tuple(int(v) for v in aa[H - 5, W - 5])
    out = act.copy()

    # 実物の上段右の枠の大きさ・位置を測り、生成の2枠をそれに合わせる
    cell = aa[0:CH, 2 * CW:3 * CW]
    wm = cell.min(-1) > 235                      # 白い付箋形の枠
    ys, xs = np.where(wm)
    x0, y0, x1, y1 = xs.min(), ys.min(), xs.max(), ys.max()
    card_w, card_h = int(x1 - x0 + 1), int(y1 - y0 + 1)   # 実測 343×343
    top_margin = int(y0)
    print("actual white card", card_w, card_h, "left margin", x0, "top margin", top_margin)
    # 生成の中央の枠は元が縦長なので、横を少し狭くして縦横の縮みの差を6%台に抑える
    widths = [card_w, card_w - 13]

    # 上段左の2マス（生成の枠を置く場所）を実物の緑で塗りつぶしてから置く
    d0 = ImageDraw.Draw(out)
    d0.rectangle([0, 0, 2 * CW - 1, CH - 1], fill=act_green)

    gen = Image.open(GEN).convert("RGB")
    ga = np.array(gen).astype(int)
    for ci, (sx0, sx1) in enumerate(GEN_COLS):
        sy0, sy1 = GEN_ROW
        crop = ga[sy0:sy1 + 1, sx0:sx1 + 1].copy()
        # 枠の外側（生成画像の緑）を実物の緑に塗り替える（緑の色味が少し違うため）
        m = (abs(crop - GEN_GREEN).sum(-1) < 45)
        # 枠の中の緑の文字を巻き込まないよう、外周20pxと左上の切り落とし（120px四方）だけ対象にする
        keep = np.zeros(m.shape, bool)
        keep[:20, :] = True
        keep[-20:, :] = True
        keep[:, :20] = True
        keep[:, -20:] = True
        keep[:120, :120] = True
        m &= keep
        crop[m] = act_green
        card = Image.fromarray(crop.astype("uint8"))
        sw, sh = card.size
        tw, th = widths[ci], card_h
        print(f"gen card {ci}: {sw}x{sh} -> {tw}x{th}  squash {(tw/sw)/(th/sh):.3f}")
        card = card.resize((tw, th), Image.LANCZOS)
        x = ci * CW + (CW - tw) // 2
        out.paste(card, (x, top_margin))
    out.save(OUT)
    g = out.copy()
    d = ImageDraw.Draw(g)
    for i in (1, 2):
        d.line([(i * CW, 0), (i * CW, H)], fill=(255, 0, 0), width=3)
    d.line([(0, CH), (W, CH)], fill=(255, 0, 0), width=3)
    g.save(OUT_GRID)
    print(OUT, out.size, "green", act_green)


if __name__ == "__main__":
    main()
