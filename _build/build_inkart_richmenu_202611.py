#!/usr/bin/env python3
"""インクアート 11月リッチメニュー（11/1〜）の登録用画像を組む。

生成画像 v4（_images/inkart_richmenu_202611_generated_v4.webp・1526×1030）は上段が高さの約52%で、
LINEの6分割テンプレ（上下ちょうど半分）とずれる。6枚の付箋形の枠を切り出し、1200×810の格子に置き直す。
（2026-10-22 殿村さん「生成した画像の方がええな」＝6枠とも生成物を使う。10月の実物と混ぜる版は --mix）

使い方: python3 _build/build_inkart_richmenu_202611.py [--mix]
※2026-10-22 最終＝枠を切り出さず、上段と下段を帯ごと縮める方式（下の rows_mode）。切り出し方式は縁に線が出て「汚い」と言われた。
出力:   _images/inkart_richmenu_202611_1200x810.png（登録用）
        _images/inkart_richmenu_202611_1200x810_gridcheck.png（格子を重ねた確認用）
"""
import sys
from PIL import Image, ImageDraw
import numpy as np

GEN = "_images/inkart_richmenu_202611_generated_v4.webp"
ACTUAL = "_data/printearth/richmenu_202610_actual.jpg"
OUT = "_images/inkart_richmenu_202611_1200x810.png"
OUT_GRID = "_images/inkart_richmenu_202611_1200x810_gridcheck.png"

W, H = 1200, 810
CW, CH = W // 3, H // 2            # 400 × 405

# 置き先の枠の大きさ（列ごとに幅を変えて、縦横の縮みの差を数%に収める）
TOP_W = [360, 350, 360]
TOP_H = 393
BOT_W = [360, 350, 360]
BOT_H = 310
TOP_Y = 6
BOT_Y = 440


def card_boxes(a, green, tol=60):
    """生成画像を上下2段×左右3列に分け、各マスの中の枠（緑でない部分）の箱を返す。"""
    h, w = a.shape[:2]
    m = (abs(a - green).sum(-1) >= tol)
    rowg = (~m).mean(1)
    # 上下の境目＝画像の中ほどで緑の割合がいちばん高い行
    mid = int(np.argmax(rowg[int(h * 0.4):int(h * 0.7)])) + int(h * 0.4)
    boxes = []
    for (y0, y1) in [(0, mid), (mid, h)]:
        row = []
        colg = (~m[y0:y1]).mean(0)
        # 列の境目＝左右の3分の1付近で緑の割合がいちばん高い列
        c1 = int(np.argmax(colg[int(w * 0.25):int(w * 0.42)])) + int(w * 0.25)
        c2 = int(np.argmax(colg[int(w * 0.58):int(w * 0.75)])) + int(w * 0.58)
        for (x0, x1) in [(0, c1), (c1, c2), (c2, w)]:
            sub = m[y0:y1, x0:x1]
            ys, xs = np.where(sub)
            row.append((x0 + xs.min(), y0 + ys.min(), x0 + xs.max(), y0 + ys.max()))
        boxes.append(row)
    return boxes


def recolor_margin(crop, src_green, dst_green):
    """枠の外周と左上の切り落としに残る生成画像の緑を、下地の緑に揃える。"""
    m = (abs(crop - src_green).sum(-1) < 45)
    keep = np.zeros(m.shape, bool)
    keep[:20, :] = True
    keep[-20:, :] = True
    keep[:, :20] = True
    keep[:, -20:] = True
    keep[:120, :120] = True
    m &= keep
    crop[m] = dst_green
    return crop


def main(mix=False):
    gen = Image.open(GEN).convert("RGB")
    ga = np.array(gen).astype(int)
    gen_green = ga[5, 5]
    boxes = card_boxes(ga, gen_green)
    for r in boxes:
        print("boxes", [(int(b[0]), int(b[1]), int(b[2]), int(b[3])) for b in r])

    if mix:
        base = Image.open(ACTUAL).convert("RGB").resize((W, H), Image.LANCZOS)
        bg = tuple(int(v) for v in np.array(base)[H - 5, W - 5])
        out = base.copy()
        ImageDraw.Draw(out).rectangle([0, 0, 2 * CW - 1, CH - 1], fill=bg)
        use = [(0, 0), (0, 1)]
        tw_list, th, ty = [343, 330], 343, 31
    else:
        bg = tuple(int(v) for v in gen_green)
        out = Image.new("RGB", (W, H), bg)
        use = [(r, c) for r in range(2) for c in range(3)]

    for (r, c) in use:
        x0, y0, x1, y1 = boxes[r][c]
        crop = ga[y0:y1 + 1, x0:x1 + 1].copy()
        crop = recolor_margin(crop, gen_green, bg)
        card = Image.fromarray(crop.astype("uint8"))
        sw, sh = card.size
        if mix:
            tw, th_, ty = tw_list[c], th, ty
        else:
            tw = (TOP_W if r == 0 else BOT_W)[c]
            th_ = TOP_H if r == 0 else BOT_H
            ty = TOP_Y if r == 0 else BOT_Y
        print(f"card r{r}c{c}: {sw}x{sh} -> {tw}x{th_}  squash {(tw/sw)/(th_/sh):.3f}")
        card = card.resize((tw, th_), Image.LANCZOS)
        out.paste(card, (c * CW + (CW - tw) // 2, ty))

    out.save(OUT)
    g = out.copy()
    d = ImageDraw.Draw(g)
    for i in (1, 2):
        d.line([(i * CW, 0), (i * CW, H)], fill=(255, 0, 0), width=3)
    d.line([(0, CH), (W, CH)], fill=(255, 0, 0), width=3)
    g.save(OUT_GRID)
    print(OUT, out.size, "bg", bg)


def rows_mode():
    """枠を切り出さず、上段・下段の帯ごと縮める（縁がきれいに残る）。登録用はこちら。"""
    im = Image.open(GEN).convert("RGB")
    w, h = im.size
    green = tuple(int(v) for v in np.array(im)[5, 5])
    sx = W / w
    top = im.crop((0, 14, w, 565)).resize((W, CH), Image.LANCZOS)      # 縦550→405（横0.786に対し0.736＝6%の縮み）
    # 下段は10月の実物の下段をそのまま使う（殿村さん「狭くて窮屈」「これ（実物）をまねて」）
    act = Image.open(ACTUAL).convert("RGB").resize((W, H), Image.LANCZOS)
    act_green = tuple(int(v) for v in np.array(act)[H - 5, W - 5])
    bot = act.crop((0, CH, W, H))
    # 生成物は左の枠が広すぎ（マスからはみ出す）、真ん中の枠が細い。上段左2つを源画像から切り直して同じ幅に置く（殿村さん 10/22）
    tw = 376
    for ci, (sx0, sx1) in enumerate([(19, 533), (541, 1019), (1031, 1510)]):   # 3枠とも同じ幅・同じ高さ・同じ位置に
        card = im.crop((sx0, 14, sx1 + 1, 565)).resize((tw, CH), Image.LANCZOS)
        ImageDraw.Draw(top).rectangle([ci * CW, 0, (ci + 1) * CW - 1, CH - 1], fill=green)
        top.paste(card, (ci * CW + (CW - tw) // 2, 0))
    # 上段の緑（生成）を実物の緑に揃える
    ta = np.array(top).astype(int)
    ta[(abs(ta - np.array(green)).sum(-1) < 45)] = act_green
    top = Image.fromarray(ta.astype("uint8"))
    green = act_green
    out = Image.new("RGB", (W, H), green)
    out.paste(top, (0, 0))
    out.paste(bot, (0, CH))
    # 上段右（商品一覧／ご注文）：枠（オレンジの縁＋白）はそのまま、枠の中に残る緑・暗い縁だけオレンジに。
    # 枠の外は緑のまま（殿村さん 10/22「缶バッジとの間の緑をもう少し広く」）
    a = np.array(out).astype(int)
    cell = a[0:CH, 2 * CW:W]
    om = (abs(cell - np.array([240, 145, 45])).sum(-1) < 90)
    orange = np.median(cell[om], axis=0).astype(int)
    nong = (abs(cell - np.array(green)).sum(-1) >= 60)
    ys, xs = np.where(nong)
    x0, y0, x1, y1 = xs.min(), ys.min(), xs.max(), ys.max()
    inside = np.zeros(cell.shape[:2], bool)
    inside[y0:y1 + 1, x0:x1 + 1] = True
    wm = cell.min(-1) > 235
    wys, wxs = np.where(wm)
    white_box = np.zeros(cell.shape[:2], bool)
    white_box[wys.min():wys.max() + 1, wxs.min():wxs.max() + 1] = True
    fix = inside & ~white_box                      # 枠の縁の帯（白い部分の外側）
    cell[fix] = orange
    a[0:CH, 2 * CW:W] = cell
    out = Image.fromarray(a.astype("uint8"))
    out.save(OUT)
    print(OUT, out.size)


if __name__ == "__main__":
    if "--cut" in sys.argv or "--mix" in sys.argv:
        main(mix="--mix" in sys.argv)
    else:
        rows_mode()
