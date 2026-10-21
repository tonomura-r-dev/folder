# -*- coding: utf-8 -*-
"""横浜シミズ 前後検索「単発バイト」を3分類で色分け（2026-10-21）。タナカホーム版（make_tanaka_zengo_color.py）と同じ作り。
分類は応募の前後に調べる順番。橙・赤・黄は使わない（CLAUDE.md 2026-10-19・10-20）。
  A=求人サービスを比べる（青）／B=応募・面接の準備（紫）／C=給与・働き方を調べる（ティール）
分類に入らないラベルは色なし（薄いまま）。
分類データ：_data/zengo/横浜シミズ_単発バイト_クエリ分類.csv（x, y＝点の中心の目安。近くの点に吸着させる。
  mode=edge は右端で左に寄せられたラベル（x＝文字の左端・点は描かない）、mode=fix は文字に隠れた点（吸着しない））
出力：_images/yokohama_shimizu_zengo_tanpatsu_color.png（凡例なし＝資料用）
      引数 --legend <出力> で、上に凡例を付けた確認用の画像も出す
"""
import csv
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from scipy import ndimage

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_data/zengo/横浜シミズ_前後検索_単発バイト.webp"
CSV = ROOT / "_data/zengo/横浜シミズ_単発バイト_クエリ分類.csv"
OUT = ROOT / "_images/yokohama_shimizu_zengo_tanpatsu_color.png"
COL = {"A": (52, 103, 178), "B": (142, 78, 198), "C": (0, 137, 123)}
LEGEND = [("A", "求人サービスを比べる"), ("B", "応募・面接の準備"),
          ("C", "給与・働き方を調べる")]
CHW = 20.0          # 全角1文字の幅（px）
RIGHT = 1984        # グラフの右端（ラベルはここで止まるよう左に寄せて描かれる）
FONT = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"


def dot_centers(a):
    teal = (a[..., 1] > 150) & (a[..., 0] < 120) & (a[..., 1] - a[..., 0] > 80)
    return teal, ndimage.distance_transform_edt(teal)


def snap(dist, x, y, r=18):
    """目安の(x,y)に一番近い点の中心"""
    y0, x0 = max(0, y - r), max(0, x - r)
    win = dist[y0:y + r + 1, x0:x + r + 1]
    if win.max() < 2.5:
        return x, y
    yy, xx = np.nonzero(win >= 2.5)                     # 文字に一部隠れた点も拾う
    i = np.argmin((xx + x0 - x) ** 2 + (yy + y0 - y) ** 2)
    return int(xx[i] + x0), int(yy[i] + y0)


def label_box(a, x, y, text):
    est = int(CHW * sum(1 if ord(ch) > 0x2000 else 0.55 for ch in text))
    if x + 10 + est > RIGHT + 4:                       # 右端で左に寄せられたラベル
        return (RIGHT - est - 8, y - 15, RIGHT + 4, y + 15)
    dark = a[max(0, y - 10):y + 11].sum(axis=2) < 420
    col = dark.any(axis=0)
    x0 = x + 7
    while x0 < x + 22 and not col[x0]:
        x0 += 1
    x1, gap = x0, 0
    for xx in range(x0, min(x0 + est + 60, a.shape[1] - 1)):
        if col[xx]:
            x1, gap = xx, 0
        else:
            gap += 1
            if gap > 14:
                break
    if x1 < x0 + 10 or x1 > x0 + est + 30:
        x1 = x0 + est
    return (x - 10, y - 15, min(x1 + 6, a.shape[1] - 1), y + 15)


def render():
    src = Image.open(SRC).convert("RGB")
    a = np.array(src).astype(int)
    _, dist = dot_centers(a)
    rows = list(csv.DictReader(open(CSV, encoding="utf-8")))
    base = Image.blend(src, Image.new("RGB", src.size, "white"), 0.55)   # 分類外の点・ラベルは控えめに
    over = Image.new("RGBA", src.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(over)
    items = []
    for r in rows:
        c, mode = COL[r["cat"]], r.get("mode") or ""
        x, y = int(r["x"]), int(r["y"])
        if mode == "edge":                              # 右端で左に寄せられたラベル：x＝文字の左端。点は描かない
            box, xy = (x - 8, y - 15, RIGHT + 4, y + 15), None
        else:
            if mode != "fix":                           # fix＝文字に隠れて吸着できない点。目安をそのまま使う
                x, y = snap(dist, x, y)
            ly = y + 12 if y < 40 else y                # グラフ上端の点（タイミー）はラベルが少し下
            box, xy = label_box(a, x, ly, r["label"]), (x, y)
        od.rounded_rectangle(box, radius=7, fill=c + (85,))
        items.append((box, c, xy))
    img = Image.alpha_composite(base.convert("RGBA"), over)
    d = ImageDraw.Draw(img)
    for box, c, _ in items:
        mask = src.crop(box).convert("L").point(lambda v: 255 if v < 150 else 0)   # 文字は原画から濃く戻す
        img.paste((40, 40, 40, 255), box, mask)
    for box, c, xy in items:
        if xy is None:
            continue
        x, y = xy
        d.ellipse((x - 9, y - 9, x + 9, y + 9), fill=c + (255,), outline=(255, 255, 255, 255), width=2)
    img = img.convert("RGB")
    img.save(OUT)
    print("ok", OUT.name, len(items))
    return img


def with_legend(img, out):
    f = ImageFont.truetype(FONT, 30)
    band = 90
    canvas = Image.new("RGB", (img.width, img.height + band), "white")
    canvas.paste(img, (0, band))
    d = ImageDraw.Draw(canvas)
    x = 40
    for k, t in LEGEND:
        d.ellipse((x, 31, x + 28, 59), fill=COL[k])
        d.text((x + 38, 29), t, font=f, fill=(51, 51, 51))
        x += 38 + int(f.getlength(t)) + 50
    canvas.save(out)
    print("ok", out)


if __name__ == "__main__":
    im = render()
    if "--legend" in sys.argv:
        with_legend(im, sys.argv[sys.argv.index("--legend") + 1])
