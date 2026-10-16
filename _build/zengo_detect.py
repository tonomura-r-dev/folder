# -*- coding: utf-8 -*-
"""LINEヤフー前後検索の散布図（画像）から、点（緑）とラベルの位置を拾い、OCR（tesseract・日本語）で文字を読む。

    python3 _build/zengo_detect.py <画像> <出力.csv> <確認用の画像.png>

  出力のCSV（label,x,y,x0,x1,side）を、色分けのスクリプト（make_juku_zengo.py など）が読む。
  OCRは誤読がある（学園→学国 など）。色分けの側で読み替えてから分類する。
  クラウド環境の初回は `apt-get install -y tesseract-ocr tesseract-ocr-jpn` が要る。
"""
import csv
import subprocess
import sys
import numpy as np
from PIL import Image

src, out_csv, out_dbg = sys.argv[1:4]
im = Image.open(src).convert("RGB")
a = np.asarray(im).astype(int)
H, W, _ = a.shape
r, g, b = a[..., 0], a[..., 1], a[..., 2]
teal = (g > r + 60) & (b > r + 40) & (g > 140)
dark = (r < 130) & (g < 130) & (b < 150) & ~teal

# 点：緑の画素の連結成分（4近傍・素朴な探索）
seen = np.zeros_like(teal)
dots = []
ys, xs = np.nonzero(teal)
for y0, x0 in zip(ys, xs):
    if seen[y0, x0]:
        continue
    stack = [(y0, x0)]
    seen[y0, x0] = True
    pts = []
    while stack:
        y, x = stack.pop()
        pts.append((y, x))
        for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            ny, nx = y + dy, x + dx
            if 0 <= ny < H and 0 <= nx < W and teal[ny, nx] and not seen[ny, nx]:
                seen[ny, nx] = True
                stack.append((ny, nx))
    if len(pts) >= 40:
        py, px = zip(*pts)
        if max(py) - min(py) > 22:
            # 縦に重なった点はひとかたまりに見えるので、高さをずらしながら候補を置く
            for yy in range(min(py) + 7, max(py) - 5, 4):
                dots.append((int(np.mean(px)), yy, min(px), max(px)))
        else:
            dots.append((int(np.mean(px)), int(np.mean(py)), min(px), max(px)))


def span(x_from, step, cy, me):
    """x_from から step 方向に、文字（暗い画素）の続く範囲を返す（同じ高さの別の点の手前で止める）"""
    band = dark[max(cy - 13, 0):cy + 14]
    cols = band.sum(axis=0)
    stops = sorted(d[0] for d in dots if d is not me and d[0] != me[0] and abs(d[1] - cy) <= 9 and (d[0] - x_from) * step > 0)
    stop = (stops[0] - 10) if step > 0 and stops else ((stops[-1] + 10) if step < 0 and stops else None)
    x, first, last, gap = x_from, None, None, 0
    while 0 <= x < W:
        if stop is not None and ((step > 0 and x >= stop) or (step < 0 and x <= stop)):
            break
        if cols[x] > 0:
            if first is None:
                if abs(x - x_from) > 14:
                    return None
                first = x
            last, gap = x, 0
        elif first is not None:
            gap += 1
            if gap > 22 or abs(x - first) > 330:      # 長すぎる＝隣の文字までつながった
                break
        elif abs(x - x_from) > 14:
            return None
        x += step
    if first is None or abs(last - first) < 12:
        return None
    x0, x1 = min(first, last), max(first, last)
    # 文字の上下の中心が点の高さに合っていること（よその文字を拾わない）
    blk = dark[max(cy - 13, 0):cy + 14, x0:x1 + 1]
    yy = np.nonzero(blk.sum(axis=1))[0]
    if len(yy) == 0 or abs((yy.min() + yy.max()) / 2 - 13) > 5:
        return None
    return (x0, x1)


rows = []
for cx, cy, xl, xr in dots:
    me = next(d for d in dots if d[0] == cx and d[1] == cy)
    s = None
    for ty in (cy, cy + 12, cy - 12):          # 端の近くでは、文字が点の少し下（上）に出る
        s = span(xr + 3, 1, ty, me)
        side = "R"
        if s is None:
            s = span(xl - 3, -1, ty, me)
            side = "L"
        if s is not None:
            break
    if s is None:
        continue
    py, cy = cy, ty
    x0, x1 = s
    crop = im.crop((x0 - 2, cy - 15, x1 + 3, cy + 16)).convert("L")
    crop = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS).point(lambda v: 0 if v < 160 else 255)
    crop = Image.eval(crop, lambda v: v)
    pad = Image.new("L", (crop.width + 40, crop.height + 40), 255)
    pad.paste(crop, (20, 20))
    crop = pad
    crop.save("/tmp/ocr.png")
    txt = subprocess.run(["tesseract", "/tmp/ocr.png", "-", "-l", "jpn", "--psm", "7"],
                         capture_output=True, text=True).stdout.strip().replace(" ", "")
    rows.append(dict(label=txt, x=cx, y=cy, py=py, x0=x0, x1=x1, side=side))  # y＝文字の高さ／py＝点の高さ

with open(out_csv, "w", encoding="utf-8", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["label", "x", "y", "py", "x0", "x1", "side"])
    w.writeheader()
    w.writerows(rows)

from PIL import ImageDraw
d = ImageDraw.Draw(im)
for i, rw in enumerate(rows):
    d.rectangle((rw["x0"], rw["y"] - 14, rw["x1"], rw["y"] + 14), outline=(255, 0, 0), width=2)
    d.text((rw["x0"], rw["y"] - 26), str(i), fill=(255, 0, 0))
im.save(out_dbg)
print(len(dots), "dots", len(rows), "labels")
