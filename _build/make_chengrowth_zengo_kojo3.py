# -*- coding: utf-8 -*-
"""前後検索（工場求人ナビ）を3分類で色分け（2026-09-30）。カーディーラー版（make_chengrowth_zengo4.py）と同じ作り。
分類：C=条件（年齢・夜勤・寮・収入・雇用形態など）／R=評判・不安／A=応募の準備（志望動機・〜とは・電話番号）。
サイト名・求人検索は色を付けない（薄いグレーのまま）。
分類データ：_data/zengo/工場求人ナビ_クエリ分類.csv（画像のラベル位置を目視で転記。lx＝ラベル左端, y＝行の中心）
出力：_images/chengrowth_zengo_kojo_3color.png
"""
import csv
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from scipy import ndimage

ROOT = Path(__file__).resolve().parent.parent
COL = {"C": (52, 103, 178), "R": (217, 102, 31), "A": (142, 78, 198)}
src = Image.open(ROOT / "_data/zengo/工場求人ナビ_前後検索_20260928.png").convert("RGB")

# 点（ティール色の丸）の位置を拾う
a = np.array(src).astype(int)
m = (a[..., 1] > 150) & (a[..., 2] > 130) & (a[..., 0] < 120) & (a[..., 1] - a[..., 0] > 60)
lab, _ = ndimage.label(m)
dots = [((s[1].start + s[1].stop) // 2, (s[0].start + s[0].stop) // 2) for s in ndimage.find_objects(lab)
        if 8 <= s[0].stop - s[0].start <= 14 and 8 <= s[1].stop - s[1].start <= 14]

base = Image.blend(src, Image.new("RGB", src.size, "white"), 0.55)   # 分類外の点・ラベルは控えめに
over = Image.new("RGBA", src.size, (0, 0, 0, 0))
od = ImageDraw.Draw(over)
font = ImageFont.truetype("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", 20)
rows = list(csv.DictReader(open(ROOT / "_data/zengo/工場求人ナビ_クエリ分類.csv", encoding="utf-8")))
items = []
for r in rows:
    lx, y, c = int(r["lx"]), int(r["y"]), COL[r["cat"]]
    w = int(font.getlength(r["label"]))
    box = (lx - 6, y - 15, min(lx + w + 8, 1999), y + 15)
    near = min(dots, key=lambda d: (d[0] - (lx - 12)) ** 2 + (d[1] - y) ** 2)
    dot = near if abs(near[0] - (lx - 12)) <= 30 and abs(near[1] - y) <= 12 and near[0] < lx - 4 else None   # ラベルに重なる点は塗らない
    items.append((box, c, dot, r["label"]))
    od.rounded_rectangle(box, radius=6, fill=c + (80,))
img = Image.alpha_composite(base.convert("RGBA"), over)
d = ImageDraw.Draw(img)
for box, c, dot, label in items:
    # ラベル文字は原画から濃く戻す
    crop = src.crop(box).convert("L")
    mask = crop.point(lambda v: 255 if v < 150 else 0)
    img.paste((40, 40, 40, 255), box, mask)
    if dot:
        x, y = dot
        d.ellipse((x - 9, y - 9, x + 9, y + 9), fill=c + (255,), outline=(255, 255, 255, 255), width=2)
    else:
        print("点が見つからない:", label)
img.convert("RGB").save(ROOT / "_images/chengrowth_zengo_kojo_3color.png")
print("ok", len(rows))
