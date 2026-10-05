# -*- coding: utf-8 -*-
"""前後検索（カーディーラー）を4分類で色分け（2026-09-24）。
分類データ：_data/zengo/チェングロウス_カーディーラー_クエリ分類.csv（画像のラベル位置を目視で転記）
出力：_images/chengrowth_zengo_4color.png
"""
import csv
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
COL = {"J": (217, 102, 31), "M": (52, 103, 178), "S": (142, 78, 198), "C": (0, 137, 123)}
src = Image.open(ROOT / "_data/zengo/チェングロウス_カーディーラー.webp").convert("RGB")
base = Image.blend(src, Image.new("RGB", src.size, "white"), 0.55)   # 分類外の点・ラベルは控えめに
over = Image.new("RGBA", src.size, (0, 0, 0, 0))
od = ImageDraw.Draw(over)
font = ImageFont.truetype("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", 21)
rows = list(csv.DictReader(open(ROOT / "_data/zengo/チェングロウス_カーディーラー_クエリ分類.csv", encoding="utf-8")))
boxes = []
for r in rows:
    x, y, c = int(r["x"]), int(r["y"]), COL[r["cat"]]
    w = int(font.getlength(r["label"])) + 26
    if x + w > 1999:   # 右端はラベルが点の左に出る
        box = (max(x - w + 12, 0), y - 15, x + 10, y + 15)
    else:
        box = (x - 9, y - 15, x + w, y + 15)
    boxes.append((box, c, (x, y)))
    a = 150 if r["cat"] == "J" else 70
    od.rounded_rectangle(box, radius=6, fill=c + (a,))
img = Image.alpha_composite(base.convert("RGBA"), over)
d = ImageDraw.Draw(img)
for box, c, (x, y) in boxes:
    # ラベル文字は原画から濃く戻す（背景色の上に黒文字が乗るように）
    crop = src.crop(box).convert("L").point(lambda v: 0 if v < 150 else 255).convert("L")
    mask = crop.point(lambda v: 255 if v == 0 else 0)
    img.paste((40, 40, 40, 255), box, mask)
    d.ellipse((x - 9, y - 9, x + 9, y + 9), fill=c + (255,), outline=(255, 255, 255, 255), width=2)
img.convert("RGB").save(ROOT / "_images/chengrowth_zengo_4color.png")
print("ok", len(rows))
