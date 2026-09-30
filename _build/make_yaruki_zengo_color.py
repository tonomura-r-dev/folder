# -*- coding: utf-8 -*-
"""やる気スイッチ S8・S9 の前後検索（子ども 習い事／習い事）を分類ごとに色分け（2026-09-30）。
チェングロウスの工場求人ナビ版（make_chengrowth_zengo_kojo3.py）と同じ作り。
分類：A=習い事の比較・検討／B=子育て・学び／C=家族のお出かけ・イベント／D=お金・制度。
E（大人の習い事・趣味）とX（その他）は色を付けない（薄いグレーのまま）。
分類データ：_data/zengo/やる気スイッチ_クエリ分類_{子ども習い事,習い事}.csv（dot_id, x, y, label, cat）
  x, y は点の中心。ラベルは点の右に同じ高さで書かれているので、右へ文字の塊を拾って囲む。
出力：_images/yaruki_zengo_{kodomo,naraigoto}_color.png
"""
import csv
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
COL = {"A": (52, 103, 178), "B": (142, 78, 198), "C": (0, 137, 123), "D": (217, 102, 31)}
CHARTS = {"kodomo": "子ども習い事", "naraigoto": "習い事"}


def label_box(a, dots_xy, x, y, text):
    """点(x,y)の右にあるラベル文字の範囲を、濃い文字の画素を右へたどって決める"""
    h, w = a.shape[:2]
    dark = a[max(0, y - 9):y + 10].sum(axis=2) < 360          # 文字（濃いグレー）
    col = dark.any(axis=0)
    others = [dx for dx, dy in dots_xy if abs(dy - y) <= 6 and dx > x + 8]
    stop = min(others) - 8 if others else w - 1               # 同じ行の次の点の手前で止める
    x0 = x + 8
    while x0 < stop and not col[x0]:
        x0 += 1
    x1, gap = x0, 0
    for xx in range(x0, stop):
        if col[xx]:
            x1, gap = xx, 0
        else:
            gap += 1
            if gap > 13:                                      # 全角スペース（約16px）より短い切れ目は同じラベル
                break
    exp = x0 + int(17 * sum(1 if ord(ch) > 0x2000 else 0.55 for ch in text))
    if x1 < x0 + 10 or x1 > exp + 40:                           # 拾えない・拾いすぎのときは文字数から
        x1 = min(exp, stop)
    return (x0 - 5, y - 14, min(x1 + 6, w - 1), y + 14)


def render(key):
    name = CHARTS[key]
    src = Image.open(ROOT / f"_data/zengo/やる気スイッチ_前後検索_{name}.png").convert("RGB")
    a = np.array(src).astype(int)
    rows = list(csv.DictReader(open(ROOT / f"_data/zengo/やる気スイッチ_クエリ分類_{name}.csv", encoding="utf-8")))
    dots_xy = [(int(r["x"]), int(r["y"])) for r in rows]
    base = Image.blend(src, Image.new("RGB", src.size, "white"), 0.55)   # 分類外の点・ラベルは控えめに
    over = Image.new("RGBA", src.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(over)
    items = []
    for r in rows:
        if r["cat"] not in COL:
            continue
        x, y, c = int(r["x"]), int(r["y"]), COL[r["cat"]]
        box = label_box(a, dots_xy, x, y, r["label"])
        od.rounded_rectangle(box, radius=6, fill=c + (80,))
        items.append((box, c, (x, y)))
    img = Image.alpha_composite(base.convert("RGBA"), over)
    d = ImageDraw.Draw(img)
    for box, c, (x, y) in items:
        crop = src.crop(box).convert("L")                      # ラベル文字は原画から濃く戻す
        mask = crop.point(lambda v: 255 if v < 150 else 0)
        img.paste((40, 40, 40, 255), box, mask)
        d.ellipse((x - 9, y - 9, x + 9, y + 9), fill=c + (255,), outline=(255, 255, 255, 255), width=2)
    out = ROOT / f"_images/yaruki_zengo_{key}_color.png"
    img.convert("RGB").save(out)
    print("ok", out.name, len(items), "/", len(rows))


for k in (sys.argv[1:] or CHARTS):
    render(k)
