# -*- coding: utf-8 -*-
"""チェングロウス提案書 P6・P7 用の画像（2026-09-24）。
  _images/chengrowth_zengo_colored.png … 前後検索（カーディーラー）で職業・転職系クエリをオレンジ強調、他は薄く
  _images/chengrowth_trend_seibishi_5y_L.png … Googleトレンド「整備士」5年。文字大きめ・3月=山/12月=谷を帯で表示
"""
from pathlib import Path
from PIL import Image, ImageDraw
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parent.parent
ORANGE = (237, 125, 49)

# ---------- 前後検索：色付け ----------
src = Image.open(ROOT / "_data/zengo/チェングロウス_カーディーラー.webp").convert("RGB")
white = Image.new("RGB", src.size, "white")
base = Image.blend(src, white, 0.62)          # 買う人系は薄く
# 職業・転職系クエリ（点の位置〜ラベル右端）: (x0, y0, x1, y1)
JOB = [
    (1020, 555, 1252, 587),   # カーディーラー 受付嬢
    (1020, 752, 1232, 784),   # カーディーラー 年収
    (440, 906, 652, 938),     # ディーラー 営業マン
    (1020, 664, 1186, 696),   # ディーラーとは
    (1330, 1082, 1472, 1112), # ホワイト企業
    (410, 1020, 512, 1052),   # 施工管理
]
d = ImageDraw.Draw(base)
for x0, y0, x1, y1 in JOB:
    base.paste(src.crop((x0, y0, x1, y1)), (x0, y0))
    d.rounded_rectangle((x0 - 6, y0 - 4, x1 + 6, y1 + 4), radius=8, outline=ORANGE, width=5)
    # 点をオレンジに
    cy = (y0 + y1) // 2
    d.ellipse((x0 + 4, cy - 8, x0 + 20, cy + 8), fill=ORANGE)
base.save(ROOT / "_images/chengrowth_zengo_colored.png")

# ---------- 検索トレンド ----------
fp = font_manager.FontProperties(fname="/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf")
for cand in ["/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"]:
    if Path(cand).exists():
        font_manager.fontManager.addfont(cand)
        plt.rcParams["font.family"] = font_manager.FontProperties(fname=cand).get_name()
        break
df = pd.read_csv(ROOT / "_data/trends/chengrowth_seibishi_5year.csv", parse_dates=["date"])
col = df.columns[1]
fig, ax = plt.subplots(figsize=(16, 5.2), dpi=150)
for y in range(df.date.dt.year.min(), df.date.dt.year.max() + 1):
    ax.axvspan(pd.Timestamp(y, 3, 1), pd.Timestamp(y, 4, 1), color="#D9661F", alpha=0.16, lw=0)
    ax.axvspan(pd.Timestamp(y, 12, 1), pd.Timestamp(y + 1, 1, 1), color="#7F7F7F", alpha=0.14, lw=0)
ax.plot(df.date, df[col], color="#3467B2", lw=2.4)
ax.set_xlim(df.date.min(), df.date.max())
ax.set_ylim(0, 105)
ax.set_ylabel("検索インタレスト（相対値）", fontsize=15, color="#333333")
ax.tick_params(labelsize=15, colors="#333333")
for s in ("top", "right"): ax.spines[s].set_visible(False)
ax.grid(axis="y", color="#E5E5E5", lw=1)
ax.text(0.005, 1.03, "Googleトレンド「整備士」（日本・過去5年・週次）", transform=ax.transAxes, fontsize=16, color="#333333")
ax.text(0.62, 1.03, "■ 3月（山）", transform=ax.transAxes, fontsize=16, color="#D9661F", fontweight="bold")
ax.text(0.76, 1.03, "■ 12月（谷）", transform=ax.transAxes, fontsize=16, color="#7F7F7F", fontweight="bold")
fig.tight_layout()
fig.savefig(ROOT / "_images/chengrowth_trend_seibishi_5y_L.png")
print("ok", df.shape)
