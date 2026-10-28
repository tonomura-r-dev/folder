# -*- coding: utf-8 -*-
"""ワクトリ Googleトレンド「リゾートバイト」「スキー場バイト」の月ごとの平均（5年分）を、1月〜12月の折れ線にする（2026-10-28）。
提案FMT（MAX版）の「施策の考え方（長期：企画配信）」の画像の枠（7.4×2.7インチ）に入れる。
元データ：_data/trends/wakutori_{リゾートバイト,スキー場バイト}_5year_weekly.csv（キーワードごとに取得＝それぞれ最も多い週を100）
色は 3467B2／00897B（CLAUDE.md 2026-08-27）。冬（10〜2月）と夏（6〜7月）に薄い帯。出力：_images/wakutori_trend_month.png
"""
import csv
from collections import defaultdict
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "_images/wakutori_trend_month.png"
SERIES = [("リゾートバイト", "3467B2"), ("スキー場バイト", "00897B")]
FONT = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"


def monthly(kw):
    acc = defaultdict(list)
    for r in csv.DictReader(open(ROOT / f"_data/trends/wakutori_{kw}_5year_weekly.csv", encoding="utf-8")):
        acc[int(r["週の初日"][5:7])].append(float(r[kw]))
    return [sum(acc[m]) / len(acc[m]) for m in range(1, 13)]


def main():
    font_manager.fontManager.addfont(FONT)
    plt.rcParams["font.family"] = font_manager.FontProperties(fname=FONT).get_name()
    fig, ax = plt.subplots(figsize=(7.4, 2.7), dpi=250)
    x = list(range(1, 13))
    for a, b in ((0.5, 2.5), (9.5, 12.5), (5.5, 7.5)):          # 冬（10〜2月）・夏（6〜7月）
        ax.axvspan(a, b, color="#E3E8F3", zorder=0)
    for kw, c in SERIES:
        ys = monthly(kw)
        ax.plot(x, ys, color="#" + c, lw=2.4, marker="o", ms=4, label=kw, zorder=3)
    ax.set_xlim(0.5, 12.5)
    ax.set_ylim(0, 100)
    ax.set_xticks(x, [f"{m}月" for m in x], fontsize=9, color="#333333")
    ax.set_yticks([0, 50, 100], ["0", "50", "100"], fontsize=8, color="#777777")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#BBBBBB")
    ax.grid(axis="y", color="#E5E5E5", lw=0.6)
    ax.legend(loc="upper left", bbox_to_anchor=(0.0, 1.17), ncol=2, frameon=False, fontsize=9.5)
    ax.text(0.5, -26, "※Googleトレンド（全国・2021年10月〜2026年10月）の月ごとの平均。最も多い週＝100",
            ha="left", fontsize=6.5, color="#777777")
    fig.subplots_adjust(left=0.06, right=0.99, top=0.86, bottom=0.2)
    fig.savefig(OUT)
    print("ok", OUT)


if __name__ == "__main__":
    main()
