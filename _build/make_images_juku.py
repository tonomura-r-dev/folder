# -*- coding: utf-8 -*-
"""教育（塾）業界の提案資料に貼る図版を生成する。

  python _build/make_images_juku.py

出力：_images/juku_*.png
入力：_data/trends/教育塾/*.csv（Googleトレンド実測・2026年年初来）

【ルール（CLAUDE.md より）】
- Googleトレンドは相対指標。スケールの違うKWを1枚に混ぜない（小さい方が0に潰れる）
  → 「塾」は単独グラフ。他3語（個別指導/予備校/家庭教師）は別グラフ
- 系列色に DYMのネイビー 1F285A・炭黒 333333 は使わない（配色バリデータFAIL）
  → 3467B2 / D9661F / 00897B / 8E4EC6 を使う。文字・軸・注釈はインク色のまま
"""
import csv
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import matplotlib.dates as mdates

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "_data" / "trends" / "教育塾"
OUTDIR = ROOT / "_images"
OUTDIR.mkdir(exist_ok=True)

# ---- フォント ----
avail = {f.name for f in font_manager.fontManager.ttflist}
JP = "IPAPGothic" if "IPAPGothic" in avail else ("IPAGothic" if "IPAGothic" in avail else "DejaVu Sans")
plt.rcParams.update({
    "font.family": JP,
    "axes.unicode_minus": False,
    "axes.edgecolor": "#B0B0B0",
    "axes.labelcolor": "#333333",
    "text.color": "#333333",
    "xtick.color": "#555555",
    "ytick.color": "#555555",
    "figure.dpi": 200,
})

SERIES = ["#3467B2", "#D9661F", "#00897B", "#8E4EC6"]
INK = "#333333"
NAVY = "#1F285A"
MUT = "#7F7F7F"


def load(fname):
    rows = list(csv.reader(open(DATA / fname, encoding="utf-8-sig")))
    hdr = rows[0][1:]
    dates, cols = [], [[] for _ in hdr]
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        dates.append(datetime.strptime(r[0], "%Y-%m-%d"))
        for i, v in enumerate(r[1:]):
            cols[i].append(int(v))
    return dates, hdr, cols


def style(ax, dates):
    ax.set_ylim(0, 108)
    ax.set_ylabel("検索インタレスト（相対値・最大100）", fontsize=9)
    ax.grid(axis="y", color="#E8E8E8", lw=0.8)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m月"))
    ax.set_xlim(dates[0], dates[-1])
    ax.tick_params(labelsize=9)


def mark_peak(ax, dates, vals, color, label):
    i = vals.index(max(vals))
    ax.plot(dates[i], vals[i], "o", color=color, ms=8, zorder=5)
    ax.annotate(f"{label}\n{dates[i]:%-m/%-d}　{vals[i]}",
                xy=(dates[i], vals[i]), xytext=(0, 14), textcoords="offset points",
                ha="center", fontsize=10, color=color, fontweight="bold", zorder=6)
    return dates[i]


def mark_lead(ax, peak_date, days, text):
    """ピークのN日前に「仕込み開始」の縦破線を引く"""
    from datetime import timedelta
    d = peak_date - timedelta(days=days)
    ax.axvline(d, color="#D9661F", lw=1.6, ls=(0, (5, 3)), zorder=4)
    ax.annotate(text, xy=(d, 96), xytext=(-8, 0), textcoords="offset points",
                ha="right", va="top", fontsize=9.5, color="#D9661F", fontweight="bold")


# ============================================================
# 図1 「塾」単独（他3語とスケールが8倍違うので必ず分ける）
# ============================================================
dates, hdr, cols = load("図1_塾_個別指導_予備校_家庭教師.csv")
i_juku = hdr.index("塾")
fig, ax = plt.subplots(figsize=(9.2, 3.5))
ax.plot(dates, cols[i_juku], color=SERIES[0], lw=2.4, label="塾")
ax.fill_between(dates, cols[i_juku], color=SERIES[0], alpha=0.10)
style(ax, dates)
peak = mark_peak(ax, dates, cols[i_juku], SERIES[0], "ピーク")
mark_lead(ax, peak, 26, "仕込み開始\n（26日前）")
ax.set_title("「塾」の検索推移｜ピークは2月上旬。検討はその26日前に始まっている",
             fontsize=12, fontweight="bold", color=NAVY, pad=26, loc="left")
fig.tight_layout()
fig.savefig(OUTDIR / "juku_trend_juku.png", bbox_inches="tight", facecolor="white")
plt.close(fig)

# ============================================================
# 図1b 他3語（「塾」とはスケールが違うので別グラフ）
# ============================================================
fig, ax = plt.subplots(figsize=(9.2, 3.2))
for k, name in enumerate(hdr):
    if name == "塾":
        continue
    ax.plot(dates, cols[hdr.index(name)], color=SERIES[(k % 3) + 1], lw=2.0, label=name)
ax.set_ylim(0, 22)
ax.set_ylabel("検索インタレスト（相対値）", fontsize=9)
ax.grid(axis="y", color="#E8E8E8", lw=0.8)
ax.set_axisbelow(True)
for spine in ("top", "right"):
    ax.spines[spine].set_visible(False)
ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m月"))
ax.set_xlim(dates[0], dates[-1])
ax.tick_params(labelsize=9)
ax.legend(fontsize=9.5, frameon=False, ncol=3, loc="upper right")
ax.set_title("学習形態別｜「塾」とは桁が違う（同じグラフに混ぜると潰れる）",
             fontsize=11.5, fontweight="bold", color=NAVY, pad=10, loc="left")
fig.tight_layout()
fig.savefig(OUTDIR / "juku_trend_keitai.png", bbox_inches="tight", facecolor="white")
plt.close(fig)

# ============================================================
# 図2 季節講習（募集タイミングの本命）
# ============================================================
dates2, hdr2, cols2 = load("図2_夏期講習_冬期講習_春期講習.csv")
fig, ax = plt.subplots(figsize=(9.2, 3.5))
for k, name in enumerate(hdr2):
    ax.plot(dates2, cols2[k], color=SERIES[k], lw=2.2, label=name)
style(ax, dates2)
i_natsu = hdr2.index("夏期講習")
peak2 = mark_peak(ax, dates2, cols2[i_natsu], SERIES[i_natsu], "夏期講習ピーク")
mark_lead(ax, peak2, 26, "仕込み開始\n（26日前）")
ax.legend(fontsize=9.5, frameon=False, ncol=3, loc="upper left")
ax.set_title("季節講習の検索推移｜夏が圧倒的（冬期・春期の3〜5倍）",
             fontsize=12, fontweight="bold", color=NAVY, pad=26, loc="left")
fig.tight_layout()
fig.savefig(OUTDIR / "juku_trend_koushuu.png", bbox_inches="tight", facecolor="white")
plt.close(fig)

# ============================================================
# 図3 受験3種
# ============================================================
dates3, hdr3, cols3 = load("図3_中学受験_大学受験_高校受験.csv")
fig, ax = plt.subplots(figsize=(9.2, 3.2))
for k, name in enumerate(hdr3):
    ax.plot(dates3, cols3[k], color=SERIES[k], lw=2.2, label=name)
style(ax, dates3)
ax.axvspan(datetime(2026, 1, 11), datetime(2026, 2, 8), color="#D9661F", alpha=0.10, zorder=0)
ax.annotate("1月中旬〜2月上旬に集中", xy=(datetime(2026, 1, 25), 103),
            ha="center", fontsize=10, color="#D9661F", fontweight="bold")
ax.legend(fontsize=9.5, frameon=False, ncol=3, loc="upper right")
ax.set_title("受験区分別の検索推移", fontsize=11.5, fontweight="bold",
             color=NAVY, pad=18, loc="left")
fig.tight_layout()
fig.savefig(OUTDIR / "juku_trend_juken.png", bbox_inches="tight", facecolor="white")
plt.close(fig)

print("saved:")
for p in sorted(OUTDIR.glob("juku_*.png")):
    print("  ", p.relative_to(ROOT))
