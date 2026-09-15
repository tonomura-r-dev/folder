# -*- coding: utf-8 -*-
"""学校法人（大学・専門学校）業界の提案資料に貼る図版を生成する。

  python3 _build/make_images_gakkou.py

出力：_images/gakkou_*.png
入力：_data/trends/学校法人/*.csv（Googleトレンド実測・2026-09-15取得）

【データの素性（スライドに載せるときの注記）】
- 5年月次＝2021-09〜2026-09。**2026-09は月途中（9/15まで）なので参考値**
- 週次＝2024-12-29〜2025-12-28。**2025年の実測**（2026年ではない）

【ルール（CLAUDE.md より）】
- Googleトレンドは相対指標。スケールの違うKWを1枚に混ぜない
  → 「共通テスト」は1月に100、「大学受験」は6〜17で約17倍差 → 必ず分ける
- 系列色に DYMのネイビー 1F285A・炭黒 333333 は使わない（配色バリデータFAIL）
  → 3467B2 / D9661F / 00897B / 8E4EC6 を使う。文字・軸・注釈はインク色のまま
"""
import csv
from datetime import datetime, timedelta
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import matplotlib.dates as mdates

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "_data" / "trends" / "学校法人"
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
NAVY = "#1F285A"
ORANGE = "#D9661F"
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


def base(ax, dates, ymax=108, ylabel="検索インタレスト（相対値・最大100）"):
    ax.set_ylim(0, ymax)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.grid(axis="y", color="#E8E8E8", lw=0.8)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.set_xlim(dates[0], dates[-1])
    ax.tick_params(labelsize=9)


def yearly_axis(ax):
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=[7]))


def monthly_axis(ax):
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m月"))


def peak_dot(ax, d, v, color, label, dy=14):
    ax.plot(d, v, "o", color=color, ms=8, zorder=5)
    ax.annotate(label, xy=(d, v), xytext=(0, dy), textcoords="offset points",
                ha="center", fontsize=10, color=color, fontweight="bold", zorder=6)


def vline(ax, d, text, y=96, ha="right", dx=-8):
    ax.axvline(d, color=ORANGE, lw=1.6, ls=(0, (5, 3)), zorder=4)
    ax.annotate(text, xy=(d, y), xytext=(dx, 0), textcoords="offset points",
                ha=ha, va="top", fontsize=9.5, color=ORANGE, fontweight="bold")


def save(fig, name):
    fig.tight_layout()
    fig.savefig(OUTDIR / name, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ============================================================
# 図1 「オープンキャンパス」5年月次｜毎年まったく同じ形
# ============================================================
d1, h1, c1 = load("図1_オープンキャンパス_5年月次.csv")
v = c1[0]
fig, ax = plt.subplots(figsize=(9.6, 3.5))
ax.plot(d1, v, color=SERIES[0], lw=2.2)
ax.fill_between(d1, v, color=SERIES[0], alpha=0.10)
base(ax, d1)
yearly_axis(ax)
# 各年の7月にマーカー
for y in range(2022, 2027):
    i = next(k for k, dt in enumerate(d1) if dt.year == y and dt.month == 7)
    ax.plot(d1[i], v[i], "o", color=ORANGE, ms=6, zorder=5)
    ax.annotate(f"{v[i]}", xy=(d1[i], v[i]), xytext=(0, 9), textcoords="offset points",
                ha="center", fontsize=9.5, color=ORANGE, fontweight="bold")
ax.annotate("毎年7月がピーク（5年連続）", xy=(d1[1], 72), fontsize=10.5,
            color=ORANGE, fontweight="bold")
ax.annotate("谷は12〜1月（3〜5）＝ピークの約1/20", xy=(d1[15], 30), fontsize=9.5, color=MUT)
ax.set_title("「オープンキャンパス」検索の5年推移｜形がまったく変わらない＝予定を立てられる需要",
             fontsize=12, fontweight="bold", color=NAVY, pad=24, loc="left")
ax.annotate("※2026年9月は月途中（9/15まで）の参考値", xy=(0.995, -0.20), xycoords="axes fraction",
            ha="right", fontsize=8, color=MUT)
save(fig, "gakkou_trend_oc_5y.png")

# ============================================================
# 図2 「オープンキャンパス」2025年週次｜仕込みはいつ始めるべきか
# ============================================================
d2, h2, c2 = load("図2_オープンキャンパス_2025年週次.csv")
v = c2[0]
fig, ax = plt.subplots(figsize=(9.6, 3.6))
ax.plot(d2, v, color=SERIES[0], lw=2.4)
ax.fill_between(d2, v, color=SERIES[0], alpha=0.10)
base(ax, d2)
monthly_axis(ax)
ipk = v.index(max(v))
peak_dot(ax, d2[ipk], v[ipk], SERIES[0], f"ピーク\n{d2[ipk]:%-m/%-d}週　100", dy=10)
# 50%到達点
i50 = next(k for k, x in enumerate(v) if x >= 50)
ax.plot(d2[i50], v[i50], "o", color=ORANGE, ms=7, zorder=5)
ax.annotate(f"{d2[i50]:%-m/%-d}週で半分（{v[i50]}）", xy=(d2[i50], v[i50]),
            xytext=(-6, 13), textcoords="offset points", ha="right",
            fontsize=10, color=ORANGE, fontweight="bold")
# 仕込み開始＝20を超える週
i20 = next(k for k, x in enumerate(v) if x >= 20)
vline(ax, d2[i20], f"動き出し\n{d2[i20]:%-m/%-d}週", y=92)
ax.axvspan(d2[i20], d2[i50], color=ORANGE, alpha=0.07, zorder=0)
ax.annotate("仕込みの窓＝5週間", xy=(d2[i20] + (d2[i50] - d2[i20]) / 2, 84),
            ha="center", fontsize=10, color=ORANGE, fontweight="bold")
ax.annotate("お盆で一度落ちる", xy=(d2[v.index(25)], 25), xytext=(4, 14),
            textcoords="offset points", fontsize=9, color=MUT)
ax.set_title("「オープンキャンパス」2025年の週次｜6月末には年間ピークの半分まで来ている",
             fontsize=12, fontweight="bold", color=NAVY, pad=24, loc="left")
ax.annotate("※データは2025年（2024-12-29〜2025-12-28）の実測", xy=(0.995, -0.20),
            xycoords="axes fraction", ha="right", fontsize=8, color=MUT)
save(fig, "gakkou_trend_oc_weekly.png")

# ============================================================
# 図3 入試方式3種 5年月次｜総合型が指定校推薦を抜いた
# ============================================================
d3, h3, c3 = load("図3_総合型_学校推薦型_指定校推薦_5年月次.csv")
fig, ax = plt.subplots(figsize=(9.6, 3.6))
for k, name in enumerate(h3):
    ax.plot(d3, c3[k], color=SERIES[k], lw=2.2, label=name)
base(ax, d3)
yearly_axis(ax)
i_so, i_si = h3.index("総合型選抜"), h3.index("指定校推薦")
# 9月値の推移を注記
for y, dy in [(2021, 12), (2026, 12)]:
    i = next(k for k, dt in enumerate(d3) if dt.year == y and dt.month == 9)
    ax.annotate(f"{c3[i_so][i]}", xy=(d3[i], c3[i_so][i]), xytext=(0, dy),
                textcoords="offset points", ha="center", fontsize=10,
                color=SERIES[i_so], fontweight="bold")
    ax.annotate(f"{c3[i_si][i]}", xy=(d3[i], c3[i_si][i]), xytext=(0, -20),
                textcoords="offset points", ha="center", fontsize=10,
                color=SERIES[i_si], fontweight="bold")
# 逆転点（2023-09 で 56 : 56）
ix = next(k for k, dt in enumerate(d3) if dt.year == 2023 and dt.month == 9)
ax.plot(d3[ix], c3[i_so][ix], "o", color=ORANGE, ms=9, zorder=6)
ax.annotate("2023年9月に逆転", xy=(d3[ix], c3[i_so][ix]), xytext=(0, 20),
            textcoords="offset points", ha="center", fontsize=10.5,
            color=ORANGE, fontweight="bold")
ax.legend(fontsize=9.5, frameon=False, ncol=3, loc="upper left")
ax.set_title("入試方式の検索推移（5年）｜総合型選抜は5年で3.3倍。指定校推薦は横ばい",
             fontsize=12, fontweight="bold", color=NAVY, pad=24, loc="left")
ax.annotate("※9月値の比較：総合型 30→100（3.3倍）／指定校推薦 64→57（▲11%）　※2026年9月は月途中",
            xy=(0.995, -0.20), xycoords="axes fraction", ha="right", fontsize=8, color=MUT)
save(fig, "gakkou_trend_nyushi_5y.png")

# ============================================================
# 図4 入試方式3種 2025年週次｜3つとも山が別の時期にある
# ============================================================
d4, h4, c4 = load("図4_総合型_学校推薦型_指定校推薦_2025年週次.csv")
fig, ax = plt.subplots(figsize=(9.6, 3.6))
for k, name in enumerate(h4):
    ax.plot(d4, c4[k], color=SERIES[k], lw=2.2, label=name)
base(ax, d4)
monthly_axis(ax)
j_so = h4.index("総合型選抜")
ip = c4[j_so].index(max(c4[j_so]))
peak_dot(ax, d4[ip], c4[j_so][ip], SERIES[j_so], f"{d4[ip]:%-m/%-d}週 100\n（9/1出願解禁の週）", dy=8)
# 11月の学校推薦型ピーク
j_ga = h4.index("学校推薦型選抜")
ig = c4[j_ga].index(max(c4[j_ga]))
ax.plot(d4[ig], c4[j_ga][ig], "o", color=SERIES[j_ga], ms=7, zorder=5)
ax.annotate(f"{d4[ig]:%-m/%-d}週\n（11/1出願解禁の直後）", xy=(d4[ig], c4[j_ga][ig]),
            xytext=(0, 12), textcoords="offset points", ha="center",
            fontsize=9.5, color=SERIES[j_ga], fontweight="bold")
ax.legend(fontsize=9.5, frameon=False, ncol=3, loc="upper left")
ax.set_title("入試方式の検索推移（2025年・週次）｜総合型＝8月末／指定校＝8月末と11月／学校推薦型＝11月",
             fontsize=11.5, fontweight="bold", color=NAVY, pad=24, loc="left")
ax.annotate("※データは2025年の実測", xy=(0.995, -0.20), xycoords="axes fraction",
            ha="right", fontsize=8, color=MUT)
save(fig, "gakkou_trend_nyushi_weekly.png")

# ============================================================
# 図5 「共通テスト」単独｜1月中旬の2週間しかない
#   （大学受験6〜17に対し最大100＝約17倍。混ぜると潰れるので分離）
# ============================================================
d5, h5, c5 = load("図5_大学受験_専門学校_共通テスト_2025年週次.csv")
j_kt = h5.index("共通テスト")
v = c5[j_kt]
fig, ax = plt.subplots(figsize=(9.6, 3.0))
ax.plot(d5, v, color=SERIES[3], lw=2.4)
ax.fill_between(d5, v, color=SERIES[3], alpha=0.10)
base(ax, d5)
monthly_axis(ax)
ip = v.index(max(v))
peak_dot(ax, d5[ip], v[ip], SERIES[3], f"{d5[ip]:%-m/%-d}週　100", dy=10)
ax.annotate("1年のうち、動くのはこの2週間だけ", xy=(d5[ip], 60), xytext=(14, 0),
            textcoords="offset points", fontsize=10.5, color=ORANGE, fontweight="bold")
ax.set_title("「共通テスト」2025年の週次｜一般選抜層は1月中旬にしか動かない",
             fontsize=12, fontweight="bold", color=NAVY, pad=22, loc="left")
ax.annotate("※「大学受験」（6〜17）とはスケールが約17倍違うため、同じグラフに混ぜていない",
            xy=(0.995, -0.22), xycoords="axes fraction", ha="right", fontsize=8, color=MUT)
save(fig, "gakkou_trend_kyotsu.png")

# ============================================================
# 図6 「専門学校」×「大学受験」5年月次｜専門学校は年中フラット
# ============================================================
d6, h6, c6 = load("図6_大学受験_専門学校_共通テスト_5年月次.csv")
fig, ax = plt.subplots(figsize=(9.6, 3.2))
for k, name in enumerate(h6):
    if name == "共通テスト":
        continue
    ax.plot(d6, c6[k], color=SERIES[k], lw=2.2, label=name)
base(ax, d6, ymax=46, ylabel="検索インタレスト（相対値）")
yearly_axis(ax)
j_sen = h6.index("専門学校")
ax.axhspan(22, 40, color=SERIES[j_sen], alpha=0.06, zorder=0)
ax.annotate("「専門学校」は5年間ずっと22〜40のレンジ＝季節性がほぼない",
            xy=(d6[6], 42), fontsize=10.5, color=SERIES[j_sen], fontweight="bold")
ax.annotate("「大学受験」は1〜2月にだけ小さな山", xy=(d6[30], 18), xytext=(4, 0),
            textcoords="offset points", fontsize=9.5, color=SERIES[h6.index("大学受験")])
ax.legend(fontsize=9.5, frameon=False, ncol=2, loc="lower left")
ax.set_title("「専門学校」と「大学受験」の5年推移｜専門学校の需要は年中フラット（大学とは別設計が要る）",
             fontsize=11.5, fontweight="bold", color=NAVY, pad=22, loc="left")
ax.annotate("※「共通テスト」（1月に100）は同一グラフから除外。※2026年9月は月途中",
            xy=(0.995, -0.22), xycoords="axes fraction", ha="right", fontsize=8, color=MUT)
save(fig, "gakkou_trend_senmon.png")

print("saved:")
for p in sorted(OUTDIR.glob("gakkou_*.png")):
    print("  ", p.relative_to(ROOT))
