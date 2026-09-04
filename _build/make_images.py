# -*- coding: utf-8 -*-
"""PPTX埋め込み用の画像を生成する。

  pip install matplotlib pillow
  python _build/make_images.py              # 全部生成
  python _build/make_images.py trend        # トレンドグラフだけ
  python _build/make_images.py phone menu   # スマホモックとリッチメニューだけ

出力先: _images/
スライドサイズが 27.52 × 19.05 cm なので、本文エリアに収まる幅で書き出す。
python-pptx 側では `slide.shapes.add_picture(path, Cm(x), Cm(y), width=Cm(w))` で貼る。

──────────────────────────────────────────────────────────────
【系列色について】
DYMブランドの ネイビー 1F285A / 炭黒 333333 は **インク色であってデータ色ではない**。
そのまま折れ線に使うと、明度が暗すぎ・彩度が低すぎて系列として読めない
（配色バリデータで Lightness band と Chroma floor が FAIL する）。

そこで **同じ色相のまま、データ用の明度・彩度に持ち上げた4色** を系列色に使う。
下の CAT は色覚多様性チェック込みで全項目 PASS 済み（CVD ΔE 10.9）。
文字・軸・注釈は今までどおりブランドのインク色を使うので、資料の見た目は揃う。
──────────────────────────────────────────────────────────────
"""
import csv
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "_data" / "trends"
OUT = ROOT / "_images"
OUT.mkdir(exist_ok=True)

# ---- 系列色（バリデータ全項目PASS・CVD ΔE 10.9）----
CAT = ["#3467B2", "#D9661F", "#00897B", "#8E4EC6"]
# ---- インク・面（DYMブランドのまま）----
INK = "#333333"        # 本文
INK_STRONG = "#1F285A"  # 見出し
MUTED = "#7F8798"       # 補助・脚注
GRID = "#E4E7EE"
SURFACE = "#FFFFFF"
ACCENT = "#D9661F"      # 強調（オレンジ系列色と同一）

# ---- フォント（PCで再生成するならメイリオが使われる）----
# ★IPAゴシックより Noto を先に置く。**IPAには太字（Bold）が無い**ため、
#   IPAが選ばれると fontweight="bold" が黙って通常字に落ちて強調が消える。
#   クラウド環境では `apt-get install -y fonts-noto-cjk` で入る。
FONT_CANDIDATES = ["Meiryo", "メイリオ", "Noto Sans CJK JP", "IPAPGothic", "IPAGothic"]


def pick_font():
    avail = {f.name for f in font_manager.fontManager.ttflist}
    for name in FONT_CANDIDATES:
        if name in avail:
            return name
    return "DejaVu Sans"


JP = pick_font()
plt.rcParams.update({
    "font.family": JP,
    "axes.unicode_minus": False,
    "figure.facecolor": SURFACE,
    "axes.facecolor": SURFACE,
    "savefig.facecolor": SURFACE,
})


def _pil_font(size, bold=False):
    """PILで使う日本語フォントを拾う。太字が無ければ通常で代用。"""
    for p in ("/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
              "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
              "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
              "C:/Windows/Fonts/meiryob.ttc" if bold else "C:/Windows/Fonts/meiryo.ttc"):
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


# ==========================================================
# 1. Googleトレンドの折れ線
# ==========================================================
def load_csv(path):
    rows = list(csv.DictReader(open(path, encoding="utf-8-sig")))
    cols = [c for c in rows[0] if c != "Time"]
    x = [r["Time"] for r in rows]
    series = {c: [int(r[c]) for r in rows] for c in cols}
    return x, series


def trend_chart(csv_path, out_name, title, subtitle="", note="",
                only=None, annotate_peak=True, width=11.0, height=5.0):
    """Googleトレンドの折れ線グラフ。

    ★1枚に1軸だけ。**スケールが違う系列を混ぜない**（混ぜると小さい方が潰れる）。
      Ver.B の `土地活用` `アパート経営` が 0 になったのがこの事故。
    """
    x, series = load_csv(csv_path)
    if only:
        series = {k: v for k, v in series.items() if k in only}
    n = len(series)

    fig, ax = plt.subplots(figsize=(width, height), dpi=200)

    # 目盛りは控えめ（横線のみ）
    ax.set_axisbelow(True)
    ax.yaxis.grid(True, color=GRID, linewidth=0.8)
    ax.xaxis.grid(False)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(GRID)
    ax.tick_params(colors=MUTED, labelsize=9, length=0)

    xs = range(len(x))
    for i, (name, vals) in enumerate(series.items()):
        c = CAT[i % len(CAT)]
        ax.plot(xs, vals, color=c, linewidth=2.0, solid_capstyle="round",
                label=name, zorder=3)
        # 系列名を線の右端に直付け（4系列までは直付け＋凡例の二重表示）
        ax.annotate(name, xy=(len(x) - 1, vals[-1]), xytext=(6, 0),
                    textcoords="offset points", color=c, fontsize=9.5,
                    va="center", fontweight="bold", zorder=4)
        # ピークだけ点と数値を出す（全点には出さない）
        if annotate_peak:
            pi = vals.index(max(vals))
            ax.plot([pi], [vals[pi]], "o", color=c, markersize=6,
                    markeredgecolor=SURFACE, markeredgewidth=1.6, zorder=5)
            ax.annotate(f"{max(vals)}", xy=(pi, vals[pi]), xytext=(0, 9),
                        textcoords="offset points", ha="center",
                        color=INK, fontsize=9, fontweight="bold", zorder=5)

    # 横軸のラベルは間引く
    step = max(1, len(x) // 8)
    ticks = list(range(0, len(x), step))
    ax.set_xticks(ticks)
    ax.set_xticklabels([x[i][:7] for i in ticks], fontsize=8.5)
    ax.set_ylim(0, max(max(v) for v in series.values()) * 1.18)
    ax.set_xlim(-0.5, len(x) - 0.5 + len(x) * 0.10)

    # 見出しは set_title を使わない（pad とサブタイトルが衝突するため）。
    # 軸座標で「サブタイトル → タイトル」の順に積む。bbox_inches="tight" が拾ってくれる。
    if subtitle:
        ax.text(0, 1.03, subtitle, transform=ax.transAxes,
                color=MUTED, fontsize=10, va="bottom")
        ax.text(0, 1.11, title, transform=ax.transAxes,
                color=INK_STRONG, fontsize=14, fontweight="bold", va="bottom")
    else:
        ax.text(0, 1.03, title, transform=ax.transAxes,
                color=INK_STRONG, fontsize=14, fontweight="bold", va="bottom")
    if n >= 2:
        ax.legend(loc="upper left", frameon=False, fontsize=9.5,
                  labelcolor=INK, ncol=min(n, 4), bbox_to_anchor=(0, -0.10))
    if note:
        fig.text(0.012, 0.012, note, color=MUTED, fontsize=8)

    fig.tight_layout(rect=(0, 0.05 if note else 0.02, 1, 0.94))
    p = OUT / out_name
    fig.savefig(p, bbox_inches="tight")
    plt.close(fig)
    print("  ", p.name)
    return p


# ==========================================================
# 1b. CPC推移のスロープチャート（2点だけの比較）
# ==========================================================
# ★ y軸に目盛りを出さない。2点しかないので、値は点に直付けする。
#   目盛りを出すと「0起点でない＝誇張」に見えるが、直付けなら全数値が読める。
def cpc_slope(out_name, series, xlabels, title, subtitle="", note="",
              width=7.6, height=5.6):
    """series = [(系列名, [左の値, 右の値], 表示文字列2つ, 増減率の文字列, 色, 太さ)]"""
    fig, ax = plt.subplots(figsize=(width, height), dpi=200)

    ax.set_axisbelow(True)
    ax.yaxis.grid(True, color=GRID, linewidth=0.8)
    ax.xaxis.grid(False)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(GRID)
    ax.set_yticks([])
    ax.tick_params(colors=MUTED, labelsize=10, length=0)

    lo = min(min(v) for _, v, _, _, _, _ in series)
    hi = max(max(v) for _, v, _, _, _, _ in series)
    pad = (hi - lo) * 0.45

    for name, vals, texts, delta, color, lw in series:
        ax.plot([0, 1], vals, color=color, linewidth=lw,
                solid_capstyle="round", zorder=3)
        for xi in (0, 1):
            ax.plot([xi], [vals[xi]], "o", color=color, markersize=8,
                    markeredgecolor=SURFACE, markeredgewidth=2.0, zorder=5)
        # 左端＝値だけ／右端＝値と増減率を線の色で
        ax.annotate(texts[0], xy=(0, vals[0]), xytext=(-10, 0),
                    textcoords="offset points", ha="right", va="center",
                    color=color, fontsize=13, fontweight="bold", zorder=6)
        ax.annotate(texts[1], xy=(1, vals[1]), xytext=(12, 6),
                    textcoords="offset points", ha="left", va="center",
                    color=color, fontsize=13, fontweight="bold", zorder=6)
        ax.annotate(f"{name}　{delta}", xy=(1, vals[1]), xytext=(12, -11),
                    textcoords="offset points", ha="left", va="center",
                    color=color, fontsize=10.5, fontweight="bold", zorder=6)

    ax.set_xticks([0, 1])
    ax.set_xticklabels(xlabels, fontsize=10.5)
    ax.set_xlim(-0.30, 1.62)
    ax.set_ylim(lo - pad, hi + pad)

    if subtitle:
        ax.text(0, 1.04, subtitle, transform=ax.transAxes,
                color=MUTED, fontsize=10, va="bottom")
        ax.text(0, 1.12, title, transform=ax.transAxes,
                color=INK_STRONG, fontsize=14, fontweight="bold", va="bottom")
    else:
        ax.text(0, 1.04, title, transform=ax.transAxes,
                color=INK_STRONG, fontsize=14, fontweight="bold", va="bottom")
    if note:
        fig.text(0.012, 0.012, note, color=MUTED, fontsize=7.5)

    fig.tight_layout(rect=(0, 0.06 if note else 0.02, 1, 0.93))
    p = OUT / out_name
    fig.savefig(p, bbox_inches="tight")
    plt.close(fig)
    print("  ", p.name)
    return p


# ==========================================================
# 2. LINEトーク画面のモック
# ==========================================================
BEZEL, SCREEN, GREEN = "#2B2B2B", "#EFF2F7", "#06C755"


def phone_mock(out_name, header, lines, w=520, h=None, pad=18):
    """lines = [(kind, text)] / kind: in（BOT吹き出し）chip（選択肢）btn（CTA）note

    h を省略すると中身の分だけの高さで書き出す（下の余白が空かない）。
    """
    f_h = _pil_font(17, True)
    f_b = _pil_font(15)

    body = sum(18 + 22 * len(t.split("\n")) + 10 for _, t in lines)
    if h is None:
        h = 54 + 62 + body + 24 + pad

    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)

    d.rounded_rectangle([0, 0, w - 1, h - 1], radius=26, fill=BEZEL)
    sx0, sy0, sx1, sy1 = pad, 54, w - pad, h - pad
    d.rectangle([sx0, sy0, sx1, sy1], fill=SCREEN)
    d.rectangle([sx0, sy0, sx1, sy0 + 44], fill=GREEN)
    d.text(((sx0 + sx1) // 2, sy0 + 22), header, font=f_h, fill="white", anchor="mm")

    y = sy0 + 62
    iw = sx1 - sx0 - 34
    for kind, text in lines:
        rows = text.split("\n")
        bh = 18 + 22 * len(rows)
        x0 = sx0 + 17
        if kind == "in":
            d.rounded_rectangle([x0, y, x0 + iw, y + bh], radius=12,
                                fill="white", outline="#D9DEE8", width=1)
            tc = INK
        elif kind == "chip":
            d.rounded_rectangle([x0, y, x0 + iw, y + bh], radius=18,
                                fill="white", outline=GREEN, width=2)
            tc = "#0B7A3B"
        elif kind == "btn":
            d.rounded_rectangle([x0, y, x0 + iw, y + bh], radius=18, fill=GREEN)
            tc = "white"
        else:  # note
            tc = MUTED
        for i, r in enumerate(rows):
            tx = x0 + 14 if kind == "in" else x0 + iw // 2
            anc = "la" if kind == "in" else "ma"
            d.text((tx, y + 9 + 22 * i), r, font=f_b, fill=tc, anchor=anc)
        y += bh + 10
        if y > sy1 - 30:
            break

    p = OUT / out_name
    img.save(p)
    print("  ", p.name)
    return p


# ==========================================================
# 3. リッチメニュー（3タブ × 2列3行）
# ==========================================================
def richmenu_mock(out_name, tabs, active=0, w=900, h=760):
    """tabs = [(タブ名, [6ボタンの文言])]"""
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    f_t = _pil_font(20, True)
    f_b = _pil_font(19, True)
    f_s = _pil_font(14)

    tw = w // len(tabs)
    for i, (name, _) in enumerate(tabs):
        x0 = i * tw
        on = (i == active)
        d.rectangle([x0, 0, x0 + tw - 2, 62], fill="white")
        d.text((x0 + tw // 2, 30), name, font=f_t,
               fill=GREEN if on else MUTED, anchor="mm")
        if on:
            d.rectangle([x0 + 12, 58, x0 + tw - 14, 62], fill=GREEN)

    btns = tabs[active][1]
    gx0, gy0 = 16, 78
    cw, ch = (w - 32) // 2, (h - gy0 - 16) // 3
    for i, label in enumerate(btns[:6]):
        cx, cy = gx0 + (i % 2) * cw, gy0 + (i // 2) * ch
        top_left = (i == 0)
        d.rounded_rectangle([cx + 4, cy + 4, cx + cw - 8, cy + ch - 8], radius=10,
                            fill="#FCE9DC" if top_left else "#F4F7FF",
                            outline=ACCENT if top_left else "#C9D2E6",
                            width=3 if top_left else 1)
        for j, ln in enumerate(label.split("\n")):
            d.text((cx + cw // 2 - 2, cy + ch // 2 - 10 + 24 * j), ln, font=f_b,
                   fill=ACCENT if top_left else INK_STRONG, anchor="mm")
        if top_left:
            d.text((cx + cw // 2 - 2, cy + 22), "★ 最重要CTA", font=f_s,
                   fill=ACCENT, anchor="mm")

    p = OUT / out_name
    img.save(p)
    print("  ", p.name)
    return p


# ==========================================================
# 4. ファネル図
# ==========================================================
def funnel(out_name, steps, title="", w=1100, h=560):
    """steps = [(ラベル, 補足, 強調するか)] 上から下へ"""
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    f_t = _pil_font(24, True)
    f_l = _pil_font(21, True)
    f_s = _pil_font(16)

    y = 20
    if title:
        d.text((28, y), title, font=f_t, fill=INK_STRONG)
        y += 46
    n = len(steps)
    bh = (h - y - 30) // n - 12
    top_w, bot_w = w - 120, int((w - 120) * 0.56)
    for i, (label, sub, hot) in enumerate(steps):
        bw = int(top_w - (top_w - bot_w) * i / max(n - 1, 1))
        x0 = (w - bw) // 2
        d.rounded_rectangle([x0, y, x0 + bw, y + bh], radius=8,
                            fill="#FCE9DC" if hot else "#F4F7FF",
                            outline=ACCENT if hot else "#C9D2E6", width=3 if hot else 1)
        cy = y + bh // 2
        fg = ACCENT if hot else INK_STRONG
        pad = 22
        # 横に並べると重なる幅なら2行に折る（下の段ほど箱が狭くなるので必ず起きる）
        wide = d.textlength(label, font=f_l) + d.textlength(sub or "", font=f_s) \
            + pad * 2 + 28 > bw
        if sub and wide:
            d.text((x0 + pad, cy - 13), label, font=f_l, fill=fg, anchor="lm")
            d.text((x0 + pad, cy + 14), sub, font=f_s, fill=INK, anchor="lm")
        else:
            d.text((x0 + pad, cy), label, font=f_l, fill=fg, anchor="lm")
            if sub:
                d.text((x0 + bw - pad, cy), sub, font=f_s, fill=INK, anchor="rm")
        y += bh + 12
        if i < n - 1:
            cx = w // 2
            d.polygon([(cx - 9, y - 10), (cx + 9, y - 10), (cx, y - 1)], fill="#9AA3B5")

    p = OUT / out_name
    img.save(p)
    print("  ", p.name)
    return p


# ==========================================================
# 生成タスク
# ==========================================================
def gen_trend():
    print("[trend]")
    # ★Ver.A：スケールが違うので1枚に混ぜない
    trend_chart(DATA / "verA_1year.csv", "verA_trend_chintai.png",
                "「賃貸」の検索は年中フラット",
                "Googleトレンド／日本／直近1年　※相対指標（期間内の最大=100）",
                "出典：Googleトレンド（2026年8月取得）。取得日を提案時に更新すること。",
                only=["賃貸"])
    trend_chart(DATA / "verA_1year.csv", "verA_trend_hikkoshi.png",
                "山が出るのは「引っ越し」「一人暮らし」— 1〜3月",
                "Googleトレンド／日本／直近1年　※「賃貸」とはスケールが違うため別グラフ",
                "出典：Googleトレンド（2026年8月取得）。※「内見」は検索Volが小さく測定不能のため除外。",
                only=["引っ越し", "一人暮らし"])
    # ★Ver.B：確定申告と混ぜると土地活用・アパート経営が0になるので単独で出す
    trend_chart(DATA / "souzoku_5year.csv", "verB_trend_souzoku5y.png",
                "「相続」への関心は5年で上がり続けている",
                "Googleトレンド／日本／過去5年　※オーナーが動く最大のトリガー",
                "出典：Googleトレンド（2026年8月取得）。")
    trend_chart(DATA / "verB_5year.csv", "verB_trend_kakutei5y.png",
                "「確定申告」は毎年2〜3月に跳ねる",
                "Googleトレンド／日本／過去5年　※モーメント配信の設計図",
                "出典：Googleトレンド（2026年8月取得）。"
                "※「土地活用」「アパート経営」「空室」は検索Volが小さく本グラフでは測定不能。",
                only=["確定申告"])


def gen_phone():
    print("[phone]")
    phone_mock("verA_day0.png", "●●不動産", [
        ("in", "お問い合わせありがとうございます！\n担当の田中です。"),
        ("in", "✅ こちらからお電話はしません\n✅ 来店を急かすこともしません"),
        ("in", "5つだけ教えてください。\nご希望に合うお部屋をお送りします。"),
        ("chip", "希望条件を入力する"),
        ("note", "※ 反響から5分以内に自動送信"),
    ])
    phone_mock("verA_day14.png", "●●不動産", [
        ("in", "お部屋はお決まりですか？"),
        ("in", "【申込前に確認すべき5つ】\n① 審査は保証会社が見ます\n② 保証人なしでも通ります\n③ 必要書類は3点だけ"),
        ("in", "「審査が不安で動けない」\nという方が実は一番多いです。"),
        ("btn", "審査の可否を先に確認する"),
        ("chip", "今回は見送る"),
        ("note", "※ 通知冒頭15字：お部屋はお決まりで"),
    ])
    phone_mock("verB_day0.png", "●●管理", [
        ("in", "診断結果が出ました📋"),
        ("in", "【ご所有の土地】\n市街化区域／建築可能\n周辺相場との差 ▲8,000円/月"),
        ("in", "この土地に向くのは\nアパートだけではありません。\n駐車場・トランクルームも\n候補に入ります。"),
        ("btn", "3つの活用案を見る"),
        ("note", "※ 訪問・お電話はいたしません"),
    ])


def gen_menu():
    print("[menu]")
    richmenu_mock("verA_richmenu.png", [
        ("探す", ["今すぐ\n物件を探す", "エリアから\n探す", "初期費用が\n安い物件",
                "ペット可・\n家具付き", "新着を\n受け取る", "条件を\n変更する"]),
        ("お金・審査", ["初期費用を\n計算する", "家賃の目安\n（手取り比）", "審査について",
                    "保証人なしで\n借りたい", "フリーレント\nとは", "相談する"]),
        ("来店・内見", ["来店を\n予約する", "オンライン\n内見", "内見の流れ",
                    "IT重説\nについて", "店舗の場所", "日程を\n変更する"]),
    ], active=0)
    richmenu_mock("verB_richmenu.png", [
        ("まず診断", ["土地活用\nタイプ診断", "空室・家賃\n査定", "建てられるか\n調べる",
                  "活用事例を\n見る", "よくある\n失敗", "相談する"]),
        ("収支・税金", ["収支を\n試算する", "確定申告の\n備え", "相続税対策",
                   "固定資産税\nの見方", "修繕費の\n相場", "FPに相談"]),
        ("事例・セミナー", ["セミナーに\n申し込む", "オーナーの声", "エリア相場",
                     "法改正の\nまとめ", "オンライン\n相談", "資料を\n受け取る"]),
    ], active=0)


def gen_funnel():
    print("[funnel]")
    funnel("verA_funnel.png", [
        ("ポータル反響", "触れない（掲載単価は言い値）", False),
        ("初動の返信", "★5分以内 → 自動化で解く", True),
        ("LINE友だち化", "★ここが資産になる", True),
        ("来店・内見予約", "リマインドでドタキャンを減らす", True),
        ("賃貸借契約", "", False),
        ("更新・住み替え・紹介", "★併営なら流出を止められる", True),
    ], title="触るのは「反響が来た後」と「予約が入った後」だけ")
    funnel("verB_funnel.png", [
        ("DM・チラシ・訪問", "触れない", False),
        ("土地活用タイプ診断", "★訪問も電話もされずに数字が出る", True),
        ("育てる区間（1〜3年）", "★モーメント配信で捨てない", True),
        ("個別相談・セミナー予約", "★主KPIはここのCPA・CVR", True),
        ("管理受託・請負契約", "CPOは3年累積の参考値", False),
    ], title="「意思はあるが今すぐでない」区間を埋める")


def gen_cpc():
    """ジム業界 p11 用。CPC推移（全業界平均 vs ジム・フィットネス）。

    出典：WordStream / LocaliQ「Google Ads Benchmarks」（米国・Google広告）。
    レポート名の年は**発行年**で、データ期間は前年4月〜当年3月。
    ★元データは米ドル。円は 1ドル=157円（2026/9/3時点）で換算した参考値であって、
      **日本のジムのCPCではない**。スライドの脚注に必ず明記すること。
    """
    print("[cpc]")
    rate = 157
    usd = {"all": (5.26, 5.42), "gym": (5.00, 6.17)}
    yen = {k: (round(a * rate), round(b * rate)) for k, (a, b) in usd.items()}

    cpc_slope(
        "gym_cpc_yen.png",
        [
            ("全業界平均", list(yen["all"]),
             (f"¥{yen['all'][0]:,}", f"¥{yen['all'][1]:,}"), "+3.0%", CAT[0], 2.2),
            ("ジム・フィットネス", list(yen["gym"]),
             (f"¥{yen['gym'][0]:,}", f"¥{yen['gym'][1]:,}"), "+23.4%", ACCENT, 3.4),
        ],
        ["2025年版", "2026年版"],
        title="ジムのCPCは、全業界平均の約8倍のペースで上昇",
        subtitle="検索広告の平均クリック単価（円換算）",
        note="※WordStream / LocaliQ「Google Ads Benchmarks」（米国・Google広告）を"
             f"1ドル={rate}円で換算した参考値。日本のジム業界の実測CPCではない。"
             "レポート年は発行年で、データ期間は前年4月〜当年3月。",
    )


TASKS = {"trend": gen_trend, "phone": gen_phone, "menu": gen_menu,
         "funnel": gen_funnel, "cpc": gen_cpc}

if __name__ == "__main__":
    want = sys.argv[1:] or list(TASKS)
    print(f"フォント: {JP}　出力先: {OUT}")
    for t in want:
        if t in TASKS:
            TASKS[t]()
        else:
            print(f"  不明なタスク: {t}（有効: {', '.join(TASKS)}）")
    print("完了")
