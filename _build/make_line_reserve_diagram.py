#!/usr/bin/env python3
"""LINE予約システム 自作構成図（1枚もの）を描く。

    python3 _build/make_line_reserve_diagram.py
    → _images/LINE予約_自作構成図.png（3200×1800）

設計の本文は _drafts/STORES予約×LINE_自作設計図.md。文言を直すときはこのスクリプトを触って再生成する。
フォントはメイリオが無い環境向けに IPAPゴシックへ自動フォールバック（太字は縁取りで擬似的に出す）。
"""
import math
import os
import sys

import matplotlib

matplotlib.use("Agg")
from matplotlib import font_manager, pyplot as plt  # noqa: E402
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch  # noqa: E402
from matplotlib.patheffects import withStroke  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "_images", "LINE予約_自作構成図.png")

# ---- 配色（make_images.py と同じ系統。ネイビー/炭黒はインク色として使う）----
NAVY = "#1F285A"
INK = "#333333"
MUTED = "#6B7280"
GREEN = "#06C755"   # LINE
TEAL = "#00897B"    # 無料の通知（Reply）
BLUE = "#3467B2"    # Google
ORANGE = "#D9661F"  # 通数を使う（Push）
RED = "#C00000"
BORDER = "#D9D9D9"
LANE_A = "#F4F7FF"
LANE_B = "#EEFBF3"
LANE_C = "#EEF3FB"
CHIP = "#F4F7FF"
CHIP_MUTED = "#F1F1F1"

FONT_CANDIDATES = [
    "C:/Windows/Fonts/meiryo.ttc",
    "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
]


def setup_font():
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            font_manager.fontManager.addfont(p)
            name = font_manager.FontProperties(fname=p).get_name()
            plt.rcParams["font.family"] = name
            return name
    return None


FONT = setup_font()

# キャンバスは 160×90 単位（1単位 = 20px @200dpi）
W, H = 160, 90
fig = plt.figure(figsize=(16, 9), dpi=200)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(0, H)
ax.axis("off")
fig.patch.set_facecolor("white")


# ---- 描画ヘルパー ----
def T(x, y, s, size=9, color=INK, bold=False, ha="left", va="center", z=4, **kw):
    t = ax.text(x, y, s, fontsize=size, color=color, ha=ha, va=va, zorder=z, **kw)
    if bold:
        t.set_path_effects([withStroke(linewidth=size * 0.055, foreground=color)])
    return t


def rbox(x, y, w, h, fc="white", ec=BORDER, lw=1.2, r=0.8, ls="-", z=2):
    ax.add_patch(FancyBboxPatch((x, y), w, h,
                                boxstyle=f"round,pad=0,rounding_size={r}",
                                fc=fc, ec=ec, lw=lw, ls=ls, zorder=z))


def card(x, y, w, h, title, lines, color, title_size=11, body_size=9, lh=1.95):
    """白カード。lines は str か (str, color) のタプル。"""
    rbox(x, y, w, h, fc="white", ec=color, lw=1.6)
    T(x + 1.3, y + h - 1.7, title, size=title_size, color=color, bold=True)
    ax.plot([x + 1.3, x + w - 1.3], [y + h - 3.1, y + h - 3.1],
            color=color, lw=0.8, alpha=0.45, zorder=3)
    yy = y + h - 4.4
    for ln in lines:
        s, c = (ln if isinstance(ln, tuple) else (ln, INK))
        T(x + 1.5, yy, s, size=body_size, color=c)
        yy -= lh


def chip(x, y, w, h, text, fc, ec, tc, size=9, bold=True):
    rbox(x, y, w, h, fc=fc, ec=ec, lw=1.0, r=0.6, z=3)
    T(x + 1.2, y + h / 2, text, size=size, color=tc, bold=bold, z=4)


def arrow(p0, p1, color, lw=1.8, dashed=False, style="-|>", z=5):
    if dashed:
        # 破線は線と矢尻を分けて描く（FancyArrowPatch の破線は矢尻まで欠けるため）
        (x0, y0), (x1, y1) = p0, p1
        L = math.hypot(x1 - x0, y1 - y0)
        ux, uy = (x1 - x0) / L, (y1 - y0) / L
        xm, ym = x1 - ux * 1.6, y1 - uy * 1.6
        ax.plot([x0, xm], [y0, ym], color=color, lw=lw, ls=(0, (4, 2.2)), zorder=z,
                solid_capstyle="butt")
        ax.add_patch(FancyArrowPatch((xm, ym), p1, arrowstyle=style, mutation_scale=15,
                                     color=color, lw=lw, shrinkA=0, shrinkB=0, zorder=z))
    else:
        ax.add_patch(FancyArrowPatch(p0, p1, arrowstyle=style, mutation_scale=15,
                                     color=color, lw=lw, shrinkA=0, shrinkB=0, zorder=z))


def label(x, y, s, color, size=8.5, rot=0, ha="center", va="center"):
    T(x, y, s, size=size, color=color, bold=True, ha=ha, va=va, rotation=rot,
      rotation_mode="anchor", z=6,
      bbox=dict(boxstyle="round,pad=0.22", fc="white", ec="none", alpha=0.92))


# ============ タイトル ============
T(2, 86.9, "LINE予約システム 自作構成図", size=20, color=NAVY, bold=True, va="center")
T(48.5, 86.6, "STORES 予約 × LINE（LINEミニアプリ連携）の代替", size=11, color=MUTED, va="center")
# 右上バッジ
rbox(122, 84.2, 36, 5.2, fc=GREEN, ec=GREEN, r=1.0)
T(125.5, 87.4, "月額 0円", size=18, color="white", bold=True, va="center")
T(140, 87.6, "STORES版 14,190円〜 → 0円", size=9, color="white", bold=True, va="center")
T(140, 85.6, "通数を使うのは前日リマインドだけ", size=8, color="white", va="center")

T(2, 83.5, "構成：LINE公式アカウント ＋ LIFF（GitHub Pages） ＋ GAS ＋ スプレッドシート　／　"
           "審査なし・申請なし。LINE公式は無料のコミュニケーションプラン（月200通）で足りる",
  size=9.5, color=MUTED, va="center")
T(2, 81.2, "予約の流れ：① リッチメニュー → ② LIFFで予約 → ③ シートに記録・カレンダー・店へメール → "
           "④ 客の投稿を合図に → ⑤ 確認Flexを返信（無料） → ⑥ 前日18:00にリマインドをPush",
  size=9, color=INK, bold=True, va="center")
# 凡例
lx = 116
ax.plot([lx, lx + 3.2], [81.2, 81.2], color=NAVY, lw=1.8, zorder=4)
T(lx + 3.8, 81.2, "操作", size=8.5, color=MUTED, va="center")
lx = 124
ax.plot([lx, lx + 3.2], [81.2, 81.2], color=TEAL, lw=1.8, zorder=4)
T(lx + 3.8, 81.2, "Reply（無料）", size=8.5, color=MUTED, va="center")
lx = 137.5
ax.plot([lx, lx + 3.2], [81.2, 81.2], color=ORANGE, lw=1.8, ls=(0, (3, 1.5)), zorder=4)
T(lx + 3.8, 81.2, "Push（通数を使う）", size=8.5, color=MUTED, va="center")

# ============ レーン ============
lanes = [
    (2, 42, LANE_A, NAVY, "お客さま（スマホのLINEアプリ内で完結）"),
    (54, 42, LANE_B, GREEN, "LINE（すべて無料枠）"),
    (106, 52, LANE_C, BLUE, "Google（GAS／スプレッドシート・無料）"),
]
for x, w, fc, hc, name in lanes:
    rbox(x, 18.5, w, 60.5, fc=fc, ec="none", r=1.2, z=1)
    rbox(x, 75.5, w, 3.5, fc=hc, ec=hc, r=0.9, z=2)
    T(x + w / 2, 77.25, name, size=12, color="white", bold=True, ha="center", z=4)

# ---- レーンA：お客さま ----
card(5, 62, 28, 11.5, "LINE公式アカウントのトーク画面", [
    "・リッチメニュー「予約する」「予約確認」",
    "・あいさつメッセージにも予約導線",
    ("・確認Flex／前日リマインドがここに届く", NAVY),
], NAVY, body_size=8.8, lh=1.9)

card(5, 22, 36, 30, "LIFF 予約画面（LINEを閉じない・審査なし）", [], NAVY)
steps = [
    "STEP1　メニュー・担当スタッフを選ぶ",
    "STEP2　日程（週カレンダーの ○／×）",
    "STEP3　氏名・電話（2回目以降は自動補完）",
    "確認 →「予約を確定する」",
]
top = 45.9
for i, s in enumerate(steps):
    y = top - i * 4.4
    chip(6.5, y, 33, 3.0, s, CHIP, NAVY, NAVY)
    if i < len(steps) - 1:
        T(23, y - 0.75, "▼", size=7, color=NAVY, ha="center")
chip(6.5, 27.6, 33, 3.0, "予約確認タブ：自分の予約を見る／キャンセルする", "white", MUTED, INK, bold=False)
chip(6.5, 23.4, 33, 3.0, "置き場：GitHub Pages（静的HTML 1枚・0円）", CHIP_MUTED, CHIP_MUTED, MUTED, size=8.5, bold=False)

# ---- レーンB：LINE ----
rbox(55.5, 20.5, 39, 54, fc="none", ec=GREEN, lw=1.2, r=1.0, ls=(0, (5, 3)), z=2)
T(75, 22.4, "★ 3つとも同じプロバイダーに入れる（userIdが揃う → Pushが届く）",
  size=8.5, color=GREEN, bold=True, ha="center")

card(58, 63, 34, 10, "LINE公式アカウント", [
    "コミュニケーションプラン 0円（月200通まで）",
    "応答メッセージ OFF／あいさつメッセージ ON",
    ("通数を使うのは ⑥ の Push だけ", ORANGE),
], GREEN, body_size=8.8, lh=1.85)

card(58, 46, 34, 14, "Messaging API チャネル", [
    "Webhook → GAS（客の投稿・友だち追加を受ける）",
    ("Reply：確認Flex（無料・通数にカウントされない）", TEAL),
    ("Push：前日リマインド・再来店案内（1通ずつ消費）", ORANGE),
    "トークン・シークレットは GAS のプロパティに保管",
], GREEN, body_size=8.8, lh=1.9)

card(58, 24, 34, 15.5, "LINEログイン チャネル ＋ LIFF", [
    "LIFFエンドポイント＝GitHub Pages のURL",
    "scope：profile・openid（idTokenをGASで検証）",
    "友だち追加＝aggressive（外から来た人も友だちに）",
    "リンク先＝上のLINE公式アカウント",
], GREEN, body_size=8.8, lh=1.9)

# ---- レーンC：Google ----
card(109, 41, 46, 27, "GAS Webアプリ（doPost 1本で全部受ける）", [
    "予約API（LIFFから fetch）：getMaster／getSlots／",
    "　createBooking／getMyBookings／cancelBooking",
    "Webhook受け口：X-Line-Signature を検証してから処理",
    "createBooking：idToken検証 → LockService → 二重予約チェック → 書込",
    ("定時トリガー：毎日18:00 前日リマインド／毎月1日 再来店案内（90日未来店）", ORANGE),
    "デプロイ：ウェブアプリ・実行＝自分・アクセス＝全員（更新は「新バージョン」）",
    "秘密情報はスクリプトプロパティ。コードにも Git にも書かない",
], BLUE, body_size=8.8, lh=2.0)

card(109, 22, 21.75, 14, "スプレッドシート＝DB", [
    "7シート：設定／メニュー／スタッフ",
    "　　　　　休業／予約／顧客／ログ",
    "店側はここで予約を見る",
    "「来店」に変えると顧客シート更新",
    "個人情報はここだけ（共有は最小限）",
], BLUE, title_size=10.5, body_size=8.4, lh=1.75)

card(133.25, 22, 21.75, 14, "Googleカレンダー ＋ Gmail", [
    "カレンダー：スタッフ予定に自動登録",
    "　（スマホのカレンダーで見える）",
    "Gmail：店へ予約・キャンセル通知",
    "どちらも無料（メール100通/日）",
], BLUE, title_size=10.5, body_size=8.4, lh=1.75)

# ============ 矢印 ============
# ① トーク画面 → LIFF
arrow((19, 62), (19, 52), NAVY)
label(20.3, 57, "① タップで起動", NAVY, ha="left")

# ② LIFF → GAS（レーンBの隙間を通す）
arrow((41, 42.5), (109, 42.5), NAVY, lw=2.2)
label(75, 43.4, "② fetch：空き枠の取得／予約の登録（idToken付き）", NAVY, va="bottom")

# ③ GAS ↔ DB／カレンダー・メール
arrow((120, 41), (120, 36), BLUE, style="<|-|>")
label(121.3, 38.5, "③ 読み書き", BLUE, ha="left")
arrow((144.5, 41), (144.5, 36), BLUE)
label(145.8, 38.5, "③ 登録・通知", BLUE, ha="left")

# ④ 客の投稿 → Messaging API → Webhook → GAS
arrow((41, 50), (58, 52.5), TEAL)
ang4 = math.degrees(math.atan2(2.5, 17))
label(49.5, 48.3, "④ 客の投稿を1通\n（liff.sendMessages）", TEAL, rot=ang4)
arrow((92, 55), (109, 55), TEAL)
label(100.5, 55.9, "④ Webhook（署名検証）", TEAL, va="bottom")

# ⑤⑥ GAS → Messaging API（Reply／Push）→ トーク画面
arrow((109, 49), (92, 49), MUTED)
label(100.5, 48.1, "⑤⑥ Reply／Push API", MUTED, va="top")
ang56 = math.degrees(math.atan2(-8, 25))
arrow((58, 59), (33, 67), TEAL)
label(46.1, 64.9, "⑤ 確認Flex＝Reply（無料・通数外）", TEAL, rot=ang56)
arrow((58, 56), (33, 64), ORANGE, dashed=True)
label(44.9, 58.1, "⑥ 前日リマインド＝Push（1通／予約）", ORANGE, rot=ang56)

# ============ 下段：費用・作らないもの・落とし穴 ============
card(2, 1.5, 50, 12, "費用（月額）", [
    "STORES版：予約スモール 9,790円 ＋ LINEミニアプリ連携 4,400円 ＝ 14,190円〜",
    ("自作版：0円（GAS・スプレッドシート・GitHub Pages・LINE公式コミュニケーションプラン）", TEAL),
    "月200予約（＝リマインド200通）を超えたら LINE公式ライト 5,000円。それでも 1/3 以下",
    "本当のコスト＝保守は自分（APIの仕様変更・GAS障害に追随。保守料の根拠になる）",
], NAVY, title_size=10.5, body_size=8.4, lh=1.85)

card(55, 1.5, 49, 12, "作らないもの（要件に入るなら自作をやめて STORES）", [
    "・事前決済（現地決済のみ。要るなら Square／Stripe の決済リンク・手数料3.6%前後）",
    "・POSレジ連携",
    "・月謝・回数券の管理",
    "・設備／備品の在庫、SMS配信も対象外",
], RED, title_size=10.5, body_size=8.4, lh=1.85)

card(107, 1.5, 51, 12, "先に潰す落とし穴", [
    "① Messaging API と LINEログインは同じプロバイダーに入れる",
    "　（別だと userId がズレて Push が届かない）",
    "② そのOAが既に Lステップ等に紐づいていたら Webhook を持てない（共存不可）",
    "③ GASは「新しいデプロイ」でなく「バージョン更新」（URLが変わると Webhook が死ぬ）",
], ORANGE, title_size=10.5, body_size=8.4, lh=1.85)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
fig.savefig(OUT, dpi=200, facecolor="white")
print(f"font={FONT} -> {OUT}")
