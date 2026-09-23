# -*- coding: utf-8 -*-
"""注文住宅業界 LINEOA資料｜挿入画像を一括生成する

生成するもの（15枚）
  talk_01〜06      LINEトーク画面モック（あいさつ / Day0 / Day3 / Day6 / Day14 / 前日リマインド）
  popup_exit       離脱防止ポップアップ
  richmenu_tab1〜3 リッチメニュー 2500×1686px（3タブ）
  notif_preview    通知プレビュー（ロック画面風・4件）
  ad_cpf           LINE広告CPFクリエイティブ
  quiz_q1〜q4      60秒予算診断の設問画面

使い方:
    python -X utf8 _build/gen_line_images.py
出力先:
    C:\\Users\\<user>\\Downloads\\注文住宅_LINE画像素材\\

方針:
  - 生成AIは使わない。日本語と寸法が崩れるため、すべてPillowで描く
  - フォントはメイリオ。影・グラデーションは使わずフラットに
  - 文面はスライド本体（_build/build_chumon_jutaku.py）と同一。改行位置を厳守する
"""
import os
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

OUT = Path(os.path.expanduser("~")) / "Downloads" / "注文住宅_LINE画像素材"
OUT.mkdir(parents=True, exist_ok=True)

# ---------- 配色（スライド本体と同じ） ----------
NAVY = "#1F285A"
INK = "#333333"
MUT = "#7F7F7F"
RED = "#C00000"
ORANGE = "#ED7D31"
GREEN = "#06C755"     # LINE UI のみ
WHITE = "#FFFFFF"
PALE = "#F4F7FF"
GREY = "#F2F2F2"
BORDER = "#D9D9D9"
LINEBG = "#7494C0"    # トーク画面の背景
HEADBG = "#364764"    # トーク画面のヘッダー

# ---------- フォント ----------
FONTS = [
    ("C:/Windows/Fonts/meiryo.ttc", "C:/Windows/Fonts/meiryob.ttc"),
    ("C:/Windows/Fonts/YuGothR.ttc", "C:/Windows/Fonts/YuGothB.ttc"),
    ("C:/Windows/Fonts/msgothic.ttc", "C:/Windows/Fonts/msgothic.ttc"),
]
_REG = _BLD = None
for reg, bld in FONTS:
    if Path(reg).exists():
        _REG, _BLD = reg, (bld if Path(bld).exists() else reg)
        break
if _REG is None:
    raise SystemExit("日本語フォントが見つかりません（メイリオ / 游ゴシック / MSゴシック）")

_cache = {}


def F(size, bold=False):
    key = (size, bold)
    if key not in _cache:
        _cache[key] = ImageFont.truetype(_BLD if bold else _REG, size)
    return _cache[key]


# ---------- 描画ヘルパー ----------
def tw(d, text, f):
    """テキストの描画幅"""
    if not text:
        return 0
    b = d.textbbox((0, 0), text, font=f)
    return b[2] - b[0]


def center(d, cx, y, text, f, fill):
    d.text((cx - tw(d, text, f) / 2, y), text, font=f, fill=fill)


def wrap(d, text, f, maxw):
    """maxw（px）に収まるよう文字単位で折り返す"""
    out, cur = [], ""
    for ch in text:
        if tw(d, cur + ch, f) <= maxw:
            cur += ch
        else:
            out.append(cur)
            cur = ch
    if cur:
        out.append(cur)
    return out or [""]


def rrect(d, box, r, fill=None, outline=None, width=1):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=width)


def save(img, name):
    p = OUT / name
    img.save(p)
    print(f"  {name}  {img.width}x{img.height}")
    return p


# ============================================================
# 1. LINEトーク画面モック
# ============================================================
W_TALK = 1000
HEAD_H = 116
PAD_TOP = 34          # ヘッダー下の余白
AV = 72               # アイコン直径
AV_X = 34
BUB_X = AV_X + AV + 26
BUB_R = 22
BUB_PADX, BUB_PADY = 30, 24
LINE_H = 44
FS_BODY = 28
GAP = 34              # 吹き出し間
BOTTOM = 40


def talk_mock(name, title, bubbles):
    """bubbles = [吹き出し文字列, ...]。改行は文字列内の \n をそのまま使う"""
    tmp = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    fb = F(FS_BODY)
    bub_w = W_TALK - BUB_X - 70

    laid = []
    for b in bubbles:
        lines = []
        for ln in b.split("\n"):
            lines.extend(wrap(tmp, ln, fb, bub_w - BUB_PADX * 2) if ln else [""])
        laid.append((lines, len(lines) * LINE_H + BUB_PADY * 2))

    h = HEAD_H + PAD_TOP + sum(x[1] for x in laid) + GAP * (len(laid) - 1) + BOTTOM
    img = Image.new("RGB", (W_TALK, h), LINEBG)
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W_TALK, HEAD_H], fill=HEADBG)
    d.text((AV_X, (HEAD_H - 34) // 2), title, font=F(30, True), fill=WHITE)

    y = HEAD_H + PAD_TOP
    for lines, bh in laid:
        d.ellipse([AV_X, y, AV_X + AV, y + AV], fill=GREEN)
        rrect(d, [BUB_X, y, BUB_X + bub_w, y + bh], BUB_R, fill=WHITE)
        ty = y + BUB_PADY
        for ln in lines:
            if ln:
                d.text((BUB_X + BUB_PADX, ty), ln, font=fb, fill=INK)
            ty += LINE_H
        y += bh + GAP
    return save(img, name)


TALKS = [
    ("talk_01_greeting.png", "〇〇ホーム 家づくり相談", [
        "友だち追加ありがとうございます。\n〇〇ホームの家づくり相談窓口です。",
        "このアカウントから\nお電話することはありません。\nご質問はこのトークにどうぞ。",
        "担当：〇〇（一級建築士）\n有人対応 平日10:00〜18:00\nそれ以外は自動でご返信します。",
        "はじめに、30秒のアンケートに\nお答えいただけますか？\nご関心に合った情報だけお送りします。",
    ]),
    ("talk_02_day0.png", "〇〇ホーム 家づくり相談", [
        "【はじめまして】家づくりのお金の話\n\n〇〇ホームの〇〇です。\n"
        "家づくりで最初につまずくのは\n「結局いくらかかるのか」です。",
        "▼ 総額の内訳（目安）\n・建物本体　　　約70％\n・付帯工事　　　約20％\n・諸費用　　　　約10％\n\n"
        "広告の「本体価格」には\n外構も地盤改良も入っていません。\n\n"
        "明日は「建てた人が後悔したこと」を\nお送りします。",
    ]),
    ("talk_03_day3.png", "〇〇ホーム 家づくり相談", [
        "【土地から探す方へ】\n先に知っておくと損しない3つのこと\n\n"
        "土地を先に買ってしまうと\n建物の予算が足りなくなる\nーーこれがいちばん多い失敗です。",
        "① 土地と建物は同時に考える\n② 「建築条件付き」は\n　 建てる会社が決まっている\n"
        "③ 地盤改良で100万円前後\n　 かかることがある\n\n"
        "土地探しからご一緒もできます。\nこのトークに「土地」と\n送っていただければご返信します。",
    ]),
    ("talk_04_day6.png", "〇〇ホーム 家づくり相談", [
        "【ご案内】1組貸切の完成見学会\n\n実際にお住まいになるお宅を\n1日1組だけご案内しています。\n\n"
        "・貸切です。ほかのお客様はいません\n・その場で契約のお話はしません\n・見学後の営業のお電話もしません",
        "見ていただきたいのは\n写真では分からない天井の高さと\n冬の朝の暖かさです。\n\n"
        "▼ ご希望の日程をお選びください\n［ 今週末 ］［ 来週末 ］\n［ 平日に相談したい ］\n［ 今回は都合が合わない ］",
    ]),
    ("talk_05_day14.png", "〇〇ホーム 家づくり相談", [
        "【一度だけ】間取りのご相談を承ります\n\n2週間おつきあいいただき\nありがとうございました。\n\n"
        "ここまでお読みの方に一度だけ\nご案内させてください。",
        "ご希望をうかがって\n間取りのラフをお作りします（無料）。\nオンラインでも構いません。\n\n"
        "・所要45分／ご夫婦だけでも可\n・その場で契約のお話はしません\n\n"
        "［ 相談したい ］\n［ もう少し情報だけ受け取りたい ］",
    ]),
    ("talk_06_reminder.png", "前日リマインド", [
        "【明日のご来場】〇〇様\n\n明日14:00にお待ちしています。\n所要は約60分、駐車場はございます。\n\n"
        "お子さま連れも歓迎です。\nキッズスペースをご用意しています。",
        "［ 場所を確認する ］\n［ 日程を変更する ］\n［ キャンセルする ］",
    ]),
]


# ============================================================
# 2. 離脱防止ポップアップ
# ============================================================
def popup_exit():
    W, H = 1040, 800
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    rrect(d, [40, 40, W - 40, H - 40], 26, fill=PALE, outline=NAVY, width=4)

    center(d, W // 2, 116, "お帰りの前に、資金計画の目安だけ", F(44, True), NAVY)
    center(d, W // 2, 184, "受け取っていきませんか？", F(44, True), NAVY)

    center(d, W // 2, 300, "土地・建物・諸費用をあわせた総額の目安がわかる資料です。", F(28), INK)
    center(d, W // 2, 366, "こちらからお電話することはありません。", F(30, True), RED)

    bx0, bx1, by0, by1 = 250, W - 250, 470, 594
    rrect(d, [bx0, by0, bx1, by1], 30, fill=GREEN)
    f = F(40, True)
    d.text(((bx0 + bx1) / 2 - tw(d, "LINEで受け取る", f) / 2, by0 + 36), "LINEで受け取る", font=f, fill=WHITE)

    center(d, W // 2, 656, "［ 今はいい ］", F(28), MUT)
    return save(img, "popup_exit.png")


# ============================================================
# 3. リッチメニュー（2500×1686px・3列×2行＝LINEの6分割標準テンプレート）
# ============================================================
RM_W, RM_H = 2500, 1686
RM_COLS, RM_ROWS = 3, 2


def icon(d, cx, cy, s, kind, col, w=9, bg=WHITE):
    """線画アイコン。s = 外接半径。図形のみで描く（絵文字・外部素材は使わない）
    bg = 背景色。塗りつぶしを白抜きする箇所で使う"""
    L, T, R, B = cx - s, cy - s, cx + s, cy + s
    if kind == "yen":                                   # 円＋¥
        d.ellipse([L, T, R, B], outline=col, width=w)
        f = F(int(s * 1.5), True)
        d.text((cx - tw(d, "¥", f) / 2, cy - s * 0.78), "¥", font=f, fill=col)
    elif kind == "home":                                # 家
        d.polygon([(cx, T), (R, cy - s * 0.1), (L, cy - s * 0.1)], outline=col, width=w)
        d.rectangle([L + s * 0.28, cy - s * 0.1, R - s * 0.28, B], outline=col, width=w)
        d.rectangle([cx - s * 0.2, cy + s * 0.35, cx + s * 0.2, B], outline=col, width=w)
    elif kind == "donut":                               # ドーナツグラフ
        d.ellipse([L, T, R, B], outline=col, width=w)
        d.ellipse([cx - s * 0.54, cy - s * 0.54, cx + s * 0.54, cy + s * 0.54], outline=col, width=w)
        d.line([cx, T, cx, cy - s * 0.54], fill=col, width=w)          # 12時
        d.line([cx + s * 0.54, cy, R, cy], fill=col, width=w)          # 3時
        d.line([cx - s * 0.4, cy + s * 0.4, cx - s * 0.72, cy + s * 0.72], fill=col, width=w)
    elif kind == "pin":                                 # 地図ピン
        hr = s * 0.72                                   # 頭部は正円で描く
        d.polygon([(cx - hr * 0.62, T + hr * 1.5), (cx + hr * 0.62, T + hr * 1.5), (cx, B)], fill=col)
        d.ellipse([cx - hr, T, cx + hr, T + hr * 2], fill=col)
        d.ellipse([cx - hr + w, T + w, cx + hr - w, T + hr * 2 - w], fill=bg)
        d.ellipse([cx - hr * 0.34, T + hr * 0.66, cx + hr * 0.34, T + hr * 1.34], fill=col)
    elif kind == "help":                                # ？
        d.ellipse([L, T, R, B], outline=col, width=w)
        f = F(int(s * 1.5), True)
        d.text((cx - tw(d, "?", f) / 2, cy - s * 0.82), "?", font=f, fill=col)
    elif kind == "chat":                                # 吹き出し
        rrect(d, [L, T, R, cy + s * 0.42], int(s * 0.3), outline=col, width=w)
        d.polygon([(cx - s * 0.42, cy + s * 0.36), (cx - s * 0.06, cy + s * 0.36), (cx - s * 0.5, B)],
                  fill=col)
    elif kind == "family":                              # 親＋子
        d.ellipse([cx - s * 0.86, T + s * 0.1, cx - s * 0.16, T + s * 0.8], outline=col, width=w)
        d.arc([cx - s * 1.0, cy - s * 0.1, cx, B], 180, 360, fill=col, width=w)
        d.ellipse([cx + s * 0.24, T + s * 0.42, cx + s * 0.8, T + s * 0.98], outline=col, width=w)
        d.arc([cx + s * 0.14, cy + s * 0.18, cx + s * 0.9, B + s * 0.2], 180, 360, fill=col, width=w)
    elif kind == "doc":                                 # 書類
        d.rectangle([L + s * 0.2, T, R - s * 0.2, B], outline=col, width=w)
        for k in range(3):
            y = T + s * (0.55 + k * 0.45)
            d.line([L + s * 0.46, y, R - s * 0.46, y], fill=col, width=w)
    elif kind == "bars":                                # 棒グラフ
        d.line([L, B, R, B], fill=col, width=w)
        for k, h in enumerate((0.45, 0.85, 1.25)):
            x = L + s * (0.35 + k * 0.62)
            d.rectangle([x, B - s * h, x + s * 0.4, B], outline=col, width=w)
    elif kind == "calc":                                # 電卓
        rrect(d, [L + s * 0.24, T, R - s * 0.24, B], int(s * 0.16), outline=col, width=w)
        d.rectangle([L + s * 0.48, T + s * 0.26, R - s * 0.48, T + s * 0.72], outline=col, width=w)
        for r in range(2):
            for c in range(3):
                x = L + s * (0.52 + c * 0.42)
                y = T + s * (1.06 + r * 0.44)
                d.ellipse([x, y, x + s * 0.16, y + s * 0.16], fill=col)
    elif kind == "calendar":                            # カレンダー
        d.rectangle([L, T + s * 0.24, R, B], outline=col, width=w)
        d.line([L, T + s * 0.78, R, T + s * 0.78], fill=col, width=w)
        d.line([L + s * 0.42, T, L + s * 0.42, T + s * 0.46], fill=col, width=w)
        d.line([R - s * 0.42, T, R - s * 0.42, T + s * 0.46], fill=col, width=w)
        d.ellipse([cx - s * 0.16, cy + s * 0.12, cx + s * 0.16, cy + s * 0.44], fill=col)
    elif kind == "video":                               # オンライン相談
        rrect(d, [L, T + s * 0.3, cx + s * 0.34, B - s * 0.3], int(s * 0.18), outline=col, width=w)
        d.polygon([(cx + s * 0.44, cy - s * 0.32), (R, cy - s * 0.6),
                   (R, cy + s * 0.6), (cx + s * 0.44, cy + s * 0.32)], outline=col, width=w)
    elif kind == "gift":                                # 来場特典
        d.rectangle([L, cy - s * 0.28, R, B], outline=col, width=w)
        d.rectangle([L - s * 0.08, cy - s * 0.66, R + s * 0.08, cy - s * 0.24], outline=col, width=w)
        d.line([cx, cy - s * 0.66, cx, B], fill=col, width=w)
        d.arc([L + s * 0.1, T, cx, cy - s * 0.24], 180, 360, fill=col, width=w)
        d.arc([cx, T, R - s * 0.1, cy - s * 0.24], 180, 360, fill=col, width=w)
    elif kind == "clock":                               # 日程変更
        d.ellipse([L, T, R, B], outline=col, width=w)
        d.line([cx, cy, cx, cy - s * 0.55], fill=col, width=w)
        d.line([cx, cy, cx + s * 0.42, cy + s * 0.16], fill=col, width=w)
    elif kind == "stopwatch":                           # 60秒診断
        d.ellipse([L, T + s * 0.2, R, B], outline=col, width=w)
        d.line([cx - s * 0.3, T, cx + s * 0.3, T], fill=col, width=w)
        d.line([cx, T, cx, T + s * 0.22], fill=col, width=w)
        d.line([cx, cy + s * 0.1, cx, cy - s * 0.42], fill=col, width=w)
        d.line([cx, cy + s * 0.1, cx + s * 0.4, cy + s * 0.3], fill=col, width=w)


TABS = [
    ("richmenu_tab1.png", NAVY, "タブ① これから考える", [
        ("60秒 予算診断", "stopwatch"), ("建築実例を見る", "home"), ("費用の内訳を知る", "donut"),
        ("土地探しのコツ", "pin"), ("よくある質問", "help"), ("相談してみる", "chat"),
    ]),
    ("richmenu_tab2.png", ORANGE, "タブ② 資金を調べる", [
        ("住宅ローン相談", "yen"), ("親からの援助・贈与の話", "family"), ("使える補助金", "doc"),
        ("月々の返済シミュレーション", "bars"), ("土地＋建物の総額目安", "calc"), ("相談してみる", "chat"),
    ]),
    ("richmenu_tab3.png", NAVY, "タブ③ 見に行く", [
        ("完成見学会に申し込む", "home"), ("個別相談を予約する", "calendar"), ("オンライン相談", "video"),
        ("展示場の場所", "pin"), ("来場特典を見る", "gift"), ("日程を変更する", "clock"),
    ]),
]


# タブ色 → (アイコンチップの薄色, 主役マスのバッジ文字色)
TINT = {NAVY: "#E8EDF7", ORANGE: "#FCE9DA"}
GRID = "#E3E8F0"   # マス間の区切り（背景を透かせて細い線に見せる）
SEP = 7            # 区切りの太さ(px)


def richmenu(name, accent, _tab, cells):
    img = Image.new("RGB", (RM_W, RM_H), GRID)
    d = ImageDraw.Draw(img)
    xs = [round(RM_W * i / RM_COLS) for i in range(RM_COLS + 1)]
    ys = [round(RM_H * i / RM_ROWS) for i in range(RM_ROWS + 1)]
    tint = TINT.get(accent, PALE)

    for i, (label, kind) in enumerate(cells):
        c, r = i % RM_COLS, i // RM_COLS
        x0, x1, y0, y1 = xs[c], xs[c + 1], ys[r], ys[r + 1]
        # 区切り分だけ内側に詰める（外周は詰めない）
        ix0 = x0 + (SEP if c else 0)
        iy0 = y0 + (SEP if r else 0)
        cw, ch = x1 - ix0, y1 - iy0
        main = (i == 0)                      # 左上＝各タブの主役CTA

        if main:                             # 主役はタブ色でベタ塗り。一番強く見せる
            d.rectangle([ix0, iy0, x1, y1], fill=accent)
            icol, tcol, chip = WHITE, WHITE, None
        else:
            d.rectangle([ix0, iy0, x1, y1], fill=WHITE)
            icol, tcol, chip = accent, INK, tint

        icx, icy = ix0 + cw / 2, iy0 + ch * 0.36
        s = 78 if main else 68
        if chip:                             # アイコンを薄色の円チップに乗せる
            rad = s * 1.62
            d.ellipse([icx - rad, icy - rad, icx + rad, icy + rad], fill=chip)
        icon(d, icx, icy, s, kind, icol, w=10 if main else 8,
             bg=accent if main else chip)

        f = F(52 if main else 46, main)
        lines = wrap(d, label, f, cw - 90)
        lh = 68 if main else 62
        ty = iy0 + ch * 0.66
        for ln in lines:
            d.text((ix0 + (cw - tw(d, ln, f)) / 2, ty), ln, font=f, fill=tcol)
            ty += lh

        if main:                             # 主役だけバッジを付ける（白地にタブ色）
            bw, bh = 190, 52
            bx, by = ix0 + 30, iy0 + 30
            rrect(d, [bx, by, bx + bw, by + bh], bh // 2, fill=WHITE)
            fb = F(26, True)
            d.text((bx + (bw - tw(d, "最重要", fb)) / 2, by + 10), "最重要", font=fb, fill=accent)
    return save(img, name)


# ============================================================
# 4. 通知プレビュー（ロック画面風・冒頭15文字）
# ============================================================
NOTIFS = [
    ("Day 0", "【はじめまして】家づくりのお金"),
    ("Day 3", "【土地から探す方へ】先に知って"),
    ("Day 6", "【ご案内】1組貸切の完成見学会"),
    ("Day 14", "【一度だけ】間取りのご相談を承"),
]


def notif_preview():
    W = 1040
    card_h, gap, top = 156, 26, 40
    H = top + (card_h + gap) * len(NOTIFS) + 20
    img = Image.new("RGB", (W, H), "#4A5568")
    d = ImageDraw.Draw(img)
    y = top
    for day, text in NOTIFS:
        rrect(d, [36, y, W - 36, y + card_h], 24, fill=WHITE)
        d.ellipse([66, y + 34, 66 + 60, y + 94], fill=GREEN)
        d.text((150, y + 30), "〇〇ホーム 家づくり相談", font=F(26, True), fill=INK)
        f = F(24)
        d.text((W - 36 - 30 - tw(d, day, f), y + 32), day, font=f, fill=MUT)
        d.text((150, y + 78), text, font=F(30), fill=INK)
        y += card_h + gap
    return save(img, "notif_preview.png")


# ============================================================
# 5. LINE広告 CPFクリエイティブ
# ============================================================
def ad_cpf():
    W, H = 1000, 1000
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 430], fill=GREY)
    center(d, W // 2, 196, "［ 実例写真をここに差し替え ］", F(28), MUT)
    d.rectangle([0, 428, W, 434], fill=BORDER)

    center(d, W // 2, 486, "その土地、建てられますか", F(54, True), NAVY)

    f = F(28)
    body = "60秒の質問に答えるだけ。土地・建物・諸費用の総額目安をLINEでお送りします。"
    ty = 592
    for ln in wrap(d, body, f, W - 160):
        center(d, W // 2, ty, ln, f, INK)
        ty += 44

    bx0, bx1, by0, by1 = 190, W - 190, 780, 900
    rrect(d, [bx0, by0, bx1, by1], 28, fill=GREEN)
    fb = F(38, True)
    d.text(((bx0 + bx1) / 2 - tw(d, "友だち追加して受け取る", fb) / 2, by0 + 34),
           "友だち追加して受け取る", font=fb, fill=WHITE)
    return save(img, "ad_cpf.png")


# ============================================================
# 6. 60秒予算診断の設問画面
# ============================================================
QUIZ = [
    ("quiz_q1.png", "Q1", "建てたい時期は？",
     ["1年以内", "1〜2年", "2〜3年", "まだ決めていない"], "検討時期タグ"),
    ("quiz_q2.png", "Q2", "土地はお決まりですか？",
     ["所有している", "候補がある", "これから探す"], "土地有無タグ"),
    ("quiz_q3.png", "Q3", "ご予算の目安は？",
     ["〜2,500万", "2,500〜3,500万", "3,500万〜", "まだ分からない"], "予算帯タグ"),
    ("quiz_q4.png", "Q4", "いちばん気になることは？",
     ["費用", "間取り", "土地", "性能・断熱", "補助金"], "関心テーマタグ"),
]


def quiz(name, no, question, options, tag):
    W = 1000
    tmp = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    fq = F(30)
    bub_w = W - BUB_X - 70
    qlines = wrap(tmp, question, fq, bub_w - BUB_PADX * 2)
    bub_h = (len(qlines) + 1) * 48 + BUB_PADY * 2

    pill_h, pill_gap = 84, 22
    H = HEAD_H + PAD_TOP + bub_h + 40 + (pill_h + pill_gap) * len(options) + 110

    img = Image.new("RGB", (W, H), LINEBG)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, HEAD_H], fill=HEADBG)
    d.text((AV_X, (HEAD_H - 34) // 2), "〇〇ホーム 家づくり相談", font=F(30, True), fill=WHITE)

    y = HEAD_H + PAD_TOP
    d.ellipse([AV_X, y, AV_X + AV, y + AV], fill=GREEN)
    rrect(d, [BUB_X, y, BUB_X + bub_w, y + bub_h], BUB_R, fill=WHITE)
    ty = y + BUB_PADY
    d.text((BUB_X + BUB_PADX, ty), f"{no}／全4問", font=F(26, True), fill=NAVY)
    ty += 48
    for ln in qlines:
        d.text((BUB_X + BUB_PADX, ty), ln, font=fq, fill=INK)
        ty += 48

    # クイックリプライ（選択式・自由入力なし）
    y += bub_h + 40
    fo = F(28, True)
    for op in options:
        w = tw(d, op, fo) + 96
        x0 = W - 70 - w
        rrect(d, [x0, y, x0 + w, y + pill_h], pill_h // 2, fill=WHITE, outline=GREEN, width=4)
        d.text((x0 + 48, y + (pill_h - 34) / 2), op, font=fo, fill=GREEN)
        y += pill_h + pill_gap

    f = F(24)
    d.text((W - 70 - tw(d, f"→ 回答すると「{tag}」が付きます", f), y + 20),
           f"→ 回答すると「{tag}」が付きます", font=f, fill=WHITE)
    return save(img, name)


# ============================================================
def main():
    print("LINEトーク画面")
    for name, title, bubbles in TALKS:
        talk_mock(name, title, bubbles)
    print("離脱防止ポップアップ")
    popup_exit()
    print("リッチメニュー")
    for name, accent, tab, cells in TABS:
        richmenu(name, accent, tab, cells)
    print("通知プレビュー")
    notif_preview()
    print("LINE広告CPF")
    ad_cpf()
    print("60秒予算診断")
    for name, no, q, opts, tag in QUIZ:
        quiz(name, no, q, opts, tag)
    print(f"\n出力先: {OUT}")


if __name__ == "__main__":
    main()
