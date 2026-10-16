"""インクアート 11/2 オリジナル卓上カレンダー（プラケース）の配信バナー（1040×1040）を組む。

文字は画像生成に描かせず、ここで入れる（明朝・ゴシックで一字も崩れない）。
写真はサイトの商品画像2枚（殿村さんが 2026-10-16 に支給）。元の画像に乗っている文字
（「ハガキ サイズ」「カレンダーの台紙は7枚〜15枚までお選びいただけます。」）は消して、絵だけ使う。

使い方:
  python3 _build/build_inkart_1102_calendar_banner.py [出力]
"""
import sys
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

S = 2  # 2倍で描いて最後に縮める（文字の縁をなめらかにする）
W = H = 1040

GREEN = (0x1C, 0x8B, 0x6A)   # HPの緑
ORANGE = (0xFF, 0x88, 0x00)  # HPのオレンジ
INK = (0x33, 0x33, 0x33)
GOLD = (0xB8, 0x97, 0x5A)
TINT = (0xEE, 0xF6, 0xF2)    # 緑のごく薄い色（値段の面）
WHITE = (255, 255, 255)

CASE_PHOTO = "_data/printearth/site_calendar_table_case_20261016.png"     # ケースに入った卓上カレンダー
SHEETS_PHOTO = "_data/printearth/site_calendar_table_sheets_20261016.png"  # 台紙を広げた写真

FONT_DIR = "/usr/share/fonts/opentype/noto/"
SERIF_B = FONT_DIR + "NotoSerifCJK-Bold.ttc"
SANS_B = FONT_DIR + "NotoSansCJK-Bold.ttc"


def font(path, size):
    return ImageFont.truetype(path, size * S, index=0)  # index 0 = JP


def text_width(f, text, tracking=0):
    return sum(f.getlength(ch) for ch in text) + tracking * S * (len(text) - 1)


def text_height(f, text):
    b = f.getbbox(text)
    return (b[3] - b[1]) / S


def draw_text(d, text, f, x, top, fill, tracking=0):
    """x, top は 1040 基準。top は文字の上端。"""
    t = f.getbbox(text)[1]
    cx = x * S
    for ch in text:
        d.text((cx, top * S - t), ch, font=f, fill=fill)
        cx += f.getlength(ch) + tracking * S


def draw_text_center(d, text, f, cx, top, fill, tracking=0):
    w = text_width(f, text, tracking) / S
    draw_text(d, text, f, cx - w / 2, top, fill, tracking)


def runs_width(runs):
    return sum(text_width(f, t, tr) for t, f, _, tr in runs) / S


def draw_runs(d, runs, x, baseline):
    """大きさの違う文字を、下の線（ベースライン）をそろえて横に並べる。"""
    cx = x * S
    for t, f, fill, tr in runs:
        for ch in t:
            d.text((cx, baseline * S), ch, font=f, fill=fill, anchor="ls")
            cx += f.getlength(ch) + tr * S


def rrect(d, box, r, fill=None, outline=None, width=1):
    x1, y1, x2, y2 = box
    d.rounded_rectangle((x1 * S, y1 * S, x2 * S, y2 * S), radius=r * S,
                        fill=fill, outline=outline, width=int(width * S))


def erase(im, x1, y1, x2, y2):
    """(x1,y1)-(x2,y2) の中を、外周の色から面でなめらかにつないで塗り直す（背景がグラデーションの所の文字消し）。"""
    a = np.asarray(im).astype(float)
    top, bot = a[y1 - 1, x1:x2 + 1], a[y2 + 1, x1:x2 + 1]
    lef, rig = a[y1:y2 + 1, x1 - 1], a[y1:y2 + 1, x2 + 1]
    c00, c01 = a[y1 - 1, x1 - 1], a[y1 - 1, x2 + 1]
    c10, c11 = a[y2 + 1, x1 - 1], a[y2 + 1, x2 + 1]
    h, w = y2 - y1 + 1, x2 - x1 + 1
    v = (np.arange(h) + 1) / (h + 1)
    u = (np.arange(w) + 1) / (w + 1)
    V, U = v[:, None, None], u[None, :, None]
    patch = ((1 - V) * top[None] + V * bot[None] + (1 - U) * lef[:, None] + U * rig[:, None]
             - ((1 - U) * (1 - V) * c00 + U * (1 - V) * c01 + (1 - U) * V * c10 + U * V * c11))
    a[y1:y2 + 1, x1:x2 + 1] = patch
    return Image.fromarray(np.clip(a, 0, 255).astype(np.uint8))


def load_case():
    im = Image.open(CASE_PHOTO).convert("RGB")
    im = erase(im, 2, 190, 69, 206)     # 左下「ハガキ サイズ」
    a = np.asarray(im).copy()
    a[190:207, 0:2] = a[190:207, 2:3]   # 左端の2列に残った文字のかけらも消す
    return Image.fromarray(a)


def load_sheets():
    im = Image.open(SHEETS_PHOTO).convert("RGB")
    return erase(im, 12, 6, 122, 56)    # 左上「カレンダーの台紙は…お選びいただけます。」


def cover(im, w, h):
    """枠いっぱいに入るよう中央で切り抜いて拡大・縮小する。"""
    r = w / h
    iw, ih = im.size
    if iw / ih > r:
        nw = int(ih * r)
        im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
    else:
        nh = int(iw / r)
        im = im.crop((0, (ih - nh) // 2, iw, (ih - nh) // 2 + nh))
    im = im.resize((int(w), int(h)), Image.LANCZOS)
    return im.filter(ImageFilter.UnsharpMask(radius=2, percent=60, threshold=2))  # 元が小さいので少し締める


def paste_rounded(canvas, im, box, r):
    x1, y1, x2, y2 = (int(round(v * S)) for v in box)
    im = cover(im, x2 - x1, y2 - y1)
    mask = Image.new("L", im.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, im.size[0] - 1, im.size[1] - 1),
                                           radius=r * S, fill=255)
    canvas.paste(im, (x1, y1), mask)


def build(out):
    c = Image.new("RGB", (W * S, H * S), WHITE)
    d = ImageDraw.Draw(c)

    # 1. 上の緑の帯：誰の何の用事か
    d.rectangle((0, 0, W * S, 84 * S), fill=GREEN)
    f_top = font(SANS_B, 36)
    t = "ノベルティや、年末年始のご挨拶に"
    draw_text_center(d, t, f_top, W / 2, 42 - text_height(f_top, t) / 2, WHITE, 3)

    # 2. 主役（明朝）：2027年／オリジナル卓上カレンダー
    f_year = font(SERIF_B, 58)
    f_hero = font(SERIF_B, 76)
    y = 108
    yw = text_width(f_year, "2027年", 4) / S
    draw_text_center(d, "2027年", f_year, W / 2, y, GREEN, 4)
    ly = y + text_height(f_year, "2027年") / 2
    d.line(((W / 2 - yw / 2 - 150) * S, ly * S, (W / 2 - yw / 2 - 24) * S, ly * S), fill=GOLD, width=int(1.5 * S))
    d.line(((W / 2 + yw / 2 + 24) * S, ly * S, (W / 2 + yw / 2 + 150) * S, ly * S), fill=GOLD, width=int(1.5 * S))
    y += text_height(f_year, "2027年") + 22
    draw_text_center(d, "オリジナル卓上カレンダー", f_hero, W / 2, y, GREEN, 2)
    y += text_height(f_hero, "オリジナル卓上カレンダー") + 30

    # 3. 写真2枚（左＝ケースに入った卓上カレンダー、右＝台紙を広げた写真）
    top, bottom = y, y + 302
    lbox = (40, top, 514, bottom)
    rbox = (526, top, 1000, bottom)
    paste_rounded(c, load_case(), lbox, 8)
    paste_rounded(c, load_sheets(), rbox, 8)
    for bx in (lbox, rbox):
        rrect(d, bx, 8, outline=(0xDD, 0xDD, 0xDD), width=1)
    #   台紙は7〜15枚（右の写真の左上。元の文字を消した場所）
    f_lab = font(SANS_B, 28)
    lab = "台紙は7〜15枚"
    lw = text_width(f_lab, lab, 1) / S + 36
    lx1, ly1 = rbox[0] + 12, top + 12
    rrect(d, (lx1, ly1, lx1 + lw, ly1 + 50), 4, fill=GREEN)
    draw_text(d, lab, f_lab, lx1 + 18, ly1 + 25 - text_height(f_lab, lab) / 2, WHITE, 1)

    # 4. 30個から（オレンジの丸）＋ ケースに入った状態で届くので、すぐに配れます
    y = bottom + 24
    r = 92
    ccx, ccy = 40 + r, y + r
    d.ellipse(((ccx - r) * S, (ccy - r) * S, (ccx + r) * S, (ccy + r) * S), fill=ORANGE)
    f_num = font(SANS_B, 76)
    f_nuk = font(SANS_B, 40)
    f_kara = font(SANS_B, 36)
    num_runs = [("30", f_num, WHITE, 0), ("個", f_nuk, WHITE, 0)]
    nw = runs_width(num_runs)
    draw_runs(d, num_runs, ccx - nw / 2, ccy + 14)
    draw_text_center(d, "から", f_kara, ccx, ccy + 26, WHITE, 2)

    f_mid1 = font(SANS_B, 38)
    f_mid2 = font(SERIF_B, 62)
    tx = 40 + 2 * r + 36
    m1 = "ケースに入った状態で届くので、"
    m2 = "すぐに配れます"
    h1, h2 = text_height(f_mid1, m1), text_height(f_mid2, m2)
    ty = ccy - (h1 + 22 + h2) / 2
    draw_text(d, m1, f_mid1, tx, ty, INK, 1)
    draw_text(d, m2, f_mid2, tx, ty + h1 + 22, GREEN, 4)

    # 5. 参考価格（テキストと同じ言葉）
    py1, py2 = ccy + r + 22, ccy + r + 22 + 76
    rrect(d, (40, py1, 1000, py2), 8, fill=TINT)
    f_chip = font(SANS_B, 26)
    chip = "参考価格"
    cw = text_width(f_chip, chip, 2) / S + 28
    f_p = font(SANS_B, 30)
    f_pn = font(SANS_B, 48)
    price = [("ハガキサイズ・13枚入り　100個 ", f_p, INK, 0), ("34,920", f_pn, GREEN, 0),
             ("円～（税別）", f_p, INK, 0)]
    total = cw + 20 + runs_width(price)
    px = (W - total) / 2
    pcy = (py1 + py2) / 2
    rrect(d, (px, pcy - 22, px + cw, pcy + 22), 4, fill=GREEN)
    draw_text(d, chip, f_chip, px + 14, pcy - text_height(f_chip, chip) / 2, WHITE, 2)
    draw_runs(d, price, px + cw + 20, pcy + 16)

    # 6. ボタン（平らなオレンジ）
    bx1, by1, bx2, by2 = 70, py2 + 24, 970, py2 + 24 + 84
    rrect(d, (bx1, by1, bx2, by2), 10, fill=ORANGE)
    f_btn = font(SANS_B, 40)
    btn = "詳細・ご注文はこちら"
    draw_text_center(d, btn, f_btn, W / 2 - 14, (by1 + by2) / 2 - text_height(f_btn, btn) / 2, WHITE, 3)
    ax, ay = 905, (by1 + by2) / 2
    d.line([((ax - 8) * S, (ay - 14) * S), ((ax + 6) * S, ay * S), ((ax - 8) * S, (ay + 14) * S)],
           fill=WHITE, width=int(3.5 * S), joint="curve")
    print("ボタンの下端:", by2, "（1040まで残り", W - by2, "）")

    c = c.resize((W, H), Image.LANCZOS)
    c.save(out, quality=95)
    print(out)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "_images/inkart_1102_calendar_table_banner_v1_1040.jpg"
    build(out)
