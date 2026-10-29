"""インクアート 11/2 オリジナル缶バッジの配信バナー（1040×1040）。
絵はサイトの缶バッジ写真（_data/printearth/site_badge_types_large_20261020.png）をそのまま使い、
文字はここで組む（2026-10-29 殿村さん「缶バッジの画像は、サイトの一回送ったものでいい」）。
構成（殿村さん確定）＝上の吹き出し「オリジナル缶バッジ」／主役「販促・ノベルティを／もっと気軽な形で」／
「1個から3,000個まで」／札3つ（形式豊富／社章代わりに／1個ずつ袋入り）／写真／ボタン。

使い方: PYTHONPATH=<pylib> python3 _build/build_inkart_1102_canbadge_site_banner.py [出力]
"""
import sys
from PIL import Image, ImageDraw, ImageFont

S = 2
W = H = 1040
GREEN = (0x1C, 0x8B, 0x6A)
ORANGE = (0xFF, 0x88, 0x00)
YELLOW = (0xFF, 0xE8, 0x00)
INK = (0x33, 0x33, 0x33)
WHITE = (255, 255, 255)

FONT_DIR = "/usr/share/fonts/opentype/noto/"
SANS_B = FONT_DIR + "NotoSansCJK-Bold.ttc"
PHOTO = "_data/printearth/site_badge_types_large_20261020.png"
PHOTO_CROP = (15, 0, 835, 479)  # 左右の黒い縁を外す


def font(size):
    return ImageFont.truetype(SANS_B, size * S, index=0)


def text_width(f, text, tracking=0):
    return sum(f.getlength(ch) for ch in text) + tracking * S * (len(text) - 1)


def draw_text(d, text, f, x, top, fill, tracking=0, stroke=0, stroke_fill=WHITE):
    t = f.getbbox(text)[1]
    cx = x * S
    for ch in text:
        d.text((cx, top * S - t), ch, font=f, fill=fill,
               stroke_width=int(stroke * S), stroke_fill=stroke_fill)
        cx += f.getlength(ch) + tracking * S


def draw_text_center(d, text, f, cx, top, fill, tracking=0, stroke=0):
    w = text_width(f, text, tracking) / S
    draw_text(d, text, f, cx - w / 2, top, fill, tracking, stroke)


def text_height(f, text):
    b = f.getbbox(text)
    return (b[3] - b[1]) / S


def rrect(d, box, r, fill=None, outline=None, width=1):
    x1, y1, x2, y2 = box
    d.rounded_rectangle((x1 * S, y1 * S, x2 * S, y2 * S), radius=r * S,
                        fill=fill, outline=outline, width=int(width * S))


def cover(im, w, h):
    r = w / h
    iw, ih = im.size
    if iw / ih > r:
        nw = int(ih * r)
        im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
    else:
        nh = int(iw / r)
        im = im.crop((0, (ih - nh) // 2, iw, (ih - nh) // 2 + nh))
    return im.resize((int(w), int(h)), Image.LANCZOS)


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

    # 1. 上の黄色い吹き出し：オリジナル缶バッジ
    f_bub = font(44)
    bub = "オリジナル缶バッジ"
    bw = text_width(f_bub, bub, 2) / S + 80
    bx1, by1 = (W - bw) / 2, 30
    bx2, by2 = bx1 + bw, by1 + 76
    rrect(d, (bx1, by1, bx2, by2), 14, fill=YELLOW)
    k = 14
    d.polygon([((W / 2 - k) * S, by2 * S), ((W / 2 + k) * S, by2 * S), ((W / 2) * S, (by2 + 16) * S)], fill=YELLOW)
    draw_text_center(d, bub, f_bub, W / 2, (by1 + by2) / 2 - text_height(f_bub, bub) / 2, GREEN, 2)

    # 2. 主役（2行・緑・白フチ）
    f_m = font(98)
    y = by2 + 30
    for s in ("販促・ノベルティを", "もっと気軽な形で"):
        draw_text_center(d, s, f_m, W / 2, y, GREEN, 2, stroke=7)
        y += text_height(f_m, s) + 12

    # 3. 1個から3,000個まで
    f_sub = font(46)
    sub = "1個から3,000個まで"
    y += 6
    draw_text_center(d, sub, f_sub, W / 2, y, INK, 2)
    y += text_height(f_sub, sub) + 20

    # 4. 札3つ
    f_chip = font(32)
    chips = ["形式豊富", "社章代わりに", "1個ずつ袋入り"]
    cy1, cy2 = y, y + 60
    for i, t in enumerate(chips):
        x1 = 40 + i * (310 + 15)
        x2 = x1 + 310
        rrect(d, (x1, cy1, x2, cy2), 8, fill=WHITE, outline=GREEN, width=2.5)
        draw_text_center(d, t, f_chip, (x1 + x2) / 2, (cy1 + cy2) / 2 - text_height(f_chip, t) / 2, GREEN, 1)
    y = cy2 + 18

    # 5. 写真（サイトの缶バッジ10個）。切らずに全部入れ、余った左右は写真と同じ下地の色で埋める
    top, bottom = y, 916
    photo = Image.open(PHOTO).convert("RGB").crop(PHOTO_CROP)
    bg = photo.getpixel((5, 5))
    rrect(d, (40, top, 1000, bottom), 10, fill=bg)
    ph = bottom - top
    pw = int(photo.size[0] * ph / photo.size[1])
    px = (W - pw) / 2
    paste_rounded(c, photo, (px, top, px + pw, bottom), 10)
    rrect(d, (40, top, 1000, bottom), 10, outline=(0xDD, 0xDD, 0xDD), width=1)

    # 6. ボタン
    bx1, by1, bx2, by2 = 70, 934, 970, 1012
    rrect(d, (bx1, by1, bx2, by2), 12, fill=ORANGE)
    f_btn = font(42)
    btn = "詳細・ご注文はこちら"
    draw_text_center(d, btn, f_btn, W / 2 - 14, (by1 + by2) / 2 - text_height(f_btn, btn) / 2, WHITE, 3)
    ax, ay = 905, (by1 + by2) / 2
    d.line([((ax - 8) * S, (ay - 14) * S), ((ax + 6) * S, ay * S), ((ax - 8) * S, (ay + 14) * S)],
           fill=WHITE, width=int(3.5 * S), joint="curve")

    c = c.resize((W, H), Image.LANCZOS)
    c.save(out, quality=95)
    print(out, "photo", top, bottom)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "_images/inkart_1102_canbadge_banner_site_v3_1040.jpg")
