"""インクアート 11月2本目 クリアカード印刷の配信バナー（1040×1040）を組む。

文字は画像生成に描かせず、ここで入れる（ゴシック・一字も崩れない）。
写真はサイトのスライド2枚（殿村さんが 2026-10-28 に取得）：
  左＝机の上にカード各種（_data/printearth/lp_clear_card_photo1_desk.webp）
  右＝表面と裏面（_data/printearth/lp_clear_card_photo2_front_back.png）
文字は配信テキストと同じ言葉だけ（10/28の本文→バナーの呼び方をそろえるルール）。
型＝帯／主役／一文／札3つ／写真／ボタン（10/20の学び：訴求の札は写真より上）。

使い方: PYTHONPATH=<pylib> python3 _build/build_inkart_1109_clearcard_banner.py [出力]
"""
import sys
from PIL import Image, ImageDraw, ImageFont

S = 2
W = H = 1040
GREEN = (0x1C, 0x8B, 0x6A)
ORANGE = (0xFF, 0x88, 0x00)
INK = (0x33, 0x33, 0x33)
WHITE = (255, 255, 255)
PALE = (0xEE, 0xF7, 0xF4)

FONT_DIR = "/usr/share/fonts/opentype/noto/"
SANS_B = FONT_DIR + "NotoSansCJK-Bold.ttc"
LEFT = "_data/printearth/lp_clear_card_photo1_desk.webp"
RIGHT = "_data/printearth/lp_clear_card_photo2_front_back.png"


def font(size):
    return ImageFont.truetype(SANS_B, size * S, index=0)


def text_width(f, text, tracking=0):
    return sum(f.getlength(ch) for ch in text) + tracking * S * (len(text) - 1)


def draw_text(d, text, f, x, top, fill, tracking=0):
    t = f.getbbox(text)[1]
    cx = x * S
    for ch in text:
        d.text((cx, top * S - t), ch, font=f, fill=fill)
        cx += f.getlength(ch) + tracking * S


def draw_text_center(d, text, f, cx, top, fill, tracking=0):
    w = text_width(f, text, tracking) / S
    draw_text(d, text, f, cx - w / 2, top, fill, tracking)


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

    # 1. 上の緑の帯：札「透明素材」＋商品名
    d.rectangle((0, 0, W * S, 96 * S), fill=GREEN)
    f_tag, f_name = font(28), font(44)
    tag = "透明素材"
    name = "クリアカード印刷"
    tag_w = text_width(f_tag, tag, 2) / S + 32
    name_w = text_width(f_name, name, 3) / S
    gx = (W - (tag_w + 22 + name_w)) / 2
    rrect(d, (gx, 26, gx + tag_w, 70), 4, fill=ORANGE)
    draw_text(d, tag, f_tag, gx + 16, 48 - text_height(f_tag, tag) / 2, WHITE, 2)
    draw_text(d, name, f_name, gx + tag_w + 22, 48 - text_height(f_name, name) / 2, WHITE, 3)

    # 2. 主役（緑の特大ゴシック・2行）
    f_m1, f_m2 = font(110), font(118)
    y = 122
    draw_text_center(d, "透明なカードで、", f_m1, W / 2, y, GREEN, 4)
    y += text_height(f_m1, "透明なカードで、") + 22
    draw_text_center(d, "印象に残る一枚を", f_m2, W / 2, y, GREEN, 4)
    y += text_height(f_m2, "印象に残る一枚を") + 26

    # 3. 一文
    f_sub = font(40)
    sub = "ガラスのような透明感と光沢感"
    draw_text_center(d, sub, f_sub, W / 2, y, INK, 3)
    y += text_height(f_sub, sub) + 26

    # 4. 札3つ（写真より上）
    f_chip = font(25)
    chips = ["名刺・ショップカードに", "100枚 9,780円〜（税別）", "6営業日発送"]
    cy1, cy2 = y, y + 64
    for i, t in enumerate(chips):
        x1 = 40 + i * (310 + 15)
        x2 = x1 + 310
        rrect(d, (x1, cy1, x2, cy2), 6, fill=PALE, outline=GREEN, width=2)
        draw_text_center(d, t, f_chip, (x1 + x2) / 2,
                         (cy1 + cy2) / 2 - text_height(f_chip, t) / 2, GREEN, 0)
    y = cy2 + 22

    # 5. 写真2枚（左＝机の上・広め、右＝表と裏）
    top, bottom = y, 900
    split = 40 + 560
    paste_rounded(c, Image.open(LEFT).convert("RGB"), (40, top, split, bottom), 8)
    paste_rounded(c, Image.open(RIGHT).convert("RGB"), (split + 12, top, 1000, bottom), 8)
    for bx in ((40, top, split, bottom), (split + 12, top, 1000, bottom)):
        rrect(d, bx, 8, outline=(0xDD, 0xDD, 0xDD), width=1)

    # 6. ボタン
    bx1, by1, bx2, by2 = 70, 922, 970, 1006
    rrect(d, (bx1, by1, bx2, by2), 10, fill=ORANGE)
    f_btn = font(40)
    btn = "詳細・ご注文はこちら"
    draw_text_center(d, btn, f_btn, W / 2 - 14, (by1 + by2) / 2 - text_height(f_btn, btn) / 2, WHITE, 3)
    ax, ay = 905, (by1 + by2) / 2
    d.line([((ax - 8) * S, (ay - 14) * S), ((ax + 6) * S, ay * S), ((ax - 8) * S, (ay + 14) * S)],
           fill=WHITE, width=int(3.5 * S), joint="curve")

    c = c.resize((W, H), Image.LANCZOS)
    c.save(out, quality=95)
    print(out, "photo top", top)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "_images/inkart_1109_clearcard_banner_v1_1040.jpg")
