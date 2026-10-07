"""インクアート 10/19 紙製ランチョンマット印刷の配信バナー（1040×1040）を組む。

文字は画像生成に描かせず、ここで入れる（明朝・ゴシックで一字も崩れない）。
真ん中の写真はサイトの商品写真をはめ込む。写真が届くまでは仮の写真で組む。

使い方:
  python3 _build/build_inkart_placemat_banner.py <左の写真> <右の写真> [出力]
  左右の写真は枠に合わせて中央で切り抜く。切り抜く範囲を指定するときは
  「ファイル名:x1,y1,x2,y2」と書く（元の画像のピクセル）。
"""
import sys
from PIL import Image, ImageDraw, ImageFont

S = 2  # 2倍で描いて最後に縮める（文字の縁をなめらかにする）
W = H = 1040

GREEN = (0x1C, 0x8B, 0x6A)   # HPの緑
ORANGE = (0xFF, 0x88, 0x00)  # HPのオレンジ（ボタン）
INK = (0x33, 0x33, 0x33)
GOLD = (0xB8, 0x97, 0x5A)
WHITE = (255, 255, 255)

LEFT_W = 530  # 左の写真の幅（右の写真は残り）

FONT_DIR = "/usr/share/fonts/opentype/noto/"
SERIF_B = FONT_DIR + "NotoSerifCJK-Bold.ttc"
SANS_B = FONT_DIR + "NotoSansCJK-Bold.ttc"


def font(path, size):
    return ImageFont.truetype(path, size * S, index=0)  # index 0 = JP


def text_width(f, text, tracking=0):
    return sum(f.getlength(ch) for ch in text) + tracking * S * (len(text) - 1)


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


def text_height(f, text):
    b = f.getbbox(text)
    return (b[3] - b[1]) / S


def rrect(d, box, r, fill=None, outline=None, width=1):
    x1, y1, x2, y2 = box
    d.rounded_rectangle((x1 * S, y1 * S, x2 * S, y2 * S), radius=r * S,
                        fill=fill, outline=outline, width=int(width * S))


def load_photo(spec):
    if ":" in spec and spec.rsplit(":", 1)[1].count(",") == 3:
        path, box = spec.rsplit(":", 1)
        im = Image.open(path).convert("RGB")
        return im.crop(tuple(int(v) for v in box.split(",")))
    return Image.open(spec).convert("RGB")


def cover(im, w, h):
    """枠いっぱいに入るよう中央で切り抜いて縮める。"""
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


def build(left, right, out):
    c = Image.new("RGB", (W * S, H * S), WHITE)
    d = ImageDraw.Draw(c)

    # 1. 上の緑の帯：新商品＋商品名
    d.rectangle((0, 0, W * S, 96 * S), fill=GREEN)
    f_tag = font(SANS_B, 28)
    f_name = font(SANS_B, 42)
    tag_w = text_width(f_tag, "新商品", 2) / S + 32
    name_w = text_width(f_name, "紙製ランチョンマット印刷", 2) / S
    gx = (W - (tag_w + 22 + name_w)) / 2
    rrect(d, (gx, 26, gx + tag_w, 70), 4, fill=ORANGE)
    draw_text(d, "新商品", f_tag, gx + 16, 48 - text_height(f_tag, "新商品") / 2, WHITE, 2)
    draw_text(d, "紙製ランチョンマット印刷", f_name, gx + tag_w + 22,
              48 - text_height(f_name, "紙製ランチョンマット印刷") / 2, WHITE, 2)

    # 2. 主役（明朝）
    f_m1 = font(SERIF_B, 124)
    f_m2 = font(SERIF_B, 132)
    y = 124
    draw_text_center(d, "紙から", f_m1, W / 2, y, GREEN, 8)
    y += text_height(f_m1, "紙から") + 26
    draw_text_center(d, "おもてなしを", f_m2, W / 2, y, GREEN, 6)
    y += text_height(f_m2, "おもてなしを") + 24

    # 3. 金の細い線と小さなひし形
    ly = y
    d.line((352 * S, ly * S, 500 * S, ly * S), fill=GOLD, width=int(1.5 * S))
    d.line((540 * S, ly * S, 688 * S, ly * S), fill=GOLD, width=int(1.5 * S))
    k = 7
    d.polygon([((520) * S, (ly - k) * S), ((520 + k) * S, ly * S),
               (520 * S, (ly + k) * S), ((520 - k) * S, ly * S)], fill=GOLD)
    y = ly + 20

    # 4. 和紙の一文（明朝）
    f_sub = font(SERIF_B, 40)
    draw_text_center(d, "風合い豊かな和紙も選べます", f_sub, W / 2, y, INK, 3)
    y += text_height(f_sub, "風合い豊かな和紙も選べます") + 26

    # 5. 写真2枚（左＝商品、右＝和紙など）。左を広めにとる
    top, bottom = y, 814
    split = 40 + LEFT_W
    paste_rounded(c, left, (40, top, split, bottom), 8)
    paste_rounded(c, right, (split + 12, top, 1000, bottom), 8)
    for bx in ((40, top, split, bottom), (split + 12, top, 1000, bottom)):
        rrect(d, bx, 8, outline=(0xDD, 0xDD, 0xDD), width=1)
    #   裏面印刷も可能！（左の写真の左上の角に札として重ねる）
    f_lab = font(SANS_B, 28)
    lab = "裏面印刷も可能！"
    lw = text_width(f_lab, lab, 1) / S + 36
    lx1, ly1 = 40 + 12, top + 12
    rrect(d, (lx1, ly1, lx1 + lw, ly1 + 50), 4, fill=GREEN)
    draw_text(d, lab, f_lab, lx1 + 18, ly1 + 25 - text_height(f_lab, lab) / 2, WHITE, 1)

    # 6. 札3つ
    f_chip = font(SANS_B, 30)
    chips = ["両面カラー印刷", "50部から注文可能", "箔押し加工にも対応"]
    cy1, cy2 = 834, 902
    for i, t in enumerate(chips):
        x1 = 40 + i * (310 + 15)
        x2 = x1 + 310
        rrect(d, (x1, cy1, x2, cy2), 6, fill=WHITE, outline=GREEN, width=2)
        if i == 2:
            rrect(d, (x1 + 5, cy1 + 5, x2 - 5, cy2 - 5), 4, outline=GOLD, width=1.2)
        draw_text_center(d, t, f_chip, (x1 + x2) / 2,
                         (cy1 + cy2) / 2 - text_height(f_chip, t) / 2, GREEN, 1)

    # 7. ボタン（平らなオレンジ）
    bx1, by1, bx2, by2 = 70, 922, 970, 1006
    rrect(d, (bx1, by1, bx2, by2), 10, fill=ORANGE)
    f_btn = font(SANS_B, 40)
    btn = "詳細・ご注文はこちら"
    draw_text_center(d, btn, f_btn, W / 2 - 14, (by1 + by2) / 2 - text_height(f_btn, btn) / 2, WHITE, 3)
    ax, ay = 905, (by1 + by2) / 2
    d.line([((ax - 8) * S, (ay - 14) * S), ((ax + 6) * S, ay * S), ((ax - 8) * S, (ay + 14) * S)],
           fill=WHITE, width=int(3.5 * S), joint="curve")

    c = c.resize((W, H), Image.LANCZOS)
    c.save(out, quality=95)
    print(out)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    out = sys.argv[3] if len(sys.argv) > 3 else "_images/inkart_1019_placemat_banner_1040.jpg"
    build(load_photo(sys.argv[1]), load_photo(sys.argv[2]), out)
