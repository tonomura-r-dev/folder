"""ノースト 10/19（No.7・無料カウンセリング当日の流れ）のバナーのラフ（1040×1040）。

競合から借りた2つを入れた形：
  ①安心の事実をアイコンで横に並べる列（完全個室／男性スタッフのみ／全国に36院／カウンセリング無料）
  ②実績を月桂樹のバッジで見せる（37年目・診療実績40万件）
必ず残す4つ（緑が主・37年目・診療実績40万件・ロゴ）は入れている。
ラフなので、文字の位置と量を見るためのもの。仕上げはデザイナーか画像生成で行う。

使い方:
  python3 _build/build_norst_1019_rough.py [出力]
"""
import math
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parent))
from norst_paste_logo import white_logo  # noqa: E402

S = 2  # 2倍で描いて最後に縮める
W = H = 1040

GREEN = (0x00, 0x80, 0x6B)       # 深緑（HPのいちばん上の帯）
GREEN_DARK = (0x00, 0x5E, 0x4F)
TEAL = (0x00, 0xA9, 0x8F)        # 明るい緑（HPのボタン）
ORANGE = (0xEE, 0x78, 0x00)      # 強調のオレンジ
PALE = (0xE8, 0xF5, 0xF3)        # 淡水色
GOLD = (0xC9, 0xA2, 0x4A)
GOLD_DARK = (0xA0, 0x7E, 0x2E)
INK = (0x33, 0x33, 0x33)
GRAY = (0x66, 0x66, 0x66)
WHITE = (255, 255, 255)

SANS_B = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"

PANEL_BOTTOM = 494  # 深緑の面の下端


def font(size):
    return ImageFont.truetype(SANS_B, int(size * S), index=0)


def tw(f, text, tracking=0):
    return sum(f.getlength(ch) for ch in text) / S + tracking * (len(text) - 1)


def th(f, text):
    b = f.getbbox(text)
    return (b[3] - b[1]) / S


def text(d, s, f, x, top, fill, tracking=0, stroke=0, stroke_fill=None):
    """x, top は 1040 基準。top は文字の上端。"""
    off = f.getbbox(s)[1]
    cx = x * S
    for ch in s:
        d.text((cx, top * S - off), ch, font=f, fill=fill,
               stroke_width=int(stroke * S), stroke_fill=stroke_fill)
        cx += f.getlength(ch) + tracking * S


def text_c(d, s, f, cx, top, fill, tracking=0, **kw):
    text(d, s, f, cx - tw(f, s, tracking) / 2, top, fill, tracking, **kw)


def rrect(d, box, r, fill=None, outline=None, width=1):
    x1, y1, x2, y2 = box
    d.rounded_rectangle((x1 * S, y1 * S, x2 * S, y2 * S), radius=r * S,
                        fill=fill, outline=outline, width=int(width * S))


def ellipse(d, box, fill=None, outline=None, width=1):
    x1, y1, x2, y2 = box
    d.ellipse((x1 * S, y1 * S, x2 * S, y2 * S), fill=fill, outline=outline, width=int(width * S))


def leaf(canvas, cx, cy, length, angle_deg, color):
    """細長い葉を1枚、中心(cx,cy)・向き angle で置く。"""
    w, h = int(length * S), int(length * 0.42 * S)
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(im).ellipse((0, 0, w - 1, h - 1), fill=color)
    im = im.rotate(angle_deg, resample=Image.BICUBIC, expand=True)
    canvas.paste(im, (int(cx * S - im.width / 2), int(cy * S - im.height / 2)), im)


def laurel_badge(c, d, cx, cy, r):
    """月桂樹のバッジ：白い円の中に 37年目／診療実績40万件。周りに金の葉。"""
    # 葉（左右に1本ずつの枝）
    for side in (-1, 1):
        n = 9
        for i in range(n):
            t = i / (n - 1)
            a = math.radians(118 + t * 110)          # 左の枝は 118°→228°（下から上へ）
            if side == 1:
                a = math.pi - a
            lx = cx + (r + 16) * math.cos(a)
            ly = cy + (r + 16) * math.sin(a)
            tang = math.degrees(math.atan2(math.cos(a), -math.sin(a))) * -1
            for k, (dr, col) in enumerate(((10, GOLD), (-10, GOLD_DARK))):
                ox = lx + dr * math.cos(a)
                oy = ly + dr * math.sin(a)
                leaf(c, ox, oy, 30 - t * 6, tang + (25 if k == 0 else -25) * side, col)
    ellipse(d, (cx - r, cy - r, cx + r, cy + r), fill=WHITE, outline=GOLD, width=4)
    f1, f1s = font(52), font(26)
    f2, f3 = font(24), font(42)
    # 37年目
    w = tw(f1, "37") + tw(f1s, "年目") + 4
    x = cx - w / 2
    text(d, "37", f1, x, cy - 66, ORANGE)
    text(d, "年目", f1s, x + tw(f1, "37") + 4, cy - 66 + th(f1, "37") - th(f1s, "年目"), GREEN)
    d.line(((cx - r + 30) * S, (cy - 2) * S, (cx + r - 30) * S, (cy - 2) * S), fill=GOLD, width=2 * S)
    text_c(d, "診療実績", f2, cx, cy + 10, GREEN)
    w = tw(f3, "40") + tw(f2, "万件") + 4
    x = cx - w / 2
    text(d, "40", f3, x, cy + 42, ORANGE)
    text(d, "万件", f2, x + tw(f3, "40") + 4, cy + 42 + th(f3, "40") - th(f2, "万件"), GREEN)


def icon_door(d, cx, cy):
    rrect(d, (cx - 20, cy - 30, cx + 20, cy + 30), 3, outline=WHITE, width=4)
    rrect(d, (cx - 12, cy - 22, cx + 12, cy + 30), 2, fill=WHITE)
    ellipse(d, (cx + 4, cy + 2, cx + 9, cy + 7), fill=GREEN)


def icon_person(d, cx, cy):
    ellipse(d, (cx - 14, cy - 32, cx + 14, cy - 4), fill=WHITE)
    d.pieslice(((cx - 28) * S, (cy + 0) * S, (cx + 28) * S, (cy + 56) * S), 180, 360, fill=WHITE)
    d.polygon([(cx * S, (cy + 4) * S), ((cx - 5) * S, (cy + 12) * S), (cx * S, (cy + 28) * S),
               ((cx + 5) * S, (cy + 12) * S)], fill=GREEN)


def icon_pin(d, cx, cy):
    d.polygon([((cx - 22) * S, (cy - 6) * S), ((cx + 22) * S, (cy - 6) * S), (cx * S, (cy + 32) * S)], fill=WHITE)
    ellipse(d, (cx - 23, cy - 32, cx + 23, cy + 14), fill=WHITE)
    ellipse(d, (cx - 9, cy - 18, cx + 9, cy), fill=GREEN)


def icon_bubble(d, cx, cy):
    rrect(d, (cx - 30, cy - 26, cx + 30, cy + 14), 10, fill=WHITE)
    d.polygon([((cx - 12) * S, (cy + 12) * S), ((cx - 2) * S, (cy + 12) * S), ((cx - 16) * S, (cy + 28) * S)], fill=WHITE)
    for dx in (-14, 0, 14):
        ellipse(d, (cx + dx - 5, cy - 11, cx + dx + 5, cy - 1), fill=GREEN)


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 1. 深緑の面
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)

    # ロゴ（左上）
    logo = white_logo(230 * S)
    c.paste(logo, (36 * S, 30 * S), logo)

    # 月桂樹のバッジ（右上）
    laurel_badge(c, d, 905, 128, 88)

    # 心の声（吹き出し）
    f_v = font(34)
    v = "「行ったら、断れなさそう…」"
    vw = tw(f_v, v) + 44
    rrect(d, (40, 104, 40 + vw, 160), 28, fill=WHITE)
    d.polygon([(70 * S, 158 * S), (98 * S, 158 * S), (66 * S, 180 * S)], fill=WHITE)
    text(d, v, f_v, 62, 132 - th(f_v, v) / 2, GREEN_DARK)

    # 主役の上
    f_0 = font(40)
    text(d, "包茎治療の無料カウンセリングは、", f_0, 44, 196, WHITE, 1)

    # 主役2行
    f_1 = font(72)
    text_c(d, "その日に決めなくても、", f_1, W / 2, 262, WHITE, 1, stroke=1.5, stroke_fill=WHITE)
    f_2 = font(136)
    s2 = "大丈夫です。"
    text_c(d, s2, f_2, W / 2 + 4, 354 + 4, GREEN_DARK, 2, stroke=12, stroke_fill=GREEN_DARK)  # 影
    text_c(d, s2, f_2, W / 2, 354, ORANGE, 2, stroke=9, stroke_fill=WHITE)

    # 2. アイコンの列（4枚）
    cards = [
        (icon_door, ["完全個室"], [None]),
        (icon_person, ["男性スタッフ", "のみ"], [None, ORANGE]),
        (icon_pin, ["全国に", "36院"], [None, ORANGE]),
        (icon_bubble, ["カウンセリング", "無料"], [None, ORANGE]),
    ]
    cw, gap, top, bot = 226, 12, 514, 704
    x0 = (W - (cw * 4 + gap * 3)) / 2
    f_c, f_cb = font(27), font(34)
    for i, (icon, lines, colors) in enumerate(cards):
        x1 = x0 + i * (cw + gap)
        rrect(d, (x1, top, x1 + cw, bot), 14, fill=WHITE, outline=TEAL, width=3)
        icx, icy = x1 + cw / 2, top + 60
        ellipse(d, (icx - 46, icy - 46, icx + 46, icy + 46), fill=GREEN)
        icon(d, icx, icy)
        y = top + 122 if len(lines) == 2 else top + 136
        for s, col in zip(lines, colors):
            f = f_cb if col else f_c
            text_c(d, s, f, icx, y, col or INK)
            y += th(f, s) + 10

    # 3. 当日の流れ（細い帯）
    fy1, fy2 = 722, 792
    rrect(d, (40, fy1, 1000, fy2), 36, fill=WHITE, outline=GREEN, width=2)
    f_l, f_s = font(25), font(25)
    rrect(d, (48, fy1 + 8, 200, fy2 - 8), 28, fill=GREEN)
    text_c(d, "当日の流れ", f_l, 124, (fy1 + fy2) / 2 - th(f_l, "当日の流れ") / 2, WHITE)
    steps = ["①ご予約", "②完全個室で1対1", "③治療法やご不安をご相談"]
    arrow_w = 30
    total = sum(tw(f_s, s) for s in steps) + arrow_w * 2 + 20 * 2
    x = 200 + (1000 - 200 - total) / 2
    for i, s in enumerate(steps):
        text(d, s, f_s, x, (fy1 + fy2) / 2 - th(f_s, s) / 2, GREEN_DARK)
        x += tw(f_s, s) + 20
        if i < 2:
            ay = (fy1 + fy2) / 2
            d.polygon([(x * S, (ay - 10) * S), ((x + 14) * S, ay * S), (x * S, (ay + 10) * S)], fill=ORANGE)
            x += arrow_w

    # 4. 予約ボタン（下の中央に横長）
    bx1, by1, bx2, by2 = 60, 812, 980, 952
    rrect(d, (bx1, by1 + 8, bx2, by2 + 8), 36, fill=GREEN_DARK)  # 厚み
    grad = Image.new("RGB", (int((bx2 - bx1) * S), int((by2 - by1) * S)), TEAL)
    gd = ImageDraw.Draw(grad)
    for yy in range(grad.height):
        t = yy / grad.height
        col = tuple(int(a + (b - a) * t) for a, b in zip((0x2A, 0xC9, 0xAE), TEAL))
        gd.line((0, yy, grad.width, yy), fill=col)
    mask = Image.new("L", grad.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, grad.width - 1, grad.height - 1), radius=36 * S, fill=255)
    c.paste(grad, (bx1 * S, by1 * S), mask)
    f_b1, f_b2 = font(30), font(70)
    text_c(d, "まずは無料カウンセリング", f_b1, W / 2 - 30, by1 + 18, WHITE, 1)
    text_c(d, "予約する", f_b2, W / 2 - 30, by1 + 62, WHITE, 6, stroke=1, stroke_fill=WHITE)
    ccx, ccy = 900, (by1 + by2) / 2
    ellipse(d, (ccx - 38, ccy - 38, ccx + 38, ccy + 38), fill=WHITE)
    d.line([((ccx - 8) * S, (ccy - 16) * S), ((ccx + 10) * S, ccy * S), ((ccx - 8) * S, (ccy + 16) * S)],
           fill=TEAL, width=8 * S, joint="curve")

    # 5. 注記
    f_n = font(22)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", f_n, W / 2, 984, GRAY)

    c = c.resize((W, H), Image.LANCZOS)
    c.save(out, quality=95)
    print(out)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_1019_rough_1040.jpg")
