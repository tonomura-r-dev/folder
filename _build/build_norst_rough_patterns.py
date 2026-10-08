"""ノーストのバナーのラフ 3パターン（1040×1040）と、3つを並べた見比べ用の1枚。

A＝安心の事実をアイコンで並べる＋実績の円（10/19 No.7 のラフ。build_norst_1019_rough.py）
B＝2つを左右に並べて比べる形（ひとりで悩み続ける／無料カウンセリングで確かめる）
C＝本人の気持ちだけで押す一言（今年のうちに、自分のために一歩。）

どれも必ず残す4つ（緑が主・37年目・診療実績40万件・ロゴ）は入れている。
C の写真の場所は仮の枠（ラフなので）。

使い方:
  python3 _build/build_norst_rough_patterns.py
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_norst_1019_rough as A  # noqa: E402
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN, GREEN_DARK, TEAL, ORANGE, PALE, INK, GRAY, WHITE,
    font, tw, th, text, text_c, rrect, ellipse, jisseki_badge, white_logo,
    icon_door, icon_person, icon_pin,
)

OUT = Path("_images")
MUTED = (0x8A, 0x8A, 0x8A)
MUTED_BG = (0xEE, 0xEE, 0xEE)


def cta(c, d, by1=812, by2=952):
    """予約ボタン（下の中央に横長）。A と同じ形。"""
    bx1, bx2 = 60, 980
    rrect(d, (bx1, by1 + 8, bx2, by2 + 8), 36, fill=GREEN_DARK)
    grad = Image.new("RGB", (int((bx2 - bx1) * S), int((by2 - by1) * S)), TEAL)
    gd = ImageDraw.Draw(grad)
    for yy in range(grad.height):
        t = yy / grad.height
        gd.line((0, yy, grad.width, yy),
                fill=tuple(int(a + (b - a) * t) for a, b in zip((0x2A, 0xC9, 0xAE), TEAL)))
    mask = Image.new("L", grad.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, grad.width - 1, grad.height - 1), radius=36 * S, fill=255)
    c.paste(grad, (bx1 * S, by1 * S), mask)
    text_c(d, "まずは無料カウンセリング", font(30), W / 2 - 30, by1 + 18, WHITE, 1)
    text_c(d, "予約する", font(70), W / 2 - 30, by1 + 62, WHITE, 6, stroke=1, stroke_fill=WHITE)
    ccx, ccy = 900, (by1 + by2) / 2
    ellipse(d, (ccx - 38, ccy - 38, ccx + 38, ccy + 38), fill=WHITE)
    d.line([((ccx - 8) * S, (ccy - 16) * S), ((ccx + 10) * S, ccy * S), ((ccx - 8) * S, (ccy + 16) * S)],
           fill=TEAL, width=8 * S, joint="curve")


def footer(d):
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(22), W / 2, 984, GRAY)


def head(c, d, panel_bottom):
    """深緑の面＋ロゴ（左上）＋実績の円（右上）。"""
    d.rectangle((0, 0, W * S, panel_bottom * S), fill=GREEN)
    logo = white_logo(230 * S)
    c.paste(logo, (36 * S, 30 * S), logo)
    jisseki_badge(c, d, 905, 128, 88)


def check(d, x, y, col):
    """チェックの印（小さな丸に✓）。"""
    ellipse(d, (x, y, x + 30, y + 30), fill=col)
    d.line([((x + 8) * S, (y + 15) * S), ((x + 13) * S, (y + 21) * S), ((x + 23) * S, (y + 9) * S)],
           fill=WHITE, width=int(3.5 * S), joint="curve")


def cross(d, x, y, col):
    ellipse(d, (x, y, x + 30, y + 30), fill=col)
    d.line([((x + 10) * S, (y + 10) * S), ((x + 20) * S, (y + 20) * S)], fill=WHITE, width=int(3.5 * S))
    d.line([((x + 20) * S, (y + 10) * S), ((x + 10) * S, (y + 20) * S)], fill=WHITE, width=int(3.5 * S))


def pattern_b(out):
    """B＝2つを左右に並べて比べる形。"""
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)
    head(c, d, 392)

    text(d, "包茎のこと、ずっと気になっている方へ", font(34), 44, 120, WHITE, 1)
    text(d, "ひとりで悩み続けるか、", font(64), 44, 190, WHITE, 1, stroke=1.5, stroke_fill=WHITE)
    s2 = "一度、確かめるか。"
    f2 = font(104)
    text(d, s2, f2, 44 + 4, 276 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)
    text(d, s2, f2, 44, 276, ORANGE, 2, stroke=8, stroke_fill=WHITE)

    # 左右のカード
    top, bot = 424, 736
    lx1, lx2, rx1, rx2 = 40, 500, 540, 1000
    # 左（ひとりで悩み続ける）＝灰色で弱く
    rrect(d, (lx1, top, lx2, bot), 18, fill=WHITE, outline=(0xCC, 0xCC, 0xCC), width=3)
    rrect(d, (lx1, top, lx2, top + 74), 18, fill=MUTED)
    d.rectangle((lx1 * S, (top + 50) * S, lx2 * S, (top + 74) * S), fill=MUTED)
    text_c(d, "ひとりで悩み続ける", font(34), (lx1 + lx2) / 2, top + 20, WHITE, 1)
    left = ["ネットの情報はバラバラ", "自分に当てはまるか", "分からない", "何年も気になったまま"]
    f_l = font(29)
    ys = [top + 112, top + 182, top + 222, top + 274]
    for i, (s, y) in enumerate(zip(left, ys)):
        if i != 2:
            cross(d, lx1 + 26, y - 2, MUTED)
        text(d, s, f_l, lx1 + 70, y, (0x66, 0x66, 0x66))
    # 右（無料カウンセリングで確かめる）＝緑で強く
    rrect(d, (rx1 - 4, top - 4, rx2 + 4, bot + 4), 20, fill=ORANGE)
    rrect(d, (rx1, top, rx2, bot), 18, fill=WHITE)
    rrect(d, (rx1, top, rx2, top + 74), 18, fill=GREEN)
    d.rectangle((rx1 * S, (top + 50) * S, rx2 * S, (top + 74) * S), fill=GREEN)
    text_c(d, "無料カウンセリングで確かめる", font(29), (rx1 + rx2) / 2, top + 22, WHITE, 0)
    right = ["今の状態を確かめられる", "その日に決めなくてOK", "スタッフは男性のみ"]
    f_r = font(31)
    for s, y in zip(right, (top + 112, top + 186, top + 260)):
        check(d, rx1 + 26, y - 2, GREEN)
        text(d, s, f_r, rx1 + 70, y, GREEN_DARK)
    # 真ん中のVS
    ellipse(d, (470, 530, 570, 630), fill=ORANGE, outline=WHITE, width=5)
    text_c(d, "VS", font(40), 520, 562, WHITE, 1)

    text_c(d, "＼全国に36院開設／", font(32), W / 2, 762, GREEN_DARK, 1)
    cta(c, d, 812, 952)
    footer(d)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


def pattern_c(out):
    """C＝本人の気持ちだけで押す一言。"""
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)
    head(c, d, 470)

    text(d, "ずっと、気になっていた…", font(38), 44, 118, (0xCF, 0xEF, 0xE8), 2)
    text(d, "今年のうちに、", font(92), 44, 180, WHITE, 2, stroke=1.5, stroke_fill=WHITE)
    s2 = "自分のために一歩。"
    f2 = font(104)
    text(d, s2, f2, 44 + 4, 296 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)
    text(d, s2, f2, 44, 296, ORANGE, 2, stroke=8, stroke_fill=WHITE)
    text(d, "まずは無料カウンセリングで、今の状態を確かめてみませんか。", font(27), 46, 426, WHITE, 0)

    # 写真の場所（ラフなので仮の枠）
    px1, py1, px2, py2 = 40, 494, 1000, 704
    rrect(d, (px1, py1, px2, py2), 14, fill=(0xD9, 0xE8, 0xE5))
    for x in range(px1 + 10, px2 - 10, 24):
        d.line((x * S, (py1 + 6) * S, (x + 12) * S, (py1 + 6) * S), fill=(0x9F, 0xBF, 0xB8), width=2 * S)
        d.line((x * S, (py2 - 6) * S, (x + 12) * S, (py2 - 6) * S), fill=(0x9F, 0xBF, 0xB8), width=2 * S)
    text_c(d, "写真：窓辺に立つ男性の後ろ姿（顔は出さない）", font(28), W / 2, 580, (0x5F, 0x80, 0x7A))

    # 安心の事実（小さく3つ）
    chips = [(icon_person, "男性スタッフのみ"), (icon_door, "完全個室"), (icon_pin, "全国に36院")]
    cw, gap, top = 300, 30, 722
    x0 = (W - (cw * 3 + gap * 2)) / 2
    f = font(25)
    for i, (icon, s) in enumerate(chips):
        x1 = x0 + i * (cw + gap)
        rrect(d, (x1, top, x1 + cw, top + 70), 35, fill=WHITE, outline=TEAL, width=2.5)
        ellipse(d, (x1 + 10, top + 8, x1 + 64, top + 62), fill=GREEN)
        # アイコンは小さく描くために縮小版のキャンバスに描いて貼る
        ic = Image.new("RGBA", (120 * S, 120 * S), (0, 0, 0, 0))
        icon(ImageDraw.Draw(ic), 60, 60)
        ic = ic.resize((int(66 * S), int(66 * S)), Image.LANCZOS)
        c.paste(ic, (int((x1 + 4) * S), int((top + 2) * S)), ic)
        text(d, s, f, x1 + 80, top + 35 - th(f, s) / 2, GREEN_DARK)

    cta(c, d, 812, 952)
    footer(d)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


def sheet(paths, labels, out):
    """3枚を横に並べた見比べ用（ラベル付き）。"""
    tw_, gap, pad, lab = 520, 30, 30, 70
    sh = Image.new("RGB", (pad * 2 + tw_ * len(paths) + gap * (len(paths) - 1), pad * 2 + lab + tw_), WHITE)
    sd = ImageDraw.Draw(sh)
    f = A.ImageFont.truetype(A.SANS_B, 30, index=0)
    for i, (p, l) in enumerate(zip(paths, labels)):
        x = pad + i * (tw_ + gap)
        sd.text((x, pad + 10), l, font=f, fill=INK)
        sh.paste(Image.open(p).resize((tw_, tw_), Image.LANCZOS), (x, pad + lab))
        sd.rectangle((x, pad + lab, x + tw_ - 1, pad + lab + tw_ - 1), outline=(0xCC, 0xCC, 0xCC))
    sh.save(out, quality=92)
    return out


if __name__ == "__main__":
    a = OUT / "norst_rough_A_icon_1040.jpg"
    A.build(str(a))
    b = pattern_b(OUT / "norst_rough_B_vs_1040.jpg")
    cc = pattern_c(OUT / "norst_rough_C_jibun_1040.jpg")
    s = sheet([a, b, cc], ["A 安心の事実を並べる", "B 2つを比べる", "C 自分のための一言"],
              OUT / "norst_rough_ABC_sheet.jpg")
    for p in (a, b, cc, s):
        print(p)
