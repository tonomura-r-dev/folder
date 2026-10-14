"""ノースト 10月2本目のバナー・イラスト版（1040×1040）。

2026-10-14 殿村さん「イラストの良さが無いやん／男性のイラストを全面に押しながらも、文面は的確に反映して欲しい」。
→ 元の写真入りの画像（`_images/norst_1016_banner_photo_v2.webp`）を土台にして、男性の大きさ・位置・緑の面はそのまま。
   古い文字（ロゴ・タグ・吹き出し・主役・円・ボタン）を消してから、確定版（v9）の言葉を同じ並びで載せ直す。

載せる言葉はv9と同じ：
  タグ3つ（包茎治療／亀頭増大／ED治療）・心の声「何とかしたいなぁ…」（10/14 「何とかならないかなぁ…」から変更）・主役「その“何とか”を、医師が一緒に考えます。」・
  カウンセリング無料・円4つ（男性スタッフのみ／全国に36院／37年目・診療実績40万件／その日に決めなくても大丈夫）・
  当日の流れ・「電話なしで、このままWEB予約できます」・予約ボタン・注記。

使い方:
  python3 _build/build_norst_1016_banner_illust.py [出力]
"""
import math
import sys
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN, GREEN_DARK, TEAL, ORANGE, INK, GRAY, WHITE,
    font, tw, th, text, text_c, rrect, ellipse, white_logo, icon_person, icon_pin,
)
from build_norst_1016_banner import icon_calendar_check, icon_chart, segs_c, counsel_badge  # noqa: E402

SRC = "_images/norst_1016_banner_photo_v2.webp"  # 1254×1254
BG_GREEN = (2, 145, 117)      # 元画像の左側の緑
BG_PALE = (237, 251, 249)     # 元画像の下の淡い色
RING = (190, 235, 222)        # 円のふち


def clean_base():
    """元画像から古い文字を消す（左は緑で塗り直し、男性にかかる所は周りの色で埋める）。1040×S の画像を返す。"""
    src = cv2.imread(SRC)  # BGR
    n = src.shape[0]
    f = n / W
    b, g, r = [src[:, :, i].astype(int) for i in range(3)]
    ys, xs = np.mgrid[0:n, 0:n]
    X, Y = xs / f, ys / f  # 1040基準の座標

    white = (np.minimum(np.minimum(r, g), b) > 228)
    orange = (r > 190) & (g > 60) & (g < 190) & (b < 100)
    dark = (g < 130) & (r < 45) & (b < 115)
    # 左側（x<600）は文字の色をすべて拾う。男性にかかる帯（600〜760）は、元の文字があった高さだけ拾う
    left = (X < 600) & (Y < 600)
    # 吹き出しの右端と点（y 140〜270）は x<680 まで（それより右は男性の頬・もみあげ）。「に」（y 355〜495）は x<760 まで
    band = (X >= 600) & (((X < 680) & (Y > 140) & (Y < 270)) | ((X < 760) & (Y > 355) & (Y < 495)))
    mask = ((white | orange | dark) & left) | ((white | orange | dark) & band)
    mask = mask.astype(np.uint8) * 255
    mask = cv2.dilate(mask, np.ones((9, 9), np.uint8))
    out = cv2.inpaint(src, mask, 9, cv2.INPAINT_TELEA)

    img = Image.fromarray(cv2.cvtColor(out, cv2.COLOR_BGR2RGB)).resize((W * S, H * S), Image.LANCZOS)
    # 左側は緑一色に塗り直し、x=500〜600 で元の色（男性の手前の緑のグラデーション）へなじませる
    flat = Image.new("RGB", img.size, BG_GREEN)
    m = Image.new("L", img.size, 0)
    md = ImageDraw.Draw(m)
    for x in range(0, 600 * S):
        a = 255 if x < 500 * S else int(255 * (600 * S - x) / (100 * S))
        md.line((x, 0, x, 600 * S), fill=a)
    img.paste(flat, (0, 0), m)
    return img


def wave_bottom(x):
    """緑の面の下端（ゆるい弧）。"""
    return 598 - 16 * math.sin(math.pi * x / W)


def cta_compact(c, d, by1, by2):
    """予約ボタン（下の中央に横長・少し低め）。"""
    bx1, bx2 = 60, 980
    rrect(d, (bx1, by1 + 7, bx2, by2 + 7), 34, fill=GREEN_DARK)
    grad = Image.new("RGB", (int((bx2 - bx1) * S), int((by2 - by1) * S)), TEAL)
    gd = ImageDraw.Draw(grad)
    for yy in range(grad.height):
        t = yy / grad.height
        gd.line((0, yy, grad.width, yy), fill=tuple(int(a + (b - a) * t) for a, b in zip((0x2A, 0xC9, 0xAE), TEAL)))
    mask = Image.new("L", grad.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, grad.width - 1, grad.height - 1), radius=34 * S, fill=255)
    c.paste(grad, (bx1 * S, by1 * S), mask)
    text_c(d, "まずは無料カウンセリング", font(26), W / 2 - 30, by1 + 12, WHITE, 1)
    text_c(d, "予約する", font(58), W / 2 - 30, by1 + 44, WHITE, 6, stroke=1, stroke_fill=WHITE)
    ccx, ccy = 900, (by1 + by2) / 2
    ellipse(d, (ccx - 34, ccy - 34, ccx + 34, ccy + 34), fill=WHITE)
    d.line([((ccx - 7) * S, (ccy - 14) * S), ((ccx + 9) * S, ccy * S), ((ccx - 7) * S, (ccy + 14) * S)],
           fill=TEAL, width=7 * S, joint="curve")


def build(out):
    c = clean_base()
    d = ImageDraw.Draw(c)

    # 下の淡い面（ゆるい弧）
    pts = [(x * S, wave_bottom(x) * S) for x in range(0, W + 1, 8)]
    d.polygon(pts + [(W * S, H * S), (0, H * S)], fill=BG_PALE)

    # ロゴ（中央）。右端が写真の明るい所にかかるので、うしろに深緑のぼかし
    logo = white_logo(290 * S)
    lx, ly = (W * S - logo.width) // 2, 22 * S
    glow = Image.new("RGBA", logo.size, GREEN_DARK + (0,))
    glow.putalpha(logo.getchannel("A").filter(ImageFilter.MaxFilter(9)).filter(ImageFilter.GaussianBlur(10 * S)))
    c.paste(glow, (lx, ly), glow)
    c.paste(logo, (lx, ly), logo)

    # 診療のタグ3つ
    f_t = font(26)
    tx = 40
    for t in ["包茎治療", "亀頭増大", "ED治療"]:
        w_ = tw(f_t, t) + 32
        rrect(d, (tx, 104, tx + w_, 140), 18, fill=GREEN_DARK, outline=WHITE, width=2)
        text(d, t, f_t, tx + 16, 122 - th(f_t, t) / 2, WHITE)
        tx += w_ + 10

    # 心の声（考えごとの吹き出し・男性の頭に向けて点を2つ）
    f_v = font(36)
    v = "「何とかしたいなぁ…」"  # 2026-10-14 殿村さん「前提、なんとかしたいなぁ のほうが良い」
    vw = tw(f_v, v) + 56
    rrect(d, (40, 156, 40 + vw, 230), 37, fill=WHITE)
    text(d, v, f_v, 68, 193 - th(f_v, v) / 2, GREEN_DARK)
    ellipse(d, (40 + vw + 8, 206, 40 + vw + 28, 226), fill=WHITE)
    ellipse(d, (40 + vw + 36, 222, 40 + vw + 48, 234), fill=WHITE)

    # 主役3行（左寄せ）
    left = 40
    f_1 = font(60)
    s1 = "その“何とか”を、"
    y1 = 252
    text(d, s1, f_1, left + 3, y1 + 3, GREEN_DARK, 1, stroke=6, stroke_fill=GREEN_DARK)
    text(d, s1, f_1, left, y1, WHITE, 1, stroke=1.2, stroke_fill=WHITE)
    f_2 = font(104)
    a, b, s3 = "医師が", "一緒に", "考えます。"
    y2 = y1 + th(f_1, s1) + 14
    y3 = y2 + th(f_2, a + b) + 12
    x = left
    for s, col in ((a, WHITE), (b, ORANGE)):
        text(d, s, f_2, x + 4, y2 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text(d, s, f_2, x, y2, col, 2, stroke=8 if col == ORANGE else 1.5, stroke_fill=WHITE)
        x += tw(f_2, s, 2) + 12
    text(d, s3, f_2, left + 4, y3 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)
    text(d, s3, f_2, left, y3, WHITE, 2, stroke=1.5, stroke_fill=WHITE)
    shadow_bottom = y3 + th(f_2, s3) + 14
    assert shadow_bottom <= wave_bottom(left + 250) - 30, shadow_bottom

    # カウンセリング無料（顔にかからない右下＝肩のあたり）
    counsel_badge(d, 952, 470, 80)

    # 円4つ（元の画像と同じく、緑と淡い面の境目にまたがせる）
    dia, gap, top = 202, 14, 584
    x0 = (W - (dia * 4 + gap * 3)) / 2
    for i in range(4):
        cx, cy = x0 + dia / 2 + i * (dia + gap), top + dia / 2
        r = dia / 2
        ellipse(d, (cx - r, cy - r, cx + r, cy + r), fill=WHITE, outline=RING, width=9)
        ellipse(d, (cx - 34, cy - 88, cx + 34, cy - 20), fill=GREEN)
        icon = [icon_person, icon_pin, icon_chart, icon_calendar_check][i]
        icon(d, cx, cy - 54)
        if i == 0:
            text_c(d, "男性スタッフ", font(27), cx, cy - 4, INK)
            text_c(d, "のみ", font(36), cx, cy + 32, ORANGE)
        elif i == 1:
            text_c(d, "全国に", font(27), cx, cy - 4, INK)
            segs_c(d, [("36", 42, ORANGE), ("院", 28, ORANGE)], cx, cy + 70)
        elif i == 2:
            segs_c(d, [("37", 40, ORANGE), ("年目", 26, INK)], cx, cy + 34)
            segs_c(d, [("診療実績", 20, INK), ("40", 32, ORANGE), ("万件", 20, INK)], cx, cy + 72)
        else:
            text_c(d, "その日に", font(22), cx, cy - 6, INK)
            text_c(d, "決めなくても", font(22), cx, cy + 20, INK)
            text_c(d, "大丈夫", font(32), cx, cy + 48, ORANGE)

    # 当日の流れ
    fy1, fy2 = 796, 840
    rrect(d, (40, fy1, 1000, fy2), 22, fill=WHITE, outline=GREEN, width=2)
    f_s = font(23)
    rrect(d, (46, fy1 + 5, 190, fy2 - 5), 17, fill=GREEN)
    text_c(d, "当日の流れ", f_s, 118, (fy1 + fy2) / 2 - th(f_s, "当日の流れ") / 2, WHITE)
    steps = ["①ご予約", "②スタッフと1対1", "③治療法やご不安をご相談"]
    total = sum(tw(f_s, s_) for s_ in steps) + 28 * 2 + 18 * 2
    x = 190 + (1000 - 190 - total) / 2
    for i, s_ in enumerate(steps):
        text(d, s_, f_s, x, (fy1 + fy2) / 2 - th(f_s, s_) / 2, GREEN_DARK)
        x += tw(f_s, s_) + 18
        if i < 2:
            ay = (fy1 + fy2) / 2
            d.polygon([(x * S, (ay - 9) * S), ((x + 13) * S, ay * S), (x * S, (ay + 9) * S)], fill=ORANGE)
            x += 28

    # 予約ボタンの上の一言・予約ボタン・注記
    text_c(d, "＼電話なしで、このままWEB予約できます／", font(31), W / 2, 850, GREEN_DARK, 1)
    cta_compact(c, d, 892, 994)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(19), W / 2, 1012, GRAY)

    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_1016_banner_nantoka_illust_v2_1040.jpg"))
