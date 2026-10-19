"""ノースト 11月2本目（11/13・【3連休の予定は？】何もせずに終えますか？）のバナー（1040×1040）。

吹き出し＝「また、何もせずに終わりそう…」（心の声・アイテム名と同じ問い）。
主役＝3連休、話す機会を（オレンジ）／作りませんか？（v2・2026-10-19 殿村さん。v1は「せっかくの3連休、／話す機会を／作りませんか？」の3行）
ボタンの上＝＼予約だけ先に。電話なしでWEB予約できます／。それ以外（札4つ・当日の流れ・ボタン・注記）は1本目・3本目と同じ。

- 引数なし：イラストなしの版（2026-10-19 殿村さん「イラスト無しにしてよ」）。形は3本目のイラストなし版と同じ。
- --illust：殿村さんが生成した男性入りの版（`_images/norst_202611_2_holiday_illust_v1.webp`）に公式ロゴを貼る（左上・横270px）。

使い方:
  python3 _build/build_norst_202611_2_banner.py [--illust] [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN, GREEN_DARK, TEAL, ORANGE, PALE, INK, GRAY, WHITE,
    font, tw, th, text, text_c, rrect, ellipse, white_logo, icon_person, icon_pin,
)
from build_norst_rough_patterns import cta  # noqa: E402
from build_norst_1016_banner import (  # noqa: E402
    icon_calendar_check, icon_chart, segs_c, counsel_badge, PANEL_BOTTOM,
)

ROOT = Path(__file__).resolve().parent.parent
ILLUST = ROOT / "_images" / "norst_202611_2_holiday_illust_v1.webp"
BUBBLE = "「また、何もせずに終わりそう…」"
NOTE = "＼予約だけ先に。電話なしでWEB予約できます／"
PANEL = 452  # 深緑の面の下端（v3・2026-10-19 殿村さん「中央にある緑の線もっと上に上げろよ」。1本目・3本目は516）


def lower_half_big(c, d):
    """札4つ・当日の流れ・ボタンの上の一言・予約ボタン・注記。3本目の lower_half より札と文字を大きくした版
    （2026-10-19 殿村さん「実績の枠・文字をもう少し大きくしてほしい」）。"""
    cards = [
        (icon_person, ["男性スタッフ", "のみ"], [None, ORANGE]),
        (icon_pin, ["全国に", "36院"], [None, ORANGE]),
        (icon_chart, "JISSEKI", None),
        (icon_calendar_check, ["その日に決めなくても", "大丈夫"], ["small", ORANGE]),
    ]
    cw, gap, top, bot = 238, 8, 466, 690
    x0 = (W - (cw * 4 + gap * 3)) / 2
    f_c, f_cb, f_cs = font(32), font(42), font(22)
    for i, (icon, lines, colors) in enumerate(cards):
        x1 = x0 + i * (cw + gap)
        rrect(d, (x1, top, x1 + cw, bot), 14, fill=WHITE, outline=TEAL, width=3)
        icx, icy = x1 + cw / 2, top + 56
        ellipse(d, (icx - 44, icy - 44, icx + 44, icy + 44), fill=GREEN)
        icon(d, icx, icy)
        if lines == "JISSEKI":
            segs_c(d, [("37", 46, ORANGE), ("年目", 30, INK)], icx, top + 152)
            segs_c(d, [("診療実績", 25, INK), ("40", 42, ORANGE), ("万件", 25, INK)], icx, top + 196)
            continue
        y = top + 112
        for s_, col in zip(lines, colors):
            f = f_cs if col == "small" else (f_cb if col else f_c)
            text_c(d, s_, f, icx, y + (8 if col == "small" else 0), INK if col in (None, "small") else col)
            y += th(f_c, s_) + 12

    fy1, fy2 = 702, 758
    rrect(d, (40, fy1, 1000, fy2), 30, fill=WHITE, outline=GREEN, width=2)
    f_l = f_s = font(25)
    rrect(d, (48, fy1 + 7, 200, fy2 - 7), 24, fill=GREEN)
    text_c(d, "当日の流れ", f_l, 124, (fy1 + fy2) / 2 - th(f_l, "当日の流れ") / 2, WHITE)
    steps = ["①ご予約", "②スタッフと1対1", "③治療法やご不安をご相談"]
    total = sum(tw(f_s, s_) for s_ in steps) + 30 * 2 + 20 * 2
    x = 200 + (1000 - 200 - total) / 2
    for i, s_ in enumerate(steps):
        text(d, s_, f_s, x, (fy1 + fy2) / 2 - th(f_s, s_) / 2, GREEN_DARK)
        x += tw(f_s, s_) + 20
        if i < 2:
            ay = (fy1 + fy2) / 2
            d.polygon([(x * S, (ay - 10) * S), ((x + 14) * S, ay * S), (x * S, (ay + 10) * S)], fill=ORANGE)
            x += 30

    text_c(d, NOTE, font(34), W / 2, 774, GREEN_DARK, 1)
    cta(c, d, 832, 966)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(22), W / 2, 996, GRAY)


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面・ロゴ（中央）・カウンセリング無料の円（右上）
    d.rectangle((0, 0, W * S, PANEL * S), fill=GREEN)
    logo = white_logo(290 * S)
    c.paste(logo, ((W * S - logo.width) // 2, 22 * S), logo)
    counsel_badge(d, 905, 128, 88)

    # 吹き出し＝読む人の心の声
    f_v = font(34)
    vw = tw(f_v, BUBBLE) + 44
    rrect(d, (40, 104, 40 + vw, 160), 28, fill=WHITE)
    d.polygon([(70 * S, 158 * S), (98 * S, 158 * S), (66 * S, 180 * S)], fill=WHITE)
    text(d, BUBBLE, f_v, 62, 132 - th(f_v, BUBBLE) / 2, GREEN_DARK)

    # 診療のタグ3つ（真ん中にそろえる。2026-10-19 殿村さん「診療例は真ん中にして」）
    f_t = font(26)
    tags = ["包茎治療", "亀頭増大", "ED治療"]
    widths = [tw(f_t, t) + 32 for t in tags]
    tx = (W - (sum(widths) + 10 * (len(tags) - 1))) / 2
    for t, tw_ in zip(tags, widths):
        rrect(d, (tx, 172, tx + tw_, 208), 18, fill=GREEN_DARK, outline=WHITE, width=2)
        text(d, t, f_t, tx + 16, 190 - th(f_t, t) / 2, WHITE)
        tx += tw_ + 10

    # 主役2行：3連休、話す機会を（オレンジ）／作りませんか？（2026-10-19 殿村さん「少し上に寄せて」）
    f_2 = font(86)
    s2, s3 = "3連休、話す機会を", "作りませんか？"
    y2 = 226
    y3 = y2 + th(f_2, s2) + 12
    for s, y, col in ((s2, y2, ORANGE), (s3, y3, WHITE)):
        text_c(d, s, f_2, W / 2 + 4, y + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text_c(d, s, f_2, W / 2, y, col, 2, stroke=8 if col == ORANGE else 1.5, stroke_fill=WHITE)
    assert tw(f_2, s2, 2) <= W - 60, tw(f_2, s2, 2)
    shadow_bottom = y3 + th(f_2, s3) + 14
    assert shadow_bottom <= PANEL - 20, shadow_bottom

    lower_half_big(c, d)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


def paste_logo(out):
    img = Image.open(ILLUST).convert("RGB").resize((W * S, H * S), Image.LANCZOS).convert("RGBA")
    logo = white_logo(270 * S)
    lx, ly = 40 * S, 22 * S  # タグ3つの左端にそろえる
    glow = Image.new("RGBA", logo.size, GREEN_DARK + (0,))
    glow.putalpha(logo.getchannel("A").filter(ImageFilter.MaxFilter(7)).filter(ImageFilter.GaussianBlur(6 * S)))
    img.alpha_composite(glow, (lx, ly))
    img.alpha_composite(logo, (lx, ly))
    img.convert("RGB").resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if "--illust" in sys.argv:
        print(paste_logo(args[0] if args else "_images/norst_202611_2_holiday_illust_v1_logo_1040.jpg"))
    else:
        print(build(args[0] if args else "_images/norst_202611_2_holiday_v1_1040.jpg"))
