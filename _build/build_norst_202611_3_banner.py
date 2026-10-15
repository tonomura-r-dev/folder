"""ノースト 11月3本目（勤労感謝の日の3連休・配信日は未定）のバナー（1040×1040）。

アイテム名＝【年内最後の3連休】悩みと向き合う、絶好のチャンスです！（2026-10-15 殿村さん）。
吹き出し＝「休日に行ってみるか」（殿村さん。生成した画像の「平日は、仕事で時間が取れない…」から変更）。
主役＝年内最後の3連休に、／無料カウンセリング（オレンジ）／受けてみませんか？

- 引数なし：イラストなしの版。形は11月1本目のイラストなし版（`build_norst_202611_1_banner.py`）と同じで、上の深緑の面の言葉だけ差し替え。
- --illust：殿村さんが生成した男性入りの版（`_images/norst_202611_3_holiday_illust_v1.webp`）の吹き出しの文字を差し替え、公式ロゴを貼る。

- --worry：イラストなしの版で、吹き出しを悩みの心の声「平日は、仕事で時間が取れない…」にする（案A）。
- --band：イラストなしの版で、吹き出しをやめて「11/21〜23は、勤労感謝の日の3連休」の帯にする（案B）。

使い方:
  python3 _build/build_norst_202611_3_banner.py [--illust | --worry | --band] [出力]
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
ILLUST = ROOT / "_images" / "norst_202611_3_holiday_illust_v1.webp"
BUBBLE = "「休日に行ってみるか」"


def lower_half(c, d):
    """札4つ・当日の流れ・ボタンの上の一言・予約ボタン・注記（11月1本目と同じ）。"""
    cards = [
        (icon_person, ["男性スタッフ", "のみ"], [None, ORANGE]),
        (icon_pin, ["全国に", "36院"], [None, ORANGE]),
        (icon_chart, "JISSEKI", None),
        (icon_calendar_check, ["その日に決めなくても", "大丈夫"], ["small", ORANGE]),
    ]
    cw, gap, top, bot = 226, 12, 526, 694
    x0 = (W - (cw * 4 + gap * 3)) / 2
    f_c, f_cb, f_cs = font(27), font(34), font(21)
    for i, (icon, lines, colors) in enumerate(cards):
        x1 = x0 + i * (cw + gap)
        rrect(d, (x1, top, x1 + cw, bot), 14, fill=WHITE, outline=TEAL, width=3)
        icx, icy = x1 + cw / 2, top + 46
        ellipse(d, (icx - 37, icy - 37, icx + 37, icy + 37), fill=GREEN)
        icon(d, icx, icy)
        if lines == "JISSEKI":
            segs_c(d, [("37", 36, ORANGE), ("年目", 24, INK)], icx, top + 128)
            segs_c(d, [("診療実績", 21, INK), ("40", 34, ORANGE), ("万件", 21, INK)], icx, top + 162)
            continue
        y = top + 92
        for s_, col in zip(lines, colors):
            f = f_cs if col == "small" else (f_cb if col else f_c)
            text_c(d, s_, f, icx, y + (6 if col == "small" else 0), INK if col in (None, "small") else col)
            y += th(f_c, s_) + 7

    fy1, fy2 = 706, 764
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

    text_c(d, "＼電話なしで、このままWEB予約できます／", font(34), W / 2, 778, GREEN_DARK, 1)
    cta(c, d, 832, 966)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(22), W / 2, 996, GRAY)


def build(out, bubble=BUBBLE, band=False):
    """bubble＝吹き出しの言葉。band=True なら吹き出しのしっぽを付けず、事実の帯として置く。"""
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面・ロゴ（中央）・カウンセリング無料の円（右上）
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)
    logo = white_logo(290 * S)
    c.paste(logo, ((W * S - logo.width) // 2, 22 * S), logo)
    counsel_badge(d, 905, 128, 88)

    # 吹き出し＝読む人の心の声（band=True のときは、しっぽの無い白い帯）
    f_v = font(34)
    vw = tw(f_v, bubble) + 44
    rrect(d, (40, 104, 40 + vw, 160), 28, fill=WHITE)
    if not band:
        d.polygon([(70 * S, 158 * S), (98 * S, 158 * S), (66 * S, 180 * S)], fill=WHITE)
    text(d, bubble, f_v, 62, 132 - th(f_v, bubble) / 2, GREEN_DARK)

    # 診療のタグ3つ（吹き出しの下）
    f_t = font(26)
    tx = 118
    for t in ["包茎治療", "亀頭増大", "ED治療"]:
        tw_ = tw(f_t, t) + 32
        rrect(d, (tx, 172, tx + tw_, 208), 18, fill=GREEN_DARK, outline=WHITE, width=2)
        text(d, t, f_t, tx + 16, 190 - th(f_t, t) / 2, WHITE)
        tx += tw_ + 10

    # 主役3行：年内最後の3連休に、／無料カウンセリング（オレンジ）／受けてみませんか？
    f_1 = font(56)
    s1 = "年内最後の3連休に、"
    y1 = 222
    text_c(d, s1, f_1, W / 2, y1, WHITE, 1, stroke=1.2, stroke_fill=WHITE)

    f_2 = font(86)
    s2, s3 = "無料カウンセリング", "受けてみませんか？"
    y2 = y1 + th(f_1, s1) + 12
    y3 = y2 + th(f_2, s2) + 10
    for s, y, col in ((s2, y2, ORANGE), (s3, y3, WHITE)):
        text_c(d, s, f_2, W / 2 + 4, y + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text_c(d, s, f_2, W / 2, y, col, 2, stroke=8 if col == ORANGE else 1.5, stroke_fill=WHITE)
    assert tw(f_2, s2, 2) <= W - 60, tw(f_2, s2, 2)
    shadow_bottom = y3 + th(f_2, s3) + 14
    assert shadow_bottom <= PANEL_BOTTOM - 20, shadow_bottom

    lower_half(c, d)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


def paste_logo(out):
    """生成した男性入りの版に、吹き出しの文字の差し替えと公式ロゴの貼り込みをする。
    吹き出し（1040基準で x 35〜555・y 169〜244 の白い角丸）の中の文字を白で消して、心の声を載せ直す。
    ロゴは右側の写真にかからないよう、左の文字の列の上に置く。"""
    img = Image.open(ILLUST).convert("RGB").resize((W * S, H * S), Image.LANCZOS).convert("RGBA")
    d = ImageDraw.Draw(img)
    d.rectangle((44 * S, 184 * S, 546 * S, 230 * S), fill=WHITE)
    f_v = font(36)
    text_c(d, BUBBLE, f_v, (35 + 555) / 2, 206 - th(f_v, BUBBLE) / 2, GREEN_DARK)
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
        print(paste_logo(args[0] if args else "_images/norst_202611_3_holiday_illust_v2_logo_1040.jpg"))
    elif "--worry" in sys.argv:
        # イラストなし版の吹き出しを悩みの心の声に（2026-10-15 殿村さん「イラスト無しverの吹き出しはこれじゃまずい」）
        print(build(args[0] if args else "_images/norst_202611_3_holiday_v2a_1040.jpg", bubble="「平日は、仕事で時間が取れない…」"))
    elif "--band" in sys.argv:
        # 吹き出しをやめて、3連休の日付を事実の帯で置く
        print(build(args[0] if args else "_images/norst_202611_3_holiday_v2b_1040.jpg", bubble="11/21〜23は、勤労感謝の日の3連休", band=True))
    else:
        print(build(args[0] if args else "_images/norst_202611_3_holiday_v1_1040.jpg"))
