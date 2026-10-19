"""ノースト 11月2本目（11/13・【3連休の予定は？】何もせずに終えますか？）のバナー（1040×1040）。

吹き出し＝「また、何もせずに終わりそう…」（心の声・アイテム名と同じ問い）。
主役＝せっかくの3連休、／話す機会を（オレンジ）／作りませんか？
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
    S, W, H, GREEN, GREEN_DARK, ORANGE, PALE, WHITE, font, tw, th, text, text_c, rrect, white_logo,
)
from build_norst_1016_banner import counsel_badge, PANEL_BOTTOM  # noqa: E402
from build_norst_202611_3_banner import lower_half  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
ILLUST = ROOT / "_images" / "norst_202611_2_holiday_illust_v1.webp"
BUBBLE = "「また、何もせずに終わりそう…」"
NOTE = "＼予約だけ先に。電話なしでWEB予約できます／"


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面・ロゴ（中央）・カウンセリング無料の円（右上）
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)
    logo = white_logo(290 * S)
    c.paste(logo, ((W * S - logo.width) // 2, 22 * S), logo)
    counsel_badge(d, 905, 128, 88)

    # 吹き出し＝読む人の心の声
    f_v = font(34)
    vw = tw(f_v, BUBBLE) + 44
    rrect(d, (40, 104, 40 + vw, 160), 28, fill=WHITE)
    d.polygon([(70 * S, 158 * S), (98 * S, 158 * S), (66 * S, 180 * S)], fill=WHITE)
    text(d, BUBBLE, f_v, 62, 132 - th(f_v, BUBBLE) / 2, GREEN_DARK)

    # 診療のタグ3つ
    f_t = font(26)
    tx = 118
    for t in ["包茎治療", "亀頭増大", "ED治療"]:
        tw_ = tw(f_t, t) + 32
        rrect(d, (tx, 172, tx + tw_, 208), 18, fill=GREEN_DARK, outline=WHITE, width=2)
        text(d, t, f_t, tx + 16, 190 - th(f_t, t) / 2, WHITE)
        tx += tw_ + 10

    # 主役3行：せっかくの3連休、／話す機会を（オレンジ）／作りませんか？
    f_1 = font(56)
    s1 = "せっかくの3連休、"
    y1 = 222
    text_c(d, s1, f_1, W / 2, y1, WHITE, 1, stroke=1.2, stroke_fill=WHITE)

    f_2 = font(86)
    s2, s3 = "話す機会を", "作りませんか？"
    y2 = y1 + th(f_1, s1) + 12
    y3 = y2 + th(f_2, s2) + 10
    for s, y, col in ((s2, y2, ORANGE), (s3, y3, WHITE)):
        text_c(d, s, f_2, W / 2 + 4, y + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text_c(d, s, f_2, W / 2, y, col, 2, stroke=8 if col == ORANGE else 1.5, stroke_fill=WHITE)
    shadow_bottom = y3 + th(f_2, s3) + 14
    assert shadow_bottom <= PANEL_BOTTOM - 20, shadow_bottom

    lower_half(c, d, note=NOTE)
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
