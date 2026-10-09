"""ノースト 10月2本目のバナーのラフ（1040×1040）。C＝自分のための一言 の形。

狙う人（2026-10-08 殿村さん）＝31歳前後の会社員・自分に自信がない・趣味は筋トレやドライブ。
主役＝「体は、鍛えてきた。／でも、自信は／まだ持てない。」（読み手の気持ちを言い当てる。結果は約束しない）。
必ず残す4つ（緑が主・37年目・診療実績40万件・ロゴ）は入れている。写真の場所は仮の枠（ラフなので）。

使い方:
  python3 _build/build_norst_rough_confidence.py [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN_DARK, TEAL, ORANGE, PALE, WHITE,
    font, th, text, text_c, rrect, ellipse, icon_door, icon_person, icon_pin,
)
from build_norst_rough_patterns import head, cta, footer  # noqa: E402
from build_norst_1019_rough import GREEN  # noqa: E402

MINT = (0xCF, 0xEF, 0xE8)


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)
    head(c, d, 482)

    # 主役の上（何の話か）
    text(d, "包茎治療をお考えの方へ", font(32), 44, 112, MINT, 2)
    # 主役（3行）
    text(d, "体は、鍛えてきた。", font(64), 44, 164, WHITE, 2, stroke=1.2, stroke_fill=WHITE)
    text(d, "でも、自信は", font(92), 44, 252, WHITE, 2, stroke=1.5, stroke_fill=WHITE)
    s3 = "まだ持てない。"
    f3 = font(108)
    text(d, s3, f3, 44 + 4, 358 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)
    text(d, s3, f3, 44, 358, ORANGE, 2, stroke=8, stroke_fill=WHITE)

    # 写真の場所（ラフなので仮の枠）
    px1, py1, px2, py2 = 40, 504, 1000, 706
    rrect(d, (px1, py1, px2, py2), 14, fill=(0xD9, 0xE8, 0xE5))
    for x in range(px1 + 10, px2 - 10, 24):
        d.line((x * S, (py1 + 6) * S, (x + 12) * S, (py1 + 6) * S), fill=(0x9F, 0xBF, 0xB8), width=2 * S)
        d.line((x * S, (py2 - 6) * S, (x + 12) * S, (py2 - 6) * S), fill=(0x9F, 0xBF, 0xB8), width=2 * S)
    text_c(d, "写真：ジムのベンチに座る男性の後ろ姿", font(28), W / 2, 568, (0x5F, 0x80, 0x7A))
    text_c(d, "（肩にタオル・背中を少し丸める・顔は出さない）", font(24), W / 2, 616, (0x5F, 0x80, 0x7A))

    # 安心の事実（小さく3つ）
    chips = [(icon_person, "男性スタッフのみ"), (icon_door, "完全個室"), (icon_pin, "全国に36院")]
    cw, gap, top = 300, 30, 724
    x0 = (W - (cw * 3 + gap * 2)) / 2
    f = font(25)
    for i, (icon, s) in enumerate(chips):
        x1 = x0 + i * (cw + gap)
        rrect(d, (x1, top, x1 + cw, top + 70), 35, fill=WHITE, outline=TEAL, width=2.5)
        ellipse(d, (x1 + 10, top + 8, x1 + 64, top + 62), fill=GREEN)
        ic = Image.new("RGBA", (120 * S, 120 * S), (0, 0, 0, 0))
        icon(ImageDraw.Draw(ic), 60, 60)
        ic = ic.resize((int(66 * S), int(66 * S)), Image.LANCZOS)
        c.paste(ic, (int((x1 + 4) * S), int((top + 2) * S)), ic)
        text(d, s, f, x1 + 80, top + 35 - th(f, s) / 2, GREEN_DARK)

    cta(c, d, 812, 952)
    footer(d)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "_images/norst_rough_confidence_1040.jpg"
    print(build(out))
