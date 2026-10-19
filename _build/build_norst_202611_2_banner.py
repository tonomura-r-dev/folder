"""ノースト 11月2本目（【3連休の予定は？】何もせずに終えますか？）の男性入りバナーに公式ロゴを貼る。

生成物 `_images/norst_202611_2_holiday_illust_v1.webp`（1254×1254・文字は全部一致・上の左は空き）に、
3本目と同じ置き方（左の文字の列の上・タグ3つの左端にそろえる・横270px・深緑のぼかしを下に敷く）でロゴを貼り、1040×1040にする。

使い方:
  python3 _build/build_norst_202611_2_banner.py [出力]
"""
import sys
from PIL import Image, ImageDraw, ImageFilter

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from build_norst_202611_3_banner import white_logo, GREEN_DARK, W, H, S  # noqa: E402

ILLUST = "_images/norst_202611_2_holiday_illust_v1.webp"


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
    print(paste_logo(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_202611_2_holiday_illust_v1_logo_1040.jpg"))
