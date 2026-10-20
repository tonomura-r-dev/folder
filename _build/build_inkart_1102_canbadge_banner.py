"""インクアート 11/2 オリジナル缶バッジ（新商品紹介）の配信用バナー（1040×1040）。

生成物 `_images/inkart_1102_canbadge_banner_generated_v1.webp`（1254×1254・文字8か所OK）の文字の部分
（上の帯・主役・副・札3つ・ボタン）はそのまま残し、真ん中の生成の缶バッジ8個（2026-10-20 殿村さん「缶バッジ感がなさすぎる」）を
サイトの写真 `_data/printearth/site_badge_types_large_20261020.png`（9個を散らした物・848×487）に差し替える。
ランチョンマット（10/7）と同じやり方＝文字は残して写真だけ入れ替える。

使い方:
  python3 _build/build_inkart_1102_canbadge_banner.py [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
GEN = ROOT / "_images" / "inkart_1102_canbadge_banner_generated_v1.webp"
PHOTO = ROOT / "_data" / "printearth" / "site_badge_types_large_20261020.png"
OUT = ROOT / "_images" / "inkart_1102_canbadge_banner_site_1040.jpg"

# 生成物（1254基準）の実測：副の文字の下端 366／札3つの上端 959。缶バッジはその間（381〜約955）
BAND_TOP, BAND_BOTTOM = 374, 952


def crop_photo(img):
    """スクショの黒い枠（RGB 14,14,14）を落として、写真の部分だけにする。"""
    a = np.array(img.convert("RGB")).astype(int)
    bright = a.max(axis=2) > 60
    ys = np.where(bright.any(axis=1))[0]
    xs = np.where(bright.any(axis=0))[0]
    return img.convert("RGB").crop((xs[0], ys[0], xs[-1] + 1, ys[-1] + 1))


def rounded_mask(size, radius):
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius, fill=255)
    return m


def build(out=OUT):
    base = Image.open(GEN).convert("RGB")
    W, H = base.size  # 1254
    d = ImageDraw.Draw(base)
    d.rectangle((0, BAND_TOP, W, BAND_BOTTOM), fill=(255, 255, 255))  # 生成の缶バッジを消す

    photo = crop_photo(Image.open(PHOTO))
    band_h = BAND_BOTTOM - BAND_TOP - 24  # 上下に12pxずつ余白
    scale = band_h / photo.height
    pw, ph = int(photo.width * scale), band_h
    if pw > W - 60:  # 幅が入らなければ幅で合わせる
        scale = (W - 60) / photo.width
        pw, ph = W - 60, int(photo.height * scale)
    photo = photo.resize((pw, ph), Image.LANCZOS)
    px, py = (W - pw) // 2, BAND_TOP + 12 + (band_h - ph) // 2

    # 薄い影＋角丸の写真カード
    radius = 28
    shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.rounded_rectangle((px + 4, py + 8, px + pw + 4, py + ph + 8), radius, fill=(0, 0, 0, 70))
    shadow = shadow.filter(ImageFilter.GaussianBlur(10))
    base = Image.alpha_composite(base.convert("RGBA"), shadow)
    base.paste(photo, (px, py), rounded_mask((pw, ph), radius))

    base.convert("RGB").resize((1040, 1040), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    print(build(Path(sys.argv[1]) if len(sys.argv) > 1 else OUT))
