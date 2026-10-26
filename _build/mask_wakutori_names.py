# -*- coding: utf-8 -*-
"""ワクトリ資料のLINE画面スクショから、あいさつの宛名（殿村さんのLINEの表示名）を「〇〇」に伏せる。

    python -m _build.mask_wakutori_names

  元の画像（_data/wakutori/line/*.jpg）はそのまま。伏せた画像を _data/wakutori/line/伏せ字/ に書き出す。
  build_wakutori_v1 は、伏せ字/ に同じ名前の画像があればそちらを使う。
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SRC = Path("_data/wakutori/line")
OUT = SRC / "伏せ字"
FONT = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
NAME = (40, 96, 122, 110)            # 宛名「殿村亮太@DYM」の範囲（277×600の画面で3枚とも同じ位置）
FILES = ["ダイブ_01.jpg", "アルファ_01.jpg", "ワクトリ_01.jpg"]


def mask(path):
    im = Image.open(path).convert("RGB")
    x0, y0, x1, y1 = NAME
    right = x1                                                    # 同じ行の残り（「さん、友だち追加」など）の右端
    for x in range(x1, im.width - 20):
        if any(sum(im.getpixel((x, y))) < 450 for y in range(y0, y1)):
            right = x + 2
    rest = im.crop((x1, y0, right, y1))
    d = ImageDraw.Draw(im)
    d.rectangle((x0, y0, right, y1), fill=(255, 255, 255))
    d.text((x0 + 1, y0 + 1), "〇〇", font=ImageFont.truetype(FONT, 11), fill=(40, 40, 40))
    im.paste(rest, (x0 + 24, y0))                                  # 宛名を詰めて、「さん」を「〇〇」の直後に
    return im


def main():
    OUT.mkdir(exist_ok=True)
    for f in FILES:
        mask(SRC / f).save(OUT / f, quality=95)
        print("saved:", OUT / f)


if __name__ == "__main__":
    main()
