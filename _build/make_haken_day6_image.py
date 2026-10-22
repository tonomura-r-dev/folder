# -*- coding: utf-8 -*-
"""人材：実文面 Day0・Day6 の画面イメージの Day6 を「相談窓口」から「派遣のしくみ」に書き換える（2026-10-22 殿村さん「相談先は外して」）。

    python _build/make_haken_day6_image.py <元画像.png> <フォントのフォルダ> <出力.png>

  元画像は殿村さんのPC保存版ver2（P25）の画面イメージ（1024×1536）。フォントは Noto Sans JP（fontsource の woff2 を ttf にしたもの）。
  吹き出しの文字と、下のボタンの文字だけを塗り直して書き直す。文字の大きさ・行の間隔・色は元の画像に合わせて測った値。
"""
import sys

from PIL import Image, ImageDraw, ImageFont

LINES = ["派遣で働く前に知っておきたい「3年ルール」", "と契約期間のしくみを、かんたんにまとめ",
         "ました。お仕事選びの参考に、ご確認くだ", "さい。"]
BUTTON = "派遣のしくみを見る"
INK, GREEN, BUBBLE = (46, 46, 51), (9, 193, 96), (254, 253, 254)


def draw_mixed(d, x, base, text, jp, latin, fill, anchor_center=None):
    """日本語は日本語のフォント、英数字は英字のフォントで書く"""
    segs = []
    for ch in text:
        f = latin if ord(ch) < 0x2000 else jp
        if segs and segs[-1][1] is f:
            segs[-1][0] += ch
        else:
            segs.append([ch, f])
    width = sum(f.getlength(t) for t, f in segs)
    if anchor_center is not None:
        x = anchor_center - width / 2
    for t, f in segs:
        d.text((x, base), t, font=f, fill=fill, anchor="ls")
        x += f.getlength(t)


def main(src, fonts, out):
    im = Image.open(src).convert("RGB")
    assert im.size == (1024, 1536), im.size
    d = ImageDraw.Draw(im)
    jp = ImageFont.truetype(f"{fonts}/noto-sans-jp-japanese-500-normal.ttf", 37)
    la = ImageFont.truetype(f"{fonts}/noto-sans-jp-latin-500-normal.ttf", 37)
    d.rectangle((70, 1022, 850, 1262), fill=BUBBLE)                  # 吹き出しの文字（元は4行・行の上端1034／1092／1152／1214）
    for k, t in enumerate(LINES):
        draw_mixed(d, 81, 1064 + 60 * k, t, jp, la, INK)
    jb = ImageFont.truetype(f"{fonts}/noto-sans-jp-japanese-700-normal.ttf", 39)
    lb = ImageFont.truetype(f"{fonts}/noto-sans-jp-latin-700-normal.ttf", 39)
    d.rectangle((150, 1340, 880, 1402), fill=(254, 254, 254))       # ボタンの文字（元は x356〜666・上端1352）
    draw_mixed(d, 0, 1385, BUTTON, jb, lb, GREEN, anchor_center=511)
    im.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
