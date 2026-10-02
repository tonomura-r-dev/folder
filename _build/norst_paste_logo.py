#!/usr/bin/env python3
"""ノーストの公式ロゴ（白抜き）を、生成したバナーのいちばん上に貼り込む。

使い方:
  python3 _build/norst_paste_logo.py <生成した画像> [出力先] [--width 0.42] [--top 0.025] [--clear]

- 元のロゴ `_images/norst_logo_original.png`（244×57・緑地に白文字）から白い文字だけを抜き出して貼る。
  緑地ごと貼ると、バナーの深緑と色がずれて四角が見えるため。
- 画像は 1040×1040 にそろえてから貼る（生成物は 1254×1254 などで来ることがある）。
- --width はロゴの横幅（画面の幅に対する割合）、--top は上からの位置（画面の高さに対する割合）。
- --clear を付けると、ロゴを置く場所を周りの深緑で塗ってから貼る（生成物がそこに何か描いていたとき用）。
- 元のロゴが小さい（244×57）ので拡大すると少しぼやける。最終の配信用は先方の高解像度のロゴに差し替える。
"""
import sys
from pathlib import Path
from PIL import Image, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
LOGO = ROOT / "_images" / "norst_logo_original.png"
SIZE = 1040


def white_logo(width_px):
    src = Image.open(LOGO).convert("RGB")
    # 緑地（約 #029C84）と白の明るさの差から、白い文字の濃さ（アルファ）を作る
    gray = src.convert("L")
    bg_l = 107
    alpha = gray.point(lambda v: max(0, min(255, int((v - bg_l) * 255 / (255 - bg_l)))))
    scale = width_px / src.width
    size = (round(src.width * scale), round(src.height * scale))
    alpha = alpha.resize(size, Image.LANCZOS)
    # 拡大でぼやけた縁を少し締める
    alpha = alpha.filter(ImageFilter.UnsharpMask(radius=1.2, percent=120, threshold=0))
    logo = Image.new("RGBA", size, (255, 255, 255, 0))
    logo.putalpha(alpha)
    return logo


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    opts = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    src = Path(args[0])
    out = Path(args[1]) if len(args) > 1 else src.with_name(src.stem + "_logo_1040.jpg")
    width = float(opts[opts.index("--width") + 1]) if "--width" in opts else 0.42
    top = float(opts[opts.index("--top") + 1]) if "--top" in opts else 0.025

    img = Image.open(src).convert("RGB").resize((SIZE, SIZE), Image.LANCZOS)
    logo = white_logo(round(SIZE * width))
    x = (SIZE - logo.width) // 2
    y = round(SIZE * top)
    if "--clear" in opts:
        pad = 12
        box = (x - pad, y - pad, x + logo.width + pad, y + logo.height + pad)
        # 置き場所の外側の帯から深緑を拾って塗る
        strip = img.crop((0, box[1], SIZE, box[3]))
        colors = sorted(strip.getdata(), key=lambda c: sum(c))
        med = colors[len(colors) // 2]
        img.paste(med, box)
    img.paste(logo, (x, y), logo)
    img.save(out, quality=95)
    print(f"{out}  ロゴ {logo.width}×{logo.height}px を ({x},{y}) に貼った")


if __name__ == "__main__":
    main()
