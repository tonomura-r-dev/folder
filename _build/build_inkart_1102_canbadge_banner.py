"""インクアート 11/2 オリジナル缶バッジ（新商品紹介）の配信用バナー（1040×1040）。

生成物 `_images/inkart_1102_canbadge_banner_generated_v1.webp`（1254×1254・文字8か所OK）の文字の部分
（上の帯・主役・副・札3つ・ボタン）はそのまま残し、真ん中の生成の缶バッジ8個（2026-10-20 殿村さん「缶バッジ感がなさすぎる」）を
サイトの写真 `_data/printearth/site_badge_types_large_20261020.png`（9個を散らした物・848×487）の缶バッジに差し替える。

- v1（--card）：写真をそのまま角丸のカードにして置いた版 → 殿村さん「だめすぎる」（灰色の四角を貼った見た目・下の1個が切れている）
- v2（引数なし）：写真から缶バッジを1個ずつ切り出し、白地に影を付けて2段に並べ直す（生成物の並びと同じ見た目）。
  下で切れている「EARTH TREKKING CLUB」の1個は使わない（8個）。

- --swap：生成物v2（殿村さんのプロンプト）の訴求3つの帯を缶バッジの上に移した版（2026-10-20 「訴求軸の文字が下のほうにある」）。
  プロンプトで並びを変えようとしたv3は札が動かず、絵も硬くなった（殿村さん「君のプロンプトの方が雑魚い」）ので、v2の画像で帯を入れ替える。

- --gray：--swap に加えて、缶バッジの後ろの白を薄い灰色（#ECECEC）にして境界を出す（2026-10-20 殿村さん「ええやないか／背景少し色変えたい」）。

使い方:
  python3 _build/build_inkart_1102_canbadge_banner.py [--card | --swap | --gray] [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter
import numpy as np
import cv2

ROOT = Path(__file__).resolve().parent.parent
GEN = ROOT / "_images" / "inkart_1102_canbadge_banner_generated_v1.webp"
PHOTO = ROOT / "_data" / "printearth" / "site_badge_types_large_20261020.png"
OUT = ROOT / "_images" / "inkart_1102_canbadge_banner_site_v2_1040.jpg"
GEN2 = ROOT / "_images" / "inkart_1102_canbadge_banner_generated_v2.webp"
OUT_SWAP = ROOT / "_images" / "inkart_1102_canbadge_banner_v2_swap_1040.jpg"
OUT_GRAY = ROOT / "_images" / "inkart_1102_canbadge_banner_v2_swap_gray_1040.jpg"
GRAY = 0xEC  # 缶バッジの後ろの薄い灰色（2026-10-20 殿村さん「うすい灰色とかで背景を色付け」）
OUT_CARD = ROOT / "_images" / "inkart_1102_canbadge_banner_site_1040.jpg"

# 生成物（1254基準）の実測：副の文字の下端 366／札3つの上端 959。缶バッジはその間（381〜約955）
BAND_TOP, BAND_BOTTOM = 374, 952
BG = np.array([239, 239, 232])  # 写真の背景色

# 写真（黒い枠を除いた 820×479 の座標）の中の缶バッジ8個の枠（cv2の連結成分で実測・50thは白いので3つに割れた分を合わせた）
BADGES = {
    "aquarium_rect": (129, 12, 254, 151),
    "triangle": (411, 6, 172, 172),
    "circle_50th": (609, 41, 182, 183),
    "star": (5, 167, 204, 195),
    "uniform": (211, 182, 118, 108),
    "osaka": (355, 202, 115, 120),
    "circle_small": (495, 198, 137, 139),
    "dog_egg": (620, 242, 176, 226),
    "hexagon_70": (176, 305, 196, 169),
}
ROW1 = ["aquarium_rect", "triangle", "circle_50th", "star"]
ROW2 = ["uniform", "osaka", "circle_small", "dog_egg", "hexagon_70"]


def crop_photo(img):
    """スクショの黒い枠（RGB 14,14,14）を落として、写真の部分だけにする。"""
    a = np.array(img.convert("RGB")).astype(int)
    bright = a.max(axis=2) > 60
    ys = np.where(bright.any(axis=1))[0]
    xs = np.where(bright.any(axis=0))[0]
    return img.convert("RGB").crop((xs[0], ys[0], xs[-1] + 1, ys[-1] + 1))


def sprite(photo, box, pad=6):
    """写真から缶バッジ1個を切り出し、形の外を透明にしたRGBAを返す。
    白いバッジ（50th・OSAKA・ユニフォーム）は中身が背景色に近いので、外側の輪郭を見つけて中を塗りつぶす。"""
    x, y, w, h = box
    region = photo.crop((max(0, x - pad), max(0, y - pad), x + w + pad, y + h + pad))
    a = np.array(region).astype(int)
    dist = np.abs(a - BG).sum(axis=2)
    m = (dist > 45).astype(np.uint8)
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, np.ones((7, 7), np.uint8))
    contours, _ = cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    filled = np.zeros_like(m)
    big = [c for c in contours if cv2.contourArea(c) > 300]
    cv2.drawContours(filled, big, -1, 1, thickness=cv2.FILLED)
    filled = cv2.morphologyEx(filled, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))
    alpha = Image.fromarray((filled * 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(0.8))
    out = region.convert("RGBA")
    out.putalpha(alpha)
    return out.crop(out.getbbox())


def place_row(base, sprites, cy, scale, gap):
    """1段ぶんの缶バッジを、横の中央にそろえて置く（影付き）。"""
    W = base.width
    scaled = [s.resize((int(s.width * scale), int(s.height * scale)), Image.LANCZOS) for s in sprites]
    total = sum(s.width for s in scaled) + gap * (len(scaled) - 1)
    x = (W - total) // 2
    for s in scaled:
        y = cy - s.height // 2
        shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
        sh = Image.new("RGBA", s.size, (0, 0, 0, 0))
        sh.putalpha(s.getchannel("A").point(lambda v: int(v * 0.35)))
        shadow.paste(sh, (x + 5, y + 9), sh)
        shadow = shadow.filter(ImageFilter.GaussianBlur(8))
        base.alpha_composite(shadow)
        base.alpha_composite(s, (x, y))
        x += s.width + gap


def build(out=OUT):
    base = Image.open(GEN).convert("RGBA")
    W, H = base.size  # 1254
    ImageDraw.Draw(base).rectangle((0, BAND_TOP, W, BAND_BOTTOM), fill=(255, 255, 255, 255))  # 生成の缶バッジを消す

    photo = crop_photo(Image.open(PHOTO))
    sp = {k: sprite(photo, v) for k, v in BADGES.items()}

    band_h = BAND_BOTTOM - BAND_TOP
    row_h = (band_h - 30) // 2  # 2段
    scale = 1.12  # 写真の大きさに対する倍率（いちばん背の高い犬のたまご型 226px → 約253px）
    place_row(base, [sp[k] for k in ROW1], BAND_TOP + 15 + row_h // 2, scale, 44)
    place_row(base, [sp[k] for k in ROW2], BAND_TOP + 15 + row_h + 30 + row_h // 2, scale, 36)

    base.convert("RGB").resize((1040, 1040), Image.LANCZOS).save(out, quality=95)
    return out


def rounded_mask(size, radius):
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius, fill=255)
    return m


def build_card(out=OUT_CARD):
    """v1：写真をそのまま角丸のカードにして置く版（不採用・記録用）。"""
    base = Image.open(GEN).convert("RGB")
    W, H = base.size
    ImageDraw.Draw(base).rectangle((0, BAND_TOP, W, BAND_BOTTOM), fill=(255, 255, 255))
    photo = crop_photo(Image.open(PHOTO))
    band_h = BAND_BOTTOM - BAND_TOP - 24
    scale = band_h / photo.height
    pw, ph = int(photo.width * scale), band_h
    if pw > W - 60:
        scale = (W - 60) / photo.width
        pw, ph = W - 60, int(photo.height * scale)
    photo = photo.resize((pw, ph), Image.LANCZOS)
    px, py = (W - pw) // 2, BAND_TOP + 12 + (band_h - ph) // 2
    radius = 28
    shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle((px + 4, py + 8, px + pw + 4, py + ph + 8), radius, fill=(0, 0, 0, 70))
    shadow = shadow.filter(ImageFilter.GaussianBlur(10))
    base = Image.alpha_composite(base.convert("RGBA"), shadow)
    base.paste(photo, (px, py), rounded_mask((pw, ph), radius))
    base.convert("RGB").resize((1040, 1040), Image.LANCZOS).save(out, quality=95)
    return out


def tint_background(base, y0, y1, gray=GRAY):
    """y0〜y1の帯の中で、縁から白くつながっている所（＝背景）だけを薄い灰色にする。
    缶バッジの中の白（ユニフォーム・ハートの地）は縁とつながっていないので残る。影の外側のにじみは、マスクをぼかして段差を消す。"""
    a = np.array(base.convert("RGB"))
    band = a[y0:y1].copy()
    h, w = band.shape[:2]
    near_white = (band.min(axis=2) >= 243).astype(np.uint8)
    # 縁から白い所へ流し込む（背景の連結成分）
    n, lab = cv2.connectedComponents(near_white, connectivity=4)
    edge_labels = set(lab[0, :]) | set(lab[-1, :]) | set(lab[:, 0]) | set(lab[:, -1])
    edge_labels.discard(0)
    m = np.isin(lab, list(edge_labels)).astype(np.float32)
    m = cv2.GaussianBlur(m, (0, 0), 1.5)
    k = gray / 255.0
    f = 1.0 - m[..., None] * (1.0 - k)
    a[y0:y1] = np.clip(band.astype(np.float32) * f, 0, 255).astype(np.uint8)
    return Image.fromarray(a)


def build_swap(out=OUT_SWAP, gray=False):
    """v2の生成物（殿村さんのプロンプト・2026-10-20「訴求軸の文字が下のほうにある」）で、
    訴求3つの帯（1254基準 y998〜1096）と缶バッジの部分（y345〜997）を上下に入れ替える。
    文字の部分（帯・主役・副・ボタン）は触らない。どちらも白地なので継ぎ目は出ない。入れ替え後に収めるため両方を0.95倍にする。"""
    base = Image.open(GEN2).convert("RGB")
    W, H = base.size
    badges = base.crop((0, 345, W, 998))
    labels = base.crop((0, 998, W, 1097))
    ImageDraw.Draw(base).rectangle((0, 340, W, 1100), fill=(255, 255, 255))
    k = 0.95
    labels = labels.resize((int(W * k), int(labels.height * k)), Image.LANCZOS)
    badges = badges.resize((int(W * k), int(badges.height * k)), Image.LANCZOS)
    gap = (1100 - 340 - labels.height - badges.height) // 3
    y = 340 + gap
    base.paste(labels, ((W - labels.width) // 2, y))
    y += labels.height + gap
    base.paste(badges, ((W - badges.width) // 2, y))
    if gray:
        # 札の下から、ボタンの上までを薄い灰色の帯にする（缶バッジの後ろだけ）
        base = tint_background(base, y - gap // 2, 1100)
    base.resize((1040, 1040), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if "--card" in sys.argv:
        print(build_card(Path(args[0]) if args else OUT_CARD))
    elif "--swap" in sys.argv:
        print(build_swap(Path(args[0]) if args else OUT_SWAP))
    elif "--gray" in sys.argv:
        print(build_swap(Path(args[0]) if args else OUT_GRAY, gray=True))
    else:
        print(build(Path(args[0]) if args else OUT))
