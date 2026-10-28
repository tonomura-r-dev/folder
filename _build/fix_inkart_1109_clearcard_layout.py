"""インクアート クリアカードの確定バナー（空にかざす版）の下側の並びを直す。
殿村さん 2026-10-28：
  上＝値段（100枚 9,780円〜（税別））を中央に
  下＝「名刺・ショップカードに」「しおり・ポストカード・チケットにも」を横に並べる
生成し直さず、生成物（1254×1254）の札を切り出して置き直す。空いた所は OpenCV の inpaint で周りの色で埋める。

使い方: PYTHONPATH=<pylib> python3 _build/fix_inkart_1109_clearcard_layout.py
出力:  _images/inkart_1109_clearcard_banner_sky_v2_1040.jpg（配信用）／_sky_v2_1254.png（元の大きさ）
"""
import cv2
import numpy as np
from PIL import Image, ImageDraw

SRC = "_images/inkart_1109_clearcard_banner_sky_generated.webp"
OUT_FULL = "_images/inkart_1109_clearcard_banner_sky_v2_1254.png"
OUT = "_images/inkart_1109_clearcard_banner_sky_v2_1040.jpg"

im = Image.open(SRC).convert("RGB")
W, H = im.size

# 元の位置（白い箱の実測）
CHIP_L = (60, 878, 583, 954)      # 名刺・ショップカードに
CHIP_R = (628, 878, 1219, 954)    # しおり・ポストカード・チケットにも
PRICE = (45, 978, 607, 1064)      # 100枚 9,780円〜（税別）
POINTER = (90, 1060, 160, 1086)   # 値段の下の三角（今回は消す）
R = 18                            # 角丸の半径


def rounded_mask(size, r):
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius=r, fill=255)
    return m


def sprite(box):
    s = im.crop(box)
    return s, rounded_mask(s.size, R)


chip_l, m_l = sprite(CHIP_L)
chip_r, m_r = sprite(CHIP_R)
price, m_p = sprite(PRICE)

# 新しい位置
pw, ph = price.size
NEW_PRICE = ((W - pw) // 2, 876)                      # 上の段・中央
cw = CHIP_L[2] - CHIP_L[0]
ch = CHIP_L[3] - CHIP_L[1]
NEW_CHIP_L = (CHIP_L[0], 978)                         # 下の段・左右は元のまま
NEW_CHIP_R = (CHIP_R[0], 978)

# 1. 消す所（元の札3つ＋三角）を inpaint で埋める
mask = np.zeros((H, W), np.uint8)
for (x1, y1, x2, y2) in (CHIP_L, CHIP_R, POINTER):
    mask[y1 - 2:y2 + 2, x1 - 2:x2 + 2] = 255
# 値段の箱はオレンジの縁と影が外側に付いているので広めに消す（ボタンの上端 1077 には掛けない）
mask[966:1076, 34:618] = 255
bgr = cv2.cvtColor(np.array(im), cv2.COLOR_RGB2BGR)
filled = cv2.inpaint(bgr, mask, 7, cv2.INPAINT_TELEA)
base = Image.fromarray(cv2.cvtColor(filled, cv2.COLOR_BGR2RGB))

# 2. 置き直す
base.paste(price, NEW_PRICE, m_p)
#   値段の箱に元と同じオレンジの縁を付け直す
ImageDraw.Draw(base).rounded_rectangle(
    (NEW_PRICE[0] - 3, NEW_PRICE[1] - 3, NEW_PRICE[0] + pw + 2, NEW_PRICE[1] + ph + 2),
    radius=R + 3, outline=(233, 121, 11), width=6)
base.paste(chip_l, NEW_CHIP_L, m_l)
base.paste(chip_r, NEW_CHIP_R, m_r)

base.save(OUT_FULL)
base.resize((1040, 1040), Image.LANCZOS).save(OUT, quality=95)
print(OUT, "price at", NEW_PRICE, "chips y", NEW_CHIP_L[1])
