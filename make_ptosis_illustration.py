# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 320, 400
SKIN_COLOR = (230, 200, 190)
SHADOW_COLOR = (200, 150, 130)
EYE_WHITE = (255, 255, 255)
IRIS_COLOR = (100, 70, 50)
LASH_COLOR = (50, 30, 20)

def draw_eye_before(d, cx, cy):
    """眼瞼下垂（症状）- まぶたが下がった目"""
    # 顔のアウトライン（簡潔）
    d.ellipse((cx-35, cy-40, cx+35, cy+50), fill=SKIN_COLOR, outline=SHADOW_COLOR, width=2)

    # 眉
    d.arc((cx-30, cy-45, cx+30, cy-30), start=0, end=180, fill=LASH_COLOR, width=3)

    # 目の周りの影（眼窩）
    d.ellipse((cx-25, cy-15, cx+25, cy+20), outline=SHADOW_COLOR, width=2)

    # 下垂したまぶた（厚く、下がった状態）
    d.polygon([(cx-22, cy-8), (cx+22, cy-8), (cx+20, cy+2), (cx-20, cy+2)], fill=SHADOW_COLOR)

    # 白目
    d.ellipse((cx-15, cy-5, cx+15, cy+12), fill=EYE_WHITE, outline=LASH_COLOR, width=1)

    # 虹彩
    d.ellipse((cx-8, cy-2, cx+8, cy+8), fill=IRIS_COLOR)

    # 瞳
    d.ellipse((cx-4, cy+1, cx+4, cy+5), fill=(30, 20, 10))

    # まつげ（下垂で隠れ気味）
    for i in range(-3, 4):
        d.line((cx+i*4, cy+12, cx+i*4-1, cy+15), fill=LASH_COLOR, width=2)

    # 下まぶた
    d.arc((cx-15, cy+5, cx+15, cy+25), start=0, end=180, fill=SHADOW_COLOR, width=2)

    # 目頭・目尻の皴
    d.line((cx-25, cy-5, cx-30, cy+5), fill=SHADOW_COLOR, width=1)
    d.line((cx+25, cy-5, cx+30, cy+5), fill=SHADOW_COLOR, width=1)

def draw_eye_after(d, cx, cy):
    """治療後 - まぶたが持ち上がった目"""
    # 顔のアウトライン
    d.ellipse((cx-35, cy-40, cx+35, cy+50), fill=SKIN_COLOR, outline=SHADOW_COLOR, width=2)

    # 眉
    d.arc((cx-30, cy-45, cx+30, cy-30), start=0, end=180, fill=LASH_COLOR, width=3)

    # 目の周りの影
    d.ellipse((cx-25, cy-20, cx+25, cy+15), outline=SHADOW_COLOR, width=2)

    # 正常なまぶた（薄く、持ち上がった状態）
    d.polygon([(cx-22, cy-18), (cx+22, cy-18), (cx+20, cy-12), (cx-20, cy-12)], fill=SHADOW_COLOR)

    # 白目（より大きく見える）
    d.ellipse((cx-15, cy-10, cx+15, cy+10), fill=EYE_WHITE, outline=LASH_COLOR, width=1)

    # 虹彩
    d.ellipse((cx-8, cy-5, cx+8, cy+8), fill=IRIS_COLOR)

    # 瞳
    d.ellipse((cx-4, cy-2, cx+4, cy+4), fill=(30, 20, 10))

    # まつげ（より見える）
    for i in range(-3, 4):
        d.line((cx+i*4, cy-12, cx+i*4-1, cy-16), fill=LASH_COLOR, width=2)

    # 下まぶた
    d.arc((cx-15, cy+5, cx+15, cy+20), start=0, end=180, fill=SHADOW_COLOR, width=2)

    # 目頭・目尻の皴（軽減）
    d.line((cx-25, cy-8, cx-28, cy+3), fill=SHADOW_COLOR, width=1)
    d.line((cx+25, cy-8, cx+30, cy+3), fill=SHADOW_COLOR, width=1)

img = Image.new("RGB", (W, H), (220, 180, 200))
d = ImageDraw.Draw(img)

# タイトル背景
d.rectangle((0, 0, W, 50), fill=(200, 150, 180))

# Before / After ラベル背景
before_box = (10, 5, W//2-5, 45)
after_box = (W//2+5, 5, W-10, 45)
d.rounded_rectangle(before_box, radius=8, fill=(100, 120, 160))
d.rounded_rectangle(after_box, radius=8, fill=(180, 80, 80))

# ラベルテキスト
d.text((20, 14), "Before", font=font(16, True), fill="#FFFFFF")
d.text((W//2+15, 14), "After", font=font(16, True), fill="#FFFFFF")

# 目のイラスト
draw_eye_before(d, W//4, 150)
draw_eye_after(d, 3*W//4, 150)

# 説明テキスト
d.text((10, H-40), "眼瞼下垂の改善例（イラスト）", font=font(10), fill=(100, 50, 80))
d.text((10, H-25), "▸ 左：症状（まぶたが下がった状態）", font=font(8), fill=(80, 40, 60))
d.text((10, H-12), "▸ 右：治療後（まぶたが持ち上がった状態）", font=font(8), fill=(80, 40, 60))

img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\ptosis_illustration.png")
print("Created: ptosis_illustration.png")
