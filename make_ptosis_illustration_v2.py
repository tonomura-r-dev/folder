# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import math

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 320, 420
SKIN_BASE = (235, 205, 190)
SKIN_SHADOW = (200, 160, 140)
SKIN_DARK = (180, 140, 120)
EYE_WHITE = (245, 245, 250)
IRIS_COLOR = (80, 60, 40)
PUPIL_COLOR = (20, 15, 10)
LASH_COLOR = (40, 25, 15)

def draw_face_base(d, cx, cy):
    """顔の基本形（共通）"""
    # 顔のメイン（楕円）
    d.ellipse((cx-50, cy-65, cx+50, cy+75), fill=SKIN_BASE, outline=SKIN_SHADOW, width=2)

    # 顔の陰影（側面）
    for i in range(20):
        t = i / 20
        col = tuple(int(SKIN_BASE[j] + (SKIN_SHADOW[j] - SKIN_BASE[j]) * t * 0.3) for j in range(3))
        d.line((cx-50, cy-65+i*7, cx-48, cy-65+i*7), fill=col, width=2)
        d.line((cx+50, cy-65+i*7, cx+48, cy-65+i*7), fill=col, width=2)

    # 鼻
    d.polygon([(cx, cy-20), (cx-6, cy+10), (cx+6, cy+10)], fill=SKIN_SHADOW)
    d.line((cx-6, cy+10, cx+6, cy+10), fill=SKIN_DARK, width=1)

    # 口（軽く）
    d.arc((cx-15, cy+35, cx+15, cy+50), start=0, end=180, fill=SKIN_DARK, width=2)

    return cx, cy

def draw_eyebrow(d, cx, cy, is_before=True):
    """眉毛"""
    # 眉の位置と形
    y_offset = -5 if is_before else -8
    d.arc((cx-20, cy-50+y_offset, cx+20, cy-40+y_offset), start=0, end=180, fill=LASH_COLOR, width=3)

def draw_eye_before(d, cx, cy):
    """眼瞼下垂（症状） - まぶたが下がった目"""
    draw_eyebrow(d, cx, cy, is_before=True)

    # 下垂したまぶた（上）
    ptosis_height = 15  # まぶたが下がっている量
    d.polygon([
        (cx-22, cy-8),
        (cx+22, cy-8),
        (cx+22, cy-8+ptosis_height),
        (cx-22, cy-8+ptosis_height)
    ], fill=SKIN_SHADOW)

    # まぶたの線
    d.line((cx-22, cy-8+ptosis_height, cx+22, cy-8+ptosis_height), fill=LASH_COLOR, width=2)

    # 白目の領域（下垂で隠れている）
    d.ellipse((cx-18, cy+2, cx+18, cy+18), fill=EYE_WHITE, outline=LASH_COLOR, width=1)

    # 虹彩（下寄り）
    d.ellipse((cx-10, cy+5, cx+10, cy+15), fill=IRIS_COLOR)

    # 瞳
    d.ellipse((cx-5, cy+7, cx+5, cy+12), fill=PUPIL_COLOR)

    # 瞳のハイライト
    d.ellipse((cx-2, cy+8, cx+1, cy+10), fill=(100, 100, 120))

    # 下まぶた
    d.arc((cx-18, cy+15, cx+18, cy+32), start=0, end=180, fill=SKIN_SHADOW, width=2)

    # まつ毛（上）- 少なく見える
    for i in range(-4, 5):
        start_x = cx + i*3.5
        d.line((start_x, cy-8+ptosis_height, start_x-0.5, cy-12+ptosis_height), fill=LASH_COLOR, width=2)

    # まつ毛（下）
    for i in range(-4, 5):
        start_x = cx + i*3.5
        d.line((start_x, cy+18, start_x-0.5, cy+22), fill=LASH_COLOR, width=1)

    # 目頭・目尻のしわ
    d.line((cx-22, cy-2, cx-28, cy+8), fill=SKIN_DARK, width=1)
    d.line((cx+22, cy-2, cx+28, cy+8), fill=SKIN_DARK, width=1)

def draw_eye_after(d, cx, cy):
    """治療後 - まぶたが持ち上がった目"""
    draw_eyebrow(d, cx, cy, is_before=False)

    # 正常なまぶた（上）- 薄く持ち上がっている
    normal_height = 6
    d.polygon([
        (cx-22, cy-18),
        (cx+22, cy-18),
        (cx+22, cy-18+normal_height),
        (cx-22, cy-18+normal_height)
    ], fill=SKIN_SHADOW)

    # まぶたの線
    d.line((cx-22, cy-18+normal_height, cx+22, cy-18+normal_height), fill=LASH_COLOR, width=2)

    # 白目の領域（より大きく見える）
    d.ellipse((cx-18, cy-12, cx+18, cy+10), fill=EYE_WHITE, outline=LASH_COLOR, width=1)

    # 虹彩（中央）
    d.ellipse((cx-10, cy-6, cx+10, cy+4), fill=IRIS_COLOR)

    # 瞳
    d.ellipse((cx-5, cy-4, cx+5, cy+1), fill=PUPIL_COLOR)

    # 瞳のハイライト
    d.ellipse((cx-2, cy-3, cx+1, cy-1), fill=(100, 100, 120))

    # 下まぶた
    d.arc((cx-18, cy+8, cx+18, cy+25), start=0, end=180, fill=SKIN_SHADOW, width=2)

    # まつ毛（上） - たくさん見える
    for i in range(-4, 5):
        start_x = cx + i*3.5
        d.line((start_x, cy-18+normal_height, start_x-0.5, cy-23+normal_height), fill=LASH_COLOR, width=2)

    # まつ毛（下）
    for i in range(-4, 5):
        start_x = cx + i*3.5
        d.line((start_x, cy+10, start_x-0.5, cy+14), fill=LASH_COLOR, width=1)

    # 目頭・目尻のしわ（軽減）
    d.line((cx-22, cy-12, cx-26, cy+2), fill=SKIN_DARK, width=1)
    d.line((cx+22, cy-12, cx+26, cy+2), fill=SKIN_DARK, width=1)

def create_illustration():
    img = Image.new("RGB", (W, H), (220, 180, 200))
    d = ImageDraw.Draw(img)

    # タイトル背景
    d.rectangle((0, 0, W, 50), fill=(200, 150, 180))

    # Before / After ラベル背景
    before_box = (10, 8, W//2-5, 42)
    after_box = (W//2+5, 8, W-10, 42)
    d.rounded_rectangle(before_box, radius=6, fill=(100, 120, 160))
    d.rounded_rectangle(after_box, radius=6, fill=(180, 80, 80))

    # ラベルテキスト
    d.text((18, 14), "Before", font=font(14, True), fill="#FFFFFF")
    d.text((W//2+13, 14), "After", font=font(14, True), fill="#FFFFFF")

    # Before 顔
    draw_face_base(d, W//4, 160)
    draw_eye_before(d, W//4, 160)

    # After 顔
    draw_face_base(d, 3*W//4, 160)
    draw_eye_after(d, 3*W//4, 160)

    # 説明テキスト
    d.text((10, H-50), "眼瞼下垂の改善（医学図解）", font=font(10, True), fill=(100, 50, 80))
    d.text((10, H-35), "左図：まぶたが下がった状態", font=font(8), fill=(80, 40, 60))
    d.text((10, H-22), "右図：治療後、まぶたが持ち上がった状態", font=font(8), fill=(80, 40, 60))

    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\ptosis_illustration.png")
    print("Created: ptosis_illustration.png (improved)")

create_illustration()
