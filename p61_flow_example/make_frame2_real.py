# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 360, 620

def icon_truck(d, cx, cy, s, color, width=3):
    d.rectangle((cx-s, cy-s*0.4, cx+s*0.3, cy+s*0.5), outline=color, width=width)
    d.polygon([(cx+s*0.3, cy-s*0.1), (cx+s*0.9, cy-s*0.1), (cx+s*1.1, cy+s*0.2), (cx+s*1.1, cy+s*0.5), (cx+s*0.3, cy+s*0.5)], outline=color, width=width)
    d.ellipse((cx-s*0.7, cy+s*0.35, cx-s*0.3, cy+s*0.75), outline=color, width=width)
    d.ellipse((cx+s*0.5, cy+s*0.35, cx+s*0.9, cy+s*0.75), outline=color, width=width)

def icon_bolt(d, cx, cy, s, color, width=3):
    pts = [(cx+s*0.15, cy-s), (cx-s*0.6, cy+s*0.15), (cx-s*0.05, cy+s*0.15), (cx-s*0.2, cy+s), (cx+s*0.6, cy-s*0.2), (cx+s*0.05, cy-s*0.2)]
    d.polygon(pts, fill=color)

def icon_house(d, cx, cy, s, color, width=3):
    d.polygon([(cx-s, cy+s*0.1), (cx, cy-s*0.8), (cx+s, cy+s*0.1)], outline=color, width=width)
    d.rectangle((cx-s*0.65, cy+s*0.1, cx+s*0.65, cy+s), outline=color, width=width)
    d.rectangle((cx-s*0.18, cy+s*0.45, cx+s*0.18, cy+s), outline=color, width=width)

def icon_glasses(d, cx, cy, s, color, width=3):
    d.ellipse((cx-s*1.1, cy-s*0.4, cx-s*0.2, cy+s*0.5), outline=color, width=width)
    d.ellipse((cx+s*0.2, cy-s*0.4, cx+s*1.1, cy+s*0.5), outline=color, width=width)
    d.line((cx-s*0.2, cy, cx+s*0.2, cy), fill=color, width=width)
    d.line((cx-s*1.1, cy-s*0.1, cx-s*1.35, cy-s*0.3), fill=color, width=width)
    d.line((cx+s*1.1, cy-s*0.1, cx+s*1.35, cy-s*0.3), fill=color, width=width)

ICONS = {"delivery": icon_truck, "utility": icon_bolt, "realestate": icon_house, "retail": icon_glasses}
LINE_GREEN = "#06C755"

def icon_flashlight(d, cx, cy, s):
    d.ellipse((cx-s, cy-s, cx+s, cy+s), fill=(70, 70, 78))
    d.line((cx-6, cy-4, cx+2, cy-4), fill="#FFFFFF", width=2)
    d.line((cx-2, cy, cx+6, cy), fill="#FFFFFF", width=2)
    d.line((cx-6, cy+4, cx+2, cy+4), fill="#FFFFFF", width=2)

def icon_camera(d, cx, cy, s):
    d.ellipse((cx-s, cy-s, cx+s, cy+s), fill=(70, 70, 78))
    d.rounded_rectangle((cx-8, cy-5, cx+8, cy+5), radius=2, outline="#FFFFFF", width=2)
    d.ellipse((cx-3, cy-3, cx+3, cy+3), outline="#FFFFFF", width=2)

def make_frame2(path, key, brand, notif_title, body_lines):
    img = Image.new("RGB", (W, H), "#1c1e26")
    d = ImageDraw.Draw(img)
    # lock screen background gradient (dark blue-purple night sky feel)
    top = (28, 30, 46); bot = (46, 40, 66)
    for i in range(H-12):
        t = i/(H-12)
        col = tuple(int(top[j] + (bot[j]-top[j])*t) for j in range(3))
        d.line((6, 6+i, W-6, 6+i), fill=col)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#3A3E47", width=2)

    # status bar
    d.text((26, 20), "SoftBank", font=font(11, True), fill="#FFFFFF")
    bx = W-92
    for i, hgt in enumerate([5, 8, 11, 14]):
        d.rectangle((bx+i*5, 30-hgt, bx+i*5+3, 30), fill="#FFFFFF")
    batx = W-48
    d.rounded_rectangle((batx, 18, batx+26, 30), radius=3, outline="#FFFFFF", width=2)
    d.rectangle((batx+27, 22, batx+29, 26), fill="#FFFFFF")
    d.rectangle((batx+2, 20, batx+22, 28), fill="#FFFFFF")

    # lock icon
    lx, ly = W/2, 62
    d.rounded_rectangle((lx-7, ly, lx+7, ly+10), radius=2, outline="#FFFFFF", width=2)
    d.arc((lx-5, ly-8, lx+5, ly+2), start=180, end=360, fill="#FFFFFF", width=2)

    d.text((W/2-58, 88), "18:02", font=font(50, True), fill="#FFFFFF")
    d.text((W/2-56, 142), "2026年7月2日（木）", font=font(13), fill="#D8D8DE")

    # push notification banner
    nb = (24, 190, W-24, 190+100)
    d.rounded_rectangle(nb, radius=20, fill=(58, 58, 68))
    icon_bg = (nb[0]+16, nb[1]+16, nb[0]+16+34, nb[1]+16+34)
    d.rounded_rectangle(icon_bg, radius=9, fill=LINE_GREEN)
    icx, icy = (icon_bg[0]+icon_bg[2])/2, (icon_bg[1]+icon_bg[3])/2
    ICONS[key](d, icx, icy, 8, "#FFFFFF", width=3)
    d.text((nb[0]+62, nb[1]+14), brand, font=font(13, True), fill="#FFFFFF")
    d.text((nb[2]-54, nb[1]+14), "いま", font=font(10), fill="#AAAAAA")
    d.text((nb[0]+62, nb[1]+36), notif_title, font=font(13, True), fill="#FFFFFF")
    yy = nb[1]+58
    for line in body_lines:
        d.text((nb[0]+62, yy), line, font=font(11), fill="#D8D8D8")
        yy += 18

    # second (older) notification, collapsed style for realism
    nb2 = (24, 300, W-24, 300+40)
    d.rounded_rectangle(nb2, radius=18, fill=(48, 48, 58))
    d.text((nb2[0]+20, nb2[1]+11), "その他の通知 2件", font=font(11), fill="#B8B8C0")

    # bottom shortcuts: flashlight / camera
    icon_flashlight(d, 56, H-72, 22)
    icon_camera(d, W-56, H-72, 22)

    # home indicator
    d.rounded_rectangle((W/2-46, H-16, W/2+46, H-11), radius=3, fill="#FFFFFF")

    d.text((10, H-30), "(2) ロック画面にプッシュ通知が届く（未タップ）", font=font(10, True), fill="#8FD8B0")
    img.save(path)

brands = [
    dict(key="delivery", brand="ハコブーン", notif_title="お荷物お届けのお知らせ",
         body=["ご注文の商品を本日18時〜20時に", "お届け予定です。"]),
    dict(key="utility", brand="でんきねっと", notif_title="今月の料金確定のお知らせ",
         body=["今月のご利用料金が確定しました。", "マイページよりご確認ください。"]),
    dict(key="realestate", brand="いえらいふ", notif_title="資料請求完了のお知らせ",
         body=["お問い合わせいただいた物件の", "資料をご用意しました。"]),
    dict(key="retail", brand="メガネのミライ", notif_title="ご購入完了のお知らせ",
         body=["ご購入ありがとうございます。", "発送準備が整い次第お知らせします。"]),
]

import os
os.chdir(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_flow_example")
for b in brands:
    make_frame2(f"flow2_{b['key']}.png", b["key"], b["brand"], b["notif_title"], b["body"])

print("done")
