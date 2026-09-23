# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 360, 620
SENT_GREEN = "#93E15D"

def icon_truck(d, cx, cy, s, color, width=4):
    d.rectangle((cx-s, cy-s*0.4, cx+s*0.3, cy+s*0.5), outline=color, width=width)
    d.polygon([(cx+s*0.3, cy-s*0.1), (cx+s*0.9, cy-s*0.1), (cx+s*1.1, cy+s*0.2), (cx+s*1.1, cy+s*0.5), (cx+s*0.3, cy+s*0.5)], outline=color, width=width)
    d.ellipse((cx-s*0.7, cy+s*0.35, cx-s*0.3, cy+s*0.75), outline=color, width=width)
    d.ellipse((cx+s*0.5, cy+s*0.35, cx+s*0.9, cy+s*0.75), outline=color, width=width)

def icon_bolt(d, cx, cy, s, color, width=4):
    pts = [(cx+s*0.15, cy-s), (cx-s*0.6, cy+s*0.15), (cx-s*0.05, cy+s*0.15), (cx-s*0.2, cy+s), (cx+s*0.6, cy-s*0.2), (cx+s*0.05, cy-s*0.2)]
    d.polygon(pts, fill=color)

def icon_house(d, cx, cy, s, color, width=4):
    d.polygon([(cx-s, cy+s*0.1), (cx, cy-s*0.8), (cx+s, cy+s*0.1)], outline=color, width=width)
    d.rectangle((cx-s*0.65, cy+s*0.1, cx+s*0.65, cy+s), outline=color, width=width)
    d.rectangle((cx-s*0.18, cy+s*0.45, cx+s*0.18, cy+s), outline=color, width=width)

def icon_glasses(d, cx, cy, s, color, width=4):
    d.ellipse((cx-s*1.1, cy-s*0.4, cx-s*0.2, cy+s*0.5), outline=color, width=width)
    d.ellipse((cx+s*0.2, cy-s*0.4, cx+s*1.1, cy+s*0.5), outline=color, width=width)
    d.line((cx-s*0.2, cy, cx+s*0.2, cy), fill=color, width=width)
    d.line((cx-s*1.1, cy-s*0.1, cx-s*1.35, cy-s*0.3), fill=color, width=width)
    d.line((cx+s*1.1, cy-s*0.1, cx+s*1.35, cy-s*0.3), fill=color, width=width)

ICONS = {"delivery": icon_truck, "utility": icon_bolt, "realestate": icon_house, "retail": icon_glasses}

def bg_gradient(d):
    top = (222, 234, 246)
    bot = (204, 221, 240)
    for i in range(H - 12):
        t = i / (H - 12)
        col = tuple(int(top[j] + (bot[j] - top[j]) * t) for j in range(3))
        d.line((6, 6 + i, W - 6, 6 + i), fill=col)

def header_icons(d, y0, y1):
    # search
    sx, sy = W-146, (y0+y1)/2-8
    d.ellipse((sx, sy, sx+16, sy+16), outline="#555555", width=2)
    d.line((sx+14, sy+14, sx+20, sy+20), fill="#555555", width=2)
    # phone
    px, py = W-106, (y0+y1)/2-9
    d.rounded_rectangle((px, py, px+18, py+18), radius=5, outline="#555555", width=2)
    # hamburger
    hx, hy = W-66, (y0+y1)/2-7
    for i in range(3):
        d.line((hx, hy+i*6, hx+18, hy+i*6), fill="#555555", width=2)

def make_frame3(path, key, accent, brand, thanks_title, thanks_body, reply_text):
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)

    # status bar (carrier + time + signal/wifi/battery)
    d.rectangle((8, 8, W-8, 48), fill="#FFFFFF")
    d.text((24, 20), "SoftBank", font=font(12, True), fill="#111111")
    d.text((W/2-24, 18), "9:41", font=font(15, True), fill="#111111")
    # signal bars
    bx = W-92
    for i, hgt in enumerate([6, 9, 12, 15]):
        d.rectangle((bx+i*5, 34-hgt, bx+i*5+3, 34), fill="#111111")
    # wifi
    wx = W-68
    d.arc((wx-8, 18, wx+8, 34), start=225, end=315, fill="#111111", width=2)
    d.arc((wx-5, 22, wx+5, 32), start=225, end=315, fill="#111111", width=2)
    d.ellipse((wx-1, 28, wx+2, 31), fill="#111111")
    # battery
    batx = W-48
    d.rounded_rectangle((batx, 20, batx+26, 32), radius=3, outline="#111111", width=2)
    d.rectangle((batx+27, 24, batx+29, 28), fill="#111111")
    d.rectangle((batx+2, 22, batx+22, 30), fill="#111111")

    # header
    header_top, header_h = 48, 56
    d.rectangle((8, header_top, W-8, header_top+header_h), fill="#FFFFFF")
    d.line((8, header_top+header_h, W-8, header_top+header_h), fill="#E5E5E5", width=1)
    d.text((22, header_top+header_h/2-13), "<", font=font(24, True), fill="#333333")
    d.ellipse((52, header_top+9, 52+38, header_top+47), fill=accent)
    ICONS[key](d, 52+19, header_top+28, 9, "#FFFFFF", width=3)
    d.text((100, header_top+header_h/2-9), brand, font=font(15, True), fill="#222222")
    header_icons(d, header_top, header_top+header_h)

    # chat background (light blue gradient)
    chat_top = header_top+header_h
    inputbar_h = 56
    chat_bottom = H-8-inputbar_h
    for i in range(int(chat_bottom-chat_top)):
        t = i/(chat_bottom-chat_top)
        col = tuple(int((222,234,246)[j] + ((202,220,242)[j]-(222,234,246)[j])*t) for j in range(3))
        d.line((8, chat_top+i, W-8, chat_top+i), fill=col)
    # subtle cloud-like texture
    import random as _r
    rnd = _r.Random(hash(key) % 1000)
    for _ in range(6):
        bx = rnd.uniform(20, W-20)
        by = rnd.uniform(chat_top+60, chat_bottom-60)
        br = rnd.uniform(25, 55)
        d.ellipse((bx-br, by-br*0.5, bx+br, by+br*0.5), fill=(214, 227, 242))

    # date chip
    chip_w = 70
    d.rounded_rectangle((W/2-chip_w/2, chat_top+14, W/2+chip_w/2, chat_top+36), radius=10, fill="#C9D6E8")
    d.text((W/2-24, chat_top+19), "7月2日", font=font(11), fill="#5B6B85")

    # OA message bubble (white, left)
    y = chat_top+52
    bbox_lines = [thanks_title, thanks_body]
    bh = 30 + 22*len(bbox_lines)
    bubble = (24, y, W-70, y+bh)
    d.rounded_rectangle(bubble, radius=4, fill="#FFFFFF")
    d.polygon([(bubble[0], y+18), (bubble[0]-8, y+22), (bubble[0], y+26)], fill="#FFFFFF")
    d.text((40, y+12), thanks_title, font=font(13, True), fill="#1a1a1a")
    d.text((40, y+36), thanks_body, font=font(11), fill="#333333")
    d.text((bubble[0], bubble[3]+4), "18:02", font=font(9), fill="#8A93A8")
    y = bubble[3] + 30

    # user reply bubble (green, right)
    bbox = d.textbbox((0,0), reply_text, font=font(12))
    tw = bbox[2]-bbox[0]
    bubble2 = (W-24-tw-30, y, W-24, y+42)
    d.rounded_rectangle(bubble2, radius=4, fill=SENT_GREEN)
    d.text((bubble2[0]+15, y+13), reply_text, font=font(12), fill="#1a1a1a")
    d.text((bubble2[0]-40, bubble2[3]-14), "既読\n18:03", font=font(8), fill="#8A93A8")
    y = bubble2[3] + 26

    # friend-added system chip
    fb_w = 200
    d.rounded_rectangle((W/2-fb_w/2, y, W/2+fb_w/2, y+30), radius=15, fill="#5B6B85")
    d.text((W/2-fb_w/2+16, y+8), "「友だち」に自動追加されました", font=font(10, True), fill="#FFFFFF")

    # input bar
    ib_top = chat_bottom
    d.rectangle((8, ib_top, W-8, H-8), fill="#FFFFFF")
    d.line((8, ib_top, W-8, ib_top), fill="#E5E5E5", width=1)
    d.ellipse((24, ib_top+13, 24+28, ib_top+41), outline="#999999", width=2)
    d.line((24+14, ib_top+19, 24+14, ib_top+35), fill="#999999", width=2)
    d.line((24+7, ib_top+27, 24+21, ib_top+27), fill="#999999", width=2)
    d.rounded_rectangle((64, ib_top+11, W-96, ib_top+43), radius=16, outline="#DADDE1", width=2)
    d.text((78, ib_top+19), "メッセージを入力", font=font(12), fill="#BBBBBB")
    # camera icon
    cx1 = W-80
    cy1 = ib_top+27
    d.rounded_rectangle((cx1-11, cy1-8, cx1+11, cy1+9), radius=4, outline="#999999", width=2)
    d.rectangle((cx1-5, cy1-12, cx1+5, cy1-8), fill="#999999")
    d.ellipse((cx1-6, cy1-3, cx1+6, cy1+9), outline="#999999", width=2)
    # mic icon
    cx2 = W-42
    cy2 = ib_top+27
    d.rounded_rectangle((cx2-6, cy2-12, cx2+6, cy2+4), radius=6, outline="#999999", width=2)
    d.arc((cx2-11, cy2-6, cx2+11, cy2+12), start=0, end=180, fill="#999999", width=2)
    d.line((cx2, cy2+12, cx2, cy2+17), fill="#999999", width=2)

    # home indicator
    d.rounded_rectangle((W/2-46, H-16, W/2+46, H-11), radius=3, fill="#1a1a1a")

    img.save(path)

brands = [
    dict(key="delivery", accent="#FF7A33", brand="ハコブーン", thanks_title="友だち追加ありがとうございます！",
         thanks_body="お届け状況はこちらから確認できます。", reply="ありがとうございます！"),
    dict(key="utility", accent="#2F80ED", brand="でんきねっと", thanks_title="友だち追加ありがとうございます！",
         thanks_body="今後の料金確定はこちらでお知らせします。", reply="ありがとうございます！"),
    dict(key="realestate", accent="#2FAE66", brand="いえらいふ", thanks_title="友だち追加ありがとうございます！",
         thanks_body="担当者よりこちらでご連絡します。", reply="よろしくお願いします！"),
    dict(key="retail", accent="#8A3FE4", brand="メガネのミライ", thanks_title="友だち追加ありがとうございます！",
         thanks_body="発送状況はこちらから確認できます。", reply="ありがとうございます！"),
]

import os
os.chdir(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_flow_example")
for b in brands:
    make_frame3(f"flow3_{b['key']}.png", b["key"], b["accent"], b["brand"], b["thanks_title"], b["thanks_body"], b["reply"])

print("done")
