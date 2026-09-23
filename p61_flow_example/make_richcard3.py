# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import random

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 360, 680
ACCENT = "#FF7A33"
ACCENT_DARK = "#C85A1E"
LINE_GREEN = "#06C755"

def icon_home(d, cx, cy, s, color):
    d.polygon([(cx-s, cy), (cx, cy-s*0.9), (cx+s, cy)], outline=color, width=3)
    d.rectangle((cx-s*0.65, cy, cx+s*0.65, cy+s*0.85), outline=color, width=3)

def phone_shell(brand="ハコブーン"):
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)

    # status bar
    d.rectangle((8, 8, W-8, 48), fill="#FFFFFF")
    d.text((24, 20), "SoftBank", font=font(12, True), fill="#111111")
    d.text((W/2-24, 18), "9:41", font=font(15, True), fill="#111111")
    bx = W-92
    for i, hgt in enumerate([6, 9, 12, 15]):
        d.rectangle((bx+i*5, 34-hgt, bx+i*5+3, 34), fill="#111111")
    wx = W-68
    d.arc((wx-8, 18, wx+8, 34), start=225, end=315, fill="#111111", width=2)
    d.arc((wx-5, 22, wx+5, 32), start=225, end=315, fill="#111111", width=2)
    d.ellipse((wx-1, 28, wx+2, 31), fill="#111111")
    batx = W-48
    d.rounded_rectangle((batx, 20, batx+26, 32), radius=3, outline="#111111", width=2)
    d.rectangle((batx+27, 24, batx+29, 28), fill="#111111")
    d.rectangle((batx+2, 22, batx+22, 30), fill="#111111")

    # LINE header
    header_top, header_h = 48, 56
    d.rectangle((8, header_top, W-8, header_top+header_h), fill="#FFFFFF")
    d.line((8, header_top+header_h, W-8, header_top+header_h), fill="#E5E5E5", width=1)
    d.text((22, header_top+header_h/2-13), "<", font=font(24, True), fill="#333333")
    d.ellipse((52, header_top+9, 52+38, header_top+47), fill=ACCENT)
    icon_home(d, 52+19, header_top+30, 9, "#FFFFFF")
    d.text((100, header_top+header_h/2-9), brand, font=font(15, True), fill="#222222")
    sx, sy = W-146, header_top+header_h/2-8
    d.ellipse((sx, sy, sx+16, sy+16), outline="#555555", width=2)
    d.line((sx+14, sy+14, sx+20, sy+20), fill="#555555", width=2)
    px, py = W-106, header_top+header_h/2-9
    d.rounded_rectangle((px, py, px+18, py+18), radius=5, outline="#555555", width=2)
    hx, hy = W-66, header_top+header_h/2-7
    for i in range(3):
        d.line((hx, hy+i*6, hx+18, hy+i*6), fill="#555555", width=2)

    # chat background (light blue)
    chat_top = header_top + header_h
    inputbar_h = 56
    chat_bottom = H - 8 - inputbar_h
    top_c = (222, 234, 246); bot_c = (202, 220, 242)
    for i in range(int(chat_bottom - chat_top)):
        t = i / (chat_bottom - chat_top)
        col = tuple(int(top_c[j] + (bot_c[j]-top_c[j])*t) for j in range(3))
        d.line((8, chat_top+i, W-8, chat_top+i), fill=col)
    rnd = random.Random(7)
    for _ in range(5):
        bx2 = rnd.uniform(20, W-20)
        by2 = rnd.uniform(chat_top+40, chat_bottom-40)
        br2 = rnd.uniform(25, 55)
        d.ellipse((bx2-br2, by2-br2*0.5, bx2+br2, by2+br2*0.5), fill=(214, 227, 242))

    # date chip
    chip_w = 70
    d.rounded_rectangle((W/2-chip_w/2, chat_top+12, W/2+chip_w/2, chat_top+34), radius=10, fill="#C9D6E8")
    d.text((W/2-24, chat_top+17), "7月2日", font=font(11), fill="#5B6B85")

    # input bar
    ib_top = chat_bottom
    d.rectangle((8, ib_top, W-8, H-8), fill="#FFFFFF")
    d.line((8, ib_top, W-8, ib_top), fill="#E5E5E5", width=1)
    d.ellipse((24, ib_top+13, 24+28, ib_top+41), outline="#999999", width=2)
    d.line((24+14, ib_top+19, 24+14, ib_top+35), fill="#999999", width=2)
    d.line((24+7, ib_top+27, 24+21, ib_top+27), fill="#999999", width=2)
    d.rounded_rectangle((64, ib_top+11, W-96, ib_top+43), radius=16, outline="#DADDE1", width=2)
    d.text((78, ib_top+19), "メッセージを入力", font=font(12), fill="#BBBBBB")
    cx1, cy1 = W-80, ib_top+27
    d.rounded_rectangle((cx1-11, cy1-8, cx1+11, cy1+9), radius=4, outline="#999999", width=2)
    d.rectangle((cx1-5, cy1-12, cx1+5, cy1-8), fill="#999999")
    d.ellipse((cx1-6, cy1-3, cx1+6, cy1+9), outline="#999999", width=2)
    cx2, cy2 = W-42, ib_top+27
    d.rounded_rectangle((cx2-6, cy2-12, cx2+6, cy2+4), radius=6, outline="#999999", width=2)
    d.arc((cx2-11, cy2-6, cx2+11, cy2+12), start=0, end=180, fill="#999999", width=2)
    d.line((cx2, cy2+12, cx2, cy2+17), fill="#999999", width=2)

    # home indicator
    d.rounded_rectangle((W/2-46, H-16, W/2+46, H-11), radius=3, fill="#1a1a1a")

    return img, d, chat_top

def icon_person_box(d, cx, cy):
    d.ellipse((cx-24, cy-46, cx-10, cy-32), fill="#3a3a3a")
    d.polygon([(cx-28, cy-30), (cx-8, cy-30), (cx-4, cy+6), (cx-32, cy+6)], fill="#3a3a3a")
    d.rectangle((cx-2, cy-22, cx+32, cy+8), fill=ACCENT)
    d.line((cx-2, cy-22, cx+32, cy+8), fill=ACCENT_DARK, width=2)
    d.line((cx+15, cy-22, cx+15, cy+8), fill=ACCENT_DARK, width=2)
    d.line((cx-2, cy-7, cx+32, cy-7), fill=ACCENT_DARK, width=2)

def card_with_label(d, top, height, brand):
    """OA公式アカウントラベル + 吹き出し + シャドウ付きカード"""
    # OA label (white bg + green text)
    label_text = f"{brand}●LINE公式アカウント"
    label_h = 24
    label_w = 180
    label_x = 20
    label_y = top - 26
    d.rounded_rectangle((label_x, label_y, label_x+label_w, label_y+label_h), radius=6, fill="#FFFFFF", outline=LINE_GREEN, width=2)
    d.text((label_x+12, label_y+6), label_text, font=font(9, True), fill=LINE_GREEN)

    # shadow (light gray)
    shadow_box = (22, top+3, W-18, top+height+3)
    d.rounded_rectangle(shadow_box, radius=8, fill=(230, 230, 230))

    # main card
    box = (20, top, W-20, top+height)
    d.rounded_rectangle(box, radius=8, fill="#FFFFFF", outline="#E0E0E0", width=1)

    # tail (speech bubble)
    tail_points = [(box[0]+16, top), (box[0]+8, top-10), (box[0]+24, top-4)]
    d.polygon(tail_points, fill="#FFFFFF")
    d.polygon(tail_points, outline="#E0E0E0", width=1)

    return box

def panelA():
    img, d, chat_top = phone_shell()
    y = chat_top + 50
    card_h = 372
    box = card_with_label(d, y, card_h, "ハコブーン")
    yy = y + 16
    d.text((box[0]+16, yy), "お荷物お届け予定のお知らせ", font=font(13, True), fill="#1a1a1a")
    yy += 26
    illus = (box[0]+16, yy, box[2]-16, yy+108)
    d.rounded_rectangle(illus, radius=10, fill="#FFF1E6")
    icon_person_box(d, (illus[0]+illus[2])/2-6, yy+70)
    yy += 118
    numbox = (box[0]+16, yy, box[2]-16, yy+30)
    d.rounded_rectangle(numbox, radius=6, outline="#DADDE1", width=2)
    d.text((box[0]+26, yy+7), "お問い合わせ番号：****-****-0650", font=font(9, True), fill="#333333")
    yy += 40
    lines = ["送り状番号を入力すると、LINEで通知を受け", "取れなかった場合でもお荷物の状況を確認", "できます。"]
    for line in lines:
        d.text((box[0]+16, yy), line, font=font(9), fill="#666666")
        yy += 15
    yy += 6
    d.text((box[0]+16, yy), "この通知が届く理由は？", font=font(9, True), fill="#888888")
    yy = box[3] - 46
    btn = (box[0]+16, yy, box[2]-16, yy+34)
    d.rounded_rectangle(btn, radius=8, fill=ACCENT)
    bbox = d.textbbox((0,0), "日時変更へ進む", font=font(12, True))
    tw = bbox[2]-bbox[0]
    d.text((W/2-tw/2, yy+9), "日時変更へ進む", font=font(12, True), fill="#FFFFFF")
    d.text((box[0], box[3]+12), "18:02", font=font(9), fill="#8A93A8")
    img.save("richcard1_notice.png")

def panelB():
    img, d, chat_top = phone_shell()
    y = chat_top + 50
    card_h = 400
    box = card_with_label(d, y, card_h, "ハコブーン")
    yy = y + 16
    d.text((box[0]+16, yy), "お荷物お届けのお知らせ", font=font(13, True), fill="#1a1a1a")
    yy += 20
    d.text((box[0]+16, yy), "配送センターからのお荷物をお届け予定です。", font=font(9), fill="#666666")
    yy += 26
    steps = ["集荷", "輸送中", "配達中", "完了"]
    seg_w = (box[2]-box[0]-32)/4
    for i, s in enumerate(steps):
        x0 = box[0]+16 + i*seg_w
        fill = ACCENT if i <= 2 else "#E5E5E5"
        d.rounded_rectangle((x0, yy, x0+seg_w-6, yy+8), radius=4, fill=fill)
        d.text((x0, yy+13), s, font=font(8), fill="#888888")
    yy += 38
    promo = (box[0]+16, yy, box[2]-16, yy+50)
    d.rounded_rectangle(promo, radius=10, fill="#FFF7EC")
    d.text((box[0]+28, yy+9), "ポイント山分けキャンペーン開催中", font=font(10, True), fill=ACCENT_DARK)
    d.text((box[0]+28, yy+27), "この配送センターご利用でポイントGET!", font=font(8), fill="#8A7A6A")
    yy += 62
    btn = (box[0]+16, yy, box[2]-16, yy+34)
    d.rounded_rectangle(btn, radius=8, fill=ACCENT)
    bbox = d.textbbox((0,0), "いますぐエントリー！", font=font(12, True))
    tw = bbox[2]-bbox[0]
    d.text((W/2-tw/2, yy+9), "いますぐエントリー！", font=font(12, True), fill="#FFFFFF")
    yy += 44
    d.polygon([(box[0]+18, yy+4), (box[0]+18, yy+13), (box[0]+24, yy+8.5)], fill="#3366CC")
    d.text((box[0]+30, yy), "受け取り日時・場所を指定", font=font(10, True), fill="#3366CC")
    yy += 22
    d.polygon([(box[0]+18, yy+4), (box[0]+18, yy+13), (box[0]+24, yy+8.5)], fill="#3366CC")
    d.text((box[0]+30, yy), "配達状況を確認", font=font(10, True), fill="#3366CC")
    yy += 28
    d.line((box[0]+16, yy, box[2]-16, yy), fill="#EEEEEE", width=1)
    yy += 10
    details = [("サービス名", "宅配便（配送）"), ("個数", "1"), ("配達予定日", "本日 18:00-20:00")]
    for k, v in details:
        d.text((box[0]+16, yy), k, font=font(8), fill="#999999")
        d.text((box[0]+120, yy), v, font=font(9, True), fill="#333333")
        yy += 18
    d.text((box[0], box[3]+12), "18:02", font=font(9), fill="#8A93A8")
    img.save("richcard2_status.png")

def panelC():
    img, d, chat_top = phone_shell()
    y = chat_top + 50
    card_h = 330
    box = card_with_label(d, y, card_h, "ハコブーン")
    yy = y + 16
    d.text((box[0]+16, yy), "ご希望の配達日をお選びください", font=font(12, True), fill="#1a1a1a")
    yy += 26
    days = ["今日", "明日", "それ以降"]
    dw = (box[2]-box[0]-32-16)/3
    for i, day in enumerate(days):
        x0 = box[0]+16 + i*(dw+8)
        fill = ACCENT if i == 0 else "#FFE9DA"
        d.rounded_rectangle((x0, yy, x0+dw, yy+42), radius=8, fill=fill)
        bbox = d.textbbox((0,0), day, font=font(11, True))
        tw = bbox[2]-bbox[0]
        d.text((x0+dw/2-tw/2, yy+13), day, font=font(11, True), fill="#FFFFFF" if i==0 else ACCENT_DARK)
    yy += 54
    d.text((box[0]+16, yy), "7月2日(木)、ご希望の配達時間帯を選択できます", font=font(9), fill="#666666")
    yy += 20
    slots = ["希望なし", "午前中", "14-16", "16-18", "18-20", "20-21"]
    cols = 3
    cw = (box[2]-box[0]-32-16)/cols
    ch = 48
    for i, s in enumerate(slots):
        r, c = divmod(i, cols)
        x0 = box[0]+16 + c*(cw+8)
        y0 = yy + r*(ch+8)
        d.rounded_rectangle((x0, y0, x0+cw, y0+ch), radius=8, fill="#FFE9DA")
        ccx, ccy = x0+cw/2, y0+15
        d.ellipse((ccx-7, ccy-7, ccx+7, ccy+7), outline=ACCENT_DARK, width=2)
        d.line((ccx, ccy, ccx, ccy-4), fill=ACCENT_DARK, width=2)
        d.line((ccx, ccy, ccx+3, ccy), fill=ACCENT_DARK, width=2)
        bbox = d.textbbox((0,0), s, font=font(9, True))
        tw = bbox[2]-bbox[0]
        d.text((x0+cw/2-tw/2, y0+27), s, font=font(9, True), fill=ACCENT_DARK)
    yy2 = yy + 2*(ch+8) + 12
    bbox = d.textbbox((0,0), "荷物問い合わせ・再配達依頼", font=font(9, True))
    tw = bbox[2]-bbox[0]
    d.text((W/2-tw/2, yy2), "荷物問い合わせ・再配達依頼", font=font(9, True), fill="#3366CC")
    d.text((box[0], box[3]+12), "18:02", font=font(9), fill="#8A93A8")
    img.save("richcard3_datetime.png")

import os
os.chdir(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_flow_example")
panelA()
panelB()
panelC()
print("done")
