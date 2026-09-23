# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 340, 560
ACCENT = "#FF7A33"
ACCENT_DARK = "#C85A1E"
LINE_GREEN = "#06C755"

def base(title_brand="ハコブーン"):
    img = Image.new("RGB", (W, H), "#EDEDED")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((4, 4, W-4, H-4), radius=14, outline="#DADFE3", width=2)
    # LINE OA header bar (simple, matching real reference)
    d.rectangle((4, 4, W-4, 40), fill="#FFFFFF")
    d.line((4, 40, W-4, 40), fill="#E5E5E5", width=1)
    d.ellipse((14, 9, 14+24, 9+24), fill=ACCENT)
    d.text((14+6, 9+3), "H", font=font(13, True), fill="#FFFFFF")
    d.text((46, 12), title_brand, font=font(13, True), fill="#222222")
    # top-right icons (house / menu)
    d.rounded_rectangle((W-56, 12, W-56+16, 12+16), radius=3, outline="#777777", width=2)
    hx = W-30
    for i in range(3):
        d.line((hx, 14+i*5, hx+14, 14+i*5), fill="#777777", width=2)
    return img, d

def icon_person_box(d, cx, cy):
    # simple person
    d.ellipse((cx-30, cy-58, cx-14, cy-42), fill="#3a3a3a")
    d.polygon([(cx-34, cy-40), (cx-10, cy-40), (cx-6, cy+6), (cx-38, cy+6)], fill="#3a3a3a")
    # box being carried
    d.rectangle((cx-2, cy-30, cx+40, cy+8), fill=ACCENT)
    d.line((cx-2, cy-30, cx+40, cy+8), fill=ACCENT_DARK, width=2)
    d.line((cx+19, cy-30, cx+19, cy+8), fill=ACCENT_DARK, width=2)
    d.line((cx-2, cy-11, cx+40, cy-11), fill=ACCENT_DARK, width=2)

def panelA():
    img, d = base()
    y = 52
    d.text((20, y), "お荷物お届け予定のお知らせ", font=font(15, True), fill="#1a1a1a")
    y += 30
    illus = (20, y, W-20, y+150)
    d.rounded_rectangle(illus, radius=12, fill="#FFF1E6")
    icon_person_box(d, W/2-10, y+95)
    y += 162
    box = (20, y, W-20, y+34)
    d.rounded_rectangle(box, radius=6, outline="#DADDE1", width=2)
    d.text((30, y+8), "お問い合わせ番号：****-****-0650", font=font(11, True), fill="#333333")
    y += 46
    lines = ["送り状番号を入力すると、LINEで通知を受け取れ", "なかった場合でもお荷物の状況をいつでも確認", "できます。時刻表示は目安となります。"]
    for line in lines:
        d.text((20, y), line, font=font(10), fill="#666666")
        y += 18
    y += 6
    d.text((20, y), "この通知が届く理由は？", font=font(10, True), fill="#888888")
    y = H - 70
    btn = (20, y, W-20, y+40)
    d.rounded_rectangle(btn, radius=8, fill=ACCENT)
    bbox = d.textbbox((0,0), "日時変更へ進む", font=font(13, True))
    tw = bbox[2]-bbox[0]
    d.text((W/2-tw/2, y+11), "日時変更へ進む", font=font(13, True), fill="#FFFFFF")
    img.save("richcard1_notice.png")

def panelB():
    img, d = base()
    y = 50
    d.text((20, y), "お荷物お届けのお知らせ", font=font(15, True), fill="#1a1a1a")
    y += 26
    d.text((20, y), "配送センターからのお荷物をお届け予定です。", font=font(10), fill="#666666")
    y += 30
    # progress bar (segmented)
    steps = ["集荷", "輸送中", "配達中", "完了"]
    seg_w = (W-40)/4
    for i, s in enumerate(steps):
        x0 = 20 + i*seg_w
        fill = ACCENT if i <= 2 else "#E5E5E5"
        d.rounded_rectangle((x0, y, x0+seg_w-6, y+10), radius=5, fill=fill)
        d.text((x0, y+16), s, font=font(9), fill="#888888")
    y += 46
    promo = (20, y, W-20, y+56)
    d.rounded_rectangle(promo, radius=10, fill="#FFF7EC")
    d.text((32, y+10), "ポイント山分けキャンペーン開催中", font=font(11, True), fill=ACCENT_DARK)
    d.text((32, y+30), "この配送センターご利用でポイントGET!", font=font(9), fill="#8A7A6A")
    y += 70
    btn = (20, y, W-20, y+40)
    d.rounded_rectangle(btn, radius=8, fill=ACCENT)
    bbox = d.textbbox((0,0), "いますぐエントリー！", font=font(13, True))
    tw = bbox[2]-bbox[0]
    d.text((W/2-tw/2, y+11), "いますぐエントリー！", font=font(13, True), fill="#FFFFFF")
    y += 50
    d.polygon([(22, y+4), (22, y+14), (28, y+9)], fill="#3366CC")
    d.text((34, y), "受け取り日時・場所を指定", font=font(11, True), fill="#3366CC")
    y += 24
    d.polygon([(22, y+4), (22, y+14), (28, y+9)], fill="#3366CC")
    d.text((34, y), "配達状況を確認", font=font(11, True), fill="#3366CC")
    y += 32
    d.line((20, y, W-20, y), fill="#EEEEEE", width=1)
    y += 12
    details = [("サービス名", "宅配便（配送）"), ("個数", "1"), ("配達予定日", "本日 18:00-20:00")]
    for k, v in details:
        d.text((20, y), k, font=font(9), fill="#999999")
        d.text((140, y), v, font=font(10, True), fill="#333333")
        y += 20
    img.save("richcard2_status.png")

def panelC():
    img, d = base()
    y = 50
    d.text((20, y), "ご希望の配達日をお選びください", font=font(13, True), fill="#1a1a1a")
    y += 30
    days = ["今日", "明日", "それ以降"]
    dw = (W-40-16)/3
    for i, day in enumerate(days):
        x0 = 20 + i*(dw+8)
        fill = ACCENT if i == 0 else "#FFE9DA"
        d.rounded_rectangle((x0, y, x0+dw, y+50), radius=8, fill=fill)
        bbox = d.textbbox((0,0), day, font=font(12, True))
        tw = bbox[2]-bbox[0]
        d.text((x0+dw/2-tw/2, y+16), day, font=font(12, True), fill="#FFFFFF" if i==0 else ACCENT_DARK)
    y += 66
    d.text((20, y), "7月2日(木)、ご希望の配達時間帯を選択できます", font=font(10), fill="#666666")
    y += 24
    slots = ["希望なし", "午前中", "14-16", "16-18", "18-20", "20-21"]
    cols = 3
    cw = (W-40-16)/cols
    ch = 56
    for i, s in enumerate(slots):
        r, c = divmod(i, cols)
        x0 = 20 + c*(cw+8)
        y0 = y + r*(ch+8)
        d.rounded_rectangle((x0, y0, x0+cw, y0+ch), radius=8, fill="#FFE9DA")
        # clock icon
        ccx, ccy = x0+cw/2, y0+18
        d.ellipse((ccx-9, ccy-9, ccx+9, ccy+9), outline=ACCENT_DARK, width=2)
        d.line((ccx, ccy, ccx, ccy-5), fill=ACCENT_DARK, width=2)
        d.line((ccx, ccy, ccx+4, ccy), fill=ACCENT_DARK, width=2)
        bbox = d.textbbox((0,0), s, font=font(10, True))
        tw = bbox[2]-bbox[0]
        d.text((x0+cw/2-tw/2, y0+32), s, font=font(10, True), fill=ACCENT_DARK)
    y2 = y + 2*(ch+8) + 14
    bbox = d.textbbox((0,0), "荷物問い合わせ・再配達依頼", font=font(10, True))
    tw = bbox[2]-bbox[0]
    d.text((W/2-tw/2, y2), "荷物問い合わせ・再配達依頼", font=font(10, True), fill="#3366CC")
    img.save("richcard3_datetime.png")

import os
os.chdir(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_flow_example")
panelA()
panelB()
panelC()
print("done")
