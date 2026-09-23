# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import textwrap

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 360, 480
LINE_GREEN = "#06C755"

def icon_truck(d, cx, cy, s, color, width=5):
    d.rectangle((cx-s, cy-s*0.4, cx+s*0.3, cy+s*0.5), outline=color, width=width)
    d.polygon([(cx+s*0.3, cy-s*0.1), (cx+s*0.9, cy-s*0.1), (cx+s*1.1, cy+s*0.2), (cx+s*1.1, cy+s*0.5), (cx+s*0.3, cy+s*0.5)], outline=color, width=width)
    d.ellipse((cx-s*0.7, cy+s*0.35, cx-s*0.3, cy+s*0.75), outline=color, width=width)
    d.ellipse((cx+s*0.5, cy+s*0.35, cx+s*0.9, cy+s*0.75), outline=color, width=width)

def icon_bolt(d, cx, cy, s, color, width=5):
    pts = [(cx+s*0.15, cy-s), (cx-s*0.6, cy+s*0.15), (cx-s*0.05, cy+s*0.15), (cx-s*0.2, cy+s), (cx+s*0.6, cy-s*0.2), (cx+s*0.05, cy-s*0.2)]
    d.polygon(pts, fill=color)

def icon_house(d, cx, cy, s, color, width=5):
    d.polygon([(cx-s, cy+s*0.1), (cx, cy-s*0.8), (cx+s, cy+s*0.1)], outline=color, width=width)
    d.rectangle((cx-s*0.65, cy+s*0.1, cx+s*0.65, cy+s), outline=color, width=width)
    d.rectangle((cx-s*0.18, cy+s*0.45, cx+s*0.18, cy+s), outline=color, width=width)

def icon_glasses(d, cx, cy, s, color, width=5):
    d.ellipse((cx-s*1.1, cy-s*0.4, cx-s*0.2, cy+s*0.5), outline=color, width=width)
    d.ellipse((cx+s*0.2, cy-s*0.4, cx+s*1.1, cy+s*0.5), outline=color, width=width)
    d.line((cx-s*0.2, cy, cx+s*0.2, cy), fill=color, width=width)
    d.line((cx-s*1.1, cy-s*0.1, cx-s*1.35, cy-s*0.3), fill=color, width=width)
    d.line((cx+s*1.1, cy-s*0.1, cx+s*1.35, cy-s*0.3), fill=color, width=width)

ICONS = {"delivery": icon_truck, "utility": icon_bolt, "realestate": icon_house, "retail": icon_glasses}

def frame1(path, key, domain, action_title, order_no_label, note_lines):
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)
    d.rectangle((6, 6, W-6, 44), fill="#F3F3F3")
    d.ellipse((22, 20, 32, 30), fill="#FF5F57")
    d.ellipse((36, 20, 46, 30), fill="#FEBC2E")
    d.ellipse((50, 20, 60, 30), fill="#28C840")
    d.rounded_rectangle((80, 16, W-26, 34), radius=8, outline="#DADFE3", width=1)
    d.text((90, 20), domain, font=font(10), fill="#888888")
    cy = 100
    d.ellipse((W/2-40, cy-40, W/2+40, cy+40), fill="#EAFBF1")
    ICONS[key](d, W/2, cy, 20, LINE_GREEN)
    bbox = d.textbbox((0,0), action_title, font=font(15, True))
    tw = bbox[2]-bbox[0]
    d.text((W/2-tw/2, cy+64), action_title, font=font(15, True), fill="#222222")
    bbox2 = d.textbbox((0,0), order_no_label, font=font(12))
    tw2 = bbox2[2]-bbox2[0]
    d.text((W/2-tw2/2, cy+92), order_no_label, font=font(12), fill="#666666")
    note_box = (30, cy+132, W-30, cy+132+22*len(note_lines)+18)
    d.rounded_rectangle(note_box, radius=10, fill="#F1F1F1")
    yy = cy+144
    for line in note_lines:
        d.text((44, yy), line, font=font(12), fill="#333333")
        yy += 22
    d.text((10, H-28), "① サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save(path)

def frame2(path, key, brand, notif_title, body_lines):
    img = Image.new("RGB", (W, H), "#2B2E36")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#3A3E47", width=2)
    d.text((W/2-46, 60), "18:02", font=font(46, True), fill="#FFFFFF")
    d.text((W/2-44, 112), "2026年7月2日（木）", font=font(13), fill="#CBCBCB")
    nb = (24, 170, W-24, 170+96)
    d.rounded_rectangle(nb, radius=18, fill="#3A3D45")
    icon_bg = (nb[0]+14, nb[1]+14, nb[0]+14+32, nb[1]+14+32)
    d.rounded_rectangle(icon_bg, radius=8, fill=LINE_GREEN)
    icx, icy = (icon_bg[0]+icon_bg[2])/2, (icon_bg[1]+icon_bg[3])/2
    ICONS[key](d, icx, icy, 8, "#FFFFFF", width=3)
    d.text((nb[0]+58, nb[1]+12), brand, font=font(12, True), fill="#FFFFFF")
    d.text((nb[2]-52, nb[1]+12), "いま", font=font(10), fill="#AAAAAA")
    d.text((nb[0]+58, nb[1]+32), notif_title, font=font(12, True), fill="#FFFFFF")
    yy = nb[1]+52
    for line in body_lines:
        d.text((nb[0]+58, yy), line, font=font(11), fill="#D8D8D8")
        yy += 18
    d.text((10, H-28), "② ロック画面にプッシュ通知が届く（未タップ）", font=font(11, True), fill="#8FD8B0")
    img.save(path)

def frame3(path, key, brand, thanks_title, thanks_body, reply_text):
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)
    y = 40
    d.text((26, y), f"{brand}●LINE公式アカウント", font=font(11), fill=LINE_GREEN)
    y += 22
    bubble = (26, y, W-26, y+70)
    d.rounded_rectangle(bubble, radius=4, fill="#F1F1F1")
    icon_pos = (bubble[2]-56, y+15, bubble[2]-26, y+45)
    icx, icy = (icon_pos[0]+icon_pos[2])/2, (icon_pos[1]+icon_pos[3])/2
    ICONS[key](d, icx, icy, 10, LINE_GREEN, width=3)
    d.text((42, y+14), thanks_title, font=font(13, True), fill="#1a1a1a")
    d.text((42, y+40), thanks_body, font=font(11), fill="#333333")
    y += 70+8
    d.text((26, y), "18:02", font=font(10), fill="#AAAAAA")
    y += 40
    bbox = d.textbbox((0,0), reply_text, font=font(12))
    tw = bbox[2]-bbox[0]
    bubble2 = (W-26-tw-28, y, W-26, y+44)
    d.rounded_rectangle(bubble2, radius=4, fill=LINE_GREEN)
    d.text((bubble2[0]+14, y+14), reply_text, font=font(12), fill="#FFFFFF")
    fb = (30, H-110, W-30, H-70)
    d.rounded_rectangle(fb, radius=10, fill="#EAFBF1")
    d.text((44, H-100), "「友だち」に自動追加されました", font=font(12, True), fill="#128A45")
    d.text((10, H-28), "③ タップでトークが開き自然に友だち化", font=font(11, True), fill="#1E2761")
    img.save(path)

brands = [
    dict(key="delivery", brand="ハコブーン", domain="hakobin-shop.example.com",
         action_title="ご注文ありがとうございます", order_no="注文番号：No.20260702-114",
         note=["LINEでお届け状況を", "リアルタイムにお知らせします"],
         notif_title="お荷物お届けのお知らせ",
         body=["ご注文の商品を本日18時〜20時に", "お届け予定です。"],
         thanks_title="友だち追加ありがとうございます！",
         thanks_body="お届け状況はこちらから確認できます。",
         reply="ありがとうございます！"),
    dict(key="utility", brand="でんきねっと", domain="denki-net.example.com",
         action_title="お引っ越し手続きが完了しました", order_no="お客様番号：4471-2290",
         note=["LINEで料金確定のお知らせを", "受け取れます"],
         notif_title="今月の料金確定のお知らせ",
         body=["今月のご利用料金が確定しました。", "マイページよりご確認ください。"],
         thanks_title="友だち追加ありがとうございます！",
         thanks_body="今後の料金確定はこちらでお知らせします。",
         reply="ありがとうございます！"),
    dict(key="realestate", brand="いえらいふ", domain="ielife-estate.example.com",
         action_title="資料請求フォームを送信しました", order_no="受付番号：RE-88213",
         note=["LINEで資料到着のお知らせを", "受け取れます"],
         notif_title="資料請求完了のお知らせ",
         body=["お問い合わせいただいた物件の", "資料をご用意しました。"],
         thanks_title="友だち追加ありがとうございます！",
         thanks_body="担当者よりこちらでご連絡します。",
         reply="よろしくお願いします！"),
    dict(key="retail", brand="メガネのミライ", domain="megane-mirai.example.com",
         action_title="ご購入手続きが完了しました", order_no="注文番号：MM-50213",
         note=["LINEで発送状況のお知らせを", "受け取れます"],
         notif_title="ご購入完了のお知らせ",
         body=["ご購入ありがとうございます。", "発送準備が整い次第お知らせします。"],
         thanks_title="友だち追加ありがとうございます！",
         thanks_body="発送状況はこちらから確認できます。",
         reply="ありがとうございます！"),
]

for b in brands:
    frame1(f"flow1_{b['key']}.png", b["key"], b["domain"], b["action_title"], b["order_no"], b["note"])
    frame2(f"flow2_{b['key']}.png", b["key"], b["brand"], b["notif_title"], b["body"])
    frame3(f"flow3_{b['key']}.png", b["key"], b["brand"], b["thanks_title"], b["thanks_body"], b["reply"])

print("done")
