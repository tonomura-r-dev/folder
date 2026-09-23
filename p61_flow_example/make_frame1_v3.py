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

def browser_chrome(d, domain, accent):
    d.rectangle((6, 6, W-6, 44), fill=accent)
    d.ellipse((22, 20, 32, 30), fill="#FFFFFF")
    d.ellipse((36, 20, 46, 30), fill="#FFFFFF")
    d.ellipse((50, 20, 60, 30), fill="#FFFFFF")
    d.rounded_rectangle((80, 16, W-26, 34), radius=8, fill="#FFFFFF")
    d.text((90, 20), domain, font=font(10), fill="#888888")

# ---- delivery: 配送完了レシート風 ----
def make_delivery():
    accent = "#E8622C"
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)
    browser_chrome(d, "hakobin-shop.example.com", accent)
    d.rounded_rectangle((26, 60, W-26, 90), radius=6, fill=accent)
    d.text((40, 68), "お荷物 受付完了", font=font(14, True), fill="#FFFFFF")
    icon_truck(d, W/2, 140, 26, accent)
    d.text((40, 190), "ご注文ありがとうございます", font=font(14, True), fill="#222222")
    receipt = (26, 220, W-26, 220+96)
    d.rounded_rectangle(receipt, radius=8, outline="#E5E5E5", width=1)
    d.text((40, 232), "注文番号", font=font(10), fill="#999999")
    d.text((40, 248), "No.20260702-114", font=font(12, True), fill="#333333")
    d.line((40, 274, W-40, 274), fill="#EEEEEE", width=1)
    d.text((40, 282), "お届け予定", font=font(10), fill="#999999")
    d.text((40, 298), "本日 18:00〜20:00", font=font(12, True), fill="#333333")
    note = (26, 330, W-26, 372)
    d.rounded_rectangle(note, radius=8, fill="#FDECE3")
    d.text((40, 344), "LINEでお届け状況をお知らせします", font=font(11), fill=accent)
    d.text((10, H-28), "① サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_delivery.png")

# ---- utility: 料金確定ダッシュボード風 ----
def make_utility():
    accent = "#2C6FE0"
    img = Image.new("RGB", (W, H), "#F7F9FC")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2, fill="#F7F9FC")
    browser_chrome(d, "denki-net.example.com", accent)
    d.text((26, 62), "お引っ越し手続きが完了しました", font=font(13, True), fill="#222222")
    card = (26, 96, W-26, 96+150)
    d.rounded_rectangle(card, radius=10, fill="#FFFFFF")
    icon_bolt(d, 56, 96+40, 16, accent)
    d.text((90, 96+20), "お客様番号", font=font(10), fill="#999999")
    d.text((90, 96+38), "4471-2290", font=font(13, True), fill="#333333")
    d.line((26, 96+70, W-26, 96+70), fill="#EEEEEE", width=1)
    d.text((44, 96+84), "今月の確定予定額", font=font(10), fill="#999999")
    d.text((44, 96+100), "5,280円（目安）", font=font(20, True), fill=accent)
    note = (26, 264, W-26, 306)
    d.rounded_rectangle(note, radius=8, fill="#E7EFFD")
    d.text((40, 278), "LINEで料金確定のお知らせを受け取れます", font=font(11), fill=accent)
    d.text((10, H-28), "① サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_utility.png")

# ---- realestate: 物件資料請求サンクス風 ----
def make_realestate():
    accent = "#2E9E5B"
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)
    browser_chrome(d, "ielife-estate.example.com", accent)
    photo = (26, 58, W-26, 58+110)
    d.rounded_rectangle(photo, radius=8, fill="#DDEEE2")
    icon_house(d, W/2, 58+55, 26, accent)
    d.text((26, 182), "資料請求フォームを送信しました", font=font(13, True), fill="#222222")
    card = (26, 212, W-26, 212+66)
    d.rounded_rectangle(card, radius=8, outline="#E5E5E5", width=1)
    d.text((40, 224), "受付番号", font=font(10), fill="#999999")
    d.text((40, 240), "RE-88213", font=font(12, True), fill="#333333")
    note = (26, 294, W-26, 336)
    d.rounded_rectangle(note, radius=8, fill="#E6F5EB")
    d.text((40, 308), "LINEで資料到着のお知らせを受け取れます", font=font(11), fill=accent)
    d.text((10, H-28), "① サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_realestate.png")

# ---- retail: 購入完了・商品サムネ風 ----
def make_retail():
    accent = "#7B3FE4"
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)
    browser_chrome(d, "megane-mirai.example.com", accent)
    d.text((26, 60), "ご購入手続きが完了しました", font=font(14, True), fill="#222222")
    card = (26, 92, W-26, 92+120)
    d.rounded_rectangle(card, radius=10, fill="#F5F0FC")
    thumb = (44, 92+16, 44+88, 92+16+88)
    d.rounded_rectangle(thumb, radius=8, fill="#FFFFFF")
    icon_glasses(d, thumb[0]+44, thumb[1]+44, 18, accent)
    d.text((150, 92+24), "フレーム TYPE-03", font=font(12, True), fill="#333333")
    d.text((150, 92+46), "注文番号：MM-50213", font=font(10), fill="#999999")
    d.text((150, 92+68), "¥12,800", font=font(15, True), fill=accent)
    note = (26, 244, W-26, 286)
    d.rounded_rectangle(note, radius=8, fill="#F0E9FB")
    d.text((40, 258), "LINEで発送状況のお知らせを受け取れます", font=font(11), fill=accent)
    d.text((10, H-28), "① サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_retail.png")

make_delivery()
make_utility()
make_realestate()
make_retail()
print("done")
