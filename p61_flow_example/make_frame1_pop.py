# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import random

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 360, 480

def icon_truck(d, cx, cy, s, color, width=6):
    d.rectangle((cx-s, cy-s*0.4, cx+s*0.3, cy+s*0.5), outline=color, width=width)
    d.polygon([(cx+s*0.3, cy-s*0.1), (cx+s*0.9, cy-s*0.1), (cx+s*1.1, cy+s*0.2), (cx+s*1.1, cy+s*0.5), (cx+s*0.3, cy+s*0.5)], outline=color, width=width)
    d.ellipse((cx-s*0.7, cy+s*0.35, cx-s*0.3, cy+s*0.75), outline=color, width=width)
    d.ellipse((cx+s*0.5, cy+s*0.35, cx+s*0.9, cy+s*0.75), outline=color, width=width)

def icon_bolt(d, cx, cy, s, color, width=6):
    pts = [(cx+s*0.15, cy-s), (cx-s*0.6, cy+s*0.15), (cx-s*0.05, cy+s*0.15), (cx-s*0.2, cy+s), (cx+s*0.6, cy-s*0.2), (cx+s*0.05, cy-s*0.2)]
    d.polygon(pts, fill=color)

def icon_house(d, cx, cy, s, color, width=6):
    d.polygon([(cx-s, cy+s*0.1), (cx, cy-s*0.8), (cx+s, cy+s*0.1)], outline=color, width=width)
    d.rectangle((cx-s*0.65, cy+s*0.1, cx+s*0.65, cy+s), outline=color, width=width)
    d.rectangle((cx-s*0.18, cy+s*0.45, cx+s*0.18, cy+s), outline=color, width=width)

def icon_glasses(d, cx, cy, s, color, width=6):
    d.ellipse((cx-s*1.1, cy-s*0.4, cx-s*0.2, cy+s*0.5), outline=color, width=width)
    d.ellipse((cx+s*0.2, cy-s*0.4, cx+s*1.1, cy+s*0.5), outline=color, width=width)
    d.line((cx-s*0.2, cy, cx+s*0.2, cy), fill=color, width=width)
    d.line((cx-s*1.1, cy-s*0.1, cx-s*1.35, cy-s*0.3), fill=color, width=width)
    d.line((cx+s*1.1, cy-s*0.1, cx+s*1.35, cy-s*0.3), fill=color, width=width)

def confetti(d, box, colors, n, seed):
    rnd = random.Random(seed)
    for _ in range(n):
        x = rnd.uniform(box[0], box[2])
        y = rnd.uniform(box[1], box[3])
        r = rnd.uniform(3, 7)
        c = rnd.choice(colors)
        shape = rnd.choice(["circle", "square", "tri"])
        if shape == "circle":
            d.ellipse((x-r, y-r, x+r, y+r), fill=c)
        elif shape == "square":
            d.rectangle((x-r, y-r, x+r, y+r), fill=c)
        else:
            d.polygon([(x, y-r), (x-r, y+r), (x+r, y+r)], fill=c)

def blob_bg(d, cx, cy, r, color1, color2):
    d.ellipse((cx-r, cy-r, cx+r, cy+r), fill=color1)
    d.ellipse((cx-r*0.62, cy-r*0.62, cx+r*0.62, cy+r*0.62), fill=color2)

def hexc(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def header(d, domain, grad1, grad2):
    for i in range(38):
        t = i / 38
        col = tuple(int(grad1[j] + (grad2[j] - grad1[j]) * t) for j in range(3))
        d.line((6, 6 + i, W - 6, 6 + i), fill=col)
    d.ellipse((22, 20, 32, 30), fill="#FFFFFF")
    d.ellipse((36, 20, 46, 30), fill="#FFFFFF")
    d.ellipse((50, 20, 60, 30), fill="#FFFFFF")
    d.rounded_rectangle((80, 16, W - 26, 34), radius=8, fill="#FFFFFF")
    d.text((90, 20), domain, font=font(10), fill="#888888")

def pill(d, box, fill, text, textcolor):
    d.rounded_rectangle(box, radius=(box[3] - box[1]) // 2, fill=fill)
    bbox = d.textbbox((0, 0), text, font=font(12, True))
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    cx = (box[0] + box[2]) / 2
    cy = (box[1] + box[3]) / 2
    d.text((cx - tw / 2 - bbox[0], cy - th / 2 - bbox[1]), text, font=font(12, True), fill=textcolor)

def make_delivery():
    accent = "#FF7A33"; accent2 = "#FFB562"; light = "#FFF1E6"
    img = Image.new("RGB", (W, H), light)
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W - 6, H - 6), radius=16, outline="#DADFE3", width=2, fill=light)
    header(d, "hakobin-shop.example.com", hexc(accent), hexc(accent2))
    blob_bg(d, W / 2, 130, 64, "#FFE0C7", "#FFFFFF")
    icon_truck(d, W / 2, 130, 30, accent, width=7)
    confetti(d, (26, 60, W - 26, 200), [accent, accent2, "#FFD166"], 14, 1)
    label = "ご注文ありがとうございます！"
    bbox = d.textbbox((0, 0), label, font=font(16, True))
    tw = bbox[2] - bbox[0]
    d.text((W / 2 - tw / 2, 206), label, font=font(16, True), fill="#7A3300")
    card = (26, 244, W - 26, 244 + 112)
    d.rounded_rectangle(card, radius=16, fill="#FFFFFF")
    pill(d, (44, 258, 168, 284), accent, "注文番号", "#FFFFFF")
    d.text((44, 296), "No.20260702-114", font=font(13, True), fill="#333333")
    d.line((44, 320, W - 44, 320), fill="#F0E3D8", width=2)
    d.text((44, 330), "お届け予定：本日 18:00-20:00", font=font(11, True), fill="#555555")
    note = (26, 372, W - 26, 414)
    d.rounded_rectangle(note, radius=14, fill="#06C755")
    d.text((40, 384), "LINEでお届け状況をお知らせ！", font=font(11, True), fill="#FFFFFF")
    d.text((10, H - 28), "(1) サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_delivery.png")

def make_utility():
    accent = "#2F80ED"; accent2 = "#6FB1FC"; light = "#EAF3FF"
    img = Image.new("RGB", (W, H), light)
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W - 6, H - 6), radius=16, outline="#DADFE3", width=2, fill=light)
    header(d, "denki-net.example.com", hexc(accent), hexc(accent2))
    d.text((26, 60), "お引っ越し手続きが完了しました", font=font(13, True), fill="#123A6B")
    card = (26, 96, W - 26, 96 + 180)
    d.rounded_rectangle(card, radius=16, fill="#FFFFFF")
    blob_bg(d, 66, 96 + 44, 30, "#D6E8FF", "#FFFFFF")
    icon_bolt(d, 66, 96 + 44, 16, accent, width=6)
    d.text((104, 96 + 20), "お客様番号", font=font(10), fill="#8A93A8")
    d.text((104, 96 + 38), "4471-2290", font=font(13, True), fill="#222222")
    band = (26, 96 + 90, W - 26, 96 + 90 + 66)
    for i in range(66):
        t = i / 66
        col = tuple(int(hexc(accent)[j] + (hexc(accent2)[j] - hexc(accent)[j]) * t) for j in range(3))
        d.line((26, 96 + 90 + i, W - 26, 96 + 90 + i), fill=col)
    d.text((44, 96 + 100), "今月の確定予定額", font=font(10), fill="#EAF3FF")
    d.text((44, 96 + 118), "5,280円（目安）", font=font(20, True), fill="#FFFFFF")
    note = (26, 300, W - 26, 342)
    d.rounded_rectangle(note, radius=14, fill="#06C755")
    d.text((40, 312), "LINEで料金確定のお知らせ！", font=font(11, True), fill="#FFFFFF")
    d.text((10, H - 28), "(1) サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_utility.png")

def make_realestate():
    accent = "#2FAE66"; accent2 = "#8AE2AE"; light = "#EAFBF1"
    img = Image.new("RGB", (W, H), light)
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W - 6, H - 6), radius=16, outline="#DADFE3", width=2, fill=light)
    header(d, "ielife-estate.example.com", hexc(accent), hexc(accent2))
    photo = (26, 58, W - 26, 58 + 126)
    for i in range(126):
        t = i / 126
        col = tuple(int(hexc(accent2)[j] + (hexc("#FFFFFF")[j] - hexc(accent2)[j]) * t) for j in range(3))
        d.line((26, 58 + i, W - 26, 58 + i), fill=col)
    confetti(d, photo, ["#FFFFFF"], 10, 3)
    icon_house(d, W / 2, 58 + 63, 32, "#FFFFFF", width=8)
    label = "資料請求ありがとうございます！"
    bbox = d.textbbox((0, 0), label, font=font(15, True))
    tw = bbox[2] - bbox[0]
    d.text((W / 2 - tw / 2, 200), label, font=font(15, True), fill="#0F5C33")
    card = (26, 232, W - 26, 232 + 70)
    d.rounded_rectangle(card, radius=14, fill="#FFFFFF")
    pill(d, (44, 246, 158, 270), accent, "受付番号", "#FFFFFF")
    d.text((44, 280), "RE-88213", font=font(13, True), fill="#333333")
    note = (26, 318, W - 26, 360)
    d.rounded_rectangle(note, radius=14, fill="#06C755")
    d.text((40, 330), "LINEで資料到着をお知らせ！", font=font(11, True), fill="#FFFFFF")
    d.text((10, H - 28), "(1) サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_realestate.png")

def make_retail():
    accent = "#8A3FE4"; accent2 = "#C9A2FF"; light = "#F5EEFF"
    img = Image.new("RGB", (W, H), light)
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, W - 6, H - 6), radius=16, outline="#DADFE3", width=2, fill=light)
    header(d, "megane-mirai.example.com", hexc(accent), hexc(accent2))
    label = "ご購入ありがとうございます！"
    bbox = d.textbbox((0, 0), label, font=font(16, True))
    tw = bbox[2] - bbox[0]
    d.text((W / 2 - tw / 2, 62), label, font=font(16, True), fill="#4B1D8C")
    card = (26, 100, W - 26, 100 + 130)
    d.rounded_rectangle(card, radius=16, fill="#FFFFFF")
    thumb = (44, 100 + 16, 44 + 98, 100 + 16 + 98)
    blob_bg(d, (thumb[0] + thumb[2]) / 2, (thumb[1] + thumb[3]) / 2, 49, "#EBDCFF", "#FFFFFF")
    icon_glasses(d, (thumb[0] + thumb[2]) / 2, (thumb[1] + thumb[3]) / 2, 20, accent, width=6)
    d.text((160, 100 + 26), "フレーム TYPE-03", font=font(12, True), fill="#333333")
    d.text((160, 100 + 48), "注文番号：MM-50213", font=font(10), fill="#8A93A8")
    pill(d, (160, 100 + 70, 236, 100 + 96), accent, "12,800円", "#FFFFFF")
    note = (26, 262, W - 26, 304)
    d.rounded_rectangle(note, radius=14, fill="#06C755")
    d.text((40, 274), "LINEで発送状況をお知らせ！", font=font(11, True), fill="#FFFFFF")
    d.text((10, H - 28), "(1) サイトで手続き完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
    img.save("flow1_retail.png")

make_delivery()
make_utility()
make_realestate()
make_retail()
print("done")
