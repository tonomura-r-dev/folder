# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import textwrap

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 420, 780
RADIUS = 28
CHAT_BG = "#F7F7F7"

def rr(draw, box, radius, **kw):
    draw.rounded_rectangle(box, radius=radius, **kw)

def icon_home(draw, cx, cy, s, color):
    draw.polygon([(cx-s, cy), (cx, cy-s*0.9), (cx+s, cy)], outline=color, width=3)
    draw.rectangle((cx-s*0.65, cy, cx+s*0.65, cy+s*0.85), outline=color, width=3)

def icon_gift(draw, cx, cy, s, color):
    draw.rectangle((cx-s, cy-s*0.2, cx+s, cy+s), outline=color, width=3)
    draw.line((cx, cy-s*0.2, cx, cy+s), fill=color, width=3)
    draw.line((cx-s, cy+s*0.15, cx+s, cy+s*0.15), fill=color, width=3)
    draw.ellipse((cx-s*0.4, cy-s*0.55, cx, cy-s*0.15), outline=color, width=3)
    draw.ellipse((cx, cy-s*0.55, cx+s*0.4, cy-s*0.15), outline=color, width=3)

def icon_phone(draw, cx, cy, s, color):
    draw.rounded_rectangle((cx-s*0.5, cy-s, cx+s*0.5, cy+s), radius=int(s*0.3), outline=color, width=3)
    draw.line((cx-s*0.15, cy+s*0.75, cx+s*0.15, cy+s*0.75), fill=color, width=3)

def make_line_notif(filename, accent, brand_letter, brand_name, title, body, caption, icon_fn3):
    img = Image.new("RGB", (W, H), "#F2F2F5")
    draw = ImageDraw.Draw(img)
    phone_box = (10, 10, W-10, H-56)
    rr(draw, phone_box, RADIUS, fill=CHAT_BG)
    rr(draw, phone_box, RADIUS, outline="#DADFE3", width=2)

    draw.rectangle((12, 12, W-12, 50), fill="#FFFFFF")
    draw.text((32, 22), "9:41", font=font(15, True), fill="#111111")
    draw.text((W-88, 23), "●●● 100%", font=font(11), fill="#555555")

    header_top = 50
    header_h = 54
    draw.rectangle((12, header_top, W-12, header_top+header_h), fill="#FFFFFF")
    draw.line((12, header_top+header_h, W-12, header_top+header_h), fill="#E5E5E5", width=1)

    draw.text((26, header_top+header_h/2-12), "‹", font=font(26, True), fill="#333333")
    draw.ellipse((58, header_top+9, 58+36, header_top+45), fill=accent)
    draw.text((58+11, header_top+15), brand_letter, font=font(15, True), fill="#FFFFFF")
    draw.text((104, header_top+header_h/2-9), brand_name, font=font(15, True), fill="#222222")

    px, py = W-92, header_top+header_h/2-9
    draw.rounded_rectangle((px, py, px+18, py+18), radius=5, outline="#555555", width=2)
    hx, hy = W-56, header_top+header_h/2-7
    for i in range(3):
        draw.line((hx, hy+i*6, hx+18, hy+i*6), fill="#555555", width=2)

    chat_top = header_top+header_h
    richmenu_h = 132
    inputbar_h = 52
    chat_bottom = phone_box[3] - richmenu_h - inputbar_h
    draw.rectangle((12, chat_top, W-12, chat_bottom), fill=CHAT_BG)

    avy = chat_top+22
    draw.ellipse((26, avy, 26+32, avy+32), fill=accent)
    draw.text((26+9, avy+6), brand_letter, font=font(14, True), fill="#FFFFFF")

    body_lines = textwrap.wrap(body, width=15)
    bh = 66 + len(body_lines)*23
    bubble_box = (66, chat_top+18, W-46, chat_top+18+bh)
    rr(draw, bubble_box, 4, fill="#FFFFFF")
    draw.polygon([(bubble_box[0], chat_top+30), (bubble_box[0]-8, chat_top+34), (bubble_box[0], chat_top+40)], fill="#FFFFFF")

    draw.text((84, chat_top+32), title, font=font(14, True), fill="#1a1a1a")
    yy = chat_top+60
    for line in body_lines:
        draw.text((84, yy), line, font=font(12), fill="#444444")
        yy += 22

    draw.text((bubble_box[2]+6, bubble_box[3]-28), "18:02", font=font(9), fill="#AAAAAA")
    draw.text((bubble_box[2]+6, bubble_box[3]-14), "既読", font=font(9), fill="#AAAAAA")

    # rich menu
    rm_top = chat_bottom
    draw.rectangle((12, rm_top, W-12, rm_top+richmenu_h), fill="#FAFAFA")
    draw.line((12, rm_top, W-12, rm_top), fill="#E5E5E5", width=1)
    icons = [(icon_home, "サイトTOP"), (icon_gift, "クーポン"), (icon_fn3, "お問合せ")]
    cell_w = (W-24)/3
    for i,(icon_fn,label) in enumerate(icons):
        cx0 = 12 + i*cell_w
        if i>0:
            draw.line((cx0, rm_top+8, cx0, rm_top+richmenu_h-8), fill="#E5E5E5", width=1)
        icon_fn(draw, cx0+cell_w/2, rm_top+42, 16, accent)
        bbox2 = draw.textbbox((0,0), label, font=font(12))
        lw = bbox2[2]-bbox2[0]
        draw.text((cx0+cell_w/2-lw/2, rm_top+80), label, font=font(12), fill="#333333")

    ib_top = rm_top+richmenu_h
    draw.rectangle((12, ib_top, W-12, phone_box[3]), fill="#FFFFFF")
    draw.line((12, ib_top, W-12, ib_top), fill="#E5E5E5", width=1)
    draw.ellipse((26, ib_top+12, 26+28, ib_top+40), outline="#999999", width=2)
    draw.line((26+14, ib_top+18, 26+14, ib_top+34), fill="#999999", width=2)
    draw.line((26+7, ib_top+26, 26+21, ib_top+26), fill="#999999", width=2)
    rr(draw, (66, ib_top+10, W-70, ib_top+42), 16, outline="#DADDE1", width=2)
    draw.text((78, ib_top+18), "メッセージを入力", font=font(12), fill="#BBBBBB")
    draw.ellipse((W-58, ib_top+10, W-58+32, ib_top+42), outline="#999999", width=2)

    draw.text((10, H-42), caption, font=font(15, True), fill="#1E2761")
    draw.text((10, H-20), "※イメージ図（架空のサービス名です）", font=font(11), fill="#999999")

    img.save(filename)

make_line_notif("01_delivery.png", "#F2994A", "配", "ハコブーン",
    "お荷物お届けのお知らせ",
    "ご注文の商品を本日18時〜20時にお届け予定です。不在の場合は再配達手続きをご案内します。",
    "ハコブーン（仮）でのお荷物お届けのお知らせ", icon_phone)

make_line_notif("02_utility.png", "#2F80ED", "電", "でんきねっと",
    "今月の料金確定のお知らせ",
    "今月のご利用料金が確定しました。詳しい内訳はマイページよりご確認いただけます。",
    "でんきねっと（仮）での料金確定のお知らせ", icon_phone)

make_line_notif("03_realestate.png", "#27AE60", "住", "いえらいふ",
    "資料請求完了のお知らせ",
    "お問い合わせいただいた物件の資料をご用意しました。担当者よりご連絡いたします。",
    "いえらいふ（仮）での資料請求完了時のお知らせ", icon_phone)

make_line_notif("04_retail.png", "#9B51E0", "眼", "メガネのミライ",
    "ご購入完了のお知らせ",
    "ご購入ありがとうございます。商品の発送準備が整い次第、追跡番号をお送りします。",
    "メガネのミライ（仮）での購入完了のお知らせ", icon_phone)

print("done")
