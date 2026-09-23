# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import textwrap

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 420, 660
RADIUS = 28
LINE_GREEN = "#06C755"

def make_line_notif(filename, brand_letter, brand_name, title, body, caption):
    img = Image.new("RGB", (W, H), "#F2F2F5")
    draw = ImageDraw.Draw(img)
    phone_box = (10, 10, W-10, H-56)
    draw.rounded_rectangle(phone_box, radius=RADIUS, fill="#FFFFFF")
    draw.rounded_rectangle(phone_box, radius=RADIUS, outline="#DADFE3", width=2)

    # status bar
    draw.text((36, 26), "9:41", font=font(16, True), fill="#111111")
    draw.text((W-92, 26), "●●● 100%", font=font(12), fill="#555555")

    # LINE green header (rounded top corners only)
    header_top = 52
    header_h = 56
    draw.rectangle((12, header_top, W-12, header_top+header_h), fill=LINE_GREEN)
    # round the very top corners to match card
    draw.pieslice((12, header_top, 12+2*RADIUS, header_top+2*RADIUS), 180, 270, fill=LINE_GREEN)
    draw.pieslice((W-12-2*RADIUS, header_top, W-12, header_top+2*RADIUS), 270, 360, fill=LINE_GREEN)
    draw.rectangle((12, header_top, 12+RADIUS, header_top+RADIUS), fill=LINE_GREEN)
    draw.rectangle((W-12-RADIUS, header_top, W-12, header_top+RADIUS), fill=LINE_GREEN)

    draw.text((32, header_top+header_h/2-12), "‹", font=font(26, True), fill="#FFFFFF")
    draw.ellipse((64, header_top+8, 64+40, header_top+48), fill="#FFFFFF")
    draw.text((64+13, header_top+16), brand_letter, font=font(18, True), fill=LINE_GREEN)
    draw.text((116, header_top+header_h/2-10), brand_name, font=font(15, True), fill="#FFFFFF")

    # chat background
    chat_top = header_top+header_h
    draw.rectangle((12, chat_top, W-12, phone_box[3]), fill="#E9E9E9")

    # avatar
    avy = chat_top+26
    draw.ellipse((30, avy, 30+34, avy+34), fill=LINE_GREEN)
    draw.text((30+10, avy+6), brand_letter, font=font(15, True), fill="#FFFFFF")

    # bubble
    body_lines = textwrap.wrap(body, width=16)
    bh = 70 + len(body_lines)*24
    bubble_box = (74, chat_top+18, W-36, chat_top+18+bh)
    draw.rounded_rectangle(bubble_box, radius=14, fill="#FFFFFF")
    # tail
    draw.polygon([(74, chat_top+40), (64, chat_top+46), (74, chat_top+52)], fill="#FFFFFF")

    draw.text((92, chat_top+34), title, font=font(15, True), fill="#1a1a1a")
    yy = chat_top+64
    for line in body_lines:
        draw.text((92, yy), line, font=font(13), fill="#444444")
        yy += 24

    draw.text((78, bubble_box[3]+8), "18:02", font=font(11), fill="#999999")

    # caption below phone card
    draw.text((10, H-40), caption, font=font(15, True), fill="#1E2761")
    draw.text((10, H-18), "※イメージ図（架空のサービス名です）", font=font(11), fill="#999999")

    img.save(filename)

make_line_notif(
    "01_delivery.png", "配", "ハコブーン",
    "お荷物お届けのお知らせ",
    "ご注文の商品を本日18時〜20時にお届け予定です。不在の場合は再配達手続きをご案内します。",
    "ハコブーン（仮）でのお荷物お届けのお知らせ"
)
make_line_notif(
    "02_utility.png", "電", "でんきねっと",
    "今月の料金確定のお知らせ",
    "今月のご利用料金が確定しました。詳しい内訳はマイページよりご確認いただけます。",
    "でんきねっと（仮）での料金確定のお知らせ"
)
make_line_notif(
    "03_realestate.png", "住", "いえらいふ",
    "資料請求完了のお知らせ",
    "お問い合わせいただいた物件の資料をご用意しました。担当者よりご連絡いたします。",
    "いえらいふ（仮）での資料請求完了時のお知らせ"
)
make_line_notif(
    "04_retail.png", "眼", "メガネのミライ",
    "ご購入完了のお知らせ",
    "ご購入ありがとうございます。商品の発送準備が整い次第、追跡番号をお送りします。",
    "メガネのミライ（仮）での購入完了のお知らせ"
)
print("done")
