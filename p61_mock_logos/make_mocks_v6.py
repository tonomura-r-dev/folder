# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import textwrap

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 420, 320
LINE_GREEN = "#06C755"

def make_card(filename, brand_name, title, body, caption):
    img = Image.new("RGB", (W, H), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((8, 8, W-8, H-8), radius=16, outline="#DADFE3", width=2)

    y = 28
    draw.text((26, y), f"{brand_name}●LINE公式アカウント", font=font(11), fill=LINE_GREEN)
    y += 22

    body_lines = textwrap.wrap(body, width=17)
    h = 34 + 24*len(body_lines)
    bubble = (26, y, W-26, y+h)
    draw.rounded_rectangle(bubble, radius=4, fill="#F1F1F1")
    draw.text((42, y+14), title, font=font(14, True), fill="#1a1a1a")
    yy = y+40
    for line in body_lines:
        draw.text((42, yy), line, font=font(13), fill="#333333")
        yy += 24
    y += h + 8
    draw.text((26, y), "18:02", font=font(10), fill="#AAAAAA")

    draw.text((10, H-30), caption, font=font(12, True), fill="#1E2761")
    draw.text((10, H-16), "※イメージ図（架空のサービス名です）", font=font(9), fill="#999999")

    img.save(filename)

make_card("01_delivery.png", "ハコブーン",
    "お荷物お届けのお知らせ",
    "ご注文の商品を本日18時〜20時にお届け予定です。不在の場合は再配達手続きをご案内します。",
    "ハコブーン（仮）でのお荷物お届けのお知らせ")

make_card("02_utility.png", "でんきねっと",
    "今月の料金確定のお知らせ",
    "今月のご利用料金が確定しました。詳しい内訳はマイページよりご確認いただけます。",
    "でんきねっと（仮）での料金確定のお知らせ")

make_card("03_realestate.png", "いえらいふ",
    "資料請求完了のお知らせ",
    "お問い合わせいただいた物件の資料をご用意しました。担当者よりご連絡いたします。",
    "いえらいふ（仮）での資料請求完了時のお知らせ")

make_card("04_retail.png", "メガネのミライ",
    "ご購入完了のお知らせ",
    "ご購入ありがとうございます。商品の発送準備が整い次第、追跡番号をお送りします。",
    "メガネのミライ（仮）での購入完了のお知らせ")

print("done")
