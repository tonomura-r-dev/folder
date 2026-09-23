# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = r"C:\Windows\Fonts"
def font(path, size):
    return ImageFont.truetype(path, size)

F_BOLD = FONT_DIR + r"\YuGothB.ttc"
F_MED  = FONT_DIR + r"\YuGothM.ttc"
F_REG  = FONT_DIR + r"\YuGothR.ttc"

W, H = 420, 620
RADIUS = 28

def rounded_rect(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)

def make_card(filename, accent, logo_letter, brand_name, notif_title, notif_body, caption):
    img = Image.new("RGB", (W, H), "#F2F2F5")
    draw = ImageDraw.Draw(img)

    # phone frame
    phone_box = (10, 10, W-10, H-70)
    rounded_rect(draw, phone_box, RADIUS, "#FFFFFF")
    draw.rounded_rectangle(phone_box, radius=RADIUS, outline="#DADFE3", width=2)

    # status bar
    draw.text((36, 30), "9:41", font=font(F_BOLD, 18), fill="#111111")
    draw.text((W-95, 30), "●●● 100%", font=font(F_REG, 14), fill="#555555")

    # app header
    header_top = 70
    header_h = 70
    draw.rounded_rectangle((28, header_top, W-28, header_top+header_h), radius=18, fill="#FAFAFA")
    # logo circle
    lc = 44
    lx, ly = 48, header_top + header_h//2
    draw.ellipse((lx-lc//2, ly-lc//2, lx+lc//2, ly+lc//2), fill=accent)
    bbox = draw.textbbox((0,0), logo_letter, font=font(F_BOLD, 22))
    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
    draw.text((lx-tw/2-bbox[0], ly-th/2-bbox[1]), logo_letter, font=font(F_BOLD, 22), fill="#FFFFFF")
    draw.text((lx+38, ly-14), brand_name, font=font(F_BOLD, 20), fill="#222222")

    # notification bubble
    bubble_top = header_top + header_h + 24
    bubble_box = (44, bubble_top, W-44, bubble_top+150)
    rounded_rect(draw, bubble_box, 20, "#F5F6F8")
    draw.rounded_rectangle(bubble_box, radius=20, outline="#E3E6E9", width=1)
    draw.text((64, bubble_top+20), notif_title, font=font(F_BOLD, 18), fill=accent)
    # wrap body
    import textwrap
    wrapped = textwrap.wrap(notif_body, width=17)
    yy = bubble_top+56
    for line in wrapped[:4]:
        draw.text((64, yy), line, font=font(F_REG, 15), fill="#333333")
        yy += 24

    # timestamp
    draw.text((W-100, bubble_box[3]-26), "18:02", font=font(F_REG, 12), fill="#999999")

    # bottom caption (like the slide caption)
    draw.text((10, H-46), caption, font=font(F_MED, 17), fill="#1E2761")
    draw.text((10, H-22), "※イメージ図（架空のサービス名です）", font=font(F_REG, 12), fill="#999999")

    img.save(filename)

make_card(
    "01_delivery.png", "#F2994A", "配",
    "ハコブーン",
    "お荷物お届けのお知らせ",
    "ご注文の商品を本日18時〜20時にお届け予定です。不在の場合は再配達手続きをご案内します。",
    "ハコブーン（仮）でのお荷物お届けのお知らせ"
)

make_card(
    "02_utility.png", "#2F80ED", "電",
    "でんきねっと",
    "今月の料金確定のお知らせ",
    "今月のご利用料金が確定しました。詳しい内訳はマイページよりご確認いただけます。",
    "でんきねっと（仮）での料金確定のお知らせ"
)

make_card(
    "03_realestate.png", "#27AE60", "住",
    "いえらいふ",
    "資料請求完了のお知らせ",
    "お問い合わせいただいた物件の資料をご用意しました。担当者よりご連絡いたします。",
    "いえらいふ（仮）での資料請求完了時のお知らせ"
)

make_card(
    "04_retail.png", "#9B51E0", "眼",
    "メガネのミライ",
    "ご購入完了のお知らせ",
    "ご購入ありがとうございます。商品の発送準備が整い次第、追跡番号をお送りします。",
    "メガネのミライ（仮）での購入完了のお知らせ"
)

print("done")
