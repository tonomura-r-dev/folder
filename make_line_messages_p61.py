# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import random

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 280, 500
LINE_GREEN = "#06C755"

def phone_base(title_brand, icon_color):
    """基本のLINEスクリーン"""
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((4, 4, W-4, H-4), radius=14, outline="#DADFE3", width=2)

    # ステータスバー
    d.rectangle((6, 6, W-6, 36), fill="#FFFFFF")
    d.text((14, 12), "9:41", font=font(11, True), fill="#111111")
    d.text((W-50, 13), "●●● 100%", font=font(9), fill="#555555")

    # LINEヘッダー
    d.rectangle((6, 36, W-6, 80), fill="#FFFFFF")
    d.line((6, 80, W-6, 80), fill="#E5E5E5", width=1)
    d.text((14, 48), "<", font=font(20, True), fill="#333333")
    d.ellipse((32, 42, 62, 72), fill=icon_color)
    d.text((80, 50), title_brand, font=font(13, True), fill="#222222")

    # トークエリア背景
    for i in range(H-86-56):
        t = i / (H-86-56)
        col = (int(222 + (202-222)*t), int(234 + (220-234)*t), int(246 + (242-246)*t))
        d.line((6, 80+i, W-6, 80+i), fill=col)

    # 入力エリア
    d.rectangle((6, H-56, W-6, H-6), fill="#FFFFFF")
    d.line((6, H-56, W-6, H-56), fill="#E5E5E5", width=1)

    return img, d

def add_oa_label(d, y, brand):
    """OA公式アカウントラベル"""
    label_text = f"{brand}●LINE公式アカウント"
    d.text((14, y), label_text, font=font(8), fill=LINE_GREEN)

def add_message_bubble(d, y, text, is_user=False):
    """テキストメッセージ吹き出し"""
    lines = text.split('\n')
    h = len(lines) * 18 + 12
    if is_user:
        box = (W-150, y, W-14, y+h)
        d.rounded_rectangle(box, radius=6, fill="#93E15D")
        text_x = W - 140
    else:
        box = (14, y, 200, y+h)
        d.rounded_rectangle(box, radius=6, fill="#FFFFFF", outline="#E0E0E0", width=1)
        text_x = 24

    yy = y + 8
    for line in lines:
        d.text((text_x, yy), line, font=font(9), fill="#333333" if not is_user else "#1a1a1a")
        yy += 18
    return y + h + 12

def add_button_area(d, y, label):
    """ボタンエリア"""
    btn = (20, y, W-20, y+36)
    d.rounded_rectangle(btn, radius=8, fill="#FF7A33")
    bbox = d.textbbox((0,0), label, font=font(11, True))
    tw = bbox[2] - bbox[0]
    d.text((W/2 - tw/2, y+10), label, font=font(11, True), fill="#FFFFFF")
    return y + 46

def add_rich_card(d, y, title, content_lines, btn_label, accent_color):
    """リッチメッセージカード"""
    card_h = 50 + len(content_lines) * 16
    card = (14, y, W-14, y+card_h)
    d.rounded_rectangle(card, radius=8, fill="#FFFFFF", outline=accent_color, width=2)
    d.text((24, y+10), title, font=font(11, True), fill=accent_color)
    yy = y + 30
    for line in content_lines:
        d.text((24, yy), line, font=font(9), fill="#666666")
        yy += 16
    return y + card_h + 12

# ===== ハコブーン（配送）=====
def hakobin_msg1():
    img, d = phone_base("ハコブーン", "#FF7A33")
    y = 90
    add_oa_label(d, y, "ハコブーン")
    y += 20
    y = add_message_bubble(d, y, "お荷物お届け予定のお知らせ\n本日18:00～20:00予定です")
    y = add_button_area(d, y, "配達予定を変更")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_hakobin_1.png")
    print("hakobin_1")

def hakobin_msg2():
    img, d = phone_base("ハコブーン", "#FF7A33")
    y = 90
    add_oa_label(d, y, "ハコブーン")
    y += 20
    y = add_message_bubble(d, y, "配送状況")
    # 進捗バー
    steps = ["集荷", "輸送中", "配達中", "完了"]
    seg_w = (W-40) / 4
    for i, s in enumerate(steps):
        x0 = 20 + i * seg_w
        fill = "#FF7A33" if i <= 2 else "#E5E5E5"
        d.rounded_rectangle((x0, y, x0+seg_w-4, y+6), radius=3, fill=fill)
        d.text((x0, y+12), s, font=font(8), fill="#888888")
    y += 36
    y = add_button_area(d, y, "詳細を見る")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_hakobin_2.png")
    print("hakobin_2")

def hakobin_msg3():
    img, d = phone_base("ハコブーン", "#FF7A33")
    y = 90
    add_oa_label(d, y, "ハコブーン")
    y += 20
    y = add_message_bubble(d, y, "配達日時を選択してください")
    y += 8
    days = ["今日", "明日", "明後日"]
    for day in days:
        d.rounded_rectangle((20, y, W-20, y+28), radius=6, fill="#FFE9DA")
        bbox = d.textbbox((0,0), day, font=font(10, True))
        tw = bbox[2] - bbox[0]
        d.text((W/2 - tw/2, y+8), day, font=font(10, True), fill="#FF7A33")
        y += 34
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_hakobin_3.png")
    print("hakobin_3")

# ===== でんきねっと（電力）=====
def denki_msg1():
    img, d = phone_base("でんきねっと", "#2F80ED")
    y = 90
    add_oa_label(d, y, "でんきねっと")
    y += 20
    y = add_message_bubble(d, y, "今月の料金が確定しました\n5,280円（目安）")
    y = add_button_area(d, y, "詳細を確認")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_denki_1.png")
    print("denki_1")

def denki_msg2():
    img, d = phone_base("でんきねっと", "#2F80ED")
    y = 90
    add_oa_label(d, y, "でんきねっと")
    y += 20
    y = add_message_bubble(d, y, "ご利用状況")
    y = add_rich_card(d, y, "8月度ご使用量", ["電気使用量：450kWh", "前月比：+5%"], "グラフを見る", "#2F80ED")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_denki_2.png")
    print("denki_2")

def denki_msg3():
    img, d = phone_base("でんきねっと", "#2F80ED")
    y = 90
    add_oa_label(d, y, "でんきねっと")
    y += 20
    y = add_message_bubble(d, y, "9月の料金予測は\n約5,600円です")
    y = add_button_area(d, y, "節電のコツを見る")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_denki_3.png")
    print("denki_3")

# ===== いえらいふ（不動産）=====
def ielife_msg1():
    img, d = phone_base("いえらいふ", "#2FAE66")
    y = 90
    add_oa_label(d, y, "いえらいふ")
    y += 20
    y = add_message_bubble(d, y, "資料請求ありがとうございます\n担当者よりご連絡します")
    y = add_button_area(d, y, "他の物件を見る")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_ielife_1.png")
    print("ielife_1")

def ielife_msg2():
    img, d = phone_base("いえらいふ", "#2FAE66")
    y = 90
    add_oa_label(d, y, "いえらいふ")
    y += 20
    y = add_message_bubble(d, y, "新着物件")
    y = add_rich_card(d, y, "渋谷区 高級マンション", ["3LDK・120㎡", "¥95,000,000"], "内覧予約", "#2FAE66")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_ielife_2.png")
    print("ielife_2")

def ielife_msg3():
    img, d = phone_base("いえらいふ", "#2FAE66")
    y = 90
    add_oa_label(d, y, "いえらいふ")
    y += 20
    y = add_message_bubble(d, y, "内覧日時を選択してください")
    y += 8
    times = ["10:00", "14:00", "16:00"]
    for time in times:
        d.rounded_rectangle((20, y, W-20, y+28), radius=6, fill="#EAFBF1")
        d.text((W/2-20, y+8), time, font=font(10, True), fill="#2FAE66")
        y += 34
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_ielife_3.png")
    print("ielife_3")

# ===== メガネのミライ（眼鏡小売）=====
def megane_msg1():
    img, d = phone_base("メガネのミライ", "#8A3FE4")
    y = 90
    add_oa_label(d, y, "メガネのミライ")
    y += 20
    y = add_message_bubble(d, y, "ご購入ありがとうございます\nフレーム TYPE-03")
    y = add_button_area(d, y, "発送状況を確認")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_megane_1.png")
    print("megane_1")

def megane_msg2():
    img, d = phone_base("メガネのミライ", "#8A3FE4")
    y = 90
    add_oa_label(d, y, "メガネのミライ")
    y += 20
    y = add_message_bubble(d, y, "発送準備中")
    # ステータス
    statuses = ["検品中", "梱包中", "配送待ち"]
    for status in statuses:
        d.rounded_rectangle((20, y, W-20, y+24), radius=6, fill="#F5EEFF")
        d.text((W/2-30, y+6), status, font=font(9, True), fill="#8A3FE4")
        y += 30
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_megane_2.png")
    print("megane_2")

def megane_msg3():
    img, d = phone_base("メガネのミライ", "#8A3FE4")
    y = 90
    add_oa_label(d, y, "メガネのミライ")
    y += 20
    y = add_message_bubble(d, y, "本日発送されました！\n配送予定日：明日")
    y = add_button_area(d, y, "追跡番号を確認")
    img.save(r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\p61_megane_3.png")
    print("megane_3")

# 全て生成
hakobin_msg1()
hakobin_msg2()
hakobin_msg3()
denki_msg1()
denki_msg2()
denki_msg3()
ielife_msg1()
ielife_msg2()
ielife_msg3()
megane_msg1()
megane_msg2()
megane_msg3()

print("\n✅ 12枚完成！")
