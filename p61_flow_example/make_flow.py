# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import textwrap

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 360, 480
LINE_GREEN = "#06C755"

def frame_base():
    img = Image.new("RGB", (W, H), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((6, 6, W-6, H-6), radius=16, outline="#DADFE3", width=2)
    return img, draw

# ---- Frame 1: サイトでの注文完了画面 ----
img1, d1 = frame_base()
# browser chrome
d1.rectangle((6, 6, W-6, 44), fill="#F3F3F3")
d1.ellipse((22, 20, 32, 30), fill="#FF5F57")
d1.ellipse((36, 20, 46, 30), fill="#FEBC2E")
d1.ellipse((50, 20, 60, 30), fill="#28C840")
d1.rounded_rectangle((80, 16, W-26, 34), radius=8, outline="#DADFE3", width=1)
d1.text((90, 20), "hakobin-shop.example.com", font=font(10), fill="#888888")

cy = 90
d1.ellipse((W/2-28, cy, W/2+28, cy+56), fill="#EAFBF1")
d1.line((W/2-14, cy+28, W/2-4, cy+40), fill=LINE_GREEN, width=5)
d1.line((W/2-4, cy+40, W/2+16, cy+14), fill=LINE_GREEN, width=5)

d1.text((W/2-70, cy+72), "ご注文ありがとうございます", font=font(15, True), fill="#222222")
d1.text((W/2-96, cy+100), "注文番号：No.20260702-114", font=font(12), fill="#666666")

note_box = (30, cy+140, W-30, cy+200)
d1.rounded_rectangle(note_box, radius=10, fill="#F1F1F1")
d1.text((44, cy+152), "LINEでお届け状況を", font=font(12), fill="#333333")
d1.text((44, cy+174), "リアルタイムにお知らせします", font=font(12), fill="#333333")

d1.text((10, H-28), "① サイトでご注文完了（LINE友だち未登録）", font=font(11, True), fill="#1E2761")
img1.save("flow1_order.png")

# ---- Frame 2: LINE通知メッセージが届く ----
img2, d2 = frame_base()
y = 110
d2.text((26, y), "ハコブーン●LINE公式アカウント", font=font(11), fill=LINE_GREEN)
y += 22
body_lines = ["ご注文の商品を本日18時〜20時に", "お届け予定です。不在の場合は", "再配達手続きをご案内します。"]
h = 34 + 24*len(body_lines)
bubble = (26, y, W-26, y+h)
d2.rounded_rectangle(bubble, radius=4, fill="#F1F1F1")
d2.text((42, y+14), "お荷物お届けのお知らせ", font=font(14, True), fill="#1a1a1a")
yy = y+40
for line in body_lines:
    d2.text((42, yy), line, font=font(12), fill="#333333")
    yy += 24
y += h + 8
d2.text((26, y), "18:02", font=font(10), fill="#AAAAAA")

badge = (W-70, 50, W-30, 78)
d2.rounded_rectangle(badge, radius=14, fill="#FF3B30")
d2.text((W-62, 56), "NEW", font=font(11, True), fill="#FFFFFF")
d2.text((10, H-28), "② 友だち未登録でも通知メッセージが届く", font=font(11, True), fill="#1E2761")
img2.save("flow2_notify.png")

# ---- Frame 3: タップしてトークが開き友だちに ----
img3, d3 = frame_base()
y = 40
d3.text((26, y), "ハコブーン●LINE公式アカウント", font=font(11), fill=LINE_GREEN)
y += 22
bubble = (26, y, W-26, y+70)
d3.rounded_rectangle(bubble, radius=4, fill="#F1F1F1")
d3.text((42, y+14), "友だち追加ありがとうございます！", font=font(13, True), fill="#1a1a1a")
d3.text((42, y+40), "お届け状況はこちらから確認できます。", font=font(11), fill="#333333")
y += 70+8
d3.text((26, y), "18:02", font=font(10), fill="#AAAAAA")
y += 40

reply = "ありがとうございます！"
bbox = d3.textbbox((0,0), reply, font=font(12))
tw = bbox[2]-bbox[0]
bubble2 = (W-26-tw-28, y, W-26, y+44)
d3.rounded_rectangle(bubble2, radius=4, fill=LINE_GREEN)
d3.text((bubble2[0]+14, y+14), reply, font=font(12), fill="#FFFFFF")

# friend badge
fb = (30, H-110, W-30, H-70)
d3.rounded_rectangle(fb, radius=10, fill="#EAFBF1")
d3.text((44, H-100), "「友だち」に自動追加されました", font=font(12, True), fill="#128A45")

d3.text((10, H-28), "③ タップでトークが開き自然に友だち化", font=font(11, True), fill="#1E2761")
img3.save("flow3_friend.png")

print("done")
