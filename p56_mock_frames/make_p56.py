# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import textwrap

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 420, 700
RADIUS = 28
LINE_GREEN = "#06C755"

def base_card():
    img = Image.new("RGB", (W, H), "#F2F2F5")
    draw = ImageDraw.Draw(img)
    phone_box = (10, 10, W-10, H-10)
    draw.rounded_rectangle(phone_box, radius=RADIUS, fill="#FFFFFF")
    draw.rounded_rectangle(phone_box, radius=RADIUS, outline="#DADFE3", width=2)
    # status bar
    draw.text((36, 28), "9:41", font=font(18, True), fill="#111111")
    draw.text((W-98, 28), "●●● 100%", font=font(13), fill="#555555")
    return img, draw

# ---- Frame 1: トークルーム ----
img, draw = base_card()

# LINE chat header
header_h = 60
draw.rounded_rectangle((10, 58, W-10, 58+header_h), radius=0, fill=LINE_GREEN)
draw.rectangle((10, 58+header_h-2, W-10, 58+header_h), fill=LINE_GREEN)
draw.text((36, 58+header_h/2-11), "‹", font=font(28, True), fill="#FFFFFF")
draw.ellipse((66, 58+10, 66+40, 58+50), fill="#FFFFFF")
draw.text((66+11, 58+18), "S", font=font(20, True), fill=LINE_GREEN)
draw.text((118, 58+18), "Sample Provider公式アカウント", font=font(14, True), fill="#FFFFFF")

# chat background
chat_top = 58+header_h+2
draw.rectangle((12, chat_top, W-12, H-12), fill="#EDEDED")

# bot avatar + message bubble
avy = chat_top+30
draw.ellipse((30, avy, 30+36, avy+36), fill=LINE_GREEN)
draw.text((30+11, avy+7), "S", font=font(16, True), fill="#FFFFFF")

bubble_box = (76, chat_top+20, W-40, chat_top+230)
draw.rounded_rectangle(bubble_box, radius=16, fill="#FFFFFF")
draw.text((96, chat_top+40), "友だち追加ありがとうございます！", font=font(14, True), fill="#222222")
body = textwrap.wrap("会員登録がまだの方は、LINEアカウントで簡単に連携できます。今すぐIDを連携して限定クーポンを受け取ろう！", width=17)
yy = chat_top+72
for line in body:
    draw.text((96, yy), line, font=font(13), fill="#444444")
    yy += 22

# CTA button (highlighted like original red box)
btn_box = (96, yy+10, W-60, yy+52)
draw.rounded_rectangle(btn_box, radius=10, fill="#FFF3E0")
draw.rounded_rectangle(btn_box, radius=10, outline="#E85D2F", width=3)
draw.text((btn_box[0]+18, yy+20), "今すぐIDを連携する ›", font=font(14, True), fill="#E85D2F")

draw.text((80, chat_top+240), "18:02", font=font(11), fill="#999999")

img.save("frame1_talkroom.png")

# ---- Frame 3: サービスログイン遷移画面 ----
img2, draw2 = base_card()
cx, cy = W//2, H//2 - 40

# spinner ring (simple dashed circle approximation)
r = 30
draw2.arc((cx-r, cy-r, cx+r, cy+r), start=0, end=270, fill=LINE_GREEN, width=6)
draw2.arc((cx-r, cy-r, cx+r, cy+r), start=270, end=360, fill="#DADFE3", width=6)

# app icon above spinner
icon_r = 34
draw2.rounded_rectangle((cx-icon_r, cy-120, cx+icon_r, cy-120+icon_r*2), radius=16, fill=LINE_GREEN)
draw2.text((cx-9, cy-120+18), "S", font=font(30, True), fill="#FFFFFF")

msg = "Sample Login Appに移動しています"
bbox = draw2.textbbox((0,0), msg, font=font(15, True))
tw = bbox[2]-bbox[0]
draw2.text((cx-tw/2, cy+50), msg, font=font(15, True), fill="#333333")
sub = "しばらくお待ちください…"
bbox2 = draw2.textbbox((0,0), sub, font=font(12))
tw2 = bbox2[2]-bbox2[0]
draw2.text((cx-tw2/2, cy+78), sub, font=font(12), fill="#999999")

img2.save("frame3_transition.png")
print("done")
