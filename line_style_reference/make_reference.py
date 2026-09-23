# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import textwrap

FONT_DIR = r"C:\Windows\Fonts"
def font(size, bold=False):
    return ImageFont.truetype(FONT_DIR + (r"\YuGothB.ttc" if bold else r"\YuGothR.ttc"), size)

W, H = 420, 480
LINE_GREEN = "#06C755"

img = Image.new("RGB", (W, H), "#FFFFFF")
draw = ImageDraw.Draw(img)
draw.rounded_rectangle((8, 8, W-8, H-8), radius=16, outline="#DADFE3", width=2)

y = 30

# --- OA message 1: sender label + white bubble, left aligned ---
draw.text((28, y), "サンプル公式●LINE公式アカウント", font=font(11), fill=LINE_GREEN)
y += 20
body1 = "友だち追加ありがとうございます！"
bubble1 = (28, y, 300, y+56)
draw.rounded_rectangle(bubble1, radius=4, fill="#F1F1F1")
draw.text((44, y+18), body1, font=font(13), fill="#1a1a1a")
y += 56 + 10
draw.text((28, y), "18:02", font=font(10), fill="#AAAAAA")
y += 30

# --- OA message 2 (no repeated name label, consecutive from same sender) ---
body2_lines = ["最新情報をご案内します。", "気になるところはこちらから", "チェック！"]
h2 = 26*len(body2_lines) + 20
bubble2 = (28, y, 300, y+h2)
draw.rounded_rectangle(bubble2, radius=4, fill="#F1F1F1")
yy = y+16
for line in body2_lines:
    draw.text((44, yy), line, font=font(13), fill="#1a1a1a")
    yy += 26
y += h2 + 10
draw.text((28, y), "18:02", font=font(10), fill="#AAAAAA")
y += 40

# --- User reply: green bubble, right aligned ---
reply = "料金はいくらくらいになりますか？"
bbox = draw.textbbox((0,0), reply, font=font(13))
tw = bbox[2]-bbox[0]
bubble3 = (W-32-tw-32, y, W-32, y+50)
draw.rounded_rectangle(bubble3, radius=4, fill=LINE_GREEN)
draw.text((bubble3[0]+16, y+16), reply, font=font(13), fill="#FFFFFF")
draw.text((bubble3[0]-46, y+34), "既読\n18:03", font=font(9), fill="#AAAAAA")

# caption
draw.text((10, H-30), "LINEトーク画面の正しいスタイル（本資料 p43・p51準拠）", font=font(12, True), fill="#1E2761")

img.save("line_style_reference.png")
print("done")
