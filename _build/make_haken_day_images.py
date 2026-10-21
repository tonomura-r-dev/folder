# 人材：実文面の画面イメージの「Day3→Day6」「Day7→Day14」「※ Day3→Day6は求人を送りません」を書き換えた画像の作り方（2026-10-21）。
# 元画像（day_s24.png / day_s25.png）と Noto Sans JP（fontsource の woff2 を ttf にしたもの）を fonts/ に置いた作業フォルダで実行した。出力は _images/haken_day_s24.png / haken_day_s25.png。
import sys
from PIL import Image, ImageFont, ImageDraw
F="fonts/noto-sans-jp-%s-%d-normal.ttf"
def font(script,w,s): return ImageFont.truetype(F%(script,w),s)

def fit_size(text_ref, width, cap, w=500):
    best=None
    for s in range(20,40):
        f=font("latin",w,s); bb=f.getbbox(text_ref,anchor="ls"); c=-f.getbbox("D",anchor="ls")[1]
        score=abs((bb[2]-bb[0])-width)+abs(c-cap)
        if best is None or score<best[0]: best=(score,s)
    return best[1]

def rounded_pill(img, box, fill, scale=4):
    x0,y0,x1,y1=box
    w,h=x1-x0,y1-y0
    big=Image.new("L",(w*scale,h*scale),0)
    ImageDraw.Draw(big).rounded_rectangle((0,0,w*scale-1,h*scale-1),radius=h*scale//2,fill=255)
    m=big.resize((w,h),Image.LANCZOS)
    img.paste(Image.new("RGB",(w,h),fill),(x0,y0),m)

# --- S24: Day3 -> Day6 (pill), note
a=Image.open("day_s24.png").convert("RGB")
d=ImageDraw.Draw(a)
PILL=(117,145,180); WHITE=(248,250,251); BG24=(140,177,221)
s=fit_size("Day3",74,24)
d.rectangle((447,923,578,966),fill=PILL)
f=font("latin",500,s); t="Day6"; bb=f.getbbox(t,anchor="ls")
d.text((511.5-(bb[2]+bb[0])/2,956),t,font=f,fill=WHITE,anchor="ls")
# note: repaint and redraw (calibrated to the original "※ Day3は求人を送りません" width 344px, ※ top 1469)
d.rectangle((40,1460,400,1502),fill=BG24)
NOTE=(81,100,124)
def segs(day): return [("※","fonts/kome-500.ttf"),(" Day%s"%day,"fonts/noto-sans-jp-latin-500-normal.ttf"),("は求人を送りません","fonts/noto-sans-jp-japanese-500-normal.ttf")]
def width(day,s): return sum(ImageFont.truetype(p,s).getlength(t) for t,p in segs(day))
ns=min(range(18,34),key=lambda s: abs(width("3",s)-344))
kf=ImageFont.truetype("fonts/kome-500.ttf",ns); top=kf.getbbox("※",anchor="ls")[1]
base=1469-top
x=47
for t,p in segs("6"):
    ff=ImageFont.truetype(p,ns); d.text((x,base),t,font=ff,fill=NOTE,anchor="ls"); x+=ff.getlength(t)
print("note size",ns,"base",base)
a.save("day_s24_edit.png")

# --- S25: Day7 -> Day14 (wider pill)
b=Image.open("day_s25.png").convert("RGB")
d=ImageDraw.Draw(b)
LPILL=(200,208,224); DARK=(49,50,59); BG25=(137,171,216)
s7=fit_size("Day7",66,21)
f=font("latin",500,s7); t="Day14"; bb=f.getbbox(t,anchor="ls")
tw=bb[2]-bb[0]
old=(441,99,584,156)
extra=tw-66
new=(int(512.5-(143+extra)/2), old[1], int(512.5+(143+extra)/2), old[3])
d.rectangle((old[0]-4,old[1]-4,old[2]+4,old[3]+4),fill=BG25)
rounded_pill(b,new,LPILL)
d=ImageDraw.Draw(b)
d.text((512.5-(bb[2]+bb[0])/2,139),t,font=f,fill=DARK,anchor="ls")
b.save("day_s25_edit.png")
print("sizes",s,s7,"new pill",new)
