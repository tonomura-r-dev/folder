"""ノースト 10月3本目（仮に10/23）のバナー（1040×1040）デモ。写真・イラストなしのシンプルな形。

論理（2026-10-14 殿村さん）：アイテム名「「自分だけなのかな？」その悩み、聞かせてください。」で興味を引く →
画像で「うちにはこれだけの実績がある」を数字で見せる → 一度話（無料カウンセリング）をさせてください。
必ず残す4つ（緑が主・37年目・診療実績40万件・ロゴ）と注記は2本目と同じ。

使い方:
  python3 _build/build_norst_1023_banner.py [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN, GREEN_DARK, TEAL, ORANGE, PALE, INK, GRAY, WHITE,
    font, tw, th, text, text_c, rrect, white_logo, icon_person, icon_pin, ellipse,
)
from build_norst_rough_patterns import cta  # noqa: E402

PANEL_BOTTOM = 560


def stat_card(d, x1, y1, x2, y2, label_top, num, unit, label_bottom=None):
    """白い札に大きな数字。label_top＝数字の上の小さな言葉。"""
    rrect(d, (x1 + 6, y1 + 6, x2 + 6, y2 + 6), 22, fill=GREEN_DARK)  # 影
    rrect(d, (x1, y1, x2, y2), 22, fill=WHITE)
    cx = (x1 + x2) / 2
    f_l, f_n, f_u = font(30), font(112), font(46)
    text_c(d, label_top, f_l, cx, y1 + 22, GREEN_DARK)
    w = tw(f_n, num) + tw(f_u, unit) + 6
    x = cx - w / 2
    base = y1 + 70 + th(f_n, num)
    text(d, num, f_n, x, y1 + 70, ORANGE)
    text(d, unit, f_u, x + tw(f_n, num) + 6, base - th(f_u, unit), GREEN_DARK)
    if label_bottom:
        text_c(d, label_bottom, f_l, cx, y2 - 22 - th(f_l, label_bottom), GREEN_DARK)


def chip(d, x1, y1, icon, s, num=None):
    """淡い面の札（アイコン＋言葉）。num は言葉の中でオレンジにする部分。"""
    f = font(30)
    w = 70 + tw(f, s) + 30
    rrect(d, (x1, y1, x1 + w, y1 + 64), 32, fill=WHITE, outline=TEAL, width=3)
    ellipse(d, (x1 + 10, y1 + 8, x1 + 58, y1 + 56), fill=GREEN)
    icon(d, x1 + 34, y1 + 34)
    x = x1 + 70
    top = y1 + 32 - th(f, s) / 2
    if num and num in s:
        a, b = s.split(num, 1)
        for part, col in ((a, INK), (num, ORANGE), (b, INK)):
            text(d, part, f, x, top, col)
            x += tw(f, part)
    else:
        text(d, s, f, x, top, INK)
    return w


def small_icon(d, cx, cy, icon):
    """アイコンを札の円に収まる大きさで描く（元は直径74の円用）。"""
    icon(d, cx, cy)


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面＋ロゴ（中央。2本目と同じ）
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)
    logo = white_logo(290 * S)
    c.paste(logo, ((W * S - logo.width) // 2, 22 * S), logo)

    # 心の声（アイテム名と同じ言葉）
    f_v = font(38)
    v = "「自分だけなのかな？」"
    vw = tw(f_v, v) + 56
    bx1 = (W - vw) / 2
    rrect(d, (bx1, 104, bx1 + vw, 168), 32, fill=WHITE)
    d.polygon([((W / 2 - 18) * S, 166 * S), ((W / 2 + 18) * S, 166 * S), (W / 2 * S, 190 * S)], fill=WHITE)
    text_c(d, v, f_v, W / 2, 136 - th(f_v, v) / 2, GREEN_DARK)

    # 主役：同じ悩みに、向き合ってきました。
    f_1, f_2 = font(44), font(70)
    text_c(d, "同じ悩みに、", f_1, W / 2, 208, WHITE, 2)
    s2 = "向き合ってきました。"
    text_c(d, s2, f_2, W / 2 + 4, 266 + 4, GREEN_DARK, 2, stroke=8, stroke_fill=GREEN_DARK)  # 影
    text_c(d, s2, f_2, W / 2, 266, WHITE, 2, stroke=1.2, stroke_fill=WHITE)

    # 数字の札2枚：37年目／診療実績40万件
    cy1, cy2 = 362, 530
    stat_card(d, 60, cy1, 500, cy2, "おかげさまで", "37", "年目")  # 「開院から」は確かめていないので使わない
    stat_card(d, 540, cy1, 980, cy2, "診療実績", "40", "万件")

    # 淡い面：全国に36院／男性スタッフのみ
    f_c = font(30)
    w1 = 70 + tw(f_c, "全国に36院") + 30
    w2 = 70 + tw(f_c, "男性スタッフのみ") + 30
    gap = 24
    x = (W - (w1 + gap + w2)) / 2
    chip(d, x, 594, icon_pin_s, "全国に36院", "36院")
    chip(d, x + w1 + gap, 594, icon_person_s, "男性スタッフのみ")

    # 一度、話をさせてください
    text_c(d, "＼一度、お話を聞かせてください／", font(42), W / 2, 694, GREEN_DARK, 1)

    cta(c, d, 760, 894)
    text_c(d, "電話なしで、このままWEB予約できます", font(28), W / 2, 922, GREEN_DARK)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(22), W / 2, 996, GRAY)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


def icon_pin_s(d, cx, cy):
    """ピン（札の小さい円用）。"""
    d.polygon([((cx - 13) * S, (cy - 4) * S), ((cx + 13) * S, (cy - 4) * S), (cx * S, (cy + 19) * S)], fill=WHITE)
    ellipse(d, (cx - 14, cy - 19, cx + 14, cy + 8), fill=WHITE)
    ellipse(d, (cx - 5, cy - 10, cx + 5, cy), fill=GREEN)


def icon_person_s(d, cx, cy):
    """人（札の小さい円用）。"""
    ellipse(d, (cx - 8, cy - 18, cx + 8, cy - 2), fill=WHITE)
    d.pieslice(((cx - 16) * S, (cy + 1) * S, (cx + 16) * S, (cy + 33) * S), 180, 360, fill=WHITE)


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_1023_banner_demo_v1_1040.jpg"))
