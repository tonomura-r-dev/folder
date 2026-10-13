"""ノースト 10月2本目（仮に10/16）のバナー（1040×1040）。ラフAの形のまま、言葉だけ決定版にした。

心の声「何とかならないかなぁ…」→ 主役「その“何とか”、／医師が一緒に／考えます。」（2026-10-09 決定）。
写真は使わない（2026-10-13 殿村さん「バナー画像は添付のようなもの」＝ラフA・Cの形）。
下半分（札4つ・当日の流れ・予約ボタン・注記）はラフAと同じ。言葉はテキスト（499字）と同じ。

使い方:
  python3 _build/build_norst_1016_banner.py [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN, GREEN_DARK, TEAL, ORANGE, PALE, INK, WHITE,
    font, tw, th, text, text_c, rrect, ellipse, jisseki_badge, white_logo,
    icon_door, icon_person, icon_pin, icon_bubble,
)
from build_norst_rough_patterns import cta, footer  # noqa: E402


def icon_calendar_check(d, cx, cy):
    """カレンダーにチェック（その日に決めなくても大丈夫）。"""
    rrect(d, (cx - 26, cy - 22, cx + 26, cy + 26), 5, outline=WHITE, width=4)
    d.rectangle(((cx - 26) * S, (cy - 22) * S, (cx + 26) * S, (cy - 10) * S), fill=WHITE)
    for dx in (-13, 13):
        d.line(((cx + dx) * S, (cy - 30) * S, (cx + dx) * S, (cy - 16) * S), fill=WHITE, width=5 * S)
    d.line([((cx - 12) * S, (cy + 7) * S), ((cx - 3) * S, (cy + 16) * S), ((cx + 14) * S, (cy - 2) * S)],
           fill=WHITE, width=5 * S, joint="curve")

PANEL_BOTTOM = 516


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面・ロゴ・実績の円
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)
    logo = white_logo(230 * S)
    c.paste(logo, (36 * S, 30 * S), logo)
    jisseki_badge(c, d, 905, 128, 88)

    # 心の声（吹き出し）
    f_v = font(34)
    v = "「何とかならないかなぁ…」"
    vw = tw(f_v, v) + 44
    rrect(d, (40, 104, 40 + vw, 160), 28, fill=WHITE)
    d.polygon([(70 * S, 158 * S), (98 * S, 158 * S), (66 * S, 180 * S)], fill=WHITE)
    text(d, v, f_v, 62, 132 - th(f_v, v) / 2, GREEN_DARK)

    # 何の話か（吹き出しの右に小さなタグ）
    f_t = font(28)
    tx1 = 40 + vw + 22
    rrect(d, (tx1, 112, tx1 + tw(f_t, "包茎治療") + 40, 152), 20, fill=GREEN_DARK, outline=WHITE, width=2)
    text(d, "包茎治療", f_t, tx1 + 20, 132 - th(f_t, "包茎治療") / 2, WHITE)

    # 主役3行
    f_1 = font(66)
    text_c(d, "その“何とか”、", f_1, W / 2, 190, WHITE, 1, stroke=1.2, stroke_fill=WHITE)

    f_2 = font(112)
    a, b = "医師が", "一緒に"
    w2 = tw(f_2, a + b, 2) + 12
    x = W / 2 - w2 / 2
    for s, col in ((a, WHITE), (b, ORANGE)):
        text(d, s, f_2, x + 4, 266 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text(d, s, f_2, x, 266, col, 2, stroke=8 if col == ORANGE else 1.5,
             stroke_fill=WHITE)
        x += tw(f_2, s, 2) + 12
    s3 = "考えます。"
    text_c(d, s3, f_2, W / 2 + 4, 390 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)
    text_c(d, s3, f_2, W / 2, 390, WHITE, 2, stroke=1.5, stroke_fill=WHITE)

    # 札4つ（2026-10-13 殿村さん「完全個室は省いたほうが良い」→ その日に決めなくても大丈夫 に差し替え）
    cards = [
        (icon_person, ["男性スタッフ", "のみ"], [None, ORANGE]),
        (icon_pin, ["全国に", "36院"], [None, ORANGE]),
        (icon_bubble, ["カウンセリング", "無料"], [None, ORANGE]),
        (icon_calendar_check, ["その日に決めなくても", "大丈夫"], ["small", ORANGE]),
    ]
    cw, gap, top, bot = 226, 12, 526, 694
    x0 = (W - (cw * 4 + gap * 3)) / 2
    f_c, f_cb, f_cs = font(27), font(34), font(21)
    for i, (icon, lines, colors) in enumerate(cards):
        x1 = x0 + i * (cw + gap)
        rrect(d, (x1, top, x1 + cw, bot), 14, fill=WHITE, outline=TEAL, width=3)
        icx, icy = x1 + cw / 2, top + 46
        ellipse(d, (icx - 37, icy - 37, icx + 37, icy + 37), fill=GREEN)
        icon(d, icx, icy)
        y = top + 92
        for s_, col in zip(lines, colors):
            f = f_cs if col == "small" else (f_cb if col else f_c)
            text_c(d, s_, f, icx, y + (6 if col == "small" else 0), INK if col in (None, "small") else col)
            y += th(f_c, s_) + 7

    # 当日の流れ（②の「完全個室で」を外した）
    fy1, fy2 = 706, 764
    rrect(d, (40, fy1, 1000, fy2), 30, fill=WHITE, outline=GREEN, width=2)
    f_l = f_s = font(25)
    rrect(d, (48, fy1 + 7, 200, fy2 - 7), 24, fill=GREEN)
    text_c(d, "当日の流れ", f_l, 124, (fy1 + fy2) / 2 - th(f_l, "当日の流れ") / 2, WHITE)
    steps = ["①ご予約", "②スタッフと1対1", "③治療法やご不安をご相談"]
    total = sum(tw(f_s, s_) for s_ in steps) + 30 * 2 + 20 * 2
    x = 200 + (1000 - 200 - total) / 2
    for i, s_ in enumerate(steps):
        text(d, s_, f_s, x, (fy1 + fy2) / 2 - th(f_s, s_) / 2, GREEN_DARK)
        x += tw(f_s, s_) + 20
        if i < 2:
            ay = (fy1 + fy2) / 2
            d.polygon([(x * S, (ay - 10) * S), ((x + 14) * S, ay * S), (x * S, (ay + 10) * S)], fill=ORANGE)
            x += 30

    # 時期の一言（予約ボタンの上）
    text_c(d, "＼年内に、とお考えの方へ／", font(30), W / 2, 774, GREEN_DARK, 1)

    cta(c, d, 812, 952)
    footer(d)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_1016_banner_nantoka_v2_1040.jpg"))
