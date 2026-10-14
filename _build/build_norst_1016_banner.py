"""ノースト 10月2本目（仮に10/16）のバナー（1040×1040）。ラフAの形のまま、言葉だけ決定版にした。

心の声「何とかしたいなぁ…」（10/14 「何とかならないかなぁ…」から変更）→ 主役「その“何とか”、／医師が一緒に／考えます。」（2026-10-09 決定）。
写真は使わない（2026-10-13 殿村さん「バナー画像は添付のようなもの」＝ラフA・Cの形）。
下半分（札4つ・当日の流れ・予約ボタン・注記）はラフAと同じ。言葉はテキスト（499字）と同じ。

男性のイラスト入りの版は `_build/build_norst_1016_banner_illust.py`。

使い方:
  python3 _build/build_norst_1016_banner.py [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN, GREEN_DARK, TEAL, ORANGE, PALE, INK, GRAY, WHITE,
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

def icon_chart(d, cx, cy):
    """棒グラフ（37年目・診療実績40万件）。"""
    for i, h in enumerate((16, 28, 40)):
        x = cx - 24 + i * 18
        d.rectangle((x * S, (cy + 20 - h) * S, (x + 12) * S, (cy + 20) * S), fill=WHITE)
    d.line(((cx - 30) * S, (cy + 22) * S, (cx + 30) * S, (cy + 22) * S), fill=WHITE, width=3 * S)


def segs_c(d, segs, cx, base):
    """色と大きさの違う文字を1行に並べて中央に置く。segs=[(文字, 大きさ, 色)]。base＝行の下端。"""
    fs = [(s_, font(sz), col) for s_, sz, col in segs]
    total = sum(tw(f, s_) for s_, f, _ in fs)
    x = cx - total / 2
    for s_, f, col in fs:
        text(d, s_, f, x, base - th(f, s_), col)
        x += tw(f, s_)


def counsel_badge(d, cx, cy, r):
    """右上の円：カウンセリング無料（2026-10-13 殿村さん「上に、カウンセリング無料を入れたい」）。"""
    k = r / 88  # 円の大きさに合わせて文字も縮める
    ellipse(d, (cx - r, cy - r, cx + r, cy + r), fill=WHITE, outline=ORANGE, width=8 * k)
    text_c(d, "カウンセリング", font(20 * k), cx, cy - 34 * k, GREEN_DARK)
    text_c(d, "無料", font(56 * k), cx, cy - 4 * k, ORANGE)


PANEL_BOTTOM = 516

def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面・ロゴ・実績の円
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)
    logo = white_logo(290 * S)  # 2026-10-13 殿村さん「ロゴもう少し大きく」230→290
    c.paste(logo, ((W * S - logo.width) // 2, 22 * S), logo)  # 2026-10-13 殿村さん「ロゴは中央に」
    counsel_badge(d, 905, 128, 88)

    # 心の声（吹き出し）
    f_v = font(34)
    v = "「何とかしたいなぁ…」"  # 2026-10-14 殿村さん「前提、なんとかしたいなぁ のほうが良い」
    vw = tw(f_v, v) + 44
    rrect(d, (40, 104, 40 + vw, 160), 28, fill=WHITE)
    d.polygon([(70 * S, 158 * S), (98 * S, 158 * S), (66 * S, 180 * S)], fill=WHITE)
    text(d, v, f_v, 62, 132 - th(f_v, v) / 2, GREEN_DARK)

    # 診療のタグ3つ（吹き出しの下。2026-10-14 殿村さん「包茎治療以外に、2つくらい治療法も記載したい」。
    # 9/16の1本目で先方が通したのと同じ3つ）
    f_t = font(26)
    tags = ["包茎治療", "亀頭増大", "ED治療"]
    tx = 118
    for t in tags:
        tw_ = tw(f_t, t) + 32
        rrect(d, (tx, 172, tx + tw_, 208), 18, fill=GREEN_DARK, outline=WHITE, width=2)
        text(d, t, f_t, tx + 16, 190 - th(f_t, t) / 2, WHITE)
        tx += tw_ + 10

    # 主役3行（2026-10-14 殿村さん「考えますが中央のラインに被らないように」→ 少し小さくして上から積み、
    # 「考えます。」の影の下端が緑の面の下端（中央のライン）より20px以上上に来るようにした）
    f_1 = font(56)
    s1 = "その“何とか”を、"  # 2026-10-14 殿村さん「“何とか”を、に変えておいて」
    y1 = 220
    text_c(d, s1, f_1, W / 2, y1, WHITE, 1, stroke=1.2, stroke_fill=WHITE)

    f_2 = font(94)
    a, b = "医師が", "一緒に"
    s3 = "考えます。"
    y2 = y1 + th(f_1, s1) + 12
    y3 = y2 + th(f_2, a + b) + 12
    w2 = tw(f_2, a + b, 2) + 12
    x = W / 2 - w2 / 2
    for s, col in ((a, WHITE), (b, ORANGE)):
        text(d, s, f_2, x + 4, y2 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text(d, s, f_2, x, y2, col, 2, stroke=8 if col == ORANGE else 1.5,
             stroke_fill=WHITE)
        x += tw(f_2, s, 2) + 12
    text_c(d, s3, f_2, W / 2 + 4, y3 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)
    text_c(d, s3, f_2, W / 2, y3, WHITE, 2, stroke=1.5, stroke_fill=WHITE)
    shadow_bottom = y3 + th(f_2, s3) + 14  # 影のずれ4＋影のフチ10
    assert shadow_bottom <= PANEL_BOTTOM - 20, shadow_bottom

    # 札4つ（2026-10-13 殿村さん「完全個室は省いたほうが良い」→ その日に決めなくても大丈夫 に差し替え）
    cards = [
        (icon_person, ["男性スタッフ", "のみ"], [None, ORANGE]),
        (icon_pin, ["全国に", "36院"], [None, ORANGE]),
        (icon_chart, "JISSEKI", None),
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
        if lines == "JISSEKI":
            # 37年目／診療実績40万件（2026-10-13 殿村さん「中央の札に入れたい」）
            segs_c(d, [("37", 36, ORANGE), ("年目", 24, INK)], icx, top + 128)
            segs_c(d, [("診療実績", 21, INK), ("40", 34, ORANGE), ("万件", 21, INK)], icx, top + 162)
            continue
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

    # 予約ボタンの上の一言（2026-10-13 殿村さん「年内に は外したい」→ 案1に決定）
    text_c(d, "＼電話なしで、このままWEB予約できます／", font(34), W / 2, 778, GREEN_DARK, 1)

    cta(c, d, 832, 966)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(22), W / 2, 996, GRAY)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_1016_banner_nantoka_v10_1040.jpg"))
