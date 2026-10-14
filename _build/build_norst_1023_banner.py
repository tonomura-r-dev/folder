"""ノースト 10月3本目（仮に10/23）のバナー（1040×1040）デモ。写真・イラストなし。

論理（2026-10-14 殿村さん）：アイテム名「「自分だけなのかな？」その悩み、聞かせてください。」で興味を引く →
画像で「うちにはこれだけの実績がある」を数字で見せる → 一度話（無料カウンセリング）をさせてください。
v3は殿村さん「ラフにしては適当すぎる／情報量それだけでいいと？」→ v4で2本目（v10）と同じくらいの情報量にした。
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
    font, tw, th, text, text_c, rrect, ellipse, white_logo,
)
from build_norst_rough_patterns import cta  # noqa: E402
from build_norst_1016_banner import counsel_badge  # noqa: E402

PANEL_BOTTOM = 502


def check(d, x, cy, col=ORANGE):
    """チェックの印（四角に✓）。"""
    rrect(d, (x, cy - 14, x + 28, cy + 14), 5, fill=WHITE, outline=col, width=3)
    d.line([((x + 6) * S, cy * S), ((x + 12) * S, (cy + 7) * S), ((x + 24) * S, (cy - 9) * S)],
           fill=col, width=4 * S, joint="curve")


def stat_card(d, x1, y1, x2, y2, label, num, unit):
    """白い札に大きな数字。label＝数字の上の小さな言葉。"""
    rrect(d, (x1 + 5, y1 + 5, x2 + 5, y2 + 5), 18, fill=GREEN_DARK)  # 影
    rrect(d, (x1, y1, x2, y2), 18, fill=WHITE)
    cx = (x1 + x2) / 2
    f_l, f_n, f_u = font(28), font(104), font(44)
    text_c(d, label, f_l, cx, y1 + 16, GREEN_DARK)
    w = tw(f_n, num) + tw(f_u, unit) + 4
    x = cx - w / 2
    top = y1 + 58
    text(d, num, f_n, x, top, ORANGE)
    text(d, unit, f_u, x + tw(f_n, num) + 4, top + th(f_n, num) - th(f_u, unit), GREEN_DARK)


def list_card(d, x1, y1, x2, y2, head, items, col):
    """見出しの帯＋チェックの並び。"""
    rrect(d, (x1, y1, x2, y2), 16, fill=WHITE, outline=col, width=3)
    hh = 54
    rrect(d, (x1, y1, x2, y1 + hh), 16, fill=col)
    d.rectangle((x1 * S, (y1 + 30) * S, x2 * S, (y1 + hh) * S), fill=col)  # 帯の下の角は四角に
    f_h = font(29)
    text_c(d, head, f_h, (x1 + x2) / 2, y1 + hh / 2 - th(f_h, head) / 2, WHITE)
    f_i = font(30)  # v6：2026-10-14 殿村さん「中の文をもう少し大きく」26→30
    step = (y2 - y1 - hh) / len(items)
    for i, s in enumerate(items):
        cy = y1 + hh + step * (i + 0.5)
        check(d, x1 + 16, cy, col)
        text(d, s, f_i, x1 + 54, cy - th(f_i, s) / 2, INK)
        assert x1 + 54 + tw(f_i, s) <= x2 - 20, s  # 枠に触れない


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面＋ロゴ（中央）＋右上の円「カウンセリング無料」（2本目と同じ）
    # v5：2026-10-14 殿村さん「吹き出しを2行に／治療例をもう少し真ん中に／全体的に上に」
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)
    logo = white_logo(290 * S)
    c.paste(logo, ((W * S - logo.width) // 2, 18 * S), logo)
    counsel_badge(d, 920, 112, 78)

    # 心の声（2行。アイテム名と同じ言葉）。v6：殿村さん「吹き出しと中の文字はもう少し小さく」38→30
    f_v = font(30)
    v1, v2 = "「自分だけ、", "　なのかな？」"
    vw = max(tw(f_v, v1), tw(f_v, v2)) + 40
    bx1, by1, by2 = 40, 58, 144  # v9：殿村さん「吹き出しもう少し上に」
    rrect(d, (bx1, by1, bx1 + vw, by2), 26, fill=WHITE)
    d.polygon([(66 * S, (by2 - 2) * S), (90 * S, (by2 - 2) * S), (62 * S, (by2 + 18) * S)], fill=WHITE)
    text(d, v1, f_v, bx1 + 18, by1 + 13, GREEN_DARK)
    text(d, v2, f_v, bx1 + 18, by1 + 49, GREEN_DARK)

    # 診療のタグ3つ。v6：殿村さん「中心に添えたい／亀頭増大が真ん中に来るように」＝真ん中のタグの中心を横の中央に
    f_t = font(26)
    tags = ["包茎治療", "亀頭増大", "ED治療"]
    widths = [tw(f_t, t) + 32 for t in tags]
    tx = W / 2 - widths[1] / 2 - 10 - widths[0]
    assert tx >= bx1 + vw + 16, (tx, bx1 + vw)  # 吹き出しに重ならない
    ty1, ty2 = 100, 136  # v7：殿村さん「治療例もっと上に／ロゴとの間無駄」＝ロゴのすぐ下
    for t, w_ in zip(tags, widths):
        rrect(d, (tx, ty1, tx + w_, ty2), 18, fill=GREEN_DARK, outline=WHITE, width=2)
        text(d, t, f_t, tx + 16, (ty1 + ty2) / 2 - th(f_t, t) / 2, WHITE)
        tx += w_ + 10

    # 主役：一度、お話を／聞かせてください。（「聞かせて」をオレンジに）
    # v8：殿村さん「全体的に上にできるよね／都度文字も大きくできるよね」＝空きを詰めて、浮いた分で文字を大きく
    f_1, f_2 = font(56), font(98)
    y1 = 152
    text_c(d, "一度、お話を", f_1, W / 2, y1, WHITE, 2, stroke=1, stroke_fill=WHITE)
    a_, b_ = "聞かせて", "ください。"
    y2 = y1 + th(f_1, "一度、お話を") + 10
    w2 = tw(f_2, a_ + b_, 2) + 2
    x = W / 2 - w2 / 2
    for s_, col in ((a_, ORANGE), (b_, WHITE)):
        text(d, s_, f_2, x + 4, y2 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text(d, s_, f_2, x, y2, col, 2, stroke=8 if col == ORANGE else 1.5, stroke_fill=WHITE)
        x += tw(f_2, s_, 2) + 2
    main_bottom = y2 + th(f_2, a_ + b_) + 14

    # 実績の数字3つ（これくらいの数値感で）
    cw, gap, top, bot = 296, 16, 330, 480
    assert main_bottom <= top - 10, main_bottom
    x0 = (W - (cw * 3 + gap * 2)) / 2
    for i, (label, num, unit) in enumerate([("おかげさまで", "37", "年目"),
                                            ("診療実績", "40", "万件"),
                                            ("全国に", "36", "院")]):
        x1 = x0 + i * (cw + gap)
        stat_card(d, x1, top, x1 + cw, bot, label, num, unit)

    # 淡い面：こんなお悩み／安心のカウンセリング
    ly1, ly2 = 522, 732
    list_card(d, 40, ly1, 512, ly2, "こんなお悩み、ありませんか？",
              ["見た目が気になる", "においや蒸れが気になる", "人に見せるのをためらう"], ORANGE)
    list_card(d, 528, ly1, 1000, ly2, "安心のカウンセリング",
              ["費用は無料です", "男性スタッフのみで対応", "その日に決めなくても大丈夫"], GREEN)

    text_c(d, "＼電話なしで、このままWEB予約できます／", font(36), W / 2, 750, GREEN_DARK, 1)
    cta(c, d, 800, 934)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(24), W / 2, 966, GRAY)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_1023_banner_demo_v9_1040.jpg"))
