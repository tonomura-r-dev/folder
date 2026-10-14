"""ノースト 10月4本目（仮に10/30）のバナー（1040×1040）デモ。3本目（v9）と同じ形の横展開。

訴求（2026-10-14 殿村さん「②と③の融合が理想」）＝費用の不安（②）と、人に知られたくない不安（③）を1本で。
アイテム名「「費用は？バレない？」その2つ、先にお答えします」→ 画像で2つの不安に事実で答える → 無料カウンセリングの予約。
費用の金額は書かない（自由診療の費用を出すと必須表示が要る）。「費用の目安を先にご説明」はノーストのコラム（カウンセリングで
状態を見て合う治療と費用の目安を説明）に基づく。「バレない」と約束はしない（事実＝男性スタッフのみ・電話なしでWEB予約・LINEで問い合わせ）。

使い方:
  python3 _build/build_norst_1030_banner.py [出力]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_norst_1019_rough import (  # noqa: E402
    S, W, H, GREEN, GREEN_DARK, ORANGE, PALE, GRAY, WHITE,
    font, tw, th, text, text_c, rrect, white_logo,
)
from build_norst_rough_patterns import cta  # noqa: E402
from build_norst_1016_banner import counsel_badge  # noqa: E402
from build_norst_1023_banner import stat_card, list_card  # noqa: E402

PANEL_BOTTOM = 518


def stat_card_big(d, x1, y1, x2, y2, label, num, unit):
    """実績の札（3本目より数字を大きく）。v2：殿村さん「もう少し実績を押したバナーに」"""
    rrect(d, (x1 + 6, y1 + 6, x2 + 6, y2 + 6), 20, fill=GREEN_DARK)  # 影
    rrect(d, (x1, y1, x2, y2), 20, fill=WHITE, outline=ORANGE, width=4)
    cx = (x1 + x2) / 2
    f_l, f_n, f_u = font(30), font(124), font(50)
    text_c(d, label, f_l, cx, y1 + 16, GREEN_DARK)
    w = tw(f_n, num) + tw(f_u, unit) + 4
    x = cx - w / 2
    top = y1 + 58
    text(d, num, f_n, x, top, ORANGE)
    text(d, unit, f_u, x + tw(f_n, num) + 4, top + th(f_n, num) - th(f_u, unit), GREEN_DARK)
    assert top + th(f_n, num) <= y2 - 12, (label, top + th(f_n, num), y2)


def build(out):
    c = Image.new("RGB", (W * S, H * S), PALE)
    d = ImageDraw.Draw(c)

    # 深緑の面＋ロゴ（中央）＋右上の円「カウンセリング無料」（3本目と同じ）
    d.rectangle((0, 0, W * S, PANEL_BOTTOM * S), fill=GREEN)
    logo = white_logo(290 * S)
    c.paste(logo, ((W * S - logo.width) // 2, 18 * S), logo)
    counsel_badge(d, 920, 112, 78)

    # 心の声（2行。アイテム名と同じ言葉）
    f_v = font(30)
    v1, v2 = "「費用は？", "　バレない？」"
    vw = max(tw(f_v, v1), tw(f_v, v2)) + 40
    bx1, by1, by2 = 40, 58, 144
    rrect(d, (bx1, by1, bx1 + vw, by2), 26, fill=WHITE)
    d.polygon([(66 * S, (by2 - 2) * S), (90 * S, (by2 - 2) * S), (62 * S, (by2 + 18) * S)], fill=WHITE)
    text(d, v1, f_v, bx1 + 18, by1 + 13, GREEN_DARK)
    text(d, v2, f_v, bx1 + 18, by1 + 49, GREEN_DARK)

    # 診療のタグ3つ（亀頭増大の中心が横の中央）
    f_t = font(26)
    tags = ["包茎治療", "亀頭増大", "ED治療"]
    widths = [tw(f_t, t) + 32 for t in tags]
    tx = W / 2 - widths[1] / 2 - 10 - widths[0]
    assert tx >= bx1 + vw + 16, (tx, bx1 + vw)
    ty1, ty2 = 100, 136
    for t, w_ in zip(tags, widths):
        rrect(d, (tx, ty1, tx + w_, ty2), 18, fill=GREEN_DARK, outline=WHITE, width=2)
        text(d, t, f_t, tx + 16, (ty1 + ty2) / 2 - th(f_t, t) / 2, WHITE)
        tx += w_ + 10

    # 主役：その2つの不安には、／実績でお応えします。（「実績」をオレンジに）
    # v2：殿村さん「もう少し実績を押したバナーに」。v1の「その2つ、先にお答えします。」は画像の中で答えていなかった
    f_1, f_2 = font(48), font(90)
    y1 = 152
    s1 = "その2つの不安には、"
    text_c(d, s1, f_1, W / 2, y1, WHITE, 2, stroke=1, stroke_fill=WHITE)
    a_, b_ = "実績", "でお応えします。"
    y2 = y1 + th(f_1, s1) + 10
    w2 = tw(f_2, a_ + b_, 2) + 2
    assert w2 <= W - 80, w2
    x = W / 2 - w2 / 2
    for s_, col in ((a_, ORANGE), (b_, WHITE)):
        text(d, s_, f_2, x + 4, y2 + 4, GREEN_DARK, 2, stroke=10, stroke_fill=GREEN_DARK)  # 影
        text(d, s_, f_2, x, y2, col, 2, stroke=8 if col == ORANGE else 1.5, stroke_fill=WHITE)
        x += tw(f_2, s_, 2) + 2
    main_bottom = y2 + th(f_2, a_ + b_) + 14

    # 実績の数字3つ（3本目より大きく・オレンジの枠）
    cw, gap, top, bot = 304, 12, 322, 498
    assert main_bottom <= top - 8, main_bottom
    x0 = (W - (cw * 3 + gap * 2)) / 2
    for i, (label, num, unit) in enumerate([("おかげさまで", "37", "年目"),
                                            ("診療実績", "40", "万件"),
                                            ("全国に", "36", "院")]):
        x1 = x0 + i * (cw + gap)
        stat_card_big(d, x1, top, x1 + cw, bot, label, num, unit)

    # 淡い面：2つの不安に1枠ずつ答える
    ly1, ly2 = 536, 742
    list_card(d, 40, ly1, 512, ly2, "費用が心配な方へ",
              ["カウンセリングは無料", "費用の目安を先にご説明", "その日に決めなくても大丈夫"], ORANGE)
    list_card(d, 528, ly1, 1000, ly2, "人に知られたくない方へ",
              ["男性スタッフのみで対応", "電話なしでWEB予約できる", "LINEでも問い合わせできる"], GREEN)

    text_c(d, "＼費用を聞くだけでも、大丈夫です／", font(36), W / 2, 758, GREEN_DARK, 1)
    cta(c, d, 804, 938)
    text_c(d, "※当院は自由診療です。　※診療実績は2026年9月時点", font(24), W / 2, 970, GRAY)
    c.resize((W, H), Image.LANCZOS).save(out, quality=95)
    return out


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "_images/norst_1030_banner_demo_v2_1040.jpg"))
