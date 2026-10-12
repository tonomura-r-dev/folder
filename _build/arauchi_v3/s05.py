# -*- coding: utf-8 -*-
"""元S5「市場分析」（3枚のカード＋出典）。S6も同じ部品なので market_cards() を共有する。

文字：上の要旨 12→14pt／カード見出し 13→14pt／カード本文 14pt（据え置き）／
      出典の見出し 10→11pt・出典の行 9→10pt。
場所：カードを少し上へ詰め、出典の枠を縦に広げて10ptの出典が枠に収まるようにする。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403

LEAD_PT = 14
CARD_TITLE_PT = 14
CARD_BODY_PT = 14
SRC_HEAD_PT = 11
SRC_LINE_PT = 10


def market_cards(slide):
    sh = ids(slide)

    # 上の要旨（帯の中で上下中央）
    lead = sh[32]
    set_size_all(lead, LEAD_PT)
    set_autofit_none(lead)
    move(lead, x=0.466, y=0.695, w=9.918, h=0.764)
    set_anchor(lead, "ctr")

    # 3枚のカード：番号 → 見出し → 本文
    for num_id, title_id, body_id in ((33, 34, 35), (36, 37, 38), (39, 40, 41)):
        move(sh[num_id], y=1.66, h=0.62)
        move(sh[title_id], y=2.28, h=0.52)
        set_size_all(sh[title_id], CARD_TITLE_PT)
        move(sh[body_id], y=2.80, h=2.46)
        set_size_all(sh[body_id], CARD_BODY_PT)

    # 出典（見出し11pt・各行10pt）
    src = sh[42]
    move(src, y=5.40, h=1.48)
    tb = txbody(src)
    for i, p_el in enumerate(tb.findall(qn("a:p"))):
        pt = SRC_HEAD_PT if i == 0 else SRC_LINE_PT
        for rpr in list(p_el.iter(qn("a:rPr"))) + list(p_el.iter(qn("a:endParaRPr"))):
            rpr.set("sz", str(int(pt * 100)))


def apply(ctx):
    market_cards(ctx.orig[5])
