# -*- coding: utf-8 -*-
"""元S11「要件定義」の課題（3枚のカード＋結論）。S5・S6のカードと大きさを揃える。

文字：カード本文 12→14pt／カード見出し 14pt（据え置き）／結論 12→14pt。
場所：カードの下の空きを使って本文の枠を縦に伸ばし、結論の枠を少し上げて広げる。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403


def apply(ctx):
    s = ctx.orig[11]
    sh = ids(s)

    for num_id, title_id, body_id in ((6, 10, 14), (17, 19, 21), (23, 25, 27)):
        move(sh[body_id], h=3.03)
        set_size_all(sh[body_id], 14)

    con = sh[15]
    set_size_all(con, 14)
    move(con, y=5.76, h=1.10)
