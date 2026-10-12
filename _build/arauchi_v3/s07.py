# -*- coding: utf-8 -*-
"""元S7「要件定義｜LINE運用ツールの選定」（左右2枚のカード＋結論）。

文字：カード本文 12→14pt／見出し横の注記 10・11→12pt／結論 12→14pt。
場所：見出しの左右余白を少し詰めて注記を1行に収める。結論の枠を縦に広げる。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403


def apply(ctx):
    s = ctx.orig[7]
    sh = ids(s)

    # 見出し（14ptの見出し＋注記）：注記を12ptに
    for hid in (7, 11):
        bump_sizes(sh[hid], mapping={10.0: 12.0, 11.0: 12.0, 14.0: 14.0})
        set_insets(sh[hid], l=0.08, r=0.08)
        move(sh[hid], h=0.46)

    # 本文（特徴・メリット・デメリット）
    for bid in (9, 13):
        move(sh[bid], y=1.904, h=3.90)
        set_size_all(sh[bid], 14)
        # 左右の余白を 0.194→0.14in に詰め、「…工数がかかる」の「る」だけが次の行に落ちないようにする
        set_insets(sh[bid], l=0.14, r=0.14)

    # 結論
    con = sh[15]
    set_size_all(con, 14)
    move(con, y=5.93, h=1.13)
