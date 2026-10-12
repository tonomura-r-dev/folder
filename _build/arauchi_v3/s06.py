# -*- coding: utf-8 -*-
"""元S6「市場分析」（3枚のカード＋出典）。S5と同じ部品なので、S5と同じ直しを当てる。

文字：上の要旨 12→14pt／カード見出し 13→14pt／カード本文 14pt（据え置き）／
      出典の見出し 10→11pt・出典の行 9→10pt。
"""
from _build.arauchi_v3.s05 import market_cards


def apply(ctx):
    market_cards(ctx.orig[6])
