# -*- coding: utf-8 -*-
"""元S11「要件定義」の課題（3枚のカード＋結論）。S5・S6のカードと大きさを揃える。

文字：カード本文 12→14pt／カード見出し 14pt（据え置き）／結論 12→14pt。
場所：カードの下の空きを使って本文の枠を縦に伸ばし、結論の枠を少し上げて広げる。
      03の2つ目の項目は「中心で、」の後で改行（「…伝わりにくい」の「い」だけが次の行に落ちないように）。
      結論は「整え、」の後で改行（「リッチメニュー」が途中で切れないように）。
"""
import copy

from _build.arauchi_v3.common import *  # noqa: F401,F403


def break_after(shape, text_mark):
    """text_mark を含む run を text_mark の直後で2つに分け、間に改行（a:br）を入れる（書式はそのまま）"""
    for r in txbody(shape).iter(qn("a:r")):
        t = r.find(qn("a:t"))
        if t is None or text_mark not in (t.text or ""):
            continue
        head, tail = t.text.split(text_mark, 1)
        t.text = head + text_mark
        br = etree.Element(qn("a:br"))
        rpr = r.find(qn("a:rPr"))
        if rpr is not None:
            br.append(copy.deepcopy(rpr))
        r2 = copy.deepcopy(r)
        r2.find(qn("a:t")).text = tail
        r.addnext(br)
        br.addnext(r2)
        return
    raise ValueError(f"not found: {text_mark}")


def apply(ctx):
    s = ctx.orig[11]
    sh = ids(s)

    for num_id, title_id, body_id in ((6, 10, 14), (17, 19, 21), (23, 25, 27)):
        move(sh[body_id], h=3.03)
        set_size_all(sh[body_id], 14)

    break_after(sh[27], "が中心で、")

    con = sh[15]
    set_size_all(con, 14)
    move(con, y=5.76, h=1.10)
    break_after(con, "を整え、")
