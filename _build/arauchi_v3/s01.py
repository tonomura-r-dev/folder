# -*- coding: utf-8 -*-
"""S1 表紙：サブタイトルを提案の軸に差し替え、ひと目で読める大きさ・位置にする（2026-10-12 殿村さん指示）。
タイトル「LINE公式アカウント運用のご提案」と線は残す。サブタイトルは線の下に、薄い青の帯の上で紺の太字22pt・2行。
「物件問い合わせの増加」だけ赤（資料の強調色）にする。"""
from lxml import etree

from _build.arauchi_v3.common import Inches, ids, move, qn, set_anchor, set_autofit_none, set_insets, set_paras

NAVY, RED, TINT = "1F285A", "C00000", "F4F7FF"


def _fill(shape, hex_):
    sppr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill"):
        for el in sppr.findall(qn(tag)):
            sppr.remove(el)
    fill = etree.Element(qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr"), val=hex_)
    geom = sppr.find(qn("a:prstGeom"))
    geom.addnext(fill)
    ln = sppr.find(qn("a:ln"))
    if ln is None:
        ln = etree.SubElement(sppr, qn("a:ln"))
    for el in list(ln):
        ln.remove(el)
    etree.SubElement(ln, qn("a:noFill"))


def _color_runs(p_el, hexes):
    for r, hex_ in zip(p_el.findall(qn("a:r")), hexes):
        rpr = r.find(qn("a:rPr"))
        for f in rpr.findall(qn("a:solidFill")):
            rpr.remove(f)
        fill = etree.Element(qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr"), val=hex_)
        rpr.insert(0, fill)
        rpr.set("b", "1")


def apply(ctx):
    s = ctx.orig[1]
    b = ids(s)
    assert "最上位パートナー" in b[6].text_frame.text
    # タイトルと線を少し上げ、線の下にサブタイトルの帯を置く
    move(b[2], y=2.30)
    move(b[5], y=3.03)
    sub = b[6]
    set_paras(sub, [
        [("LINEへの流入と友だち追加後の導線を整備し、", None)],
        [("物件問い合わせの増加", None), ("を実現する", None)],
    ], size=22)
    ps = sub.text_frame._txBody.findall(qn("a:p"))
    _color_runs(ps[0], [NAVY])
    _color_runs(ps[1], [RED, NAVY])
    for p_el in ps:
        p_el.find(qn("a:pPr")).set("algn", "ctr") if p_el.find(qn("a:pPr")) is not None else None
    set_autofit_none(sub)
    move(sub, x=1.46, y=3.28, w=7.91, h=1.15)
    set_insets(sub, l=0.2, t=0.1, r=0.2, b=0.1)
    set_anchor(sub, "ctr")
    _fill(sub, TINT)
