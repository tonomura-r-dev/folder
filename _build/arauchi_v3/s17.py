# -*- coding: utf-8 -*-
"""S16・S17を1枚に統合する（2026-10-12 殿村さん指示）。

元のS16（章の扉「施策の詳細（新規友だち獲得）」）は build.py の最後で消す。
S17のタイトルを「施策詳細｜新規友だちの獲得」にし、2つの導線を左右2列で整理する：
  左：① 離脱防止ポップアップ → 右：② サイト内の「LINEでお問い合わせ」ボタン
各列＝見出し → 役割 → 対象 → 訴求 → LINE友だち追加 → 注記（費用・リンク先の変更）。
元の訴求・対象物件の注意書き・離脱防止ポップアップの費用はそのまま残す。"""
from lxml import etree

from _build.arauchi_v3.common import (Inches, clone_into, ids, move, qn, set_anchor, set_color, set_insets,
                                      set_line_spacing, set_paras, set_size_all, set_spacing)

NAVY, BLUE = "1F2A5C", "4A6FA5"


def _recolor_fill(shape, hex_):
    sppr = shape._element.spPr
    for el in sppr.findall(qn("a:solidFill")):
        el.find(qn("a:srgbClr")).set("val", hex_)
    ln = sppr.find(qn("a:ln"))
    if ln is not None:
        for el in ln.findall(qn("a:solidFill")):
            el.find(qn("a:srgbClr")).set("val", hex_)


def _rotate_down(shape):
    """右向きの矢印を下向きにする（中心を軸に90度回す）"""
    xfrm = shape._element.spPr.find(qn("a:xfrm"))
    xfrm.set("rot", "5400000")


def _place_down_arrow(shape, cx, top, length=0.26, thick=0.20):
    """下向き（90度回転）の矢印を、見た目の中心 cx・上端 top に置く。回転前の幅＝長さ、高さ＝太さ"""
    w, h = length, thick
    # 回転は中心まわり。回転後の見た目：幅 h・高さ w。中心を (cx, top + w/2) に合わせる
    move(shape, x=cx - w / 2, y=top + w / 2 - h / 2, w=w, h=h)
    _rotate_down(shape)


def apply(ctx):
    s = ctx.orig[17]
    b = ids(s)
    assert b[3].text_frame.text == "提供施策について"
    set_paras(b[3], [[("施策詳細｜新規友だちの獲得", None)]])

    set_paras(b[10], [[("サイトの2つの接点から、LINEの友だち追加へつなげます", None)]], size=16)
    move(b[10], x=0.47, y=1.00, w=9.90, h=0.45)

    cols = (0.47, 5.57)          # 左右の列の左端
    W = 4.80                      # 列の幅
    cx = [c + W / 2 for c in cols]

    # 見出し（元の「LINE追加」の紺の箱を複製して使う）
    h1 = clone_into(s, b[28])
    h2 = clone_into(s, b[28])
    _recolor_fill(h2, BLUE)
    set_paras(h1, [[("① 離脱防止ポップアップ", None)]], size=15)
    set_paras(h2, [[("② サイト内の「LINEでお問い合わせ」ボタン", None)]], size=15)
    for h, c in ((h1, cols[0]), (h2, cols[1])):
        move(h, x=c, y=1.58, w=W, h=0.52)
        for p in h.text_frame._txBody.findall(qn("a:p")):
            p.find(qn("a:pPr")).set("algn", "ctr") if p.find(qn("a:pPr")) is not None else None

    # 役割（元の要旨の枠を複製）
    r1 = clone_into(s, b[10])
    r2 = clone_into(s, b[10])
    set_paras(r1, [[("役割：", True), ("サイトを離れようとした方を引き留め、LINEの友だち追加につなげる", False)]], size=14)
    set_paras(r2, [[("役割：", True), ("問い合わせを考えている方を、LINEの友だち追加につなげる", False)]], size=14)
    for r, c in ((r1, cols[0]), (r2, cols[1])):
        move(r, x=c, y=2.18, w=W, h=0.70)
        set_anchor(r, "ctr")
        set_insets(r, l=0.05, r=0.05)

    # 対象（上）・訴求（中）・LINE追加（下）を縦に並べる
    targets = ((15, ["サイトを離れようとした方", "（離脱防止ポップアップ）"]),
               (29, ["サイト内の「LINEでお問い合わせ」", "ボタン（全ページ）"]))
    appeals = ((26, ["物件のURLを送るだけで仲介手数料100円", "（100円賃貸が対応できる物件に限ります）"]),
               (31, ["物件のURL・物件名・画像を送るだけで、", "空室確認から内見の手配まで"]))
    for (sid, paras), c in zip(targets, cols):
        set_paras(b[sid], paras, size=15)
        move(b[sid], x=c, y=2.98, w=W, h=0.88)
    for (sid, paras), c in zip(appeals, cols):
        set_paras(b[sid], paras, size=15)
        move(b[sid], x=c, y=4.26, w=W, h=0.88)
    for sid, c in ((28, cols[0]), (33, cols[1])):
        set_paras(b[sid], [[("LINEの友だち追加", None)]], size=15)
        move(b[sid], x=c + 0.6, y=5.54, w=W - 1.2, h=0.52)
    # 矢印：対象→訴求、訴求→LINE追加（元の右向き矢印を下向きにして使う）
    for sid, i, top in ((18, 0, 3.92), (30, 1, 3.92), (27, 0, 5.20), (32, 1, 5.20)):
        _place_down_arrow(b[sid], cx[i], top, length=0.30, thick=0.22)

    # 注記：元の1文を、それぞれの列の下に分けて置く（文言はそのまま）
    set_paras(b[45], [[("※離脱防止ポップアップは、LINE公式アカウントとは別に導入する施策です（初期1.5万円・月3万円）。", None)]], size=13)
    move(b[45], x=cols[0], y=6.22, w=W, h=0.66)
    n2 = clone_into(s, b[45])
    set_paras(n2, [[("※サイトのボタンは、リンク先を友だち追加用のURLに変更します。", None)]], size=13)
    move(n2, x=cols[1], y=6.22, w=W, h=0.66)
    for n in (b[45], n2):
        set_color(n, 0, "7F7F7F")          # 元の薄い灰色では読みにくいので、同系統の濃い灰色に
        set_anchor(n, "t")
        set_insets(n, l=0.05, r=0.05)
