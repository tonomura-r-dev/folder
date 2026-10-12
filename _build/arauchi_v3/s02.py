# -*- coding: utf-8 -*-
"""元S2「現状把握」（スマホ画面2枚＋WEBサイト／LINEのステータス表）。

文字：上の要旨 10.5→14pt（S5・S6と揃える）／表の見出し帯 12→14pt／
      表の項目名・値 10→12pt／表の下の注記 9→11pt（薄い灰色A6A6A6→595959）／下の説明の箱 10.5→12pt。
場所：スマホ画面を約82%に縮め、表の列を広げる（項目名1.40in・値1.83in）。行の間隔を0.355inにして全体を少し下へ伸ばす。
      右下の説明は12ptで1行に入らないため「。」の後ろで改行する（文言はそのまま）。
      右上の白の半透明の飾り（文字なし）がスライドの外にはみ出していたので、スライドの内側だけに切り詰める。
"""
import copy

from lxml import etree

from _build.arauchi_v3.common import *  # noqa: F401,F403

LEAD_PT = 14
HEAD_PT = 14
CELL_PT = 12
NOTE_PT = 11
BOTTOM_PT = 12

ROW_TOP = 2.10
ROW_PITCH = 0.355
ROW_H = 0.32
LABEL_W = 1.40
VALUE_W = 1.83

LEFT_X = 0.40           # 左（WEBサイト）
LEFT_LABEL_X = 1.95
RIGHT_X = 5.45          # 右（LINE）
RIGHT_LABEL_X = 7.12
PHONE_SCALE = 0.82


def box(shape, x=None, y=None, w=None, h=None, anchor="ctr", lr=None):
    """大きさを固定し（自動調整を外す）、位置・大きさ・上下位置・左右の余白をそろえる"""
    set_autofit_none(shape)
    move(shape, x=x, y=y, w=w, h=h)
    if anchor:
        set_anchor(shape, anchor)
    if lr is not None:
        set_insets(shape, l=lr, r=lr)


def add_break(shape, para_idx, after):
    """段落 para_idx の中の文字 after の直後で改行する（run を分けて a:br を入れる。書式はそのまま）"""
    p_el = txbody(shape).findall(qn("a:p"))[para_idx]
    for r in p_el.findall(qn("a:r")):
        t = r.find(qn("a:t"))
        if t.text and after in t.text:
            pos = t.text.index(after) + len(after)
            head, tail = t.text[:pos], t.text[pos:]
            t.text = head
            rpr = r.find(qn("a:rPr"))
            br = etree.Element(qn("a:br"))
            if rpr is not None:
                b_rpr = copy.deepcopy(rpr)
                b_rpr.tag = qn("a:rPr")
                br.append(b_rpr)
            r.addnext(br)
            if tail:
                r2 = copy.deepcopy(r)
                r2.find(qn("a:t")).text = tail
                br.addnext(r2)
            return
    raise ValueError(f"not found: {after}")


def scale_group(shapes, ox, oy, nx, ny, s):
    """(ox, oy) を基準に図形の集まりを s 倍に縮め、基準を (nx, ny) に移す（EMUで計算）"""
    OX, OY, NX, NY = Inches(ox), Inches(oy), Inches(nx), Inches(ny)
    for sh in shapes:
        l, t, w, h = sh.left, sh.top, sh.width, sh.height
        sh.left = Emu(int(round(NX + (l - OX) * s)))
        sh.top = Emu(int(round(NY + (t - OY) * s)))
        sh.width = Emu(int(round(w * s)))
        sh.height = Emu(int(round(h * s)))


def apply(ctx):
    slide = ctx.orig[2]
    sh = ids(slide)

    # ---- 上の要旨（帯の中で上下中央。S5・S6と同じ枠）
    lead = sh[26]
    set_size_all(lead, LEAD_PT)
    box(lead, x=0.466, y=0.695, w=9.918, h=0.764)

    # ---- 表の見出し帯（WEBサイトステータス／LINEステータス）
    set_size_all(sh[31], HEAD_PT)
    box(sh[31], x=LEFT_X, y=1.64, w=VALUE_W + LABEL_W + (LEFT_LABEL_X - LEFT_X) + 0.02, h=0.38)
    set_size_all(sh[29], HEAD_PT)
    right_w = VALUE_W + LABEL_W + (RIGHT_LABEL_X - RIGHT_X) + 0.02
    box(sh[29], x=RIGHT_X, y=1.64, w=right_w, h=0.38)

    # ---- 表（項目名＝濃い灰色の箱、値＝薄い灰色の箱）。x の範囲で左右・列を分け、上から順に並べ直す
    def col(x_lo, x_hi):
        return sorted([s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()
                       and x_lo <= s.left / 914400 < x_hi and 2.1 < s.top / 914400 < 5.8
                       and s.height / 914400 < 0.35], key=lambda s: s.top)

    cols = [
        (col(2.3, 2.4), LEFT_LABEL_X, LABEL_W),                     # 左の項目名（10行）
        (col(3.5, 3.6), LEFT_LABEL_X + LABEL_W + 0.02, VALUE_W),    # 左の値
        (col(7.2, 7.3), RIGHT_LABEL_X, LABEL_W),                    # 右の項目名（11行）
        (col(8.4, 8.5), RIGHT_LABEL_X + LABEL_W + 0.02, VALUE_W),   # 右の値
    ]
    assert [len(c[0]) for c in cols] == [10, 10, 11, 11], [len(c[0]) for c in cols]
    for cells, x, w in cols:
        for i, cell in enumerate(cells):
            set_size_all(cell, CELL_PT)
            box(cell, x=x, y=ROW_TOP + i * ROW_PITCH, w=w, h=ROW_H, lr=0.05)

    # ---- スマホ画面（左：1枚の画像／右：枠＋白地＋画面の3枚重ね）を縮める
    phone_l = sh[9]
    scale_group([phone_l], phone_l.left / 914400, phone_l.top / 914400, LEFT_X, ROW_TOP, PHONE_SCALE)
    frame = sh[55]
    for s_ in (sh[61],):
        set_autofit_none(s_)            # 白地（空の段落入り）が自動で伸びないように
    scale_group([frame, sh[61], sh[24]], frame.left / 914400, frame.top / 914400, RIGHT_X, ROW_TOP, PHONE_SCALE)

    # ---- 注記（左の表の下）
    note = sh[67]
    set_size_all(note, NOTE_PT)
    for i in range(2):
        set_color(note, i, "595959")
    last_row_bottom = ROW_TOP + 9 * ROW_PITCH + ROW_H
    box(note, x=LEFT_LABEL_X, y=last_row_bottom + 0.05, w=LABEL_W + VALUE_W + 0.02, h=0.50, anchor="t")

    # ---- 下の説明の箱
    bottom_y = ROW_TOP + 10 * ROW_PITCH + ROW_H + 0.22
    set_size_all(sh[43], BOTTOM_PT)
    box(sh[43], x=LEFT_X, y=bottom_y, w=sh[31].width / 914400, h=0.60, lr=0.05)
    set_size_all(sh[53], BOTTOM_PT)
    add_break(sh[53], 0, "ボタンなし。")
    box(sh[53], x=RIGHT_X, y=bottom_y, w=right_w, h=0.60, lr=0.05)

    # ---- 右上の白の半透明の飾り（文字なし）：スライドの外にはみ出した分を切り詰める
    deco = sh[23]
    set_autofit_none(deco)
    deco.width = Emu(ctx.prs.slide_width - deco.left)

