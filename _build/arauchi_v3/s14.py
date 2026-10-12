# -*- coding: utf-8 -*-
"""元S14「実施施策」（ステップ配信3行×6列・企画配信4行×5列の箱のフロー）。

文字：上の要旨 12.5→14pt（S5・S6・S13と揃える）／「[ステップ配信・全10通]」「[企画配信]」12.5→14pt／
      列の見出し（ターゲット層・セグメント方法・配信コンテンツ・ねらい）9.5→12pt／
      箱の中（日数ラベル・各箱・問い合わせへ）9.5→12pt。
場所：1枚のまま収める（分割しない）。箱の列幅を中身の長さに合わせて組み替え（ステップ配信は
      日数ラベル1.08in＋1.30/1.60/1.74/1.74/1.42in、企画配信は1.90/1.90/1.80/1.74/1.60in）、
      矢印を0.32→0.26inに詰めて左右いっぱい（0.42〜10.40in）に広げる。12ptで元の改行どおり
      1行に収まる幅にしてある（最長＝「物件を送っていない方」「問い合わせ済みを除く」の10文字）。
      箱の高さ0.50→0.53in・行の間隔0.61in。2行の箱はメイリオの行の高さでも収まるよう行間90%・上下の余白を詰める。
      列の見出しは各列の真上に置き直す。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403

BOX_PT = 12
HEAD_PT = 12
SECTION_PT = 14
LEAD_PT = 14

L, R = 0.42, 10.40
ARROW_W, ARROW_H = 0.26, 0.159
BOX_H = 0.53
PITCH = 0.60

# ステップ配信：日数ラベル＋5つの箱
STEP_LABEL_W = 1.08
STEP_GAP = 0.06
STEP_W = (1.30, 1.60, 1.74, 1.74, 1.42)
STEP_ROWS = (
    (20, 24, 28, 32, 35, 36, 39, 40, 43, 44),
    (48, 49, 51, 52, 53, 55, 56, 57, 59, 60),
    (61, 63, 64, 65, 67, 68, 69, 71, 72, 73),
)
STEP_HEADS = (5, 8, 10, 12)          # ターゲット層・セグメント方法・配信コンテンツ・ねらい（箱1〜4の上）

# 企画配信：5つの箱
PLAN_W = (1.90, 1.90, 1.80, 1.74, 1.60)
PLAN_ROWS = (
    (75, 76, 77, 78, 79, 81, 83, 85, 87),
    (88, 89, 90, 91, 92, 93, 94, 95, 96),
    (97, 98, 99, 100, 101, 102, 103, 104, 105),
    (106, 107, 108, 109, 110, 111, 112, 113, 114),
)
PLAN_HEADS = (15, 17, 19, 22)

# 縦の位置
STEP_SEC_Y, SEC_H = 1.50, 0.32
STEP_HEAD_Y, HEAD_H = 1.82, 0.28
STEP_ROW_Y0 = 2.10
PLAN_SEC_Y = 4.06
PLAN_HEAD_Y = 4.38
PLAN_ROW_Y0 = 4.66


def _xs(x0, widths):
    """箱の左端の位置（箱と箱の間に矢印1本ぶんの幅）"""
    out, x = [], x0
    for w in widths:
        out.append(x)
        x += w + ARROW_W
    return out


def _box(shape, x, y, w):
    move(shape, x=x, y=y, w=w, h=BOX_H)
    set_size_all(shape, BOX_PT)
    lb = txbody(shape).find(qn("a:bodyPr"))
    lr = 0.0 if lb.get("lIns") == "0" else 0.04     # 日数ラベルは元から左右の余白0
    set_insets(shape, l=lr, r=lr, t=0.02, b=0.02)
    set_line_spacing(shape, 90)
    set_anchor(shape, "ctr")


def _arrow(shape, x, y):
    move(shape, x=x, y=y + (BOX_H - ARROW_H) / 2, w=ARROW_W, h=ARROW_H)


def _head(shape, cx, w, y):
    """列の見出し：列の真上に中央寄せ"""
    set_size_all(shape, HEAD_PT)
    set_autofit_none(shape)
    move(shape, x=cx - w / 2, y=y, w=w, h=HEAD_H)
    set_insets(shape, l=0.02, r=0.02, t=0.0, b=0.02)
    set_anchor(shape, "b")


def _section(shape, y):
    set_size_all(shape, SECTION_PT)
    set_autofit_none(shape)
    move(shape, x=L, y=y, w=4.0, h=SEC_H)
    set_insets(shape, l=0.05, r=0.05, t=0.0, b=0.0)
    set_anchor(shape, "ctr")


def apply(ctx):
    s = ctx.orig[14]
    sh = ids(s)

    # 上の要旨（帯の中で上下中央。S5・S6・S13と同じ）
    lead = sh[6]
    set_size_all(lead, LEAD_PT)
    set_autofit_none(lead)
    move(lead, x=0.466, y=0.695, w=9.918, h=0.764)
    set_anchor(lead, "ctr")

    # ---- ステップ配信
    _section(sh[13], STEP_SEC_Y)
    box_x0 = L + STEP_LABEL_W + STEP_GAP
    xs = _xs(box_x0, STEP_W)
    for i, row in enumerate(STEP_ROWS):
        y = STEP_ROW_Y0 + i * PITCH
        label, boxes, arrows = row[0], row[1::2], row[2::2]
        _box(sh[label], L, y, STEP_LABEL_W)
        for bid, x, w in zip(boxes, xs, STEP_W):
            _box(sh[bid], x, y, w)
        for aid, x, w in zip(arrows, xs, STEP_W):
            _arrow(sh[aid], x + w, y)
    for hid, x, w in zip(STEP_HEADS, xs, STEP_W):
        _head(sh[hid], x + w / 2, w, STEP_HEAD_Y)

    # ---- 企画配信
    _section(sh[74], PLAN_SEC_Y)
    xs = _xs(L, PLAN_W)
    for i, row in enumerate(PLAN_ROWS):
        y = PLAN_ROW_Y0 + i * PITCH
        boxes, arrows = row[0::2], row[1::2]
        for bid, x, w in zip(boxes, xs, PLAN_W):
            _box(sh[bid], x, y, w)
        for aid, x, w in zip(arrows, xs, PLAN_W):
            _arrow(sh[aid], x + w, y)
    for hid, x, w in zip(PLAN_HEADS, xs, PLAN_W):
        _head(sh[hid], x + w / 2, w, PLAN_HEAD_Y)

    assert abs(_xs(box_x0, STEP_W)[-1] + STEP_W[-1] - R) < 0.005
    assert abs(_xs(L, PLAN_W)[-1] + PLAN_W[-1] - R) < 0.005
