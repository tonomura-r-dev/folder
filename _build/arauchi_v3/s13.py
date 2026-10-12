# -*- coding: utf-8 -*-
"""元S13「想定の費用対効果」（初期・月次の実施内容の表＋6か月後の数字＋注記）。

文字：上の要旨 12→14pt（S5・S6と揃える）／「初期・月次・6か月後」の見出し 12→14pt／
      ＜実施内容＞ 12→14pt／初期・月次の表 11→13pt／6か月後のラベル・数字 12→14pt／
      注記 9→10pt（S5・S6の出典と揃える）。金額（16pt）は据え置き。
場所：見出し・金額の枠を少し低くして表を上へ寄せ、表の列幅を広げる（「リッチメニュー」が
      13ptで1行に収まるように）。6か月後のラベル枠も広げる。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403

HEAD_Y, HEAD_H = 1.66, 0.42
AMT_Y, AMT_H = 2.14, 0.74
SUB_Y, SUB_H = 2.94, 0.38
ROW_Y0, ROW_PITCH, ROW_H = 3.38, 0.535, 0.49
TABLE_PT = 13

# 表：初期（左）と月次（右）の (ラベル, 値) の id。上から順に。
INIT_ROWS = ((12, 14), (18, 22), (26, 30), (34, 38), (42, 47))
MONTH_ROWS = ((25, 29), (33, 37), (41, 46), (50, 54), (58, 62), (66, 70))
# 6か月後：(ラベル, 数字)
RESULT_ROWS = ((80, 11), (82, 16), (84, 23), (86, 27))


def apply(ctx):
    s = ctx.orig[13]
    sh = ids(s)

    # 上の要旨（帯の中で上下中央。S5・S6と同じ）
    lead = sh[88]
    set_size_all(lead, 14)
    set_autofit_none(lead)
    move(lead, x=0.466, y=0.695, w=9.918, h=0.764)
    set_anchor(lead, "ctr")

    # 見出し：初期・月次・6か月後
    for hid in (5, 7, 17):
        move(sh[hid], y=HEAD_Y, h=HEAD_H)
        set_size_all(sh[hid], 14)

    # 金額（16ptのまま）
    for aid in (8, 10):
        move(sh[aid], y=AMT_Y, h=AMT_H)

    # ＜実施内容＞
    for tid, col_x in ((19, 0.286), (21, 3.65)):
        set_size_all(sh[tid], 14)
        set_autofit_none(sh[tid])
        move(sh[tid], x=col_x, y=SUB_Y, w=3.171, h=SUB_H)
        set_anchor(sh[tid], "ctr")

    # 表（ラベル＋値）
    for rows, col_x in ((INIT_ROWS, 0.286), (MONTH_ROWS, 3.65)):
        lab_x = col_x + 0.05
        lab_w = 1.56
        val_x = lab_x + lab_w + 0.04
        val_w = 1.47
        for i, (lid, vid) in enumerate(rows):
            y = ROW_Y0 + i * ROW_PITCH
            for cid, x, w in ((lid, lab_x, lab_w), (vid, val_x, val_w)):
                move(sh[cid], x=x, y=y, w=w, h=ROW_H)
                set_size_all(sh[cid], TABLE_PT)
                set_insets(sh[cid], l=0.05, r=0.05, t=0.02, b=0.02)
    # 2行のラベル（リッチメニュー／修正変更）は行間を詰める
    set_line_spacing(sh[33], 90)

    # 矢印
    move(sh[15], y=AMT_Y, h=3.06)

    # 6か月後：ラベル（緑）と数字（赤）
    r_lab_x, r_lab_w = 7.45, 1.56
    r_val_x, r_val_w = 9.03, 1.45
    r_pitch, r_h = 0.52, 0.44
    for i, (lid, vid) in enumerate(RESULT_ROWS):
        y = AMT_Y + i * r_pitch
        move(sh[lid], x=r_lab_x, y=y, w=r_lab_w, h=r_h)
        set_size_all(sh[lid], 14)
        set_insets(sh[lid], l=0.05, r=0.05)
        v = sh[vid]
        set_size_all(v, 14)
        set_autofit_none(v)
        move(v, x=r_val_x, y=y, w=r_val_w, h=r_h)
        set_anchor(v, "ctr")
    tot = sh[31]
    set_size_all(tot, 14)
    set_autofit_none(tot)
    move(tot, x=r_val_x + r_val_w / 2 - 1.05, y=AMT_Y + 4 * r_pitch, w=2.1, h=0.40)
    set_insets(tot, l=0.02, r=0.02)
    set_anchor(tot, "ctr")

    # 注記（9→10pt）
    note = sh[87]
    set_size_all(note, 10)
    set_autofit_none(note)
    move(note, x=0.29, y=6.60, w=10.3, h=0.48)
