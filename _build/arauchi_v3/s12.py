# -*- coding: utf-8 -*-
"""元S12「要件定義」の目的・目標・手法。

文字：手法の説明 12→14pt／左の分類ラベル（入口・あいさつ／追加後の案内）11→14pt／
      手法名 14pt・目的24pt・目標20pt（据え置き）。見出しの帯（プロモーション目的・配信目標・
      弊社からのご提案・手法）は継承の18ptをそのまま明示。
場所：手法名の枠を少し細くして説明の枠を広げ、手法の行を縦に少し伸ばす（2行で収める）。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403

ROW_Y0 = 4.076
ROW_PITCH = 0.72
ROW_H = 0.66


def apply(ctx):
    s = ctx.orig[12]
    sh = ids(s)

    # 継承の18pt（見出しの帯・矢印の文字）を明示する（大きさは変えない）
    for hid in (26, 28, 33, 17):
        set_size_all(sh[hid], 18)
    # 見出しの帯（プロモーション目的・配信目標・手法）：上の余白が大きく、文字が帯の下の枠に
    # かかっていたので、上下の余白を揃えて帯の中央に置く
    set_insets(sh[26], t=0.0, b=0.0)
    for hid in (33, 17):
        set_insets(sh[hid], t=0.02, b=0.02)

    # 手法の4行：名前（14pt据え置き）＋説明（12→14pt）
    names = (2, 5, 6, 7)
    descs = (8, 9, 10, 11)
    name_x, name_w = 1.722, 2.95
    desc_x = name_x + name_w + 0.071
    desc_w = 10.225 - desc_x
    for i, (nid, did) in enumerate(zip(names, descs)):
        y = ROW_Y0 + i * ROW_PITCH
        move(sh[nid], x=name_x, y=y, w=name_w, h=ROW_H)
        move(sh[did], x=desc_x, y=y, w=desc_w, h=ROW_H)
        set_size_all(sh[did], 14)
        set_insets(sh[did], l=0.1, r=0.08, t=0.03, b=0.03)

    # 左の分類ラベル（2行ずつまとめる）
    for j, lid in enumerate((12, 13)):
        y = ROW_Y0 + 2 * j * ROW_PITCH
        move(sh[lid], x=0.651, y=y, w=1.0, h=ROW_PITCH + ROW_H)
        set_size_all(sh[lid], 14)
        set_insets(sh[lid], l=0.03, r=0.03)
