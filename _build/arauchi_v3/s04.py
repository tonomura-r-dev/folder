# -*- coding: utf-8 -*-
"""元S4「LINEアカウントの現状」（スマホ画面3枚＋下の説明の箱3つ）。

文字：上の要旨 12→14pt（S5・S6と揃える）／下の説明の箱 10.5→13pt。
場所：スマホ画面は大きさそのままで少し上・左へ寄せ、説明の箱を左右に広げる（左5.05in・右4.90in）。
      13ptで1行に入らない2か所は、読点・句点の後ろで改行する（文言はそのまま）：
      「あいさつの後にマンガ・動画を送り、／サービスの理解を促している。」
      「リッチメニューは未設置で、／友だち追加後に次へ進む導線がない。」
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403
from _build.arauchi_v3.s02 import add_break, box

LEAD_PT = 14
CAPTION_PT = 13

LEFT_X, LEFT_W = 0.45, 5.05
RIGHT_X, RIGHT_W = 5.65, 4.90
PHONE_Y = 1.63
CAP_Y = 5.58
CAP2_H = 0.60      # 2行の箱
CAP3_H = 0.84      # 3行の箱


def shift(shapes, dx, dy):
    for s_ in shapes:
        s_.left = Emu(int(round(s_.left + Inches(dx))))
        s_.top = Emu(int(round(s_.top + Inches(dy))))


def apply(ctx):
    slide = ctx.orig[4]
    sh = ids(slide)

    # ---- 上の要旨（帯の中で上下中央。S5・S6と同じ枠）
    lead = sh[28]
    set_size_all(lead, LEAD_PT)
    box(lead, x=0.466, y=0.695, w=9.918, h=0.764)

    # ---- スマホ画面（大きさはそのまま。左の2台は説明の箱の左端に、右の1台も右の箱の左端にそろえる）
    set_autofit_none(sh[46])                       # 画面の後ろの白地（空の段落入り）が自動で伸びないように
    phone_a = [sh[41], sh[46], sh[7]]              # 枠＋白地＋あいさつの画面
    phone_b = [sh[11], sh[21]]                     # 枠＋マンガの画面
    phone_c = [sh[9], sh[13]]                      # 枠＋プロフィールの画面
    gap_ab = (sh[11].left - sh[41].left) / 914400  # 左の2台の間隔（動かす前に測る）
    for grp, new_x in ((phone_a, LEFT_X), (phone_b, LEFT_X + gap_ab), (phone_c, RIGHT_X)):
        base = grp[0]
        shift(grp, new_x - base.left / 914400, PHONE_Y - base.top / 914400)

    # ---- 下の説明の箱
    a, b, c = sh[17], sh[23], sh[19]
    for s_ in (a, b, c):
        set_size_all(s_, CAPTION_PT)
    add_break(b, 0, "動画を送り、")
    add_break(c, 1, "未設置で、")
    box(a, x=LEFT_X, y=CAP_Y, w=LEFT_W, h=CAP2_H, lr=0.05)
    box(b, x=LEFT_X, y=CAP_Y + CAP2_H + 0.08, w=LEFT_W, h=CAP2_H, lr=0.05)
    box(c, x=RIGHT_X, y=CAP_Y, w=RIGHT_W, h=CAP3_H, lr=0.05)
