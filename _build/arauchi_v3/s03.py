# -*- coding: utf-8 -*-
"""元S3「想定動線」（上：カスタマージャーニーの帯と①〜⑧の階段／下：貴社のユーザー動線の図）。

文字：上の要旨 12→14pt（S5・S6と揃える。1行目は「LINE追加 →」の後ろで改行して3行に）／
      「カスタマージャーニーマップ」10.5→12pt／①〜⑧の箱 11→12pt（「(成約)」10→11pt・「(web/LINE)」9→11pt）／
      「<貴社のユーザー動線イメージ>」10.5→12pt（元は「>」だけ2行目に落ちていた）／
      動線図の箱 ①〜④ 11→13pt・問い合わせ／お申込み／ご契約 12→13pt・（鍵のお渡し・お引っ越し）10→12pt
      （薄すぎる灰色F2F2F2→595959）／線の上のラベル「電話・WEB」「LINE」9→11pt。
場所：①〜④の箱を2.04→2.30inに広げ、階段の段の高さを0.29inにそろえる。下の灰色の枠を下へ広げ、
      動線図の箱の高さを0.36inにそろえる。つなぎの線は新しい箱の位置から計算し直す
      （④→⑤の点線は③の箱に重ならないよう、③の右端の外から出す）。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403
from _build.arauchi_v3.s02 import add_break, box

LEAD_PT = 14
PHASE_MAP = {11.0: 12.0, 10.0: 11.0, 9.0: 11.0}
FLOW_PT = 13
LABEL_PT = 11

# 上の階段（①〜④＝橙、⑤〜⑧＝青）
ROWS = [2.12, 2.41, 2.70, 2.99]
STEP_H = 0.36
ORANGE_W = 2.30
BLUE_W = 2.572

# 下の動線図
FLOW_H = 0.36
LEFT_X, LEFT_W = 0.40, 2.15
Y1, Y2, Y3, Y4 = 4.08, 4.52, 5.90, 6.34          # ①〜④の上端
Q_X, Q_W = 3.874, 3.982                          # 問い合わせ
APPLY_X, APPLY_Y = 5.283, 4.98                   # お申込み
CONTRACT_X, CONTRACT_Y = 6.692, 5.55             # ご契約
KEY_X, KEY_Y, KEY_H = 8.102, 6.07, 0.32          # （鍵のお渡し・お引っ越し）


def E(v):
    return Emu(int(round(Inches(v))))


def set_xfrm(shape, x, y, w, h):
    shape.left, shape.top, shape.width, shape.height = E(x), E(y), E(w), E(h)


def apply(ctx):
    slide = ctx.orig[3]
    sh = ids(slide)

    # ---- 上の要旨（3行・帯の中で上下中央）
    lead = sh[31]
    set_size_all(lead, LEAD_PT)
    add_break(lead, 0, "LINE追加 → ")
    box(lead, x=0.466, y=0.655, w=9.918, h=0.84)

    # ---- カスタマージャーニーの帯
    move(sh[15], y=1.62, h=0.44)                      # 灰色の帯
    move(sh[16], y=1.84)                              # 両矢印の点線
    lab = sh[17]
    set_size_all(lab, 12)
    cx = lab.left / 914400 + lab.width / 914400 / 2   # 元の中心を保つ
    box(lab, x=cx - 1.45, y=1.68, w=2.90, h=0.32)

    # ---- ①〜⑧の階段
    orange = [sh[10], sh[11], sh[12], sh[13]]         # ①②③④
    blue = [sh[6], sh[8], sh[7], sh[9]]               # ⑤⑥⑦⑧
    for i, s_ in enumerate(orange):
        bump_sizes(s_, mapping=PHASE_MAP)
        box(s_, y=ROWS[i], w=ORANGE_W, h=STEP_H, anchor="t", lr=0.07)
    for i, s_ in enumerate(blue):
        bump_sizes(s_, mapping=PHASE_MAP)
        box(s_, y=ROWS[i], w=BLUE_W, h=STEP_H + (0.08 if i == 3 else 0), anchor="t")

    # ④→⑤の点線（矢印）：③の右端の外側から上がり、⑤の下辺の中央へ。
    # 回転90°・左右上下反転のカギ線なので、回転前の幅＝縦の差、高さ＝横の差。
    cxn = sh[14]
    nv = cxn._element.find(qn("p:nvCxnSpPr")).find(qn("p:cNvCxnSpPr"))
    for tag in ("a:stCxn", "a:endCxn"):          # 新しい位置で固定する（PowerPointとLibreOfficeで同じ形にする）
        for el in nv.findall(qn(tag)):
            nv.remove(el)
    sx, sy = orange[2].left / 914400 + ORANGE_W + 0.09, ROWS[3]          # ③の右端の外・④の上辺
    ex, ey = blue[0].left / 914400 + BLUE_W / 2, ROWS[0] + STEP_H        # ⑤の下辺の中央
    w_, h_ = sy - ey, ex - sx
    set_xfrm(cxn, (sx + ex) / 2 - w_ / 2, (sy + ey) / 2 - h_ / 2, w_, h_)

    # ---- 下の灰色の枠
    move(sh[28], y=3.55, h=3.40)
    lab2 = sh[38]
    set_size_all(lab2, 12)
    box(lab2, x=0.36, y=3.62, w=3.00, h=0.34)

    # ---- 動線図の箱
    for s_, y in ((sh[39], Y1), (sh[26], Y2), (sh[37], Y3), (sh[46], Y4)):
        set_size_all(s_, FLOW_PT)
        box(s_, x=LEFT_X, y=y, w=LEFT_W, h=FLOW_H)
    for s_, x, y, w in ((sh[34], Q_X, Y2, Q_W), (sh[40], APPLY_X, APPLY_Y, BLUE_W),
                        (sh[36], CONTRACT_X, CONTRACT_Y, BLUE_W)):
        set_size_all(s_, FLOW_PT)
        box(s_, x=x, y=y, w=w, h=FLOW_H)
    key = sh[74]
    set_size_all(key, 12)
    set_color(key, 0, "595959")
    box(key, x=KEY_X, y=KEY_Y, w=BLUE_W, h=KEY_H, lr=0.05)

    # ---- つなぎの線
    left_r = LEFT_X + LEFT_W
    mid2, mid4, midq = Y2 + FLOW_H / 2, Y4 + FLOW_H / 2, Y2 + FLOW_H / 2
    set_xfrm(sh[22], left_r, midq, Q_X - left_r, 0)                    # ②→問い合わせ（横一直線）
    set_xfrm(sh[47], left_r, midq, Q_X - left_r, mid4 - midq)          # ④→問い合わせ（上下反転のカギ線）
    cx_ = LEFT_X + LEFT_W / 2
    set_xfrm(sh[75], cx_, Y2 + FLOW_H, 0, Y3 - (Y2 + FLOW_H))          # ②→③（縦の矢印）
    # お申込み（下辺の中央）→ご契約（上辺の中央）：回転270°・左右反転のカギ線
    sx, sy = APPLY_X + BLUE_W / 2, APPLY_Y + FLOW_H
    ex, ey = CONTRACT_X + BLUE_W / 2, CONTRACT_Y
    w_, h_ = ey - sy, ex - sx
    set_xfrm(sh[41], (sx + ex) / 2 - w_ / 2, (sy + ey) / 2 - h_ / 2, w_, h_)

    # ---- 線の上のラベル（電話・WEB／LINE）
    vline_x = left_r + (Q_X - left_r) / 2             # ④→問い合わせの縦の線の位置
    tel = sh[33]
    set_size_all(tel, LABEL_PT)
    box(tel, x=vline_x - 0.55, y=mid2 - 0.25, w=0.58, h=0.50, lr=0.02)
    line = sh[55]
    set_size_all(line, LABEL_PT)
    box(line, x=vline_x - 0.30, y=Y3 - 0.06, w=0.60, h=0.34, lr=0.02)
