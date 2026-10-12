# -*- coding: utf-8 -*-
"""元S8「想定動線」（上段＝入口→友だち追加→配信の箱→問い合わせ、下段＝問い合わせ後の継続配信）。

文字：上の説明 12→14pt／フローの箱 10→12pt／初期設計項目の箇条書き 9→12pt／
      見出しラベル（初期設計項目・定期対応項目）9→11pt／「標準機能で運用」10→12pt。
      見出し「<対新規ユーザー領域>」等の14ptは据え置き。
場所：箱を 0.24→0.28in の高さ・0.325in 間隔に広げ、配信の箱の列を 2.51→2.66in に広げる。
      入口（左）・友だち追加（中）・配信の箱（右）・行き先（右端）の横の配分を組み直し、
      つなぎの線（カギ線）は新しい位置から計算し直す（接続は外して PowerPoint と LibreOffice で同じ形にする）。
      薄い灰色の文字（箇条書き・見出しラベル）は 808080→595959 に濃くする。
"""
from _build.arauchi_v3.common import *  # noqa: F401,F403

E = 914400

LEAD_PT = 14
BOX_PT = 12          # フローの箱
LIST_PT = 12         # 初期設計項目の箇条書き
LABEL_PT = 11        # 初期設計項目↓／定期対応項目（一部抽出）
TAG_PT = 12          # LINE公式アカウントの標準機能で運用

BOX_H = 0.28
PITCH = 0.325

# ---- 上段
UP_ROW0 = 2.05                      # 配信の箱 1行目の上端
COL_X, COL_W = 5.93, 2.66           # 配信の箱の列
HALF_W = (COL_W - 0.06) / 2         # 2つに分かれた行（リサーチ｜→配信に反映）
TGT_X, TGT_W = 8.97, 1.50           # 行き先（空室確認・内見／物件の問い合わせ）
J1 = COL_X + COL_W + 0.17           # 6本の線が物件の問い合わせへ集まる縦の位置
J2 = COL_X + COL_W + 0.28           # 個別チャット → 空室確認・内見 の縦の位置
MID_X, MID_W = 4.13, 1.50           # 友だち追加・初期設計項目
ENTRY_R = 3.89                      # 入口の箱の右端

# ---- 下段
LO_ROW0 = 5.04
LO_COL_X = 4.10
RB_X = 7.36                         # 定期対応項目の列


def cell(shape, x, y, w, h, pt, lr=0.03, tb=0.0, anchor="ctr"):
    """箱：大きさを固定し、位置・文字の大きさ・余白をそろえる"""
    set_autofit_none(shape)
    move(shape, x=x, y=y, w=w, h=h)
    set_size_all(shape, pt)
    set_insets(shape, l=lr, r=lr, t=tb, b=tb)
    set_anchor(shape, anchor)


def _unhook(cxn):
    """コネクタの接続を外す（新しい位置で固定し、PowerPointとLibreOfficeで同じ形にする）"""
    nv = cxn._element.find(qn("p:nvCxnSpPr")).find(qn("p:cNvCxnSpPr"))
    for tag in ("a:stCxn", "a:endCxn"):
        for el in nv.findall(qn(tag)):
            nv.remove(el)


def _xfrm(cxn):
    return cxn._element.find(qn("p:spPr")).find(qn("a:xfrm"))


def _set_adj(cxn, frac):
    for gd in cxn._element.find(qn("p:spPr")).find(qn("a:prstGeom")).iter(qn("a:gd")):
        if gd.get("name") == "adj1":
            gd.set("fmla", "val %d" % int(round(frac * 100000)))


def bent(cxn, sx, sy, ex, ey, jx=None):
    """回転なしのカギ線（bentConnector3）：始点(sx,sy)→ 縦の位置 jx で折れて → 終点(ex,ey)。sx<ex のみ"""
    _unhook(cxn)
    xf = _xfrm(cxn)
    for k in ("flipH", "flipV", "rot"):
        if k in xf.attrib:
            del xf.attrib[k]
    if ey < sy:
        xf.set("flipV", "1")
    cxn.left, cxn.top = Emu(int(sx * E)), Emu(int(min(sy, ey) * E))
    cxn.width, cxn.height = Emu(int((ex - sx) * E)), Emu(int(abs(ey - sy) * E))
    if jx is not None and ex > sx:
        _set_adj(cxn, (jx - sx) / (ex - sx))


def bent_rot(cxn, sx, sy, ex, ey, yh):
    """90°回転・左右上下反転のカギ線：下の始点(sx,sy)から上へ→ yh で横へ→ 右上の終点(ex,ey)へ上がる。
    回転前の幅＝縦の差、高さ＝横の差。"""
    _unhook(cxn)
    w_, h_ = sy - ey, ex - sx
    cx, cy = (sx + ex) / 2, (sy + ey) / 2
    cxn.left, cxn.top = Emu(int((cx - w_ / 2) * E)), Emu(int((cy - h_ / 2) * E))
    cxn.width, cxn.height = Emu(int(w_ * E)), Emu(int(h_ * E))
    _set_adj(cxn, (sy - yh) / w_)


def apply(ctx):
    s = ctx.orig[8]
    sh = ids(s)

    # ---- 上の説明（2行）
    lead = sh[71]
    set_size_all(lead, LEAD_PT)
    set_autofit_none(lead)
    move(lead, x=0.466, y=0.695, w=9.918, h=0.764)
    set_anchor(lead, "ctr")
    set_insets(lead, l=0.05, r=0.05)

    # ================================================================ 上段 <対新規ユーザー領域>
    move(sh[7], y=1.56)                                   # 見出し（14pt据え置き）
    cell(sh[69], 10.547 - 3.10, 1.60, 3.10, 0.30, TAG_PT, lr=0.05)   # LINE公式アカウントの標準機能で運用
    move(sh[5], y=1.93, h=2.48)                           # 上段の枠（→4.41）
    move(sh[26], x=MID_X - 0.14, y=1.99, w=COL_X + COL_W + 0.06 - (MID_X - 0.14), h=2.36)   # 薄緑の面

    # 入口
    cell(sh[9], 0.62, UP_ROW0, 1.44, BOX_H, BOX_PT, lr=0.0)            # 100円賃貸のサイト
    pop = sh[17]
    cell(pop, 2.10, 2.09, ENTRY_R - 2.10, 0.32, BOX_PT)                # 離脱防止ポップアップ
    btn = sh[15]
    cell(btn, 0.62, 2.58, ENTRY_R - 0.62, 0.32, BOX_PT)                # サイトの「LINEでお問い合わせ」ボタン

    # 友だち追加・初期設計項目
    add = sh[27]
    cell(add, MID_X, 2.10, MID_W, 0.30, BOX_PT)
    lab = sh[34]
    cell(lab, MID_X, 2.46, MID_W, 0.26, LABEL_PT)
    set_color(lab, 0, "595959")
    lst = sh[35]
    cell(lst, MID_X, 2.74, 1.60, 1.58, LIST_PT, lr=0.05, tb=0.02, anchor="t")
    set_insets(lst, r=0.02)
    for i in range(len(lst.text_frame.paragraphs)):
        set_color(lst, i, "595959")

    # 中かっこ
    move(sh[49], x=COL_X - 0.18, y=UP_ROW0 - 0.03, w=0.16, h=6 * PITCH + BOX_H + 0.06)

    # 配信の箱（7行）
    rows = [31, 32, (33, 39), 36, 38, 37, 40]
    ys = [UP_ROW0 + PITCH * i for i in range(7)]
    for r, y in zip(rows, ys):
        if isinstance(r, tuple):
            cell(sh[r[0]], COL_X, y, HALF_W, BOX_H, BOX_PT)
            cell(sh[r[1]], COL_X + COL_W - HALF_W, y, HALF_W, BOX_H, BOX_PT)
        else:
            cell(sh[r], COL_X, y, COL_W, BOX_H, BOX_PT)

    # 行き先
    cell(sh[19], TGT_X, ys[0], TGT_W, BOX_H, BOX_PT, lr=0.02)          # 空室確認・内見
    cell(sh[14], TGT_X, ys[1], TGT_W, BOX_H, BOX_PT, lr=0.02)          # 物件の問い合わせ

    # つなぎの線（上段）
    add_mid = 2.10 + 0.15
    bent(sh[30], ENTRY_R, 2.09 + 0.16, MID_X, add_mid, jx=ENTRY_R + 0.07)    # ポップアップ → 友だち追加
    bent(sh[28], ENTRY_R, 2.58 + 0.16, MID_X, add_mid, jx=ENTRY_R + 0.07)    # ボタン → 友だち追加
    col_r = COL_X + COL_W
    q_mid = ys[1] + BOX_H / 2                                           # 物件の問い合わせの左辺の中央
    for cid, row in ((48, 0), (47, 1), (46, 2), (44, 3), (43, 4), (42, 5)):
        bent(sh[cid], col_r, ys[row] + BOX_H / 2, TGT_X, q_mid, jx=J1)
    bent(sh[41], col_r, ys[6] + BOX_H / 2, TGT_X, ys[0] + BOX_H / 2, jx=J2)  # 個別チャット → 空室確認・内見

    # ================================================================ 下段 <友だち追加後・問い合わせ後の継続配信>
    move(sh[8], y=4.47)                                   # 見出し（14pt据え置き）
    move(sh[6], y=4.85, h=2.23)                           # 下段の枠（→7.08）
    move(sh[50], x=0.58, y=4.97, w=LO_COL_X + COL_W + 0.12 - 0.58, h=2.06)   # 薄緑の面

    inq = sh[52]
    cell(inq, 0.66, LO_ROW0, 0.95, BOX_H, BOX_PT, lr=0.0)              # 問い合わせ

    lrows = [57, 58, (59, 61), 60, 66, 67]
    lys = [LO_ROW0 + PITCH * i for i in range(6)]
    for r, y in zip(lrows, lys):
        if isinstance(r, tuple):
            cell(sh[r[0]], LO_COL_X, y, HALF_W, BOX_H, BOX_PT)
            cell(sh[r[1]], LO_COL_X + COL_W - HALF_W, y, HALF_W, BOX_H, BOX_PT)
        else:
            cell(sh[r], LO_COL_X, y, COL_W, BOX_H, BOX_PT)

    # 定期対応項目（一部抽出）
    lab2 = sh[25]
    cell(lab2, RB_X, LO_ROW0 - 0.01, 1.90, 0.28, LABEL_PT)
    set_color(lab2, 0, "595959")
    for i, sid in enumerate((20, 21, 23, 22, 24)):
        cell(sh[sid], RB_X, lys[1] + PITCH * i, COL_W, BOX_H, BOX_PT)

    # つなぎの線（下段）：問い合わせ → 1〜4行目
    inq_r, inq_mid = 0.66 + 0.95, LO_ROW0 + BOX_H / 2
    jx = (inq_r + LO_COL_X) / 2
    for cid, row in ((62, 0), (63, 1), (64, 2), (65, 3)):
        bent(sh[cid], inq_r, inq_mid, LO_COL_X, lys[row] + BOX_H / 2, jx=jx)

    # 物件の問い合わせ（上段右）→ 問い合わせ（下段左）の点線
    bent_rot(sh[53], 0.66 + 0.95 / 2, LO_ROW0, TGT_X + TGT_W / 2, ys[1] + BOX_H, 4.92)
