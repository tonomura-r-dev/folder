# -*- coding: utf-8 -*-
"""元S8「想定動線」（上段＝入口→友だち追加→配信の箱→問い合わせ、下段＝問い合わせ後の継続配信）。

文字：上の説明 12→14pt／フローの箱（上段・下段・定期対応項目）10→12pt／
      初期設計項目の箇条書き 9→12pt／見出しラベル（初期設計項目↓・定期対応項目(一部抽出)）9→12pt／
      「LINE公式アカウントの標準機能で運用」10→12pt。見出し「<対新規ユーザー領域>」等の14ptは据え置き。
場所：入口の3つ（100円賃貸のサイト／離脱防止ポップアップ／LINEでお問い合わせボタン）を縦に積み、
      ボタンは「サイトの」の後で改行して2行にする（横幅を空けるため）。
      空いた横幅で、配信の箱の列を 2.51→2.82in、行き先の箱を 1.48→1.50in に広げ、
      箱の高さを 0.24→0.29in（上段）・0.28in（下段）にする。上段の枠を少し縦に伸ばし、下段は詰める。
      箇条書きは行間を 14.5pt 固定にして、PowerPoint と LibreOffice で同じ高さにする。
      つなぎの線（カギ線）は新しい位置から計算し直す（つなぎ先は元のとおり残す）。
      薄い灰色の文字（箇条書き・見出しラベル）は 808080→595959 に濃くする。
"""
import copy

from _build.arauchi_v3.common import *  # noqa: F401,F403

E = 914400

LEAD_PT = 14
BOX_PT = 12          # フローの箱
LIST_PT = 12         # 初期設計項目の箇条書き
LIST_LINE_PT = 14.5  # 箇条書きの行の高さ（固定）
LABEL_PT = 12        # 初期設計項目↓／定期対応項目（一部抽出）
TAG_PT = 12          # LINE公式アカウントの標準機能で運用

# ---- 上段
F1_Y, F1_B = 1.92, 4.58             # 上段の枠
Y0, PITCH, BOX_H = 2.02, 0.355, 0.29
ENTRY_X, ENTRY_W = 0.62, 2.80       # 入口の列
MID_X, MID_W = 3.72, 1.66           # 友だち追加・初期設計項目
LIST_W = 1.72
COL_X, COL_W = 5.68, 2.82           # 配信の箱の列
HALF_W = (COL_W - 0.06) / 2         # 2つに分かれた行（リサーチ｜→配信に反映）
J1 = COL_X + COL_W + 0.17           # 6本の線が物件の問い合わせへ集まる縦の位置
J2 = COL_X + COL_W + 0.30           # 個別チャット → 空室確認・内見 の縦の位置
TGT_X, TGT_W = 8.98, 1.50           # 行き先（空室確認・内見／物件の問い合わせ）

# ---- 下段
H2_Y = 4.61
F2_Y, F2_B = 4.97, 7.07
LINK_Y = 5.03                       # 物件の問い合わせ → 問い合わせ の点線の横の位置
LY0, LPITCH, LBOX_H = 5.13, 0.315, 0.28
INQ_X, INQ_W = 0.70, 1.05
LO_COL_X = 3.95
RB_X, RB_W = 7.30, 2.82             # 定期対応項目の列


def cell(shape, x, y, w, h, pt, lr=0.03, tb=0.0, anchor="ctr"):
    """箱：大きさを固定し、位置・文字の大きさ・余白をそろえる"""
    set_autofit_none(shape)
    move(shape, x=x, y=y, w=w, h=h)
    set_size_all(shape, pt)
    set_insets(shape, l=lr, r=lr, t=tb, b=tb)
    set_anchor(shape, anchor)


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
    raise ValueError(text_mark)


def line_exact(shape, pt):
    """全段落の行の高さを pt に固定する（PowerPointとLibreOfficeで同じ高さにする）"""
    for p_el in txbody(shape).findall(qn("a:p")):
        ppr = p_el.find(qn("a:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("a:pPr"))
            p_el.insert(0, ppr)
        for old in ppr.findall(qn("a:lnSpc")):
            ppr.remove(old)
        ln = etree.Element(qn("a:lnSpc"))
        etree.SubElement(ln, qn("a:spcPts"), val=str(int(pt * 100)))
        ppr.insert(0, ln)


def _hook(cxn, st=None, end=None):
    """コネクタのつなぎ先を (図形id, 接続点) で指定し直す（None なら外す）。接続点：0上・1左・2下・3右。
    LibreOffice は、つなぎ先の無い縦長のカギ線を別の形に描いてしまうので、つなぎ先を必ず入れる。"""
    nv = cxn._element.find(qn("p:nvCxnSpPr")).find(qn("p:cNvCxnSpPr"))
    for tag in ("a:stCxn", "a:endCxn"):
        for el in nv.findall(qn(tag)):
            nv.remove(el)
    for tag, v in (("a:stCxn", st), ("a:endCxn", end)):
        if v is not None:
            etree.SubElement(nv, qn(tag), id=str(v[0]), idx=str(v[1]))


def _xfrm(cxn):
    return cxn._element.find(qn("p:spPr")).find(qn("a:xfrm"))


def _set_adj(cxn, frac):
    for gd in cxn._element.find(qn("p:spPr")).find(qn("a:prstGeom")).iter(qn("a:gd")):
        if gd.get("name") == "adj1":
            gd.set("fmla", "val %d" % int(round(frac * 100000)))


def bent(cxn, sx, sy, ex, ey, jx, st=None, end=None):
    """回転なしのカギ線（bentConnector3）：始点(sx,sy)→ 縦の位置 jx で折れて → 終点(ex,ey)。sx<ex のみ"""
    _hook(cxn, st, end)
    xf = _xfrm(cxn)
    for k in ("flipH", "flipV", "rot"):
        if k in xf.attrib:
            del xf.attrib[k]
    if ey < sy:
        xf.set("flipV", "1")
    cxn.left, cxn.top = Emu(int(sx * E)), Emu(int(min(sy, ey) * E))
    cxn.width, cxn.height = Emu(int((ex - sx) * E)), Emu(max(int(abs(ey - sy) * E), 1))
    _set_adj(cxn, (jx - sx) / (ex - sx))


def bent_up(cxn, sx, sy, ex, ey, yh, st=None, end=None):
    """縦→横→縦のカギ線：下の始点(sx,sy)から上へ→ yh で右へ→ 右上の終点(ex,ey)へ上がる。
    元の図形と同じ「90°回転＋左右上下反転」の bentConnector3 で表す（回転前の幅＝縦の差、高さ＝横の差）。"""
    _hook(cxn, st, end)
    xf = _xfrm(cxn)
    xf.set("rot", "5400000")
    xf.set("flipH", "1")
    xf.set("flipV", "1")
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
    move(lead, x=0.35, y=0.69, w=10.13, h=0.78)
    set_anchor(lead, "ctr")
    set_insets(lead, l=0.05, r=0.05)

    # ================================================================ 上段 <対新規ユーザー領域>
    move(sh[7], y=1.55)                                   # 見出し（14pt据え置き）
    cell(sh[69], 10.547 - 3.20, 1.585, 3.20, 0.30, TAG_PT, lr=0.05)   # LINE公式アカウントの標準機能で運用
    move(sh[5], y=F1_Y, h=F1_B - F1_Y)                    # 上段の枠
    ys = [Y0 + PITCH * i for i in range(7)]
    col_r = COL_X + COL_W
    move(sh[26], x=MID_X - 0.10, y=Y0 - 0.06, w=col_r + 0.08 - (MID_X - 0.10),
         h=ys[6] + BOX_H + 0.06 - (Y0 - 0.06))            # 薄緑の面

    # 入口（縦に積む）
    cell(sh[9], ENTRY_X, Y0, 1.80, BOX_H, BOX_PT, lr=0.04)            # 100円賃貸のサイト
    pop = sh[17]
    pop_y, pop_h = 2.37, 0.32
    cell(pop, ENTRY_X, pop_y, ENTRY_W, pop_h, BOX_PT)                  # 離脱防止ポップアップ
    btn = sh[15]
    btn_y, btn_h = 2.78, 0.54
    break_after(btn, "サイトの")                                        # サイトの／「LINEでお問い合わせ」ボタン
    cell(btn, ENTRY_X, btn_y, ENTRY_W, btn_h, BOX_PT)
    set_line_spacing(btn, 100)

    # 友だち追加・初期設計項目
    add = sh[27]
    cell(add, MID_X, Y0, MID_W, 0.30, BOX_PT)
    lab = sh[34]
    cell(lab, MID_X, 2.39, MID_W, 0.26, LABEL_PT)
    set_color(lab, 0, "595959")
    lst = sh[35]
    cell(lst, MID_X, 2.67, LIST_W, 1.70, LIST_PT, lr=0.05, tb=0.02, anchor="t")
    set_insets(lst, r=0.02)
    line_exact(lst, LIST_LINE_PT)
    for i in range(len(lst.text_frame.paragraphs)):
        set_color(lst, i, "595959")

    # 中かっこ
    move(sh[49], x=COL_X - 0.20, y=Y0 - 0.03, w=0.16, h=6 * PITCH + BOX_H + 0.06)

    # 配信の箱（7行）
    rows = [31, 32, (33, 39), 36, 38, 37, 40]
    for r, y in zip(rows, ys):
        if isinstance(r, tuple):
            cell(sh[r[0]], COL_X, y, HALF_W, BOX_H, BOX_PT)
            cell(sh[r[1]], col_r - HALF_W, y, HALF_W, BOX_H, BOX_PT)
        else:
            cell(sh[r], COL_X, y, COL_W, BOX_H, BOX_PT)

    # 行き先
    cell(sh[19], TGT_X, ys[0], TGT_W, BOX_H, BOX_PT, lr=0.02)          # 空室確認・内見
    cell(sh[14], TGT_X, ys[1], TGT_W, BOX_H, BOX_PT, lr=0.02)          # 物件の問い合わせ

    # つなぎの線（上段）
    add_mid = Y0 + 0.15
    jx = ENTRY_X + ENTRY_W + 0.13
    bent(sh[30], ENTRY_X + ENTRY_W, pop_y + pop_h / 2, MID_X, add_mid, jx, st=(17, 3), end=(27, 1))  # ポップアップ → 友だち追加
    bent(sh[28], ENTRY_X + ENTRY_W, btn_y + btn_h / 2, MID_X, add_mid, jx, st=(15, 3), end=(27, 1))  # ボタン → 友だち追加
    q_mid = ys[1] + BOX_H / 2                                           # 物件の問い合わせの左辺の中央
    for cid, src, row in ((48, 31, 0), (47, 32, 1), (46, 39, 2), (44, 36, 3), (43, 38, 4), (42, 37, 5)):
        bent(sh[cid], col_r, ys[row] + BOX_H / 2, TGT_X, q_mid, J1, st=(src, 3))
    bent(sh[41], col_r, ys[6] + BOX_H / 2, TGT_X, ys[0] + BOX_H / 2, J2, st=(40, 3))  # 個別チャット → 空室確認・内見

    # ================================================================ 下段 <友だち追加後・問い合わせ後の継続配信>
    move(sh[8], y=H2_Y)                                   # 見出し（14pt据え置き）
    move(sh[6], y=F2_Y, h=F2_B - F2_Y)                    # 下段の枠
    lys = [LY0 + LPITCH * i for i in range(6)]
    move(sh[50], x=0.58, y=LY0 - 0.06, w=LO_COL_X + COL_W + 0.12 - 0.58,
         h=lys[5] + LBOX_H + 0.05 - (LY0 - 0.06))         # 薄緑の面

    inq = sh[52]
    cell(inq, INQ_X, LY0, INQ_W, LBOX_H, BOX_PT, lr=0.0)               # 問い合わせ

    lrows = [57, 58, (59, 61), 60, 66, 67]
    for r, y in zip(lrows, lys):
        if isinstance(r, tuple):
            cell(sh[r[0]], LO_COL_X, y, HALF_W, LBOX_H, BOX_PT)
            cell(sh[r[1]], LO_COL_X + COL_W - HALF_W, y, HALF_W, LBOX_H, BOX_PT)
        else:
            cell(sh[r], LO_COL_X, y, COL_W, LBOX_H, BOX_PT)

    # 定期対応項目（一部抽出）
    lab2 = sh[25]
    cell(lab2, RB_X, LY0, 2.10, LBOX_H, LABEL_PT)
    set_color(lab2, 0, "595959")
    for i, sid in enumerate((20, 21, 23, 22, 24)):
        cell(sh[sid], RB_X, lys[1 + i], RB_W, LBOX_H, BOX_PT)

    # つなぎの線（下段）：問い合わせ → 1〜4行目
    inq_r, inq_mid = INQ_X + INQ_W, LY0 + LBOX_H / 2
    jx2 = (inq_r + LO_COL_X) / 2
    for cid, dst, row in ((62, 57, 0), (63, 58, 1), (64, 59, 2), (65, 60, 3)):
        bent(sh[cid], inq_r, inq_mid, LO_COL_X, lys[row] + LBOX_H / 2, jx2, st=(52, 3), end=(dst, 1))

    # 物件の問い合わせ（上段右）→ 問い合わせ（下段左）の点線
    bent_up(sh[53], INQ_X + INQ_W / 2, LY0, TGT_X + TGT_W / 2, ys[1] + BOX_H, LINK_Y, st=(52, 0))
