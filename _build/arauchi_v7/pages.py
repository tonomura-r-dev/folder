# -*- coding: utf-8 -*-
"""あらうち提案資料 v7（2026-10-13 殿村さん指示）の直し。元＝v6。直すのは次のページだけ。

  S1  表紙 … 元のDYM資料の表紙（サブタイトルを小さな灰色の1行でタイトルの上に置く形）に戻す。文言は変えない
  S3  想定動線 … 「物件情報の送信（CV）」をゴールに強調。CV以降は「あらうち様による個別対応（参考）」に簡略化
  S8  想定動線 … 運用フローの終点を「物件情報の送信（CV）」に。CV以降は参考表示（灰色）にして計測対象と分ける
  S15 問い合わせ対応の効率化 … あらうち様側を「物件確認・内見等の個別対応」に整理。契約まで対応するとは書かない
  費用4枚（v6のS23〜S26）をS13の直後へ移す → 並べ替えは build.py。ここでは、移動後のページ番号に合わせて
  「26ページ」→「17ページ」に直すだけ（S24・S25）

※ページ番号は v6 のもの（S3＝v6の3枚目）。ほかのページ・料金・数値は触らない。
"""
import copy

from lxml import etree

from _build.arauchi_v4.lib import NAVY, RED, add_box, delete, emu, ids, make_para, move, qn, set_paras

GREY_TXT = "595959"     # 資料で使っている濃い灰色の文字（S8の「定期対応項目」など）
GREY_FILL = "F2F2F2"
GREY_LINE = "A6A6A6"
GREY_BOX = "7F7F7F"     # 資料で使っている灰色の箱（S8の「定期対応項目」の中身）


# ---------------------------------------------------------------- 小さな部品
def sppr(shape):
    return shape._element.find(qn("p:spPr"))


def set_fill(shape, hex_=None, alpha=None):
    """図形の塗り。hex_=None で塗りなし"""
    sp = sppr(shape)
    for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        for e in sp.findall(qn(tag)):
            sp.remove(e)
    if hex_ is None:
        new = etree.Element(qn("a:noFill"))
    else:
        new = etree.Element(qn("a:solidFill"))
        c = etree.SubElement(new, qn("a:srgbClr"), val=hex_)
        if alpha:
            etree.SubElement(c, qn("a:alpha"), val=str(alpha))
    geom = sp.find(qn("a:prstGeom"))
    if geom is None:
        geom = sp.find(qn("a:custGeom"))
    geom.addnext(new)


def set_outline(shape, hex_=None, w_pt=0.75, dash=None):
    """図形の枠線。hex_=None で枠線なし"""
    sp = sppr(shape)
    for e in sp.findall(qn("a:ln")):
        sp.remove(e)
    ln = etree.Element(qn("a:ln"))
    if hex_ is None:
        ln.set("w", "0")
        etree.SubElement(ln, qn("a:noFill"))
    else:
        ln.set("w", str(int(w_pt * 12700)))
        f = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(f, qn("a:srgbClr"), val=hex_)
        if dash:
            etree.SubElement(ln, qn("a:prstDash"), val=dash)
    # a:ln は塗りの後ろ（効果などの前）
    after = None
    for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:prstGeom", "a:custGeom"):
        found = sp.find(qn(tag))
        if found is not None:
            after = found
            break
    after.addnext(ln)


def set_text_color(shape, hex_, paras=None):
    """図形の中の文字色（paras＝段落の番号の並び。None なら全部）"""
    tb = shape.text_frame._txBody
    for i, p_el in enumerate(tb.findall(qn("a:p"))):
        if paras is not None and i not in paras:
            continue
        for rpr in list(p_el.iter(qn("a:rPr"))) + list(p_el.iter(qn("a:endParaRPr"))):
            for f in rpr.findall(qn("a:solidFill")):
                rpr.remove(f)
            fill = etree.Element(qn("a:solidFill"))
            etree.SubElement(fill, qn("a:srgbClr"), val=hex_)
            ln = rpr.find(qn("a:ln"))
            if ln is not None:
                ln.addnext(fill)
            else:
                rpr.insert(0, fill)


def set_bold(shape, bold, paras=None):
    tb = shape.text_frame._txBody
    for i, p_el in enumerate(tb.findall(qn("a:p"))):
        if paras is not None and i not in paras:
            continue
        for rpr in p_el.iter(qn("a:rPr")):
            rpr.set("b", "1" if bold else "0")


def set_size(shape, pt, paras=None):
    tb = shape.text_frame._txBody
    for i, p_el in enumerate(tb.findall(qn("a:p"))):
        if paras is not None and i not in paras:
            continue
        for rpr in list(p_el.iter(qn("a:rPr"))) + list(p_el.iter(qn("a:endParaRPr"))):
            rpr.set("sz", str(int(round(pt * 100))))


def set_line_color(connector, hex_):
    """コネクタ（線）の色だけ変える"""
    ln = connector._element.find(qn("p:spPr")).find(qn("a:ln"))
    for f in ln.findall(qn("a:solidFill")):
        ln.remove(f)
    fill = etree.Element(qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr"), val=hex_)
    ln.insert(0, fill)


def set_body(shape, anchor=None, wrap=None, autofit=None, ins=None):
    bp = shape.text_frame._txBody.find(qn("a:bodyPr"))
    if anchor:
        bp.set("anchor", anchor)
    if wrap:
        bp.set("wrap", wrap)
    if ins:
        for k, v in zip(("lIns", "tIns", "rIns", "bIns"), ins):
            bp.set(k, str(emu(v)))
    if autofit:
        for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
            for e in bp.findall(qn(tag)):
                bp.remove(e)
        etree.SubElement(bp, qn(f"a:{autofit}"))


def replace_text_in_runs(slide, old, new):
    """文字の一部を差し替える（run の中に収まっているものだけ）。差し替えた数を返す"""
    n = 0
    for t in slide._element.iter(qn("a:t")):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            n += 1
    return n


# ---------------------------------------------------------------- S1 表紙
def p01(s):
    """元のDYM資料の表紙：宛名（左上）／サブタイトル（小さな灰色の文字でタイトルの上）／タイトル（大きな太字）／線。
    文言は変えない。v6で付けた薄い青の帯・紺の太字・赤字は外す。"""
    sh = ids(s)
    title, sub, rule, addr = sh[26], sh[27], sh[28], sh[31]

    # 位置は元の資料（殿村さんがPCで保存した47枚の1枚目）と同じ：タイトル y=2.770in／線 y=3.500in
    title.top = 2533105
    rule.top = 3199966
    rule.height = 0

    # サブタイトル：帯をやめて、宛名と同じ灰色（ただし太字にしない）16ptでタイトルの上へ。2行（読点で切る）。
    set_fill(sub, None)
    set_outline(sub, None)
    move(sub, x=1.462, y=1.66, w=7.909, h=1.14)       # 下端＝元のサブタイトルの下端（2.80in）
    set_body(sub, anchor="b", wrap="square", autofit="noAutofit", ins=(0.1, 0.05, 0.1, 0.05))
    base_rpr = addr.text_frame._txBody.find(".//" + qn("a:rPr"))
    tb = sub.text_frame._txBody
    for p_el in tb.findall(qn("a:p")):
        tb.remove(p_el)
    for line in ("LINEへの流入と友だち追加後の導線を整備し、", "物件問い合わせの増加を実現する"):
        p = etree.SubElement(tb, qn("a:p"))
        ppr = etree.SubElement(p, qn("a:pPr"), algn="ctr")
        r = etree.SubElement(p, qn("a:r"))
        rpr = copy.deepcopy(base_rpr)
        rpr.set("b", "0")
        r.append(rpr)
        t = etree.SubElement(r, qn("a:t"))
        t.text = line
        end = copy.deepcopy(rpr)
        end.tag = qn("a:endParaRPr")
        p.append(end)


# ---------------------------------------------------------------- S3 想定動線
def p03(s):
    sh = ids(s)

    # 上の要旨：「→ 問い合わせ」をやめて、CV（物件情報の送信）を終点に
    set_paras(sh[117], [
        "ポータルサイトで物件検索 → サイト閲覧 → LINE追加 → 物件情報の送信（CV）の流れで設計します。",
        "最大の改善ポイントは、LINE友だち追加後に物件情報の送信（CV）へつなげる導線の整備です。",
    ])

    # カスタマージャーニーマップ：⑤＝CVをゴール、⑥〜⑧は「あらうち様による個別対応（参考）」として薄く
    set_paras(sh[93], ["⑤ リード有効化（CV）"])
    for i in (95, 94, 96):                         # ⑥ 有効リードの再育成／⑦ マネタイズ／⑧ LTV最大化・紹介
        set_fill(sh[i], "BFBFBF", alpha=30000)    # ①〜④と同じ「薄い色」のまま、灰色に
        set_text_color(sh[i], GREY_TXT)
    add_box(s, 8.70, 2.08, 1.98, 0.56,
            ["⑥〜⑧：あらうち様による", "個別対応（参考）"],
            size=11, bold=True, color=GREY_TXT, align="l", anchor="t", ins=(0.02, 0.02, 0.0, 0.02),
            name="⑥〜⑧の注記")

    # ユーザー動線：④＝CV（赤）。CV以降はひとつの灰色の枠に簡略化
    set_paras(sh[112], ["④物件情報の送信（CV）"])
    set_fill(sh[112], RED)
    set_text_color(sh[112], "FFFFFF")
    set_size(sh[112], 12)                          # 「④物件情報の送信（CV）」を1行に収める

    ref = sh[106]                                 # もとは「問い合わせ（空室確認・内見の手配）」の紺の箱
    set_paras(ref, [
        [("あらうち様による個別対応（参考・計測対象外）", True)],
        [("空室確認・内見の手配／お申込み／ご契約／鍵のお渡し・お引っ越し など", False)],
    ])
    set_fill(ref, "FFFFFF")
    set_outline(ref, GREY_LINE, 1.0, "dash")
    set_text_color(ref, GREY_TXT)
    set_size(ref, 14, paras=[0])
    set_size(ref, 12, paras=[1])
    move(ref, x=3.874, y=4.30, w=6.60, h=0.80)    # 中心の高さ（4.70in）は変えない＝つなぐ線の終点のまま
    for i in (118, 107, 115):                      # お申込み／ご契約（赤）／鍵のお渡し・お引っ越し
        delete(sh[i])
    delete(sh[119])                                # お申込み→ご契約の線

    add_box(s, 3.874, 6.30, 6.60, 0.44,
            ["◀ ゴール：当社の計測・評価はここまで"],
            size=13, bold=True, color=RED, align="l", anchor="ctr", ins=(0.04, 0.02, 0.04, 0.02),
            name="CVの注記")


# ---------------------------------------------------------------- S8 想定動線（全体）
def p08(s):
    sh = ids(s)

    set_paras(sh[232], [
        "離脱防止ポップアップとサイトのLINEボタンから友だち追加し、物件情報の送信（CV）につなげます。",
        "LINE公式アカウントの標準機能で運用（離脱防止ポップアップは別途導入）",
    ])

    # 終点＝物件情報の送信（CV）。赤い箱は中心の高さ（2.52in）を変えない＝6本の線の終点のまま
    red = sh[185]
    set_paras(red, ["物件情報の送信", "（CV）"])
    move(red, y=2.24, h=0.56)
    add_box(s, 8.88, 2.83, 1.70, 0.44,
            ["運用フローの終点", "（当社の計測対象）"],
            size=11, bold=True, color=RED, align="ctr", anchor="ctr", ins=(0, 0, 0, 0), name="CVの注記")

    # CV以降（空室確認・内見）は参考：薄い灰色の枠
    ref = sh[188]
    set_paras(ref, [[("空室確認・内見", True)], [("あらうち様による", False)], [("個別対応（参考）", False)]])
    set_fill(ref, GREY_FILL)
    set_outline(ref, GREY_LINE, 1.0, "dash")
    set_text_color(ref, GREY_TXT)
    move(ref, y=3.32, h=0.90)
    # 個別チャット→空室確認・内見の線：終点を新しい位置へ
    cx = sh[209]._element
    xf = cx.find(".//" + qn("a:xfrm"))
    off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    bottom = int(off.get("y")) + int(ext.get("cy"))        # flipV の線：始点（下）の高さ
    new_end = emu(3.77)
    off.set("y", str(new_end))
    ext.set("cy", str(bottom - new_end))

    # 下の段（問い合わせ後）＝参考。見出し・色を灰色に。長い点線（上の終点とつないでいた線）は文字の上を通るので外す
    delete(sh[219])
    set_paras(sh[183], ["<参考：CV以降の個別対応（あらうち様）／当社の計測対象外>"])
    move(sh[183], w=6.8)
    set_fill(sh[217], GREY_FILL)
    for i in (220, 221, 222, 223, 224, 229, 230):
        set_fill(sh[i], GREY_BOX)
    for i in (225, 226, 227, 228):
        set_line_color(sh[i], GREY_BOX)


# ---------------------------------------------------------------- S15 LINEによる問い合わせ対応の効率化
def p15(s):
    sh = ids(s)
    set_paras(sh[474], ["あらうち様｜物件確認・内見等の個別対応"])
    set_paras(sh[475], [
        "・個別チャットでの返信",
        "・物件の空室確認",
        "・内見の手配",
        "・その他、お客様ごとの個別対応",
    ])
    add_box(s, 5.495, 5.19, 4.725, 0.32,
            ["※物件情報の送信（CV）以降は、当社の計測対象外です"],
            size=12, bold=False, color="444444", align="l", anchor="ctr", ins=(0.04, 0.02, 0.04, 0.02),
            name="計測範囲の注記")


# ---------------------------------------------------------------- 費用4枚の移動に伴うページ参照
def refs(prs):
    """プラン④のページが26→17になる（v6のS24・S25の中の「26ページ」）。13ページ（費用対効果）・6ページ（出典）は動かない"""
    n = 0
    for no in (24, 25):
        n += replace_text_in_runs(prs.slides[no - 1], "26ページ", "17ページ")
    return n


def apply(prs):
    p01(prs.slides[0])
    p03(prs.slides[2])
    p08(prs.slides[7])
    p15(prs.slides[14])
    n = refs(prs)
    assert n == 4, n          # S24が1か所、S25が3か所
