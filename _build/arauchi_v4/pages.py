# -*- coding: utf-8 -*-
"""あらうち提案資料 v4（2026-10-13）の各ページの直し。ページ番号は v3（46枚）の番号。

  p03 想定動線 ⑤リード有効化を強調 ／ p05 市場分析 訴求軸＋出典 ／ p06 市場分析 出典
  p07 ツールの選定 ／ p09 想定セグメント×成果（SIMと照合） ／ p11 課題 ／ p12 方針 ／ p13 費用対効果
  p24 運用プラン（3プラン比較に作り替え）
数字の正本は 株式会社あらうち御中_LINEOA施策提案【SIM】ver1.4.xlsx（2026-10-13 版）。
"""
import copy

from lxml import etree

from _build.arauchi_v4.lib import *  # noqa: F401,F403
from _build.arauchi_v4.lib import (NAVY, RED, TEXT, add_box, add_table, ids, make_para, move, qn,
                                   set_anchor, set_autofit_none, set_insets, set_paras, set_size_all,
                                   set_spacing, txbody, delete)

URL_YONDAI = "https://www.athome.co.jp/corporate/news/data/questionnaire/yondai-news-202512/"
URL_ONLINE = "https://www.athome.co.jp/corporate/news/data/questionnaire/online-chintai-202510/"
URL_100EN = "https://100en.net/"
URL_TADASUMU = "https://tadasumu.com/"
URL_TADASUMU_NAVI = "https://tadasumu.com/navi/tadasumu-toha/"


# ---------------------------------------------------------------- 小さな部品
def set_fill(shape, hex_):
    sp = shape._element.spPr
    for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:pattFill", "a:blipFill"):
        for e in sp.findall(qn(tag)):
            sp.remove(e)
    f = etree.Element(qn("a:solidFill"))
    etree.SubElement(f, qn("a:srgbClr"), val=hex_)
    geom = sp.find(qn("a:prstGeom"))
    if geom is None:
        geom = sp.find(qn("a:custGeom"))
    geom.addnext(f)


def set_outline(shape, hex_, w_pt):
    sp = shape._element.spPr
    for e in sp.findall(qn("a:ln")):
        sp.remove(e)
    ln = etree.Element(qn("a:ln"), w=str(int(w_pt * 12700)))
    f = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(f, qn("a:srgbClr"), val=hex_)
    anchor = None
    for tag in ("a:solidFill", "a:noFill", "a:prstGeom", "a:custGeom"):
        found = sp.findall(qn(tag))
        if found:
            anchor = found[-1]
            break
    anchor.addnext(ln)


def set_text_color(shape, hex_):
    for rpr in list(txbody(shape).iter(qn("a:rPr"))) + list(txbody(shape).iter(qn("a:endParaRPr"))):
        for f in rpr.findall(qn("a:solidFill")):
            rpr.remove(f)
        fill = etree.Element(qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr"), val=hex_)
        ln = rpr.find(qn("a:ln"))
        (ln.addnext(fill) if ln is not None else rpr.insert(0, fill))


def set_run_sizes(shape, para_idx, sizes):
    """段落 para_idx の run ごとの文字サイズ（pt）。足りない分は最後の値"""
    p_el = txbody(shape).findall(qn("a:p"))[para_idx]
    for i, r in enumerate(p_el.findall(qn("a:r"))):
        r.find(qn("a:rPr")).set("sz", str(int(sizes[min(i, len(sizes) - 1)] * 100)))


def add_break_in_para(shape, para_idx, first_text):
    """段落 para_idx（run が1つ）の文字を first_text の後ろで改行する"""
    p_el = txbody(shape).findall(qn("a:p"))[para_idx]
    r = p_el.findall(qn("a:r"))[0]
    t = r.find(qn("a:t"))
    full = t.text
    assert full.startswith(first_text), (full, first_text)
    t.text = first_text
    br = etree.Element(qn("a:br"))
    br.append(copy.deepcopy(r.find(qn("a:rPr"))))
    r.addnext(br)
    r2 = copy.deepcopy(r)
    r2.find(qn("a:t")).text = full[len(first_text):]
    br.addnext(r2)


def clone_para(shape, idx):
    """段落 idx を複製して最後に足す"""
    tb = txbody(shape)
    ps = tb.findall(qn("a:p"))
    new = copy.deepcopy(ps[idx])
    ps[-1].addnext(new)
    return new


def replace_all_paras(shape, p_elements):
    tb = txbody(shape)
    for p in tb.findall(qn("a:p")):
        tb.remove(p)
    for p in p_elements:
        tb.append(p)


def source_paras(slide, lines, head="出典", note=None, size=11):
    """出典ブロックの段落。lines＝[[(文字, {link:URL}), ...], ...]"""
    out = [make_para([(head, {})], size, bold=True, color=NAVY, after=3, slide=slide)]
    for runs in lines:
        out.append(make_para([("・", {})] + runs, size, color=TEXT, after=2, slide=slide))
    if note:
        out.append(make_para([(note, {})], size - 1, color="595959", before=4, slide=slide))
    return out


# ---------------------------------------------------------------- p03 想定動線
def p03(s):
    sh = ids(s)
    # 要旨：2行目を「最大の改善ポイント」に（課題仮説として。離脱地点は未計測）
    set_paras(sh[31], [
        [("ポータルサイトで物件検索 → サイト閲覧 → LINE追加 → 物件情報の送信 → 問い合わせ、の流れに沿って設計します。", None)],
        [("最大の改善ポイントは、LINE友だち追加後に物件の問い合わせへつなげる導線の整備", True), ("です。", False)],
    ])
    add_break_in_para(sh[31], 0, "ポータルサイトで物件検索 → サイト閲覧 → LINE追加 → ")
    # ⑤ リード有効化：赤地・白の太字で強調（8フェーズの並びは変えない）
    b5 = sh[6]
    set_fill(b5, RED)
    set_text_color(b5, "FFFFFF")
    set_size_all(b5, 14)
    # ⑤の右に「最重要課題（仮説）」
    add_box(s, 6.52, 2.12, 2.3, 0.3, [[("◀ 最重要課題（仮説）", {})]], size=13, bold=True, color=RED,
            align="l", anchor="ctr", ins=(0.04, 0.0, 0.04, 0.0), name="最重要課題チップ")
    # 下の動線図：友だち追加後〜問い合わせ（③④）を赤い点線で囲む
    frame = add_box(s, 0.33, 5.84, 2.32, 0.94, [[("", {})]], size=11, line=RED, line_w=1.75, dash="dash",
                    name="⑤の範囲")
    add_box(s, 1.56, 5.70, 1.09, 0.26, [[("⑤ 最重要課題", {})]], size=11, bold=True, color="FFFFFF", fill=RED,
            align="ctr", anchor="ctr", ins=(0.02, 0.0, 0.02, 0.0), name="⑤タブ")


def market_geometry(s):
    """p05・p06 共通：カードの本文を縦に広げ、出典の枠を下げる（メイリオは行間が広いので余裕を見る）"""
    sh = ids(s)
    for num_id, title_id, body_id in ((33, 34, 35), (36, 37, 38), (39, 40, 41)):
        move(sh[num_id], y=1.62, h=0.56)
        move(sh[title_id], y=2.18, h=0.50)
        move(sh[body_id], y=2.68, h=3.00)
    move(sh[42], y=5.76, h=1.28)


# ---------------------------------------------------------------- p05 市場分析（訴求軸）
def p05(s):
    market_geometry(s)
    sh = ids(s)
    set_paras(sh[32], [
        [("家賃が上がる今、仲介手数料・初期費用への負担感に応えることを、提案の軸にします。", None)],
        [("物件はポータルサイトで自分で探し、問い合わせ先は別に選ぶ流れが主流", True), ("です。", False)],
    ])
    # 01：家賃の上昇は「市場背景」。市場データと提案上の仮説を分けて書く
    set_paras(sh[34], ["家賃の上昇（市場背景）"])
    set_paras(sh[35], [
        [("【市場データ】", True), ("東京23区のシングル向き（30㎡以下）マンションの平均家賃は、2025年5月に初めて10万円を超えた", False)],
        [("【提案上の仮説】", True), ("家賃の上昇で、仲介手数料・初期費用への負担感が高まると考えられる", False)],
    ])
    set_spacing(sh[35], after=6)
    # 03：100円賃貸の強み＝仲介手数料を抑えられる点
    set_paras(sh[41], [
        [("・問い合わせる不動産会社を選ぶ基準は「取り扱っている物件数が多い」が1位（経験者26.6％・検討者47.7％）", False)],
        [("・100円賃貸は、お客様が見つけた物件の仲介に対応し、仲介手数料を抑えられる（100円賃貸が対応できる物件に限ります）", False)],
    ])
    set_spacing(sh[41], after=6)
    # 出典：調査主体・調査名を短く。URLは出典名のリンクに
    src = sh[42]
    replace_all_paras(src, source_paras(s, [
        [("アットホーム「2025年の賃貸市場における4大ニュース」", {"link": URL_YONDAI})],
        [("アットホーム「オンラインでの住まい探しに関する調査 2025 賃貸編」", {"link": URL_ONLINE})],
        [("100円賃貸 公式サイト", {"link": URL_100EN})],
    ]))


# ---------------------------------------------------------------- p06 市場分析（出典）
def p06(s):
    market_geometry(s)
    sh = ids(s)
    src = sh[42]
    replace_all_paras(src, source_paras(s, [
        [("アットホーム「オンラインでの住まい探しに関する調査 2025 賃貸編」", {"link": URL_ONLINE})],
        [("タダスム（", {}), ("公式サイト", {"link": URL_TADASUMU}), ("／", {}), ("サービス紹介", {"link": URL_TADASUMU_NAVI}), ("）", {})],
        [("100円賃貸 公式サイト", {"link": URL_100EN})],
    ]))


# ---------------------------------------------------------------- p07 ツールの選定
def p07(s):
    sh = ids(s)
    # 左右の見出し
    set_paras(sh[7], [[("標準機能で運用", True), (" ", False), ("※基本施策はこちらで運用", False)]])
    set_run_sizes(sh[7], 0, [14, 14, 12])
    set_paras(sh[11], [[("外部の運用ツール", True), (" ", False), ("※基本施策には含まない", False)]])
    set_run_sizes(sh[11], 0, [14, 14, 12])
    # 右：Lステップ等は本文へ。「取り扱い」を追加
    rb = sh[13]
    set_paras(rb, [
        [("特徴", None)],
        [("タグや流入経路を細かく管理でき、外部データ連携に強い（Lステップ等）", None)],
        [("メリット", None)],
        [("流入元ごとに、配信を細かく出し分けられる", None)],
        [("デメリット", None)],
        [("月額のツール費と、初期設定・移行の工数がかかる", None)],
    ])
    # 枠の高さを詰めて、結論を広げる
    move(sh[9], h=3.72)
    move(rb, h=3.72)
    move(sh[15], y=5.72, h=1.38)
    set_paras(sh[15], [
        [("結論", None)],
        [("基本施策は、LINE公式アカウントの標準機能を中心に運用します。", None)],
        [("プラン①では、必要に応じてAPI連携・開発を別途お見積りでご提案します（基本料金に含まず）。", None)],
        [("※離脱防止ポップアップは、LINE公式アカウントの機能ではなく、サイト側に別途導入する施策です。", None)],
    ])


# ---------------------------------------------------------------- p09 想定セグメント×成果
# SIM（2026-10-13版）：物件を送る方の割合＝問い合わせ①÷新しい友だち（5÷22.4＝22%→13÷23.0＝57%）
#                     CVR＝問い合わせ②÷配信のクリック（SIM41行：0.34%→2.0%）
def p09(s):
    sh = ids(s)
    set_paras(sh[41], [
        [("22→57", True), ("%", True)],
        [("友だち追加直後に", False)],
        [("物件を送る方（仮）", False)],
    ])
    set_run_sizes(sh[41], 0, [20, 14])
    set_paras(sh[42], [
        [("0.3→2.0", True), ("%", True)],
        [("配信クリックからの", False)],
        [("問い合わせ（CVR）", False)],
    ])
    set_run_sizes(sh[42], 0, [20, 14])
    # 指標の定義は上の要旨に（物件を送る方の割合・CVRを混同しないため）
    lead = sh[68]
    set_paras(lead, [
        [("新しい友だちには物件の送り方を、今の友だちには配信とリッチメニューで問い合わせのきっかけを届けます。", None)],
        [("「どの物件が対象か」が分からないこと", True), ("がハードルになっていると考え、100円チェッカーを前面に出します。", False)],
        [("※試算の定義：物件を送る方の割合＝問い合わせ①÷新しい友だち／CVR＝問い合わせ②÷配信のクリック数", False)],
    ])
    set_run_sizes(lead, 2, [11])
    set_color(lead, 2, "595959")
    set_autofit_none(lead)
    set_anchor(lead, "ctr")
    move(lead, y=0.66, h=0.8)


# ---------------------------------------------------------------- p11 課題
def p11(s):
    sh = ids(s)
    # 上に「どの課題が最重要か」の区分（1行ずつ）
    add_box(s, 0.42, 0.72, 3.19, 0.42, [[("新規友だち獲得（②サイト離脱対策）", {})]],
            size=12, bold=True, color=NAVY, fill="FFFFFF", line=NAVY, line_w=1.25, align="ctr", anchor="ctr",
            ins=(0.03, 0.02, 0.03, 0.02), name="課題の区分：新規友だち獲得")
    add_box(s, 3.82, 0.72, 6.59, 0.42, [[("最重要課題：LINE友だち追加後の問い合わせ転換（⑤リード有効化）", {})]],
            size=14, bold=True, color="FFFFFF", fill=RED, align="ctr", anchor="ctr",
            ins=(0.05, 0.02, 0.05, 0.02), name="課題の区分：最重要（リード有効化）")
    # カードを少し上げ、本文の高さを確保（メイリオは行間が広い）
    for num_id, title_id, body_id in ((6, 10, 14), (17, 19, 21), (23, 25, 27)):
        move(sh[num_id], y=1.22, h=0.58)
        move(sh[title_id], y=1.80, h=0.48)
        move(sh[body_id], y=2.28, h=3.40)
    move(sh[15], y=5.74, h=1.36)
    set_paras(sh[15], [
        [("結論", None)],
        [("最重要課題は、友だち追加後の問い合わせ獲得（⑤リード有効化）です。あいさつ・リッチメニュー・ステップ配信で「物件を送る」まで案内します。", None)],
        [("サイト離脱対策は、離脱防止ポップアップ・LINEボタンで新しい友だちを増やす課題です。", None)],
    ])


# ---------------------------------------------------------------- p12 方針
def p12(s):
    sh = ids(s)
    set_paras(sh[34], ["LINE経由の問い合わせ 1か月目 月7件→6か月目 月23件（試算）"])


# ---------------------------------------------------------------- p13 費用対効果
def p13(s):
    sh = ids(s)
    set_paras(sh[88], [
        [("プラン①｜総合運用：初期21.5万円・月額13.5万円で運用します。", None)],
        [("6か月目にLINE経由の問い合わせ月23件を見込みます（試算）", True), ("。", False)],
    ])
    set_paras(sh[37], ["月次運用範囲内で適宜調整"])
    set_insets(sh[37], l=0.04, r=0.04)
    move(sh[33], h=0.53)
    move(sh[37], h=0.53)
    set_paras(sh[11], ["約1,377人"])
    set_paras(sh[16], ["13件/月"])
    set_paras(sh[23], ["10件/月"])
    set_paras(sh[27], ["23件/月"])
    set_paras(sh[31], ["（6か月累計90件）"])
    set_paras(sh[87], [
        [("※初期＝コンサル20万円＋離脱防止1.5万円／月額＝コンサル10万円＋離脱防止3万円＋LINE公式アカウント0.5万円", None)],
        [("※6か月後は6か月目の月間の試算。問い合わせ①＝友だち追加直後の物件の問い合わせ、②＝ステップ配信・企画配信などからの問い合わせ", None)],
    ])
    # 試算の前提（右下の空き）
    add_box(s, 7.45, 4.85, 3.1, 1.65,
            [[("＜試算の前提＞", {"bold": True, "color": NAVY})],
             [("プラン①｜総合運用の基本施策を\n対象とした試算です。", {})],
             [("API連携・追加開発の効果は\n含みません。", {})]],
            size=12, color=TEXT, fill="F4F7FF", line="D9D9D9", line_w=0.75, align="l", anchor="t",
            ins=(0.12, 0.08, 0.12, 0.06), after=3, name="試算の前提")


# ---------------------------------------------------------------- p24 運用プラン（3プラン比較）
HEAD = ["項目", "プラン①｜総合運用", "プラン②｜獲得重視", "プラン③｜初期導入"]
ROWS = [
    ("初期費用", "21.5万円", "21.5万円", "21.5万円"),
    ("月額コンサル費", "10万円", "5万円", "なし"),
    ("月額ツール・LINE費", "3.5万円", "3.5万円", "別途実費"),
    ("月額合計", "13.5万円", "8.5万円", "運用契約なし"),
    ("企画配信", "月4〜5本", "月1〜2本", "なし"),
    ("ステップ配信", "10通", "なし", "なし"),
    ("離脱防止", "導入・運用", "導入・運用", "初期導入支援"),
    ("継続改善", "対応", "限定的", "なし"),
    ("API連携・開発", "必要に応じ別途見積", "なし", "なし"),
]
NOTES = [
    "※価格はあらうち様向けの個別提案価格です（DYMの標準価格ではありません）。初期費用21.5万円は3プラン共通です。",
    "※プラン③は月額コンサル契約がありません。ツールを継続して利用する場合の利用料は別途実費です。",
    "※API連携・開発は基本料金に含まれません。必要に応じて、プラン①で別途お見積りします。",
    "※13ページの試算（6か月目に月23件）は、プラン①｜総合運用の基本施策を対象としたものです。プラン②・③の成果は含みません。",
]


def p24(s):
    sh = ids(s)
    keep = {4, 9, 5, 17}
    for shape in list(s.shapes):
        if shape.shape_id not in keep:
            delete(shape)
    set_paras(sh[17], ["あらうち様向けプラン"])
    set_paras(sh[9], ["プラン比較"])
    set_paras(sh[5], ["ご希望の運用範囲に合わせて、3つのプランをご用意しました。"])
    GRID = "BFC8DC"
    col_w = [2.2, 2.6, 2.6, 2.6]
    row_h = [0.6] + [0.43] * len(ROWS)
    rows = []
    rows.append([dict(text=t, size=14, bold=True, color="FFFFFF", fill=NAVY, align="ctr",
                      borders={"b": (1.5, NAVY), "t": (0.75, NAVY), "l": (0.75, "FFFFFF"), "r": (0.75, "FFFFFF")})
                 for t in HEAD])
    for i, row in enumerate(ROWS):
        label, vals = row[0], row[1:]
        total = (label == "月額合計")
        top_thick = i in (4, 8)          # 費用／運用内容／追加対応 の区切り
        bot_thick = i in (3, 7)
        def bd(base=None):
            b = {"t": (1.5, NAVY) if top_thick else (0.75, GRID), "b": (1.5, NAVY) if bot_thick else (0.75, GRID),
                 "l": (0.75, GRID), "r": (0.75, GRID)}
            return b
        cells = [dict(text=label, size=13, bold=True, color=("FFFFFF" if total else NAVY),
                      fill=(NAVY if total else "E8ECF6"), align="ctr", borders=bd())]
        for v in vals:
            cells.append(dict(text=v, size=(16 if total else 14), bold=total, color=(NAVY if total else TEXT),
                              fill=("E8ECF6" if total else "FFFFFF"), align="ctr", borders=bd()))
        rows.append(cells)
    add_table(s, 0.42, 1.48, col_w, row_h, rows, name="プラン比較表")
    add_box(s, 0.42, 6.02, 10.0, 1.08, [[(t, {})] for t in NOTES], size=11, color="444444", align="l",
            anchor="t", ins=(0.02, 0.02, 0.02, 0.02), after=2, name="注記")


APPLY = {3: p03, 5: p05, 6: p06, 7: p07, 9: p09, 11: p11, 12: p12, 13: p13, 24: p24}
