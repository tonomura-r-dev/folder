# -*- coding: utf-8 -*-
"""株式会社あらうち（100円賃貸）LINE公式アカウント提案資料：S1〜S17をあらうち向けに直す（2026-10-09 殿村さん指示）。

- 元の資料（47枚）をそのまま土台にし、S1〜S17の文字・図形だけを直す。S18〜S47には一切触れない
- スライドの追加・削除・並び替えはしない（47枚のまま）
- 数字の正本は「株式会社あらうち御中_LINEOA施策提案【SIM】ver1.4.xlsx」（殿村さん版：6か月で問い合わせ77件）
- 方針：友だち獲得＝離脱防止ポップアップ＋サイト内「LINEでお問い合わせ」ボタン。CPF・サンクスLINE・通知メッセージは採用しない。
  LINE公式アカウントの標準機能で運用（外部の運用ツールなし）。CPAは載せない
- 根拠のない仮置きの数字には「（仮）」を付ける

使い方:
    python _build/patch_arauchi_s1_17.py <元のpptx> <出力.pptx>
"""
import copy
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches

SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
S = prs.slides


# ---------------------------------------------------------------- 部品
def ids(slide):
    return {sh.shape_id: sh for sh in slide.shapes}


def new_run(rpr_src):
    r = copy.deepcopy(rpr_src.getparent()) if rpr_src is not None else None
    return r


def make_run(template_rpr, text):
    from lxml import etree
    r = etree.SubElement(etree.Element(qn("a:p")), qn("a:r"))
    r = copy.deepcopy(r)
    rpr = copy.deepcopy(template_rpr) if template_rpr is not None else etree.Element(qn("a:rPr"), lang="ja-JP")
    rpr.tag = qn("a:rPr")
    r.append(rpr)
    t = etree.SubElement(r, qn("a:t"))
    t.text = text
    return r


def set_paras(shape, paras, size=None, plain=False, tmpl_rpr=None, align=None):
    """paras＝段落ごとの [(文字, 太字 or None), ...]（文字だけの段落は "文字" でも可）。
    各段落の最初の run の書式を土台にする。段落が足りなければ最後の段落を複製する。
    size＝ポイント（全 run に適用）。plain＝色の指定を外す。tmpl_rpr＝run が無い図形に使う書式。"""
    txBody = shape.text_frame._txBody
    ps = txBody.findall(qn("a:p"))
    norm = [[(p, None)] if isinstance(p, str) else p for p in paras]
    # 書式の土台（図形内の最初の rPr）
    first_rpr = txBody.find(".//" + qn("a:rPr"))
    if first_rpr is None:
        first_rpr = tmpl_rpr
    while len(ps) < len(norm):
        dup = copy.deepcopy(ps[-1])
        ps[-1].addnext(dup)
        ps = txBody.findall(qn("a:p"))
    for p_el, runs in zip(ps, norm):
        rs = p_el.findall(qn("a:r"))
        base_rpr = rs[0].find(qn("a:rPr")) if rs else first_rpr
        base_rpr = copy.deepcopy(base_rpr) if base_rpr is not None else None
        for child in list(p_el):
            if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
                p_el.remove(child)
        end = p_el.find(qn("a:endParaRPr"))
        for text, bold in runs:
            r = make_run(base_rpr, text)
            rpr = r.find(qn("a:rPr"))
            if bold is not None:
                rpr.set("b", "1" if bold else "0")
            if plain:
                for f in rpr.findall(qn("a:solidFill")):
                    rpr.remove(f)
            if end is not None:
                end.addprevious(r)
            else:
                p_el.append(r)
        if align is not None:
            ppr = p_el.find(qn("a:pPr"))
            if ppr is None:
                from lxml import etree
                ppr = etree.Element(qn("a:pPr"))
                p_el.insert(0, ppr)
            ppr.set("algn", align)
    for p_el in ps[len(norm):]:
        txBody.remove(p_el)
    if size is not None:
        for rpr in list(txBody.iter(qn("a:rPr"))) + list(txBody.iter(qn("a:endParaRPr"))):
            rpr.set("sz", str(int(size * 100)))
    if plain:
        for end in txBody.iter(qn("a:endParaRPr")):
            for f in end.findall(qn("a:solidFill")):
                end.remove(f)


def set_run_sizes(shape, sizes):
    """段落ごと・run ごとの大きさを指定（[[pt, pt], [pt]]）"""
    ps = shape.text_frame._txBody.findall(qn("a:p"))
    for p_el, sz in zip(ps, sizes):
        for r, s in zip(p_el.findall(qn("a:r")), sz):
            r.find(qn("a:rPr")).set("sz", str(int(s * 100)))


def delete(shape):
    el = shape._element
    el.getparent().remove(el)


def clone_into(dst_slide, src_shape, dy=0.0, new_id=None):
    el = copy.deepcopy(src_shape._element)
    tree = dst_slide.shapes._spTree
    tree.append(el)
    nid = new_id or (max(sh.shape_id for sh in dst_slide.shapes) + 1)
    for c in el.iter(qn("p:cNvPr")):
        c.set("id", str(nid))
        break
    sh = [s for s in dst_slide.shapes if s.shape_id == nid][-1]
    if dy:
        sh.top = sh.top + Inches(dy)
    return sh


def set_color(shape, para_idx, hex_):
    """段落 para_idx の全 run の文字色を hex_ にする（色が入れ替わらないよう明示する）"""
    from lxml import etree
    p_el = shape.text_frame._txBody.findall(qn("a:p"))[para_idx]
    for rpr in [r.find(qn("a:rPr")) for r in p_el.findall(qn("a:r"))]:
        for f in rpr.findall(qn("a:solidFill")):
            rpr.remove(f)
        fill = etree.Element(qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr"), val=hex_)
        # rPr の子は ln の後・latin の前に solidFill を置く
        ln = rpr.find(qn("a:ln"))
        if ln is not None:
            ln.addnext(fill)
        else:
            rpr.insert(0, fill)


def set_spc_before(shape, para_idx, pts100):
    """段落 para_idx の前の間隔を pts100（1/100pt）にする"""
    from lxml import etree
    p_el = shape.text_frame._txBody.findall(qn("a:p"))[para_idx]
    ppr = p_el.find(qn("a:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("a:pPr"))
        p_el.insert(0, ppr)
    for old in ppr.findall(qn("a:spcBef")):
        ppr.remove(old)
    sb = etree.Element(qn("a:spcBef"))
    etree.SubElement(sb, qn("a:spcPts"), val=str(pts100))
    ln = ppr.find(qn("a:lnSpc"))
    (ln.addnext(sb) if ln is not None else ppr.insert(0, sb))


def set_line_spacing(shape, pct):
    """全段落の行間を pct（%）にする"""
    from lxml import etree
    for p_el in shape.text_frame._txBody.findall(qn("a:p")):
        ppr = p_el.find(qn("a:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("a:pPr"))
            p_el.insert(0, ppr)
        for old in ppr.findall(qn("a:lnSpc")):
            ppr.remove(old)
        ln = etree.Element(qn("a:lnSpc"))
        etree.SubElement(ln, qn("a:spcPct"), val=str(int(pct * 1000)))
        ppr.insert(0, ln)


def set_size_all(shape, pt):
    for rpr in list(shape.text_frame._txBody.iter(qn("a:rPr"))) + list(shape.text_frame._txBody.iter(qn("a:endParaRPr"))):
        rpr.set("sz", str(int(pt * 100)))


def no_side_inset(shape):
    """左右の余白を0にする（文字を小さくせずに1行に収めるため）"""
    bp = shape.text_frame._txBody.find(qn("a:bodyPr"))
    bp.set("lIns", "0")
    bp.set("rIns", "0")


def rpr_of(shape):
    return shape.text_frame._txBody.find(".//" + qn("a:rPr"))


# 元のS3・S11（書き換える前）から部品を取っておく（S5・S6の市場分析に使う）
S3_ORIG = ids(S[2])
S11_ORIG = ids(S[10])


# ---------------------------------------------------------------- S2 現状把握
def s2():
    b = ids(S[1])
    assert b[171].text_frame.text == "18,200" and b[179].text_frame.text == "Lstep", "S2の中身が想定と違う"
    set_paras(b[26], [
        [("LINE公式アカウントの標準機能で運用中（Lステップ等のツールなし）。", None)],
        [("友だち追加の入口・追加後の動線ともに、改善の余地があると想定。", None)],
    ])
    set_paras(b[106], [[("1,737", None)]])
    set_paras(b[108], [[("0（広告なし）", None)]])
    set_paras(b[110], [[("なし", True)]])
    set_paras(b[118], [[("30件（仮）", None)]], plain=True)
    set_paras(b[120], [[("問い合わせ", None)]])
    set_paras(b[124], [[("0％", None)]])
    set_paras(b[114], [[("YouTube", None)]], plain=True)
    set_paras(b[112], [[("不明（広告なし）", None)]])
    set_paras(b[171], [[("1,241", None)]])
    set_paras(b[173], [[("約807（仮）", None)]])
    set_paras(b[175], [[("35%（仮）", None)]])
    set_paras(b[179], [[("なし（標準機能）", None)]])
    set_paras(b[92], [[("リッチメニュー", None)]], size=10)
    no_side_inset(b[92])
    set_paras(b[177], [[("なし", True)]])
    set_paras(b[169], [[("あいさつ", None)]])
    set_paras(b[191], [[("ボタンなし", True)]])
    set_paras(b[67], [
        [("※", None), ("いずれも貴社ヒヤリング、外部計測", None)],
        [("　をもとにした想定。（仮）は仮置き。", None)],
    ])
    set_paras(b[43], [
        [("全ページに「LINEでお問い合わせ」ボタンを設置。", None)],
        [("100円チェッカー・内見依頼", True), ("など、問い合わせ入口を整備。", False)],
    ])
    set_paras(b[53], [[("リッチメニュー未設置・あいさつにボタンなし。改善の余地あり。", None)]])
    set_paras(b[189], [[("有（サイト・動画）", None)]])   # プロフィールにサイト・紹介動画へのリンクあり（S4と揃える）
    # 値の文字を10ptにそろえる（WEB側は10pt、LINE側は9pt・問い合わせは8ptだった）
    for k in (120, 171, 173, 175, 179, 181, 183, 185, 187, 189):
        set_size_all(b[k], 10)
        if b[k].height < Inches(0.28):
            b[k].height = Inches(0.28)


# ---------------------------------------------------------------- S3 想定動線（カスタマージャーニー）
def s3():
    b = ids(S[2])
    set_paras(b[31], [
        [("ポータルサイトで物件検索 → サイト閲覧 → LINE追加 → 物件情報の送信 → 問い合わせ、の流れに沿って設計します。", None)],
        [("サイトを見て離れる方・友だち追加後に物件を送らない方への対策", True), ("をご提案します。", False)],
    ])
    set_paras(b[7], [[("⑦マネタイズ", None), ("(成約)", None)]])
    set_run_sizes(b[7], [[11, 10]])
    # 左の4つの箱に①〜④の番号を付け、指示の順番（物件検索→サイト閲覧→LINE追加→物件情報の送信→問い合わせ）を示す。
    # 文字は元の11ptのまま。収まるように文言を短くする（流れの全文は上の要旨に書いている）
    set_paras(b[39], [[("①ポータルで物件検索", None)]])
    set_paras(b[26], [[("②100円賃貸のサイト", None)]])
    set_paras(b[37], [[("③LINEで友だち追加", None)]])
    set_paras(b[46], [[("④物件の情報を送信", None)]])
    for k in (39, 26, 37, 46):
        no_side_inset(b[k])
    # 線：①→問い合わせの直通（44）と、③→問い合わせ（27）を外す。②→問い合わせ（22）は「電話／WEB」、④→問い合わせ（47）は「LINE」
    delete(b[44])
    delete(b[27])
    b[33].top = Inches(4.635) - b[33].height // 2      # 「電話／WEB」を②からの線の上へ
    S[2].shapes._spTree.append(b[33]._element)         # 線より手前に出す（線が文字の上を通らないように）
    for k in (33, 55):                                   # 薄い青の地に白8ptでは読めないので、9pt・紺に
        set_size_all(b[k], 9)
        for i in range(len(b[k].text_frame.paragraphs)):
            set_color(b[k], i, "1F285A")
    set_paras(b[55], ["LINE"])
    # ②→③の縦の矢印（元の点線の矢印を複製して、まっすぐ下向きにする）
    arrow = copy.deepcopy(b[22]._element)
    for ext in arrow.iter(qn("a:extLst")):
        ext.getparent().remove(ext)
        break
    S[2].shapes._spTree.append(arrow)
    arrow.find(qn("p:nvCxnSpPr")).find(qn("p:cNvPr")).set("id", str(max(sh.shape_id for sh in S[2].shapes) + 1))
    arrow.find(qn("p:nvCxnSpPr")).find(qn("p:cNvPr")).set("name", "矢印 ②→③")
    xfrm = arrow.find(qn("p:spPr")).find(qn("a:xfrm"))
    for k in ("flipH", "flipV"):
        xfrm.attrib.pop(k, None)
    cx_mid = b[26].left + b[26].width // 2
    xfrm.find(qn("a:off")).set("x", str(cx_mid))
    xfrm.find(qn("a:off")).set("y", str(b[26].top + b[26].height))
    xfrm.find(qn("a:ext")).set("cx", "0")
    xfrm.find(qn("a:ext")).set("cy", str(b[37].top - (b[26].top + b[26].height)))
    arrow.find(qn("p:spPr")).find(qn("a:prstGeom")).set("prst", "straightConnector1")
    set_paras(b[34], [[("問い合わせ（空室確認・内見の手配）", None)]])
    set_paras(b[40], [[("お申込み", None)]])
    set_paras(b[36], [[("ご契約", None)]])
    set_paras(b[74], [[("（鍵のお渡し・お引っ越し）", None)]])


# ---------------------------------------------------------------- S4 LINEアカウントの現状
def s4():
    b = ids(S[3])
    set_paras(b[45], [[("LINEアカウントの現状", None)]])
    set_paras(b[28], [
        [("あいさつで物件名・URLの送信をお願いする形は、100円賃貸の使い方に合っています。", None)],
        [("一方で、", False), ("あいさつに送るメリット・安心材料やボタンがなく、リッチメニューも未設置", True), ("です。", False)],
    ])
    set_paras(b[17], [
        "あいさつで「気になる物件名」と「物件ページのURL」の",
        "送信をお願いし、やり取りが始まる形。",
    ])
    set_paras(b[23], ["あいさつの後にマンガ・動画を送り、サービスの理解を促している。"])
    set_paras(b[19], [
        "プロフィールにサイト・紹介動画へのリンクを掲載。",
        "リッチメニューは未設置で、友だち追加後に次へ進む導線がない。",
    ])


# ---------------------------------------------------------------- S5・S6 市場分析
def market(slide, summary, cards, sources):
    b = ids(slide)
    set_paras(b[4], [[("市場分析", None)]])
    delete(b[2])
    # 上の帯（S3の部品）
    for sid in (3, 4, 2):
        clone_into(slide, S3_ORIG[sid])
    summ = clone_into(slide, S3_ORIG[31])
    set_paras(summ, summary)
    # カード3枚（S11の部品を0.36インチ下げる）
    for (num, title, body), (n_id, t_id, b_id) in zip(cards, ((6, 10, 14), (17, 19, 21), (23, 25, 27))):
        n = clone_into(slide, S11_ORIG[n_id], dy=0.36)
        t = clone_into(slide, S11_ORIG[t_id], dy=0.36)
        bd = clone_into(slide, S11_ORIG[b_id], dy=0.36)
        set_paras(n, [[(num, None)]])
        set_paras(t, [[(title, None)]], size=13)
        set_paras(bd, body, size=14)
    # 出典（S11の「結論・推奨」の枠）
    src = clone_into(slide, S11_ORIG[15])
    src.top, src.height = Inches(5.42), Inches(1.33)
    lines = []
    for name, url, two_lines in sources:     # 長いURLは名前と別の行に（URLの途中で折り返さないように）
        lines += [[(name, False)], [("　" + url, False)]] if two_lines else [[(name + "　" + url, False)]]
    set_paras(src, [[("出典（2026年10月9日確認）", True)]] + lines)
    ps = src.text_frame._txBody.findall(qn("a:p"))
    for i, p_el in enumerate(ps):
        for r in p_el.findall(qn("a:r")):
            r.find(qn("a:rPr")).set("sz", "1000" if i == 0 else "900")


def s5():
    market(S[4],
           [[("家賃が上がり、東京23区のシングル向きマンションは平均10万円を超えました。", None)],
            [("物件はポータルサイトで自分で探し、問い合わせ先は別に選ぶ流れが主流", True), ("です。", False)]],
           [("01", "家賃の上昇", [
               "・東京23区のシングル向き（30㎡以下）マンションの平均家賃は、2025年5月に初めて10万円を超えた",
           ]),
            ("02", "物件はポータルサイトで探す", [
                "・住まいの探し方は「不動産ポータルサイトで検索」がトップ（引越し経験者66.8％・検討者73.2％）",
            ]),
            ("03", "問い合わせ先は物件数で選ぶ", [
                "・問い合わせる不動産会社を選ぶ基準は「取り扱っている物件数が多い」が1位（経験者26.6％・検討者47.7％）",
                "・100円賃貸は、お客様が見つけた物件の仲介に対応（100円賃貸が対応できる物件に限ります）",
            ])],
           [("・アットホーム「2025年の賃貸市場における4大ニュース」", "https://www.athome.co.jp/corporate/news/data/questionnaire/yondai-news-202512/", True),
            ("・アットホーム「オンラインでの住まい探しに関する調査 2025 賃貸編」", "https://www.athome.co.jp/corporate/news/data/questionnaire/online-chintai-202510/", True),
            ("・100円賃貸", "https://100en.net/", False)])


def s6():
    market(S[5],
           [[("内見予約のやり取りは、電話よりメール・SMSを希望する方が多くいます。", None)],
            [("物件のURLをLINEで送るだけで相談できるサービスも出ており、LINEでの受け口が欠かせません", True), ("。", False)]],
           [("01", "電話より、文字でのやり取り", [
               "・内見予約のやり取りの希望は、電話33.4％に対し、メール・SMSが52.6％",
           ]),
            ("02", "手続きもオンラインで", [
                "・オンラインで契約したい検討者は36.4％、オンラインで重要事項説明を受けたい検討者は29.6％",
                "・100円賃貸は、テレビ電話・郵送・ネットでのお申込み・ご契約に対応",
            ]),
            ("03", "URLを送るだけの競合も", [
                "・タダスム：見つけた物件のURLをLINEで送ると、空室確認・内見調整。仲介手数料は「0円or最大50%」",
            ])],
           [("・アットホーム「オンラインでの住まい探しに関する調査 2025 賃貸編」", "https://www.athome.co.jp/corporate/news/data/questionnaire/online-chintai-202510/", True),
            ("・タダスム", "https://tadasumu.com/　https://tadasumu.com/navi/tadasumu-toha/", False),
            ("・100円賃貸", "https://100en.net/", False)])


# ---------------------------------------------------------------- S7 LINE運用ツールの選定
def s7():
    b = ids(S[6])
    set_paras(b[7], [[("標準機能で運用 ", True), ("※今回はこちらで運用", False)]])
    set_run_sizes(b[7], [[14, 10]])
    set_paras(b[9], [
        [("特徴", True)], [("ツール費がかからず、LINE公式アカウントの管理画面だけで運用できる", False)],
        [("メリット", True)], [("ステップ配信・問い合わせ済みの方を除く配信・キーワード応答・チャットタグなど、LINE上の施策はすべて標準機能で実現できる", False)],
        [("デメリット", True)], [("離脱防止ポップアップとサイトのボタンなど、同じ友だち追加URLから入った方は、経路で分けて配信できない", False)],
    ])
    set_paras(b[11], [[("外部の運用ツール", True), (" ", False), ("※Lステップ等・今回は導入しない", False)]])
    set_run_sizes(b[11], [[14, 14, 11]])
    set_paras(b[13], [
        [("特徴", True)], [("タグや流入経路を細かく管理でき、外部データ連携に強い", False)],
        [("メリット", True)], [("流入元ごとに、配信を細かく出し分けられる", False)],
        [("デメリット", True)], [("月額のツール費と、初期設定・移行の工数がかかる", False)],
    ])
    set_paras(b[15], [
        [("結論", True)],
        [("LINE公式アカウントの標準機能で運用します（外部の運用ツールは導入しません）。", False)],
        [("※離脱防止ポップアップは、LINE公式アカウントの機能ではなく、サイト側に別途導入する施策です。", False)],
    ])
    # 見出し・本文の色と段落の間隔を明示（段落を差し替えると色がずれるため）。右カードの間隔は左カードに揃える
    for i in (1, 3, 5):
        set_color(b[9], i, "333333")
        set_color(b[13], i, "333333")
    for i in (0, 2, 4):
        set_color(b[13], i, "00B050")
    set_spc_before(b[9], 2, 600)
    for pl, pr in zip(b[9].text_frame._txBody.findall(qn("a:p")), b[13].text_frame._txBody.findall(qn("a:p"))):
        ppl, ppr = pl.find(qn("a:pPr")), pr.find(qn("a:pPr"))
        if ppl is not None:
            if ppr is not None:
                pr.remove(ppr)
            pr.insert(0, copy.deepcopy(ppl))


# ---------------------------------------------------------------- S8 想定動線（全体）
def s8():
    sl = S[7]
    b = ids(sl)
    set_paras(b[71], [
        [("離脱防止ポップアップとサイトのLINEボタンから友だち追加し、物件の問い合わせと継続配信につなげます。", None)],
        [("LINE公式アカウントの標準機能で運用（離脱防止ポップアップは別途導入）", None)],
    ])
    set_paras(b[8], [[("<友だち追加後・問い合わせ後の継続配信>", None)]])
    # 入口は2つだけ：①サイト＞離脱防止ポップアップ ②サイトのLINEボタン
    set_paras(b[9], [[("100円賃貸のサイト", None)]])
    no_side_inset(b[9])
    set_paras(b[17], [[("離脱防止ポップアップ", None)]])
    set_paras(b[15], [[("サイトの「LINEでお問い合わせ」ボタン", None)]])
    for i in (0,):
        set_color(b[15], i, "FFFFFF")      # 元は「使わない項目」の薄いグレー。並びの「離脱防止ポップアップ」と同じ白に
    b[15].top = Inches(2.94)
    b[28].top, b[28].height = Inches(2.81), Emu(int(Inches(3.075) - Inches(2.81)))
    for sid in (10, 11, 18, 12, 16, 13, 29, 68, 55, 56, 51, 54):
        delete(b[sid])
    # 初期設計項目
    set_paras(b[35], ["・プロフィール", "・リッチメニュー", "・あいさつ", "・応答メッセージ",
                      "・チャットタグ", "・計測タグ", "・離脱防止（別途）", "・サイトのボタン"], size=9)
    set_line_spacing(b[35], 90)
    b[35].height = Inches(1.25)
    for k in (34, 25):
        set_size_all(b[k], 9)
        bp = b[k].text_frame._txBody.find(qn("a:bodyPr"))
        bp.set("tIns", "0")
        bp.set("bIns", "0")
        b[k].height = Inches(0.22)
    b[25].width = Inches(1.75)
    set_paras(b[31], [[("企画配信（月4本）", None)]])
    set_paras(b[32], [[("ステップ配信（全10通・14日）", None)]])
    set_paras(b[33], [[("リサーチ", None)]])
    set_paras(b[39], [[("→ 配信に反映", None)]])
    set_paras(b[36], [[("問い合わせ済みの方を除く配信", None)]])
    set_paras(b[38], [[("よくある質問（キーワード応答）", None)]])
    set_paras(b[37], [[("リッチメニュー（6ボタン）", None)]])
    set_paras(b[40], [[("個別チャット（物件の確認）", None)]])
    set_paras(b[19], [[("空室確認・内見", None)]])
    set_paras(b[14], [[("物件の問い合わせ", None)]])
    set_paras(b[69], [[("LINE公式アカウントの標準機能で運用", None)]])
    # 友だち追加後・問い合わせ後
    set_paras(b[52], [[("問い合わせ", None)]], size=10)
    no_side_inset(b[52])                # 幅0.8インチの箱。余白0で10ptを1行に収める
    set_color(b[52], 0, "595959")      # 薄いグレーの箱に白8ptでは読めない
    set_paras(b[57], [[("内見・お申込みの案内", None)]])
    set_paras(b[58], [[("別の物件URLの受付", None)]])
    set_paras(b[59], [[("ご成約", None)]])
    set_paras(b[61], [[("入居までの案内", None)]])
    set_paras(b[60], [[("チャットタグで対応状況を管理", None)]])
    set_paras(b[66], [[("企画配信（月4本）・リサーチ", None)]])
    set_paras(b[67], [[("個別チャット", None)]])
    set_paras(b[20], [[("アクセス解析での効果測定", None)]])
    set_paras(b[23], [[("キーワード応答の見直し", None)]])
    set_paras(b[22], [[("受付時間外の自動応答", None)]])
    set_paras(b[24], [[("配信結果の振り返り", None)]])


# ---------------------------------------------------------------- S9 想定セグメント×成果
def s9():
    b = ids(S[8])
    set_paras(b[68], [
        [("新しい友だちには物件の送り方を、今の友だちには配信とリッチメニューで問い合わせのきっかけを届けます。", None)],
        [("「どの物件が対象か」が分からないこと", True), ("がハードルになっていると考え、100円チェッカーを前面に出します。", False)],
    ])
    # 1・2行目：友だち追加の2つの導線
    set_paras(b[11], ["100円賃貸のサイト"])
    set_paras(b[9], ["離脱防止", "ポップアップ"])
    set_paras(b[34], ["あいさつ", "物件の送り方を案内"])
    set_paras(b[35], ["ステップ配信（全10通）", "→ 物件URLの送信"])
    set_paras(b[12], ["100円賃貸のサイト"])
    set_paras(b[10], ["LINEでお問い合わせ", "ボタン（全ページ）"])
    set_paras(b[37], ["キーワード応答", "→ よくある質問に返信"])
    set_paras(b[38], ["個別チャット", "→ 空室確認・内見の手配"])
    # 3行目：今の友だち・新しい友だちへの配信
    set_paras(b[17], ["今の友だち（1,241人）", "新しい友だち"])
    set_paras(b[48], ["企画配信（月4本）", "→ 送り方・費用など"])
    set_paras(b[49], ["リッチメニュー", "→ 100円チェッカーなど"])
    # 4〜7行目：問い合わせを後押しする対応
    set_paras(b[19], ["問い合わせた方"])
    set_paras(b[18], ["個別チャット"])
    set_paras(b[51], ["空室確認・内見の手配"])
    set_paras(b[52], ["お申込み・ご契約の案内", "（オンラインでも可）"])
    set_paras(b[26], ["受付時間外の方"])
    set_paras(b[25], ["自動応答"])
    set_paras(b[54], ["受付時間外は", "受け付けた旨を自動で返信"])
    set_paras(b[55], ["受付時間（11〜18時）", "内に担当者が確認"])
    set_paras(b[27], ["迷っている方"])
    set_paras(b[29], ["リッチメニュー"])
    set_paras(b[57], ["100円チェッカー", "（名前・電話番号は不要）"])
    set_paras(b[58], ["→ 物件の送信"])
    set_paras(b[28], ["他社で内見済みの方"])
    set_paras(b[30], ["企画配信"])
    set_paras(b[60], ["内見済みでも相談可", "（お受けできない場合も）"])
    set_paras(b[61], ["→ 物件の送信"])
    # 表の中の文字は原本の8pt・薄いグレーでは読みにくいので、10pt・濃いグレー・行間100%・上下中央に（文言は1行12字以内）
    for k in (34, 35, 37, 38, 48, 49, 51, 52, 54, 55, 57, 58, 60, 61):
        set_size_all(b[k], 10)
        for i in range(len(b[k].text_frame.paragraphs)):
            set_color(b[k], i, "595959")
        set_line_spacing(b[k], 100)
        b[k].text_frame._txBody.find(qn("a:bodyPr")).set("anchor", "ctr")
    # 右の数字（試算）。25→52％＝SIMの問い合わせ①÷新しい友だち（1か月目 6÷22.4、6か月目 12÷23.0）
    # 0.3→1.2％＝SIMのCVR（問い合わせ②÷クリック。1か月目0.34％、6か月目1.20％）
    set_paras(b[40], [[("問い合わせ率", None)], [("※試算の想定", None)]])
    set_paras(b[41], [[("25→52", None), ("%", None)], [("友だち追加直後に", False)], [("物件を送る方（仮）", False)]])
    set_run_sizes(b[41], [[20, 14], [9], [9]])
    set_paras(b[42], [[("0.3→1.2", None), ("%", None)], [("クリックからの", False)], [("問い合わせ（CVR）", False)]])
    set_run_sizes(b[42], [[20, 14], [9], [9]])
    for k in (41, 42):                  # 20ptの「0.3→1.2%」が1行に入るよう、左へ広げる（右端は10.15インチ）
        b[k].left, b[k].width = Inches(8.20), Inches(1.95)
        no_side_inset(b[k])
    for sid in (43, 44, 46, 47):
        set_paras(b[sid], [[("－", None)]])
        set_run_sizes(b[sid], [[24]])


# ---------------------------------------------------------------- S11 要件定義（課題）
def s11():
    b = ids(S[10])
    set_paras(b[10], [[("サイト訪問者の取りこぼし", None)]])
    set_paras(b[19], [[("問い合わせへの心理的ハードル", None)]])
    set_paras(b[25], [[("友だち追加後の配信・導線不足", None)]])
    set_paras(b[14], [
        "・サイトの月間訪問者（UU）は1,737。広告は出しておらず、友だちの入口はサイトの「LINEでお問い合わせ」ボタンだけ",
        "・問い合わせずにサイトを離れる方へ、もう一度案内する手段がない",
        "・友だちは1,241人で、増やす仕組みが少ない",
    ])
    set_paras(b[21], [
        "・「どの物件が対象か」「本当に100円になるか」が分からないと、物件を送るのをためらう可能性がある",
        "・内見予約のやり取りは、電話（33.4％）よりメール・SMS（52.6％）を希望する方が多い（出典は6ページ）",
    ])
    set_paras(b[27], [
        "・リッチメニューが未設置で、友だち追加後に次の行動へ進むボタンがない",
        "・あいさつは物件名・URLの送信のお願いが中心で、送るメリットや安心材料が伝わりにくい",
        "・物件を送らなかった方へ、続けて案内する配信の設計が必要",
    ])
    set_paras(b[15], [
        [("結論", True)],
        [("サイトの入口（離脱防止ポップアップ・LINEボタン）を整え、友だち追加後はあいさつ・リッチメニュー・ステップ配信で「物件を送る」まで案内します。", False)],
    ])


# ---------------------------------------------------------------- S12 要件定義（提案方針）
def s12():
    b = ids(S[11])
    set_paras(b[27], [[("LINE経由の物件問い合わせを増やす", None)]])
    set_paras(b[34], [[("LINE経由の問い合わせ 1か月目 月8件→6か月目 月18件（試算）", None)]], size=20)
    set_paras(b[2], [[("離脱防止ポップアップ", None)]])
    set_paras(b[5], [[("あいさつメッセージの改善", None)]])
    set_paras(b[6], [[("リッチメニュー・自動応答", None)]])
    set_paras(b[7], [[("ステップ配信・企画配信", None)]])
    set_paras(b[8], ["サイトを離れようとした方に、LINEでの物件の問い合わせを案内（LINE公式アカウントとは別に導入）"])
    set_paras(b[9], ["物件の送り方に加え、営業電話なし・確認だけでも可などの送る理由と、ボタンを追加"])
    set_paras(b[10], ["「物件を送る」「100円になるか確認」などを常に表示。よくある質問はキーワード応答で即時に回答"])
    set_paras(b[11], ["友だち追加から14日間・全10通で物件の送信を後押しし、月4本の企画配信で今の友だちにも案内"])
    set_paras(b[12], [[("入口・", None)], [("あいさつ", None)]])
    set_paras(b[13], [[("追加後の", None)], [("案内", None)]])
    set_run_sizes(b[13], [[11], [11]])
    delete(b[19])   # 下の2行を薄く見せていた半透明の覆い


# ---------------------------------------------------------------- S13 想定の費用対効果
def s13():
    sl = S[12]
    b = ids(sl)
    set_paras(b[8], [[("¥215,000", None)]])
    set_paras(b[10], [[("¥135,000", None)]])
    delete(b[13])   # Lステップのロゴ
    # 初期の実施内容
    for k, v in ((12, "あいさつ"), (14, "改善"), (18, "リッチメニュー"), (22, "新規作成"),
                 (26, "ステップ配信"), (30, "10通を構築"),
                 (34, "自動応答"), (38, "設定"), (42, "離脱防止"), (47, "導入")):
        set_paras(b[k], [[(v, None)]])
    # 月次の実施内容
    for k, v in ((25, "企画配信"), (29, "月4本"), (58, "離脱防止"), (62, "運用"), (70, "10通の改善")):
        set_paras(b[k], [[(v, None)]])
    # 6か月後（試算・6か月目の月間）
    set_paras(b[82], [[("問い合わせ①", None)]])
    set_paras(b[84], [[("問い合わせ②", None)]])
    set_paras(b[86], [[("合計", None)]])
    set_paras(b[11], [[("1,377人", None)]])
    set_paras(b[16], [[("12件/月", None)]])
    set_paras(b[23], [[("6件/月", None)]])
    set_paras(b[27], [[("18件/月", None)]])
    set_paras(b[31], [[("（6か月で77件）", None)]])
    b[31].left, b[31].width = Inches(8.75), Inches(1.80)   # 元の幅だと折り返す。右端は「6か月後」の帯に揃える
    note = clone_into(sl, b[19])
    note.left, note.top, note.width, note.height = Inches(0.29), Inches(6.58), Inches(10.26), Inches(0.5)
    set_paras(note, [
        "※初期＝コンサル20万円＋離脱防止1.5万円／月額＝コンサル10万円＋離脱防止3万円＋LINE公式アカウント0.5万円",
        "※6か月後は6か月目の月間の試算。問い合わせ①＝友だち追加直後の物件の問い合わせ（送る方の割合は仮）、②＝配信・リッチメニューから",
    ], size=9, align="l")


def s13_summary():
    """S13のタイトル下の帯（要旨の欄）が空なので、S3の要旨の枠を複製して入れる"""
    summ = clone_into(S[12], S3_ORIG[31])
    set_paras(summ, [[("初期21.5万円・月額13.5万円で運用します。", None)],
                     [("6か月目にLINE経由の問い合わせ月18件を見込みます（試算）", True), ("。", False)]])


# ---------------------------------------------------------------- S14 実施施策（配信）
def s14():
    sl = S[13]
    b = ids(sl)
    tmpl = copy.deepcopy(rpr_of(b[40]))
    set_paras(b[6], [[("友だち追加から14日間のステップ配信（全10通）と月4本の企画配信で、物件URLの送信を後押しします", None)]])
    set_paras(b[13], [[("[ステップ配信・全10通]", None)]])
    # 日数は元の9.5ptのまま。「9〜14日後」が収まるよう左右の余白を0にする
    set_paras(b[20], [[("1〜3日後", None)]])
    set_paras(b[48], [[("4〜7日後", None)]])
    set_paras(b[61], [[("9〜14日後", None)]])
    for k in (20, 48, 61):
        no_side_inset(b[k])
    set_paras(b[12], [[("ねらい", None)]], size=9.5)
    set_paras(b[22], [[("ねらい", None)]], size=9.5)
    step = (
        (24, "新しい友だち", 32, "追加後に自動配信", 36, ["仲介手数料のしくみ", "100円チェッカー"], 40, "物件URLの送信"),
        (49, "新しい友だち", 52, "追加後に自動配信", 55, ["ポータルでの探し方", "内見済みでも相談可"], 57, "比較中の方の後押し"),
        (63, "新しい友だち", 65, "追加後に自動配信", 68, ["初期費用の交渉", "入居までの流れ"], 71, "お申込みの後押し"),
    )
    plan = (
        (75, "友だち全員", 77, "絞り込みなし", 79, "物件URLの送り方", 83, "物件の送信"),
        (88, "友だち全員", 90, "絞り込みなし", 92, ["初期費用の内訳", "（家賃10万円の例）"], 94, "費用の不安を解消"),
        (97, "物件を送っていない方", 99, ["問い合わせ済みを除く", "（チャットタグ）"], 101, ["内見済みでも相談可", "100円チェッカー"], 103, "物件の送信"),
        (106, "友だち全員", 108, "絞り込みなし", 110, ["お申込みの準備", "入居までの流れ"], 112, "お申込みの後押し"),
    )
    for row in step + plan:
        for k, text in zip(row[0::2], row[1::2]):
            paras = [[(t, None)] for t in text] if isinstance(text, list) else [[(text, None)]]
            set_paras(b[k], paras, size=9.5, tmpl_rpr=tmpl, align="ctr")
    for k in (44, 60, 73, 87, 96, 105, 114):
        set_paras(b[k], [[("問い合わせへ", None)]])


# ---------------------------------------------------------------- S15 その他（ユーザー満足度・管理工数）
def s15():
    b = ids(S[14])
    set_paras(b[2], [[("ユーザー対応の質を高める3つの施策（物件の問い合わせを前提に整理）", None)]])
    set_paras(b[5], [[("よくある質問の自動回答", None)]])
    set_paras(b[6], ["・初期費用やお申込みの流れなどの質問に自動で回答",
                     "・リッチメニューのボタンと同じ言葉で応答を用意"])
    set_paras(b[9], ["・送られた物件（URL・物件名・画像）の空室確認・内見を手配",
                     "・チャットタグで受付・確認中・返信済を管理"])
    set_paras(b[13], ["・受付時間外は、受け付けた旨を自動で返信",
                      "・リッチメニューから100円チェッカーへ案内"])
    set_paras(b[19], ["・キーワード応答で一次対応を自動化し、物件の確認に集中",
                      "・受付時間外の自動返信は、時間を指定して時間外だけに返信"])
    set_paras(b[23], ["・チャットタグの付け替えで、確認漏れ・返信漏れを防止",
                      "・問い合わせ済みの方を除いた再案内で、配信の重複を防止"])


# ---------------------------------------------------------------- S17 新規友だち獲得
def s17():
    b = ids(S[16])
    set_paras(b[10], [[("サイトの2つの接点から、LINEの友だち追加へつなげます", None)]])
    set_paras(b[15], ["サイトを離れようとした方", "（離脱防止ポップアップ）"])
    set_paras(b[26], ["物件のURLを送るだけで仲介手数料100円", "（100円賃貸が対応できる物件に限ります）"], size=10)
    set_paras(b[29], ["サイト内の「LINEでお問い合わせ」", "ボタン（全ページ）"])
    set_paras(b[31], ["物件のURL・物件名・画像を送るだけで、", "空室確認から内見の手配まで"])
    for sid in (34, 35, 36, 43, 44):
        delete(b[sid])
    set_paras(b[45], [[("※離脱防止ポップアップは、LINE公式アカウントとは別に導入する施策です（初期1.5万円・月3万円）。サイトのボタンは、リンク先を友だち追加用のURLに変更します。", None)]])


# S5・S6は元のS3・S11の部品を使うので、S3・S11を書き換える前に作る
s5()
s6()
s2()
s3()
s4()
s7()
s8()
s9()
s11()
s12()
s13()
s13_summary()
s14()
s15()
s17()
assert len(prs.slides) == 47
prs.save(OUT)
print("saved:", OUT)
