# -*- coding: utf-8 -*-
"""チェングロウス｜サイトリニューアル提案からのブリッジ資料（19枚）。2026-10-04 殿村さん指示の構成。
土台は ver4.1（殿村さんがPCで保存した版）。流用するページは文言だけ差し替え、残りは新規に組む。
SIMの数値・費用は ver4.1 のものをそのまま使う。画像は既存の画面イメージ（_images/chengrowth_v11_*.png）だけ使う。
  python3 _build/build_chengrowth_site_bridge.py <ver4.1.pptx>
"""
import copy
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20261004_チェングロウス_サイトリニューアルとLINE同時導入のご提案ver1.0.pptx"
IMG = ROOT / "_images"

NAVY, TITLE, INK, GRAY = "1F285A", "002060", "333333", "7F7F7F"
CARD, GREEN, GREEN_BG, GREEN_TX = "F4F7FF", "06C755", "E8F8EE", "0B7A3B"
GRAY_BG, ORANGE_BG, ORANGE = "F2F2F2", "FCE4D6", "ED7D31"
WHITE = "FFFFFF"

prs = Presentation(sys.argv[1])
SRC = list(prs.slides)  # ver4.1 の34枚（添字は 番号-1）
LAYOUT = SRC[22].slide_layout


def rgb(h):
    return RGBColor.from_string(h)


# ---------------- 文字・図形の部品 ----------------
def style_run(r, size, bold=False, color=INK):
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = rgb(color)
    rpr = r._r.get_or_add_rPr()
    for t in ("a:latin", "a:ea", "a:cs"):
        for e in rpr.findall(qn(t)):
            rpr.remove(e)
    after = [qn(x) for x in ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")]
    nxt = next((c for c in rpr if c.tag in after), None)
    for t in ("a:latin", "a:ea", "a:cs"):
        e = rpr.makeelement(qn(t), {"typeface": "メイリオ"})
        (nxt.addprevious(e) if nxt is not None else rpr.append(e))


def fill_text(tf, paras, align=PP_ALIGN.LEFT):
    """paras = [(text, size, bold, color, space_after_pt)]（色・太さ省略可）"""
    first = True
    for p_ in paras:
        text, size = p_[0], p_[1]
        bold = p_[2] if len(p_) > 2 else False
        color = p_[3] if len(p_) > 3 else INK
        sa = p_[4] if len(p_) > 4 else 4
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(sa)
        p.line_spacing = 1.2
        r = p.add_run()
        r.text = text
        style_run(r, size, bold, color)


def box(s, x, y, w, h, paras, fill=CARD, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, adj=0.06, line=None,
        margins=(0.4, 0.25, 0.4, 0.15)):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    sh.adjustments[0] = adj
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(1.25)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    st = sh._element.find(qn("p:style"))
    if st is not None:
        st.find(qn("a:effectRef")).set("idx", "0")
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Cm(m) for m in margins]
    fill_text(tf, paras, align)
    return sh


def label(s, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(0.1)
    tf.margin_top = tf.margin_bottom = Cm(0.05)
    fill_text(tf, paras, align)
    return tb


def arrow(s, x, y, w, h, kind="right", color=NAVY):
    shp = MSO_SHAPE.RIGHT_ARROW if kind == "right" else MSO_SHAPE.DOWN_ARROW
    a = s.shapes.add_shape(shp, Cm(x), Cm(y), Cm(w), Cm(h))
    a.fill.solid()
    a.fill.fore_color.rgb = rgb(color)
    a.line.fill.background()
    a.shadow.inherit = False
    st = a._element.find(qn("p:style"))
    if st is not None:
        st.find(qn("a:effectRef")).set("idx", "0")
    return a


def band(s, y, text, h=1.2):
    return box(s, 2.2, y, 23.11, h, [(text, 14, True, WHITE, 0)], fill=NAVY, align=PP_ALIGN.CENTER,
               anchor=MSO_ANCHOR.MIDDLE, adj=0.1, margins=(0.4, 0, 0.4, 0))


def note(s, y, text):
    return label(s, 2.2, y, 23.11, 0.8, [(text, 9, False, GRAY, 0)])


def pic(s, name, x, y, w=None, h=None):
    kw = {}
    if w:
        kw["width"] = Cm(w)
    if h:
        kw["height"] = Cm(h)
    return s.shapes.add_picture(str(IMG / name), Cm(x), Cm(y), **kw)


def set_para(sh, k, text):
    """既存の段落 k の文言だけ差し替える（1つ目の書式を残す）。"""
    while len(sh.text_frame.paragraphs) <= k:
        last = sh.text_frame.paragraphs[-1]._p
        last.addnext(copy.deepcopy(last))
    p = sh.text_frame.paragraphs[k]
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def drop_para(sh, k):
    p = sh.text_frame.paragraphs[k]
    p._p.getparent().remove(p._p)


def find(slide, name):
    return next(sh for sh in slide.shapes if sh.name == name)


def retitle(slide, title, lead1, lead2=None):
    set_para(find(slide, "TextBox 1"), 0, title)
    t2 = find(slide, "TextBox 2")
    set_para(t2, 0, lead1)
    if lead2 is None:
        if len(t2.text_frame.paragraphs) > 1:
            drop_para(t2, 1)
    else:
        set_para(t2, 1, lead2)


def new_slide(title, lead1, lead2=None):
    """ver4.1 の23枚目と同じ見出し部（タイトル・リード・区切り線）で空のページを作る。"""
    s = prs.slides.add_slide(LAYOUT)
    for sh in list(s.shapes):
        sh._element.getparent().remove(sh._element)
    for nm in ("TextBox 1", "TextBox 2", "Connector 3"):
        s.shapes._spTree.append(copy.deepcopy(find(SRC[22], nm)._element))
    retitle(s, title, lead1, lead2)
    return s


# ================= 1 表紙（ver4.1の1枚目を流用）=================
s1 = SRC[0]
sub, ttl = find(s1, "正方形/長方形 5"), find(s1, "正方形/長方形 1")
set_para(sub, 0, "―新サイトの集客力と応募率を最大化する「同時導入」プラン―")
set_para(ttl, 0, "Webサイトリニューアルに伴う")
ttl.left, ttl.width = Cm(2.7), Cm(22.1)
sub.left, sub.width = Cm(2.7), Cm(22.1)
sub.top, ttl.top = Cm(5.0), Cm(6.0)
ttl.height = Cm(3.1)
p0 = ttl.text_frame.paragraphs[0]
p1 = copy.deepcopy(p0._p)
p0._p.addnext(p1)
ttl.text_frame.paragraphs[1].runs[0].text = "「公式LINE活用・顧客ストック化」のご提案"
for p in ttl.text_frame.paragraphs:
    for r in p.runs:
        r.font.size = Pt(26)
ttl.text_frame.word_wrap = True

# ================= 2 アジェンダ =================
s2 = new_slide("本日のアジェンダ", "Webサイト改善の「先」にある、求職者を逃さないLINE連携の全体像")
agenda = [
    ("1", "Web改修後の課題とLINEの立ち位置", "サイトを新しくしても残る課題と、LINEを同時に入れる理由（3〜5・9〜11ページ）"),
    ("2", "大手（タウンワーク）事例", "求人業界でLINEが応募を増やした実例と、自動車求人サイトへの応用（6〜8ページ）"),
    ("3", "導線とユーザー体験", "新サイトとLINEをつなぐ導線・画面イメージ・自動応答（12〜15ページ）"),
    ("4", "効果と費用", "友だち獲得の試算・費用の考え方・お見積もり・スケジュール（16〜19ページ）"),
]
for i, (n, h, d) in enumerate(agenda):
    y = 4.6 + i * 3.05
    box(s2, 2.2, y, 2.6, 2.7, [(n, 28, True, WHITE, 0)], fill=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.1, 0, 0.1, 0))
    box(s2, 5.0, y, 20.31, 2.7, [(h, 17, True, NAVY, 6), (d, 13, False, INK, 0)], anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.6, 0.1, 0.4, 0.1))

# ================= 3 サイトリニューアルの「次の課題」 =================
s3 = new_slide("Webサイトリニューアルの「次の課題」",
               "サイトを新しくしても、フォーム入力の途中で離脱する求職者は必ず発生する",
               "サイト改修だけでは届かない「検討層・離脱層」への対策が必要")
box(s3, 2.2, 4.6, 11.0, 1.0, [("サイト改修で良くなること", 13, True, WHITE, 0)], fill=GRAY, align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE, margins=(0.2, 0, 0.2, 0))
box(s3, 2.2, 5.6, 11.0, 4.3, [("・見やすさ・使いやすさ", 14, False, INK, 6), ("・求人の探しやすさ", 14, False, INK, 6),
                              ("・応募フォームの入力しやすさ", 14, False, INK, 6), ("→ 応募する方が増える", 14, True, NAVY, 0)],
    fill=GRAY_BG, margins=(0.6, 0.4, 0.4, 0.1))
arrow(s3, 13.4, 7.0, 0.8, 0.8)
box(s3, 14.3, 4.6, 11.0, 1.0, [("改修しても残ること", 13, True, WHITE, 0)], fill=ORANGE, align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE, margins=(0.2, 0, 0.2, 0))
box(s3, 14.3, 5.6, 11.0, 4.3, [("・比較検討中で、まだ応募しない方", 14, False, INK, 6), ("・フォームの途中で離れた方", 14, False, INK, 6),
                               ("・応募しなかった方との接点が、その後は残らない", 14, False, INK, 0)],
    fill=ORANGE_BG, margins=(0.6, 0.4, 0.4, 0.1))
label(s3, 2.2, 10.2, 23.1, 0.8, [("整備士は、複数の経路を比べながら仕事を探している（仕事を探した経路・複数回答）", 13, True, NAVY, 0)])
for i, (nm, v) in enumerate([("転職エージェント", "35.3%"), ("総合型の求人サイト", "34.1%"), ("業界特化型の求人サイト", "33.0%")]):
    box(s3, 2.2 + i * 7.9, 11.1, 7.3, 3.4, [(nm, 13, False, INK, 4), (v, 30, True, NAVY, 0)], align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE)
band(s3, 15.0, "応募に至らなかった方を、後から追える仕組みが必要")
note(s3, 16.4, "出典：モビリア総研調査（直近3年に転職した整備士1,249人・Response 2026年9月23日）")

# ================= 4 解決策としてのWeb×LINE =================
s4 = new_slide("解決策｜Web × LINEの同時導入",
               "LINEを同時に導入し、「友だち追加＝会員・見込み客のストック」として獲得",
               "応募まで至らない方を貯めて、後から案内できる仕組みにする")
box(s4, 2.2, 4.6, 4.2, 6.6, [("新サイトに", 16, True, WHITE, 2), ("来訪", 16, True, WHITE, 0)], fill=NAVY, align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE, margins=(0.1, 0, 0.1, 0))
arrow(s4, 6.55, 5.5, 0.7, 0.7)
arrow(s4, 6.55, 9.6, 0.7, 0.7)
box(s4, 7.4, 4.6, 5.3, 2.8, [("応募する方", 15, True, NAVY, 0)], fill=CARD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
box(s4, 7.4, 8.4, 5.3, 2.8, [("応募に至らない方", 15, True, INK, 4), ("（検討層・離脱層）", 12, False, INK, 0)], fill=GRAY_BG,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s4, 12.85, 5.5, 0.7, 0.7)
arrow(s4, 12.85, 9.3, 0.7, 0.7)
box(s4, 13.7, 4.6, 5.6, 2.8, [("サイトで応募", 15, True, NAVY, 4), ("（CV②）", 12, False, INK, 0)], fill=CARD, align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE)
box(s4, 13.7, 8.4, 5.6, 2.8, [("LINEで友だち追加", 14.5, True, WHITE, 4), ("＝ストック", 12, False, WHITE, 0)], fill=GREEN,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s4, 19.45, 9.3, 0.7, 0.7)
box(s4, 20.3, 8.4, 5.0, 2.8, [("新着求人を配信", 13, True, GREEN_TX, 2), ("→ 応募", 13, True, GREEN_TX, 0)], fill=GREEN_BG,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for i, (h, d) in enumerate([("① 後から案内できる", "友だちに、新着求人や条件に合う求人をいつでも届けられる"),
                            ("② 会員の予備軍が貯まる", "今すぐ応募しない方も、転職を検討する方のリストに残る"),
                            ("③ 興味が分かる", "配信のクリックで、整備士・営業など興味別に分けられる")]):
    box(s4, 2.2 + i * 7.9, 11.9, 7.3, 3.0, [(h, 14, True, NAVY, 6), (d, 12.5, False, INK, 0)])
band(s4, 15.4, "応募まで至らない方をストックし、後から応募へつなげる")

# ================= 5 なぜ「今（サイトオープン時）」同時設計か =================
s5 = new_slide("なぜ「サイトのオープン時」に同時設計するのか",
               "オープン後のアクセス増加期を逃さない",
               "初日からデータを蓄積するのが、最も効率的")
for i, (h, d1, d2) in enumerate([
        ("① 初日からデータを蓄積", "オープン直後から、友だち・会員・クリックの履歴を蓄積。", "半年で約540人の友だちに（弊社シミュレーション）"),
        ("② 作り直しを避ける", "会員登録のLINE化を、リニューアルの要件に入れておけば、", "サイトの開発が二度手間にならない。"),
        ("③ 取りこぼしを防ぐ", "後から導入すると、それまでに来訪した方とは", "LINEでつながれず、機会損失になる。")]):
    box(s5, 2.2 + i * 7.9, 4.6, 7.3, 5.2, [(h, 15, True, NAVY, 8), (d1, 13.5, False, INK, 6), (d2, 13.5, False, INK, 0)])
box(s5, 2.2, 10.5, 3.6, 1.1, [("同時に導入", 12, True, WHITE, 0)], fill=GREEN, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    margins=(0.1, 0, 0.1, 0))
box(s5, 6.0, 10.5, 19.3, 1.1, [("サイトのオープン　────　友だち・会員が増え続ける　────▶", 12, True, WHITE, 0)], fill=GREEN,
    anchor=MSO_ANCHOR.MIDDLE, adj=0.2, margins=(0.5, 0, 0.3, 0))
box(s5, 2.2, 12.3, 3.6, 1.1, [("後付けで導入", 12, True, WHITE, 0)], fill=GRAY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    margins=(0.1, 0, 0.1, 0))
box(s5, 6.0, 12.3, 10.6, 1.1, [("この期間の来訪者とは、LINEでつながれない", 12, True, INK, 0)], fill=GRAY_BG, anchor=MSO_ANCHOR.MIDDLE,
    adj=0.2, margins=(0.5, 0, 0.3, 0))
box(s5, 16.8, 12.3, 8.5, 1.1, [("導入後から蓄積", 12, True, WHITE, 0)], fill=GREEN, anchor=MSO_ANCHOR.MIDDLE, adj=0.2,
    margins=(0.5, 0, 0.3, 0))
band(s5, 14.4, "サイトと同時に設計するのが、最も効率的で、取りこぼしがない")
note(s5, 15.9, "※友だち数は弊社シミュレーション（リニューアル前に貯まる友だちは含めない）")

# ================= 6 タウンワーク事例（ver4.1の4枚目を流用）=================
retitle(SRC[3], "【事例】求人業界におけるLINE活用のインパクト",
        "「タウンワーク」は、LINEリッチメニューからの検索強化で、応募数が79%増",
        "総合大手の求人サービスは、数十万〜数百万人規模でLINEを活用している")
t = find(SRC[3], "TextBox 14")
set_para(t, 0, "出典：LINEヤフー for Business 導入事例（タウンワーク・2019年公開）／友だち数：page.line.me（公式）2026-09-21実測")

# ================= 7 タウンワークの勝因（ver4.1の5枚目を流用）=================
retitle(SRC[4], "タウンワークの勝因｜常設の入口（リッチメニュー）",
        "勝因は、LINE上でワンタップで求人を探せる常設の入口",
        "LINE経由の応募は、配信よりリッチメニュー経由が最多")

# ================= 8 自動車求人サイトへの応用 =================
s8 = new_slide("自動車求人サイトへの応用｜LINEのワンタップ検索",
               "「資格」「職種」「こだわり条件」は、LINEのワンタップ検索と相性が良い",
               "「整備士2級・3級」「未経験OK」「土日休み」など、すぐに条件で引ける")
for i, (h, d) in enumerate([("資格", "整備士1級・2級・3級・無資格"), ("職種", "整備士／営業／受付・事務"),
                            ("こだわり条件", "未経験OK・資格取得支援あり・土日休みなど")]):
    box(s8, 2.2 + i * 7.9, 4.6, 7.3, 3.6, [(h, 16, True, NAVY, 8), (d, 14, False, INK, 0)], align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE)
box(s8, 2.2, 8.9, 6.6, 2.4, [("リッチメニューの", 13, True, WHITE, 0), ("ボタンをタップ", 13, True, WHITE, 0)], fill=GREEN,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s8, 9.0, 9.7, 0.8, 0.8)
box(s8, 9.9, 8.9, 7.6, 2.4, [("条件を入れた状態の", 13, True, NAVY, 0), ("検索結果ページが開く", 13, True, NAVY, 0)], fill=CARD,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s8, 17.7, 9.7, 0.8, 0.8)
box(s8, 18.6, 8.9, 6.7, 2.4, [("気になる求人で", 13, True, NAVY, 0), ("サイトから応募", 13, True, NAVY, 0)], fill=CARD,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
box(s8, 2.2, 12.0, 23.11, 2.6, [("掲載中の求人は、全国2,134件", 15, True, NAVY, 6),
                                ("（うち正社員・整備士946件、正社員・営業180件）。LINEから、条件別に探せる形にする", 13.5, False, INK, 0)],
    anchor=MSO_ANCHOR.MIDDLE, margins=(0.7, 0.1, 0.4, 0.1))
band(s8, 15.1, "資格・職種・条件を、LINEのワンタップで求人へ直結")
note(s8, 16.5, "出典：自動車求人Navi 検索画面（2026年9月28日時点の掲載件数）")

# ================= 9 現状のボトルネック =================
s9 = new_slide("現状のWebサイトにおけるボトルネック",
               "現行サイトは「その場の応募」のみで、長期的な求職者データベースが残らない構造",
               "会員登録の入力の手間も、登録の壁になっている")
box(s9, 2.2, 4.6, 4.6, 2.2, [("サイトに来訪", 14, True, WHITE, 0)], fill=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s9, 7.0, 5.3, 0.8, 0.8)
box(s9, 8.0, 4.6, 8.6, 2.2, [("入口は「応募」「転職支援の申込」", 13, True, NAVY, 0)], fill=CARD, align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE)
arrow(s9, 16.8, 5.3, 0.8, 0.8)
box(s9, 17.8, 4.6, 7.5, 2.2, [("応募した方だけ", 14, True, NAVY, 2), ("連絡先が残る", 14, True, NAVY, 0)], fill=CARD,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
box(s9, 2.2, 7.5, 4.6, 2.2, [("応募しなかった方", 13, True, INK, 0)], fill=GRAY_BG, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s9, 7.0, 8.2, 0.8, 0.8)
box(s9, 8.0, 7.5, 17.3, 2.2, [("接点が残らず、後から案内できない（再来訪を待つしかない）", 14, True, ORANGE, 0)], fill=ORANGE_BG,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for i, (h, d) in enumerate([("① 検討層の受け皿がない", "「応募」か「転職支援の申込」しか入口がなく、比較検討中の方と接点を持てない"),
                            ("② 登録が使われにくい", "氏名・経歴・パスワードの設定など、入力の手間が登録の壁になる"),
                            ("③ 名簿が残らない", "応募しなかった方の情報が蓄積されず、長期的に育てられない")]):
    box(s9, 2.2 + i * 7.9, 10.4, 7.3, 4.5, [(h, 14, True, NAVY, 8), (d, 13, False, INK, 0)])
band(s9, 15.4, "「その場の応募」の先に、求職者のデータベースを残す必要がある")

# ================= 10 エントリーハードルの低減（ver4.1の8枚目を流用）=================
retitle(SRC[7], "LINEによるエントリーハードルの低減",
        "「LINEでログイン」とLINE Profile+で、登録の入力は「確認して押すだけ」に",
        "ワンタップの友だち追加で、登録の心理的ハードルを大きく下げる")

# ================= 11 ストック顧客へのアプローチ =================
s11 = new_slide("獲得したストック顧客へのアプローチ",
                "蓄積した友だちに、新着求人や条件に合う求人を、LINEのメッセージで直接届ける",
                "メルマガより反応されやすく、じっくり検討する方を応募へ引き上げる")
steps = [("① 友だち（ストック）", "応募前の方も含めて、LINEで接点が残る"),
         ("② クリックで自動分類", "整備士・営業・未経験OKなど、配信のクリックで自動的に分かれる"),
         ("③ 興味に合う求人を配信", "新着求人・条件に合う求人を届け、応募へ")]
for i, (h, d) in enumerate(steps):
    x = 2.2 + i * 8.0
    box(s11, x, 4.6, 7.0, 3.8, [(h, 14.5, True, NAVY, 6), (d, 12.5, False, INK, 0)])
    if i < 2:
        arrow(s11, x + 7.1, 6.0, 0.8, 0.8)
label(s11, 2.2, 8.9, 23.1, 0.8, [("他業種・就職支援での、LINE配信の反応（LINEヤフー公式事例）", 13, True, NAVY, 0)])
for i, (cap, v, sub) in enumerate([("バス会社｜メッセージの開封率", "約70%", "LINE通知メッセージの配信後"),
                                   ("バス会社｜クリック率", "約5倍", "メルマガのクリック率との比較"),
                                   ("就職支援｜LINE問い合わせ後", "55%", "の方が、面談を予約")]):
    box(s11, 2.2 + i * 7.9, 9.8, 7.3, 4.2, [(cap, 12.5, False, INK, 6), (v, 30, True, NAVY, 6), (sub, 11, False, GRAY, 0)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
band(s11, 14.6, "ストックした友だちに、求人を何度でも届けられる")
note(s11, 16.0, "出典：LINEヤフー for Business 導入事例（琴平バス・UZUZ）")

# ================= 12 新サイト⇄LINEの連動導線 =================
s12 = new_slide("新サイトとLINEの連動導線（全体フロー）",
                "新サイトのあらゆる接点から、自然にLINEへ誘導",
                "求人詳細・キープ画面・相談窓口などから、行き来しながら応募へ")
cols = [("求人詳細", ["離脱防止ポップアップ", "「LINEで登録」"]),
        ("キープ画面", ["「気になる求人を", "LINEでキープ」"]),
        ("相談窓口", ["「LINEで適正年収を", "相談」"]),
        ("応募の完了画面", ["サンクスLINE", "（オプション）"])]
cw, gap = 5.4, 0.5
for i, (pg, ls) in enumerate(cols):
    x = 2.2 + i * (cw + gap)
    box(s12, x, 4.6, cw, 1.6, [(pg, 14, True, NAVY, 0)], fill=CARD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        line=NAVY)
    arrow(s12, x + cw / 2 - 0.3, 6.3, 0.6, 0.6, "down")
    box(s12, x, 7.0, cw, 2.4, [(t, 12, True, GREEN_TX, 2) for t in ls], fill=GREEN_BG, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE)
    arrow(s12, x + cw / 2 - 0.3, 9.5, 0.6, 0.6, "down")
box(s12, 2.2, 10.2, 23.11, 1.4, [("LINE公式アカウント　友だち追加＝会員・見込み客のストック", 15, True, WHITE, 0)], fill=GREEN,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, adj=0.1, margins=(0.4, 0, 0.4, 0))
arrow(s12, 13.5, 11.7, 0.6, 0.6, "down")
for i, (h, d) in enumerate([("あいさつ", "会員連携を案内"), ("リッチメニュー", "条件検索・相談"), ("配信", "新着求人・条件別")]):
    box(s12, 2.2 + i * 6.0, 12.4, 5.5, 2.0, [(h, 13.5, True, NAVY, 2), (d, 12, False, INK, 0)], align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE)
arrow(s12, 20.3, 13.1, 0.7, 0.7)
box(s12, 21.2, 12.4, 4.1, 2.0, [("サイトへ戻って", 12.5, True, WHITE, 2), ("応募", 12.5, True, WHITE, 0)], fill=NAVY,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
band(s12, 15.0, "サイトとLINEを行き来しながら、応募へつなげる")
note(s12, 16.4, "※キープ画面・相談窓口は、新サイトに追加する接点の想定。シミュレーションの数字に含めるのは動線00・01のみ（動線02は含めない）")

# ================= 13 ユーザー体験イメージ① =================
s13 = new_slide("ユーザー体験イメージ①｜サイト訪問からLINE保存まで",
                "気軽なポップアップ・ボタンで、サイトからその場でLINEへ",
                "「LINEでキープ」「LINEで相談」など、求人に合わせた訴求")
pic(s13, "chengrowth_v11_popup.png", 2.2, 4.6, w=11.2)
pic(s13, "chengrowth_v11_register.png", 14.1, 4.6, w=11.2)
label(s13, 2.2, 11.8, 11.2, 2.2, [("離脱防止ポップアップ（動線00）", 14, True, NAVY, 4),
                                  ("ページを離れようとした方に、「条件に合う新着求人をLINEでお届け」と案内", 12, False, INK, 0)])
label(s13, 14.1, 11.8, 11.2, 2.2, [("「LINEで登録」ボタン（動線01）", 14, True, NAVY, 4),
                                   ("LINEの情報が自動で入り、会員登録と友だち追加が同時に完了", 12, False, INK, 0)])
label(s13, 2.2, 14.2, 23.1, 0.8, [("ポップアップ・ボタンの訴求の例", 12.5, True, NAVY, 0)])
for i, tx in enumerate(["「気になる求人をLINEでキープ」", "「LINEで適正年収を相談」", "「新着求人をLINEでお届け」"]):
    box(s13, 2.2 + i * 7.9, 15.0, 7.3, 1.3, [(tx, 11, True, GREEN_TX, 0)], fill=GREEN_BG, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE, adj=0.2, margins=(0.2, 0, 0.2, 0))
note(s13, 16.6, "※画面はイメージです")

# ================= 14 ユーザー体験イメージ② =================
s14 = new_slide("ユーザー体験イメージ②｜LINEトーク画面・リッチメニュー",
                "自動車求人に特化したリッチメニューを常設",
                "希望条件の検索から相談まで、LINEの中で完結")
pic(s14, "chengrowth_v12_richmenu.png", 2.2, 4.6, w=11.4)
label(s14, 2.2, 12.3, 11.4, 2.6, [("リッチメニュー（トーク画面の下に常設）", 14, True, NAVY, 4),
                                  ("整備士の求人／営業・受付の求人／近くの求人／資格から探す／未経験OK・資格取得支援／新着求人を受け取る", 12, False, INK, 0)])
pic(s14, "chengrowth_v11_talk_step.png", 14.4, 4.4, h=10.4)
box(s14, 20.7, 4.6, 4.6, 10.2, [("LINE内で完結", 14, True, NAVY, 8), ("① 希望の職種をタップで選ぶ", 12.5, False, INK, 8),
                                ("② 条件に合う求人が、カードで届く", 12.5, False, INK, 8),
                                ("③ 求人ページへそのまま移動し、サイトで応募", 12.5, False, INK, 0)], margins=(0.4, 0.4, 0.3, 0.1))
note(s14, 15.2, "※画面はイメージです")

# ================= 15 自動応答・アンケート =================
s15 = new_slide("自動応答・アンケートシナリオの活用",
                "友だち追加直後の自動アンケートで、「保有資格」「希望時期」を把握",
                "回答に合う求人を、自動でセグメント配信")
flow = [("友だち追加", None), ("あいさつ（自動）", "会員連携を案内"), ("アンケート", "ボタンをタップするだけ"),
        ("回答をタグで保存", "資格・希望時期"), ("条件に合う求人を配信", "自動")]
fw, fg = 4.0, 0.78
for i, (h, d) in enumerate(flow):
    x = 2.2 + i * (fw + fg)
    paras = [(h, 12.5, True, WHITE if i in (0, 4) else NAVY, 2)] + ([(d, 10.5, False, INK, 0)] if d else [])
    if i in (0, 4):
        paras = [(h, 12.5, True, WHITE, 0)]
    box(s15, x, 4.6, fw, 2.4, paras, fill=GREEN if i in (0, 4) else CARD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.15, 0.1, 0.15, 0.1))
    if i < 4:
        arrow(s15, x + fw + 0.08, 5.5, 0.62, 0.62)
box(s15, 2.2, 7.7, 9.6, 6.9, [("アンケートの例", 14, True, NAVY, 8),
                              ("保有資格：1級／2級／3級／無資格", 12.5, False, INK, 6),
                              ("希望時期：すぐ／3か月以内／情報収集中", 12.5, False, INK, 6),
                              ("希望の職種：整備士／営業／受付・事務", 12.5, False, INK, 0)])
box(s15, 12.0, 7.7, 9.6, 6.9, [("回答に合わせた配信の例", 14, True, NAVY, 8),
                               ("整備士2級・すぐ → 整備士の新着求人を優先して配信", 12.5, False, INK, 6),
                               ("未経験・情報収集中 → 資格取得支援の求人と、働き方の情報を配信", 12.5, False, INK, 6),
                               ("友だち追加から14日間、全10通のステップ配信で応募へ", 12.5, True, NAVY, 0)])
pic(s15, "chengrowth_v11_talk_step.png", 21.8, 7.5, h=7.2)
band(s15, 15.2, "追加直後の数十秒で、求める条件が分かり、最適な求人を届けられる")
note(s15, 16.6, "※画面・選択肢はイメージです")

# ================= 16 友だち登録数・獲得シミュレーション =================
s16 = new_slide("友だち登録数・獲得シミュレーション",
                "新サイトの来訪（月約6,100人）に応じて、友だちは半年で542人まで増加",
                "転職を検討する方のリストが、公開直後から積み上がる")
months = ["4月", "5月", "6月", "7月", "8月", "9月"]
friends = [63, 143, 235, 335, 436, 542]
cd = CategoryChartData()
cd.categories = months
cd.add_series("友だち（累計）", friends)
gf = s16.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Cm(2.2), Cm(4.3), Cm(23.1), Cm(5.6), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
ch.font.size = Pt(12)
ch.font.name = "メイリオ"
pl = ch.plots[0]
pl.gap_width = 60
pl.has_data_labels = True
dl = pl.data_labels
dl.number_format = '#,##0"人"'
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(13)
dl.font.bold = True
dl.font.name = "メイリオ"
dl.font.color.rgb = rgb(NAVY)
ser = pl.series[0]
ser.format.fill.solid()
ser.format.fill.fore_color.rgb = rgb(GREEN)
va = ch.value_axis
va.visible = False
va.has_major_gridlines = False
va.maximum_scale = 650
va.minimum_scale = 0
ca = ch.category_axis
ca.tick_labels.font.size = Pt(12)
ca.tick_labels.font.name = "メイリオ"
ca.format.line.color.rgb = rgb("D9D9D9")
# 表（ver4.1の27枚目の表を複製し、応募単価の行を除く）
tbl_src = next(sh for sh in SRC[26].shapes if sh.has_table)
tbl = copy.deepcopy(tbl_src._element)
s16.shapes._spTree.append(tbl)
trs = tbl.findall(".//" + qn("a:tr"))
trs[-1].getparent().remove(trs[-1])
trs = tbl.findall(".//" + qn("a:tr"))
tbl.find(qn("p:xfrm")).find(qn("a:off")).set("y", str(Cm(10.2)))
tbl.find(qn("p:xfrm")).find(qn("a:ext")).set("cy", str(sum(int(t.get("h")) for t in trs)))
tbl.find(qn("p:nvGraphicFramePr")).find(qn("p:cNvPr")).set("id", "900")
note(s16, 15.1, "前提：リニューアル前に貯まる友だちは含めない／サイト来訪 月約6,100人（広告のシミュレーション確定後に差し替え）／LINEで登録＝来訪者の1.0%／会員登録（CV①）・応募（CV②）はLINEでの登録・リッチメニューと配信からの件数")

# ================= 17 導入費用の算出考え方 =================
s17 = new_slide("導入費用の考え方｜サイト制作費とは分けて明示",
                "投資対効果を正しく把握するため、サイト制作費とは分ける",
                "LINEの費用を、「LINE独自構築費」として明確化")
cols17 = [("サイト改修費", "サイト制作会社", "新サイトの見た目・使いやすさ・応募フォームを改善する", "制作会社のお見積もり（サイトの制作費）", GRAY, GRAY_BG),
          ("広告費", "各広告媒体", "サイトへ来訪する方を増やす", "媒体ごとの広告費・運用費", GRAY, GRAY_BG),
          ("LINE独自構築費", "弊社（本提案）", "来訪した方を友だち・会員として貯め、応募へつなげる",
           "初期費用＋月額の固定費／LINE公式アカウントの利用料は配信通数で変動（別途）／LINE連携の開発は都度お見積もり", NAVY, CARD)]
for i, (h, who, role, cost, hc, bg) in enumerate(cols17):
    x = 2.2 + i * 7.9
    box(s17, x, 4.6, 7.3, 1.3, [(h, 15, True, WHITE, 0)], fill=hc, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.2, 0, 0.2, 0))
    box(s17, x, 5.9, 7.3, 8.2, [("担い手", 11, True, GRAY, 2), (who, 13.5, True, INK, 10), ("役割", 11, True, GRAY, 2), (role, 13, False, INK, 10),
                                ("費用の構造", 11, True, GRAY, 2), (cost, 12, False, INK, 0)], fill=bg)
band(s17, 14.8, "費用を分けることで、LINEの効果（友だち・会員・応募）を単独で把握できる")

# ================= 18 概算お見積もり（ver4.1の28枚目を流用）=================
retitle(SRC[27], "概算お見積もり",
        "初期構築・シナリオ設計・リッチメニュー制作のパッケージ費用",
        "内訳と費用感を明示し、サイト制作費とは分けて整理")

# ================= 19 スケジュール（ver4.1の25枚目を流用）=================
retitle(SRC[24], "スケジュール｜サイトの公開日に合わせて、LINEを同時にリリース",
        "サイトの公開日（4月）に合わせ、LINEの構築・動作検証を並行して進める")


def replace_par(slide, old, new):
    n = 0
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                if p.text == old:
                    p.runs[0].text = new
                    for r in p.runs[1:]:
                        r._r.getparent().remove(r._r)
                    n += 1
    assert n, old


replace_par(SRC[24], "お申込み・申請", "要件定義・お申込み・申請")
replace_par(SRC[24], "初期構築（この期間は初期費用のみ）", "デザイン・初期構築（この期間は初期費用のみ）")
replace_par(SRC[24], "あいさつ・リッチメニュー・配信の設計、各ツールとLINEの連携、離脱防止の設置",
            "あいさつ・リッチメニュー・配信の設計とデザイン、各ツールとLINEの連携、離脱防止の設置")
replace_par(SRC[24], "LINEを追加してもらう期間", "テスト・動作検証（公開前）")
replace_par(SRC[24], "離脱防止ポップアップなどで、LINEを追加してもらう（数字はシミュレーションに含めない）",
            "友だち追加から配信までの動作を検証（離脱防止の先行稼働分は、シミュレーションに含めない）")
replace_par(SRC[24], "リニューアル公開と同時に運用を本格化（月額の固定費）", "サイト公開と同時にLINEもリリース（運用を本格化・月額の固定費）")

# ---------------- 並べ替え・不要ページの削除 ----------------
order = [SRC[0], s2, s3, s4, s5, SRC[3], SRC[4], s8, s9, SRC[7], s11, s12, s13, s14, s15, s16, s17, SRC[27], SRC[24]]
lst = prs.slides._sldIdLst
want = [s.slide_id for s in order]
elems = {int(sid.get("id")): sid for sid in lst}
for sid_val, el in list(elems.items()):
    if sid_val not in want:
        prs.part.drop_rel(el.get(qn("r:id")))
        lst.remove(el)
for el in list(lst):
    lst.remove(el)
for sid_val in want:
    lst.append(elems[sid_val])

# ---------------- 字体メイリオ・最小8pt ----------------
AFTER = {qn(x) for x in ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")}
for sl in prs.slides:
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for rpr in sl.shapes._spTree.iter(tag.replace("a:", "{http://schemas.openxmlformats.org/drawingml/2006/main}")):
            if rpr.get("sz") and int(rpr.get("sz")) < 800:
                rpr.set("sz", "800")
            fs = []
            for ft in ("a:latin", "a:ea", "a:cs"):
                e = rpr.find(qn(ft))
                if e is None:
                    e = rpr.makeelement(qn(ft), {})
                else:
                    rpr.remove(e)
                e.set("typeface", "メイリオ")
                fs.append(e)
            nxt = next((c for c in rpr if c.tag in AFTER), None)
            for e in fs:
                (nxt.addprevious(e) if nxt is not None else rpr.append(e))
prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
