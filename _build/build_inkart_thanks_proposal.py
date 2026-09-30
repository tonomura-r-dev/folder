# -*- coding: utf-8 -*-
"""インクアート様（プリントアース）10月以降のご提案 12枚（簡易版）

サンクスLINE誘導＋離脱防止をセットで提案する。前半は9月のLINE配信分析の結果。
ベースは _templates/DYM_LINEOA_BUFFF_62p.pptx（原本は編集しない。コピーして使う）。

  BUFFF 1・6〜13・27・31・62 の12枚だけ残し、
  「1, 6, 7, 8, 9, 10, 11, 31, 27, 12, 13, 62」の順に並べる。
  6〜13 は clear_slide() して作り直す（スライドの新規追加はしない）。

使い方:
    pip install python-pptx
    python _build/build_inkart_thanks_proposal.py [YYYYMMDD]
    python _build/qa_render.py <出力.pptx> _qa
"""
import re
import shutil
import sys
from copy import deepcopy
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_BUFFF_62p.pptx"
STAMP = sys.argv[1] if len(sys.argv) > 1 else date.today().strftime("%Y%m%d")
OUT = ROOT / f"{STAMP}_インクアート様_LINE公式アカウント_10月以降のご提案.pptx"

TNAVY = "002060"   # タイトル
NAVY = "1F285A"    # チップ・図の箱
INK = "333333"     # 本文
PALE = "F4F7FF"    # カード
BORDER = "D9D9D9"  # 枠
RED = "C00000"     # 課題の数字だけ
BAR = "3467B2"     # グラフの棒
MUT = "808080"     # 脚注
WHITE = "FFFFFF"
FONT = "メイリオ"

ORDER = [1, 6, 7, 8, 9, 10, 11, 31, 27, 12, 13, 62]   # BUFFFでのページ番号

shutil.copyfile(SRC, OUT)
prs = Presentation(str(OUT))
assert prs.slide_width == 9906000 and prs.slide_height == 6858000
src_slides = list(prs.slides)
assert len(src_slides) == 62, len(src_slides)

# ---- 素材キャプチャ：BUFFF 6ページの区切り線（y=1.52inch・全幅） ----
div_el = None
for sh in src_slides[5].shapes:
    if sh._element.tag.endswith("}cxnSp") and abs(sh.top - 1390675) < 20000:
        div_el = deepcopy(sh._element)
        break
assert div_el is not None, "divider not found"

# ---- 構造作業：12枚を残して並べ替え（追加はしない） ----
sld_id_lst = prs.slides._sldIdLst
ids = list(sld_id_lst)
for n, el in enumerate(ids, 1):
    if n not in ORDER:
        prs.part.drop_rel(el.rId)
        sld_id_lst.remove(el)
by_page = {n: ids[n - 1] for n in ORDER}
for el in list(sld_id_lst):
    sld_id_lst.remove(el)
for n in ORDER:
    sld_id_lst.append(by_page[n])
S = list(prs.slides)
assert len(S) == 12, len(S)


# ---------- helpers（build_special_plan.py と同じ経路） ----------
def set_ea_cs(rPr, name=FONT):
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", name)
        for attr in ("panose", "pitchFamily", "charset"):
            if attr in e.attrib:
                del e.attrib[attr]
    # latin/ea/cs の順番を揃える（スキーマ順）
    for tag in ("a:latin", "a:ea", "a:cs", "a:sym"):
        e = rPr.find(qn(tag))
        if e is not None:
            rPr.remove(e)
            rPr.append(e)


def set_font(run, size, bold=None, color=INK, name=FONT):
    f = run.font
    f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    f.color.rgb = RGBColor.from_string(color)
    set_ea_cs(run._r.get_or_add_rPr(), name)


ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCH = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def put_text(tf, paras, anchor="t", ml=0.06, mr=0.06, mt=0.03, mb=0.03, wrap=True):
    tf.word_wrap = wrap
    tf.margin_left = Inches(ml)
    tf.margin_right = Inches(mr)
    tf.margin_top = Inches(mt)
    tf.margin_bottom = Inches(mb)
    tf.vertical_anchor = ANCH[anchor]
    first = True
    for p in paras:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.alignment = ALIGN[p.get("align", "l")]
        if p.get("sa") is not None:
            para.space_after = Pt(p["sa"])
        if p.get("ls") is not None:
            para.line_spacing = p["ls"]
        for (t, sz, b, c) in p["runs"]:
            r = para.add_run()
            r.text = t
            set_font(r, sz, b, c)
    return tf


def add_text(slide, x, y, w, h, paras, anchor="t", **kw):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    put_text(box.text_frame, paras, anchor=anchor, **kw)
    return box


def add_box(slide, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = RGBColor.from_string(line)
        sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    sp.text_frame.word_wrap = True
    return sp


def chip(slide, x, y, w, h, paras, fill=NAVY):
    """紺の箱に白抜き文字。paras は put_text と同じ形"""
    c = add_box(slide, x, y, w, h, fill=fill)
    put_text(c.text_frame, paras, anchor="m", ml=0.08, mr=0.08, mt=0.02, mb=0.02)
    return c


def arrow(slide, x, y, w=0.45, h=0.50):
    return add_box(slide, x, y, w, h, fill=NAVY, shape=MSO_SHAPE.RIGHT_ARROW)


def clear_slide(slide):
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)
    # 消した画像などへの参照を外す（レイアウト・ノートは残す）
    xml = slide._element.xml
    for rId, rel in list(slide.part.rels.items()):
        if rel.reltype.endswith(("/slideLayout", "/notesSlide")):
            continue
        if f'"{rId}"' not in xml:
            slide.part.drop_rel(rId)


def frame(slide, header):
    """タイトル＋区切り線（BUFFF 27・31ページと同じ位置）"""
    clear_slide(slide)
    slide.shapes._spTree.append(deepcopy(div_el))
    add_text(slide, 0.60, 0.15, 8.10, 0.37,
             [{"runs": [(header, 16, True, TNAVY)]}],
             anchor="m", ml=0.10, mr=0.10, mt=0.05, mb=0.05)


def footnote(slide, text):
    add_text(slide, 0.60, 6.83, 9.63, 0.32,
             [{"runs": [(text, 9, False, MUT)]}],
             anchor="t", ml=0.0, mr=0.0, mt=0.0, mb=0.0)


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # GROUP
            yield from iter_shapes(sh.shapes)


def shape_by_id(slide, shape_id):
    for sh in iter_shapes(slide.shapes):
        if sh.shape_id == shape_id:
            return sh
    raise KeyError(f"shape_id={shape_id} not found")


def set_lines(slide, shape_id, text):
    """先頭runの書式を引き継いだまま文字だけ差し替える（build_generic_lineoa.py と同じ）"""
    sh = shape_by_id(slide, shape_id)
    tf = sh.text_frame
    p0 = tf.paragraphs[0]
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    for r in p0.runs[1:]:
        r._r.getparent().remove(r._r)
    run = p0.runs[0]
    run.text = text
    set_ea_cs(run._r.get_or_add_rPr())
    end = p0._p.find(qn("a:endParaRPr"))
    if end is not None:
        set_ea_cs(end)
    return sh


# ---------- 表 ----------
def cell_border(cell, color=BORDER, w_pt=1.0):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = tcPr.makeelement(qn(tag), {"w": str(int(Pt(w_pt))), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
        sf = ln.makeelement(qn("a:solidFill"), {})
        clr = sf.makeelement(qn("a:srgbClr"), {"val": color})
        sf.append(clr)
        ln.append(sf)
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
        tcPr.append(ln)


def fill_cell(cell, paras, fill, anchor="m"):
    cell_border(cell)                    # 枠 → 塗りの順（tcPrの要素順を守る）
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor.from_string(fill)
    tf = cell.text_frame
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    for r in tf.paragraphs[0].runs:
        r._r.getparent().remove(r._r)
    cell.margin_left = Inches(0.20)
    cell.margin_right = Inches(0.10)
    cell.margin_top = Inches(0.04)
    cell.margin_bottom = Inches(0.04)
    cell.vertical_anchor = ANCH[anchor]
    first = True
    for p in paras:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.alignment = ALIGN[p.get("align", "l")]
        for (t, sz, b, c) in p["runs"]:
            r = para.add_run()
            r.text = t
            set_font(r, sz, b, c)


# ---------- グラフ ----------
def chart_fonts(chart):
    """グラフ内の文字もメイリオ（latin/ea/cs）に揃える"""
    cs = chart._chartSpace
    for tag in ("a:defRPr", "a:rPr"):
        for rPr in cs.iter(qn(tag)):
            set_ea_cs(rPr)


# =====================================================================
# S1 表紙（BUFFF 1ページ。文字だけ差し替え）
# =====================================================================
s = S[0]
t = set_lines(s, 2, "LINE公式アカウント　10月以降のご提案")
# 32ptのままだと下線（幅7.91inch）より文字が長くなるため30ptに下げ、中央に置き直す
for r in t.text_frame.paragraphs[0].runs:
    r.font.size = Pt(30)
t.left, t.width = Inches(0.90), Inches(9.03)
set_lines(s, 6, "LINEの友だちを「買ったお客様」で増やす")
b = set_lines(s, 3, "インクアート株式会社 御中　2026年10月")
# 元の箱（幅3.07inch）には収まらないので、中心を保ったまま幅を広げる
cx = b.left + b.width // 2
b.width = Inches(4.60)
b.left = cx - b.width // 2

# =====================================================================
# S2 9月の結果①
# =====================================================================
s = S[1]
frame(s, "9月の結果①　絞り込み配信で、ブロックが減りました")
COL_L, COL_R = (3.15, 3.15), (6.93, 3.20)   # (x, w)
chip(s, COL_L[0] + 0.15, 1.80, COL_L[1] - 0.30, 0.50,
     [{"runs": [("全員に配信", 18, True, WHITE)], "align": "c"}], fill=MUT)
chip(s, COL_R[0] + 0.15, 1.80, COL_R[1] - 0.30, 0.50,
     [{"runs": [("絞り込み配信（9月）", 18, True, WHITE)], "align": "c"}])
ROWS = [
    (2.50, "1回の配信で増えたブロック", "19〜37人", "数人"),
    (4.55, "開封率", "17〜21%", "30〜36%"),
]
for y, label, before, after in ROWS:
    h = 1.80
    add_box(s, 0.60, y, 9.63, h, fill=PALE, line=BORDER)
    add_text(s, 0.85, y, 2.25, h,
             [{"runs": [(label, 20, True, INK)]}], anchor="m", ml=0.0)
    # 大きな数字は折り返さない（「19〜37 / 人」のように割れるのを防ぐ）
    add_text(s, COL_L[0], y, COL_L[1], h,
             [{"runs": [(before, 40, True, INK)], "align": "c"}], anchor="m", wrap=False)
    arrow(s, 6.36, y + h / 2 - 0.25, 0.45, 0.50)
    add_text(s, COL_R[0], y, COL_R[1], h,
             [{"runs": [(after, 40, True, NAVY)], "align": "c"}], anchor="m", wrap=False)
footnote(s, "出典：LINE公式アカウント管理画面。ブロックは6〜8月＝配信当日の増加、"
            "9月＝配信日から3日間の増加（配信の無い3日間でも約4人増える）。"
            "開封率は7〜8月の全員配信と9月の絞り込み配信。")

# =====================================================================
# S3 9月の結果②
# =====================================================================
s = S[2]
frame(s, "9月の結果②　申し込みは、リッチメニューから入っています")
FLOW = ["配信で開く", "リッチメニュー", "会員登録・購入"]
bw, gap = 2.65, 0.84
x = 0.60 + (9.63 - (bw * 3 + gap * 2)) / 2
for i, label in enumerate(FLOW):
    chip(s, x, 1.95, bw, 1.10, [{"runs": [(label, 20, True, WHITE)], "align": "c"}])
    if i < 2:
        arrow(s, x + bw + (gap - 0.45) / 2, 1.95 + 0.30, 0.45, 0.50)
    x += bw + gap
add_box(s, 0.60, 3.55, 9.63, 2.60, fill=PALE, line=BORDER)
add_text(s, 0.60, 3.55, 9.63, 2.60, [
    {"runs": [("LINE経由の会員登録4件・購入4件は、", 20, True, INK),
              ("すべてリッチメニューから", 20, True, NAVY)], "sa": 18},
    {"runs": [("リッチメニューが見られた回数の", 20, True, INK),
              ("87%", 20, True, NAVY),
              ("は、配信日とその翌日", 20, True, INK)]},
], anchor="m", ml=0.45, mr=0.30)
footnote(s, "出典：GA（2026年9月・暫定値）、LINE公式アカウント管理画面（9/1〜9/27）")

# =====================================================================
# S4 課題①（棒グラフ・あとで編集できる形）
# =====================================================================
s = S[3]
frame(s, "課題①　同じ方に続けて送り、開封率が下がっています")
cd = CategoryChartData(number_format="0.0%")
cd.categories = ["8/24", "8/31", "9/10", "9/18", "9/25"]
cd.add_series("開封率", (0.426, 0.371, 0.362, 0.332, 0.301))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                        Inches(1.10), Inches(1.75), Inches(8.63), Inches(3.80), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
ch.font.size = Pt(14)
ch.font.name = FONT
ch.font.color.rgb = RGBColor.from_string(INK)
plot = ch.plots[0]
plot.gap_width = 80
plot.vary_by_categories = False
ser = plot.series[0]
ser.format.fill.solid()
ser.format.fill.fore_color.rgb = RGBColor.from_string(BAR)
ser.format.line.fill.background()
plot.has_data_labels = True
dl = plot.data_labels
dl.number_format = "0.0%"
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.show_value = True
dl.font.size = Pt(14)
dl.font.bold = True
dl.font.color.rgb = RGBColor.from_string(INK)
va = ch.value_axis
va.has_major_gridlines = False
va.has_minor_gridlines = False
va.minimum_scale = 0
va.maximum_scale = 0.5
va.visible = False
ca = ch.category_axis
ca.has_major_gridlines = False
ca.tick_labels.font.size = Pt(14)
ca.tick_labels.font.color.rgb = RGBColor.from_string(INK)
ca.format.line.color.rgb = RGBColor.from_string(BORDER)
chart_fonts(ch)
add_text(s, 0.60, 5.75, 9.63, 0.60,
         [{"runs": [("約600人の配信先に続けて送った5回で、毎回下がっている", 20, True, INK)],
           "align": "c"}], anchor="m")
footnote(s, "出典：LINE公式アカウント管理画面（各回の配信先は約560〜660人）")

# =====================================================================
# S5 課題②
# =====================================================================
s = S[4]
frame(s, "課題②　新しい友だちが、ほとんど増えていません")
add_box(s, 2.67, 1.90, 5.50, 3.00, fill=PALE, line=BORDER)
add_text(s, 2.67, 2.05, 5.50, 0.60,
         [{"runs": [("9月の新しい友だち", 22, True, INK)], "align": "c"}], anchor="m")
add_text(s, 2.67, 2.65, 5.50, 2.05,
         [{"runs": [("7", 96, True, RED), ("人", 40, True, RED)], "align": "c"}], anchor="m")
add_text(s, 0.60, 5.30, 9.63, 0.70,
         [{"runs": [("届く人数は、9月だけで", 20, True, INK),
                    ("41人減りました", 20, True, RED),
                    ("（1,676人→1,635人）", 20, True, INK)], "align": "c"}], anchor="m")
footnote(s, "出典：LINE公式アカウント管理画面（2026/9/1〜9/28）")

# =====================================================================
# S6 課題③
# =====================================================================
s = S[5]
frame(s, "課題③　買ったお客様が、LINEに入ってくる道がありません")
CARDS = [
    (0.60, "Google広告経由の購入", [("月", 28, True, NAVY), ("457〜854", 40, True, NAVY), ("件", 28, True, NAVY)]),
    (5.63, "9月の新しい友だち", [("7", 66, True, RED), ("人", 28, True, RED)]),
]
for x, label, runs in CARDS:
    add_box(s, x, 1.90, 4.60, 3.00, fill=PALE, line=BORDER)
    add_text(s, x, 2.10, 4.60, 0.60,
             [{"runs": [(label, 22, True, INK)], "align": "c"}], anchor="m")
    add_text(s, x, 2.75, 4.60, 1.90,
             [{"runs": runs, "align": "c"}], anchor="m", wrap=False)
add_text(s, 0.60, 5.30, 9.63, 0.70,
         [{"runs": [("ご注文の柱は、名前で検索して戻ってくるお客様（リピート）です", 20, True, INK)],
           "align": "c"}], anchor="m")
footnote(s, "出典：広告レポート（購入＝2026年7月854件・8/1〜25 457件／"
            "名前で検索する広告＝8/1〜25 費用17万円・売上1,022万円）、LINE公式アカウント管理画面")

# =====================================================================
# S7 ご提案
# =====================================================================
s = S[6]
frame(s, "ご提案　サイトの「出口」と「入口」で、LINEにご案内します")
PROPOSALS = [
    # 文は読点で改行する（「方 / を、」「100 / 人に」のような割れ方を防ぐ）
    (0.60, "出口｜サンクスLINE誘導",
     ["購入・会員登録が終わった方を、", "そのままLINEへ"],
     ["Google広告経由の購入は月457〜854件"]),
    (5.63, "入口｜離脱防止",
     ["買わずに帰ろうとする方に、", "LINEをご案内"],
     ["広告で来た方のうち、", "その場で買うのは100人に2〜10人"]),
]
for x, head, body, note in PROPOSALS:
    add_box(s, x, 1.85, 4.60, 3.55, fill=PALE, line=BORDER)
    chip(s, x, 1.85, 4.60, 0.70, [{"runs": [(head, 20, True, WHITE)], "align": "c"}])
    add_text(s, x + 0.20, 2.75, 4.30, 1.55,
             [{"runs": [(t, 20, True, INK)], "ls": 1.15} for t in body],
             anchor="m", ml=0.0, mr=0.0)
    add_text(s, x + 0.20, 4.45, 4.30, 0.80,
             [{"runs": [(t, 14, False, INK)]} for t in note],
             anchor="m", ml=0.0, mr=0.0)
plus = add_box(s, 5.415 - 0.30, 3.33, 0.60, 0.60, fill=NAVY, line=WHITE, lw=2.0, shape=MSO_SHAPE.OVAL)
put_text(plus.text_frame, [{"runs": [("＋", 22, True, WHITE)], "align": "c"}],
         anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)
add_text(s, 0.60, 5.65, 9.63, 0.65,
         [{"runs": [("2つをセットで入れます", 24, True, NAVY)], "align": "c"}], anchor="m")
footnote(s, "出典：広告レポート（2026年7〜8月。その場で買う割合＝CVRは、名前で検索する広告を除き2〜10%）")

# =====================================================================
# S8 サンクスLINE誘導（BUFFF 31ページ。暫定措置の注記だけ削除）
# =====================================================================
s = S[7]
sh = shape_by_id(s, 10)
assert "暫定措置" in sh.text_frame.text
sh._element.getparent().remove(sh._element)

# S9 離脱防止（BUFFF 27ページのまま）／S12 裏表紙（BUFFF 62ページのまま）

# =====================================================================
# S10 期待する効果
# =====================================================================
s = S[9]
frame(s, "期待する効果　3つの課題に、それぞれ効きます")
EFFECTS = [
    ("課題①", "開封率の低下", "買った方に絞って、送り分けられる"),
    ("課題②", "友だちが増えない", "入口と出口に、友だちが増える道ができる"),
    ("課題③", "買ったお客様がいない", "次のご注文を、LINEでご案内できる"),
]
y = 1.80
for no, issue, effect in EFFECTS:
    chip(s, 0.60, y, 2.90, 1.00, [
        {"runs": [(no, 14, True, WHITE)], "align": "c"},
        {"runs": [(issue, 18, True, WHITE)], "align": "c"},
    ])
    arrow(s, 3.64, y + 0.25, 0.45, 0.50)
    add_box(s, 4.23, y, 6.00, 1.00, fill=PALE, line=BORDER)
    add_text(s, 4.23, y, 6.00, 1.00,
             [{"runs": [(effect, 20, True, INK)]}], anchor="m", ml=0.30, mr=0.10)
    y += 1.20
add_text(s, 0.60, 5.60, 9.63, 0.70,
         [{"runs": [("11月から入れれば、年賀状・カレンダーのお客様から始められます", 20, True, NAVY)],
           "align": "c"}], anchor="m")

# =====================================================================
# S11 費用とスケジュール
# =====================================================================
s = S[10]
frame(s, "費用とスケジュール")
tbl_gf = s.shapes.add_table(3, 3, Inches(0.60), Inches(1.80), Inches(9.63), Inches(2.05))
tbl = tbl_gf.table
tbl.first_row = False
tbl.horz_banding = False
for i, w in enumerate((3.20, 2.50, 3.93)):
    tbl.columns[i].width = Inches(w)
for i, h in enumerate((0.80, 0.62, 0.63)):
    tbl.rows[i].height = Inches(h)
TABLE = [
    ("サンクスLINE誘導", "初期10万円",
     [{"runs": [("月額3万円〜", 20, True, INK)]},
      {"runs": [("（＋設定地点追加）", 14, False, INK)]}], PALE, WHITE, INK),
    ("離脱防止", "初期1.5万円",
     [{"runs": [("月額3万円", 20, True, INK)]}], PALE, WHITE, INK),
    ("合計（2つセット）", "初期11.5万円",
     [{"runs": [("月額6万円〜", 20, True, WHITE)]}], NAVY, NAVY, WHITE),
]
for r, (name, init, monthly, f_name, f_val, c_val) in enumerate(TABLE):
    c_name = WHITE if f_name == NAVY else INK
    fill_cell(tbl.cell(r, 0), [{"runs": [(name, 20, True, c_name)]}], f_name)
    fill_cell(tbl.cell(r, 1), [{"runs": [(init, 20, True, c_val)]}], f_val)
    fill_cell(tbl.cell(r, 2), monthly, f_val)

STEPS = [
    ("10月", ["ご判断・タグの設置", "（貴社サイト）"]),
    ("11月", ["開始"]),
    ("12月末", ["効果の確認"]),
]
bw, gap = 2.85, 0.54
x = 0.60
for i, (month, desc) in enumerate(STEPS):
    paras = [{"runs": [(month, 22, True, WHITE)], "align": "c", "sa": 4}]
    paras += [{"runs": [(d, 18, True, WHITE)], "align": "c"} for d in desc]
    chip(s, x, 4.20, bw, 1.35, paras)
    if i < 2:
        arrow(s, x + bw + (gap - 0.40) / 2, 4.20 + 0.425, 0.40, 0.50)
    x += bw + gap
add_box(s, 0.60, 5.85, 9.63, 0.65, fill=PALE, line=BORDER)
add_text(s, 0.60, 5.85, 9.63, 0.65,
         [{"runs": [("確認する数字：", 18, True, NAVY),
                    ("友だち追加数／LINE経由のご注文（GA）／開封率", 18, True, INK)],
           "align": "c"}], anchor="m")

prs.save(str(OUT))

# ---------- 検品 ----------
done = Presentation(str(OUT))
slides = list(done.slides)
assert len(slides) == 12, len(slides)
assert done.slide_width == 9906000 and done.slide_height == 6858000


def slide_text(slide):
    parts = []
    for sh in iter_shapes(slide.shapes):
        if sh.has_text_frame:
            parts.append(sh.text_frame.text)
        if sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    parts.append(cell.text_frame.text)
        if sh.has_chart:
            parts += list(sh.chart.plots[0].categories)
    return "\n".join(parts)


# 3. 入れないもの
NG = ["101万", "0〜3日", "0～3日", "55%", "ブロック率", "年齢", "性別", "アップセル", "スコープ"]
problems = []
for i, sl in enumerate(slides, 1):
    txt = slide_text(sl)
    for w in NG:
        if w in txt:
            problems.append(f"S{i}: 入れない言葉「{w}」")

# 作り直した枚（S1〜S7・S10・S11）の数字が、指示にある数字だけか
ALLOWED = set("""
2026 10 9 19 37 17 21 30 36 6 8 7 3 4 4 87 1 27 24 31 18 25 42.6 37.1 36.2 33.2 30.1
600 5 560 660 41 1,676 1,635 28 457 854 1,022 100 2 11 12 1.5 11.5 0.0
""".split())
for i in (1, 2, 3, 4, 5, 6, 7, 10, 11):
    txt = slide_text(slides[i - 1])
    for num in re.findall(r"\d[\d,]*(?:\.\d+)?", txt):
        if num not in ALLOWED:
            problems.append(f"S{i}: 指示に無い数字「{num}」")
# グラフの値（42.6%など）も指示どおりか
vals = None
for shp in slides[3].shapes:
    if shp.has_chart:
        vals = [round(v, 3) for v in shp.chart.plots[0].series[0].values]
assert vals == [0.426, 0.371, 0.362, 0.332, 0.301], vals

# S8：暫定措置の注記が消え、料金はそのまま
t8 = slide_text(slides[7])
assert "暫定措置" not in t8 and "10万円" in t8 and "月額3万円" in t8
t9 = slide_text(slides[8])
assert "1.5万円" in t9 and "月額3万円" in t9

print(f"saved: {OUT.name}")
print(f"slides: {len(slides)}")
if problems:
    print("!! 要確認:")
    for p in problems:
        print("   ", p)
else:
    print("入れないもの・指示に無い数字: なし")
