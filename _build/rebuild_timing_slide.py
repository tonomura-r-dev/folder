# -*- coding: utf-8 -*-
"""Slide 20（配信タイミング設計）を「来場後14日間」に振り切って作り直す"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "20260804_株式会社ビイサイドプランニング御中_公式LINEのご提案v07.pptx"

TNAVY = "002060"; NAVY = "1F285A"; RED = "C00000"; INK = "333333"
MUT = "808080"; WHITE = "FFFFFF"; PALE = "F4F7FF"; BORDER = "D9D9D9"; GREY = "ECEDF3"

prs = Presentation(str(TARGET))
slides = list(prs.slides)


def set_font(run, size, bold=None, color=INK, name="メイリオ"):
    f = run.font; f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    f.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {}); rPr.append(e)
        e.set("typeface", name)
    f.color.rgb = RGBColor.from_string(color)


ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCH = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def reset_tf(tf):
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    for r in list(tf.paragraphs[0].runs):
        r._r.getparent().remove(r._r)


def put_text(tf, paras, anchor="t", ml=0.05, mr=0.05, mt=0.02, mb=0.02):
    reset_tf(tf); tf.word_wrap = True
    tf.margin_left = Inches(ml); tf.margin_right = Inches(mr)
    tf.margin_top = Inches(mt); tf.margin_bottom = Inches(mb)
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
            r = para.add_run(); r.text = t
            set_font(r, sz, b, c)


def add_text(slide, x, y, w, h, paras, anchor="t", **kw):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    put_text(box.text_frame, paras, anchor=anchor, **kw)
    return box


def add_box(slide, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = RGBColor.from_string(line); sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp


def clear_slide(slide):
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)


# ================= Slide 20 =================
s = slides[19]
clear_slide(s)
add_text(s, 0.60, 0.13, 9.60, 0.45,
         [{"runs": [("具体施策⑦ 来場後14日間のフォロー設計", 20, True, TNAVY)]}],
         anchor="m", ml=0.0, mr=0.0)
add_text(s, 0.60, 0.66, 9.60, 0.58,
         [{"runs": [("来場から10日で、求職者は他社の転職サービスを調べ始めます。それまでに次の一手を渡します。", 14, True, INK)], "ls": 1.05}],
         anchor="m", ml=0.0, mr=0.0)

# --- 上部：離脱のタイムライン ---
add_text(s, 0.60, 1.30, 9.60, 0.28,
         [{"runs": [("来場者が実際に検索していること（当社調査）", 11, True, NAVY)]}], anchor="m", ml=0.0, mr=0.0)
tl = [("+4日", "職務経歴書\nフォーマット", "書類でつまずく", PALE, NAVY),
      ("+6日", "面接 質問\n一覧", "面接が不安", PALE, NAVY),
      ("+9日", "志望動機 例文\n転職", "まだ書けない", PALE, NAVY),
      ("+10日", "ワークポート 評判\ndoda ログイン", "他社を調べ始める", RED, WHITE),
      ("+13日", "オープンワーク", "他社で企業研究", RED, WHITE)]
for i, (d, kw, state, bg, tc) in enumerate(tl):
    x = 0.60 + i * 1.96
    hb = add_box(s, x, 1.62, 1.80, 0.30, fill=NAVY if bg == PALE else RED)
    put_text(hb.text_frame, [{"runs": [(d, 10.5, True, WHITE)], "align": "c"}],
             anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)
    bd = add_box(s, x, 1.94, 1.80, 0.86, fill=bg)
    put_text(bd.text_frame,
             [{"runs": [(ln, 8.5, None, tc)], "align": "c", "ls": 1.05} for ln in kw.split("\n")],
             anchor="m", ml=0.04, mr=0.04, mt=0.0, mb=0.0)
    add_text(s, x, 2.82, 1.80, 0.26,
             [{"runs": [(state, 9, True, RED if bg == RED else MUT)], "align": "c"}],
             anchor="m", ml=0.0, mr=0.0)
    if i < 4:
        add_box(s, x + 1.82, 2.24, 0.12, 0.24, fill=MUT, shape=MSO_SHAPE.RIGHT_ARROW)
warn = add_box(s, 7.44, 3.14, 2.76, 0.34, fill=RED)
put_text(warn.text_frame,
         [{"runs": [("← ここが離脱の分岐点", 10, True, WHITE)], "align": "c"}],
         anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)

# --- 下部：配信設計 ---
add_text(s, 0.60, 3.52, 9.60, 0.28,
         [{"runs": [("検索が立つ日に、必要なものを先回りして渡す", 11, True, NAVY)]}], anchor="m", ml=0.0, mr=0.0)
plan = [
    ("D+1", "お礼＋返信文例＋次会場の案内", "面接日程 返信メール −0.2日", False),
    ("D+4", "職務経歴書テンプレートを配布", "職務経歴書 +4.0日", True),
    ("D+6", "面接の想定質問リストを配布", "面接 質問 一覧 +5.8日", True),
    ("D+9", "志望動機の型＋個別キャリア相談へ誘導", "志望動機 例文 +9.3日", True),
    ("D+13", "最終面接の対策コンテンツ", "最終面接 +13.8日", False),
]
PY, RH = 3.86, 0.44
for i, (d, body, src, hl) in enumerate(plan):
    y = PY + i * (RH + 0.05)
    db = add_box(s, 0.60, y, 0.95, RH, fill=NAVY)
    put_text(db.text_frame, [{"runs": [(d, 11, True, WHITE)], "align": "c"}],
             anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)
    bb = add_box(s, 1.60, y, 5.20, RH, fill=PALE if hl else WHITE,
                 line=NAVY if hl else BORDER, lw=1.25 if hl else 0.75)
    put_text(bb.text_frame,
             [{"runs": [(body, 11, hl or None, NAVY if hl else INK)]}],
             anchor="m", ml=0.16, mr=0.10, mt=0.0, mb=0.0)
    sb = add_box(s, 6.85, y, 3.35, RH, fill=GREY)
    put_text(sb.text_frame, [{"runs": [(src, 9, None, MUT)], "align": "c"}],
             anchor="m", ml=0.06, mr=0.06, mt=0.0, mb=0.0)

# --- 右下の注記（D+9の意味） ---
add_text(s, 0.60, 6.20, 9.60, 0.28,
         [{"runs": [("D+9は、他社検索が始まる+10日の直前。ここで個別相談を提示し、次の行動を自社内に留めます。", 10.5, True, INK)]}],
         anchor="m", ml=0.0, mr=0.0)

bar = add_box(s, 0.60, 6.50, 9.60, 0.44, fill=NAVY)
put_text(bar.text_frame,
         [{"runs": [("来場後のフォローが、出展企業の「採用できた」を作る", 13.5, True, WHITE)], "align": "c"}],
         anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)
add_text(s, 0.60, 6.94, 9.60, 0.22,
         [{"runs": [("※日数は「転職イベント」等の検索発生日を0日とした相対値。当社調査（検索需要分析）による参考値です。", 8.5, None, MUT)]}], ml=0.0)

prs.save(str(TARGET))
print("saved:", TARGET)
