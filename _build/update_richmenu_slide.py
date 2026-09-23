# -*- coding: utf-8 -*-
"""Slide 14（リッチメニュー）を3パターン出し分けの実物モックに作り直す"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "20260804_株式会社ビイサイドプランニング御中_公式LINEのご提案v03.pptx"

TNAVY = "002060"
NAVY = "1F285A"
INK = "333333"
MUT = "808080"
WHITE = "FFFFFF"
PALE = "F4F7FF"
CELL = "DDE5F5"      # 通常セル（淡い青）
BORDER = "D9D9D9"
CHATBG = "E9EDF2"

prs = Presentation(str(TARGET))
slides = list(prs.slides)


def set_font(run, size, bold=None, color=INK, name="メイリオ"):
    f = run.font
    f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    f.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", name)
    f.color.rgb = RGBColor.from_string(color)


ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCH = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def reset_tf(tf):
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    for r in list(tf.paragraphs[0].runs):
        r._r.getparent().remove(r._r)


def put_text(tf, paras, anchor="t", ml=0.05, mr=0.05, mt=0.02, mb=0.02, wrap=True):
    reset_tf(tf)
    tf.word_wrap = wrap
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
    return tf


def add_text(slide, x, y, w, h, paras, anchor="t", **kw):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    put_text(box.text_frame, paras, anchor=anchor, **kw)
    return box


def add_box(slide, x, y, w, h, fill=None, line=None, lw=1.0,
            shape=MSO_SHAPE.RECTANGLE, radius=None):
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
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE and radius is not None:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp


def clear_slide(slide):
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)


def richmenu(slide, x, y, w, cells):
    """6分割のリッチメニュー。cells=[(ラベル行のリスト, 主役か), ...] 6個"""
    gap = 0.045
    cw = (w - gap * 4) / 3
    ch = 0.74
    h = ch * 2 + gap * 3
    add_box(slide, x, y, w, h, fill=WHITE, line=NAVY, lw=1.5)
    for i, (lines, star) in enumerate(cells):
        r, c = divmod(i, 3)
        cx = x + gap + c * (cw + gap)
        cy = y + gap + r * (ch + gap)
        add_box(slide, cx, cy, cw, ch, fill=NAVY if star else CELL)
        put_text(slide.shapes[-1].text_frame,
                 [{"runs": [(t, 8.5, True, WHITE if star else NAVY)], "align": "c", "ls": 1.0}
                  for t in lines],
                 anchor="m", ml=0.02, mr=0.02, mt=0.0, mb=0.0)
    return y + h


# ================= Slide 14 =================
s = slides[13]
clear_slide(s)
add_text(s, 0.60, 0.13, 9.60, 0.45,
         [{"runs": [("具体施策② リッチメニュー設計（フェーズ別の出し分け）", 20, True, TNAVY)]}],
         anchor="m", ml=0.0, mr=0.0)
add_text(s, 0.60, 0.68, 9.60, 0.55,
         [{"runs": [("検討フェーズに合わせて3パターンを出し分け、「いま押してほしいボタン」を常に左上に置きます。", 14, True, INK)], "ls": 1.05}],
         anchor="m", ml=0.0, mr=0.0)

patterns = [
    (0.60, "① 通常時", "情報収集フェーズ",
     [(["出展企業", "一覧"], True), (["会場MAP"], False), (["開催日程"], False),
      (["事前予約"], False), (["キャリア相談"], False), (["よくある質問"], False)],
     "まずは比較検討。どんな企業が来るのかを見せ、興味を醸成する。"),
    (4.02, "② 週末（土日の夜）", "決断フェーズ",
     [(["事前予約"], True), (["特典受取"], False), (["出展企業", "一覧"], False),
      (["会場MAP"], False), (["キャリア相談"], False), (["よくある質問"], False)],
     "転職意欲のピーク。予約を左上に移し、申込への最短距離をつくる。"),
    (7.44, "③ フェア直前（3日前〜当日）", "来場フェーズ",
     [(["会場MAP", "アクセス"], True), (["当日の流れ"], False), (["持ち物・服装"], False),
      (["出展企業", "一覧"], False), (["事前予約"], False), (["よくある質問"], False)],
     "当日の迷いを消す。「行けるか不安」を潰し、来場率を守る。"),
]
PW = 3.18
for x, name, phase, cells, aim in patterns:
    hb = add_box(s, x, 1.38, PW, 0.52, fill=NAVY)
    put_text(hb.text_frame,
             [{"runs": [(name, 11.5, True, WHITE)], "align": "c", "sa": 1},
              {"runs": [(phase, 8.5, None, WHITE)], "align": "c"}],
             anchor="m", ml=0.03, mr=0.03, mt=0.0, mb=0.0)
    bottom = richmenu(s, x, 1.98, PW, cells)
    add_text(s, x, bottom + 0.06, PW, 0.26,
             [{"runs": [("▲ 左上＝いま押してほしいボタン", 8, None, MUT)], "align": "c"}],
             anchor="m", ml=0.0, mr=0.0)
    ab = add_box(s, x, bottom + 0.36, PW, 0.78, fill=PALE)
    put_text(ab.text_frame, [{"runs": [(aim, 10, None, INK)], "ls": 1.15}],
             anchor="m", ml=0.14, mr=0.10, mt=0.05, mb=0.05)

add_text(s, 0.60, 4.98, 9.60, 0.30,
         [{"runs": [("切替の運用", 12, True, NAVY)]}], anchor="m", ml=0.0, mr=0.0)
ops = [
    ("①⇄② は毎週自動で切替", "金曜の夜に②へ、日曜の深夜に①へ戻す。スケジュール設定で自動化でき、運用工数はかからない。"),
    ("③ は開催3日前に切替", "各会場の3日前に③へ、翌日に①へ戻す。秋クールは4会場なので計4回の切替。"),
    ("画像は3枚作れば使い回せる", "初期構築で3パターン制作すれば、以降の切替に追加費用は発生しない。"),
]
for i, (h, b) in enumerate(ops):
    x = 0.60 + i * 3.28
    sp = add_box(s, x, 5.30, 3.05, 1.00, fill=WHITE, line=BORDER, lw=1.0)
    put_text(sp.text_frame,
             [{"runs": [(h, 10.5, True, NAVY)], "sa": 3},
              {"runs": [(b, 9.5, None, INK)], "ls": 1.1}],
             anchor="m", ml=0.14, mr=0.10, mt=0.05, mb=0.05)

bar = add_box(s, 0.60, 6.48, 9.60, 0.46, fill=NAVY)
put_text(bar.text_frame,
         [{"runs": [("メニューは固定せず、求職者の検討フェーズに合わせて入れ替える", 13, True, WHITE)], "align": "c"}],
         anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)

prs.save(str(TARGET))
print("saved:", TARGET)
