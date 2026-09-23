# -*- coding: utf-8 -*-
"""Slide 19（新規）＝検索需要に基づく配信タイミング設計。⑦⑧を⑧⑨へ繰り下げ"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "20260804_株式会社ビイサイドプランニング御中_公式LINEのご提案v05.pptx"

TNAVY = "002060"; NAVY = "1F285A"; INK = "333333"
MUT = "808080"; WHITE = "FFFFFF"; PALE = "F4F7FF"; BORDER = "D9D9D9"
GREY = "ECEDF3"

prs = Presentation(str(TARGET))
slides = list(prs.slides)
assert len(slides) == 30, len(slides)


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


def put_text(tf, paras, anchor="t", ml=0.05, mr=0.05, mt=0.02, mb=0.02, wrap=True):
    reset_tf(tf); tf.word_wrap = wrap
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


# ================= Slide 19（新規） =================
s = slides[18]
clear_slide(s)
add_text(s, 0.60, 0.13, 9.60, 0.45,
         [{"runs": [("具体施策⑦ 検索需要に基づく配信タイミング設計", 20, True, TNAVY)]}],
         anchor="m", ml=0.0, mr=0.0)
add_text(s, 0.60, 0.68, 9.60, 0.55,
         [{"runs": [("求職者が「何を・いつ」検索するかの実データから、キーワードが立つ直前に先回りして配信します。", 14, True, INK)], "ls": 1.05}],
         anchor="m", ml=0.0, mr=0.0)

# --- 配信表 ---
rows = [
    ("事前", "D-22", "適職診断 転職　−17.4日", "希望条件診断アンケート", False),
    ("事前", "D-13", "退職代行　−13.2日", "現職の不満に寄り添う訴求＋出展企業", False),
    ("事前", "D-8", "リクルートエージェント　−8.1日", "エージェントとの違いを提示", False),
    ("事前", "D-6", "※服装KWが立つ2日前", "「服装自由・履歴書不要」を先回り", True),
    ("事前", "D-4", "合同企業説明会　−4.2日", "出展企業の確定版＋予約ラストコール", False),
    ("事前", "D-1", "転職　−2.2日", "前日リマインド＋会場MAP", False),
    ("当日", "D-0", "合同説明会 服装　0日", "当日朝：服装自由を再掲＋会場MAP", True),
    ("事後", "D+1", "面接日程 返信メール　−0.2日", "お礼＋次会場（10/3草津）の案内", False),
    ("事後", "D+4", "職務経歴書　+4.0日", "職務経歴書テンプレを配布", True),
    ("事後", "D+6", "面接 質問 一覧　+5.8日", "面接の想定質問リスト", False),
    ("事後", "D+10", "志望動機 例文　+9.3日", "志望動機・退職理由の書き方", False),
]
TX, TY, TW = 0.60, 1.42, 6.75
hdr_h, row_h = 0.32, 0.335
cols = [0.52, 0.72, 2.55, 2.96]  # 区分 / D-x / KW / 配信内容
# ヘッダー
cx = TX
for lab, w in zip(["", "時期", "立つ検索キーワード（実測）", "配信内容"], cols):
    hb = add_box(s, cx, TY, w, hdr_h, fill=NAVY)
    put_text(hb.text_frame, [{"runs": [(lab, 9.5, True, WHITE)], "align": "c"}],
             anchor="m", ml=0.02, mr=0.02, mt=0.0, mb=0.0)
    cx += w
# 明細
for i, (phase, dx, kw, body, hl) in enumerate(rows):
    y = TY + hdr_h + i * row_h
    cx = TX
    fills = [GREY if phase != "当日" else NAVY,
             PALE if hl else WHITE, PALE if hl else WHITE, PALE if hl else WHITE]
    vals = [(phase if (i == 0 or rows[i-1][0] != phase) else "", 8.5, True,
             WHITE if phase == "当日" else MUT, "c"),
            (dx, 9.5, True, NAVY, "c"),
            (kw, 9, None, INK, "l"),
            (body, 9.5, hl or None, NAVY if hl else INK, "l")]
    for (w, f, (t, sz, b, col, al)) in zip(cols, fills, vals):
        cell = add_box(s, cx, y, w, row_h, fill=f, line=BORDER, lw=0.5)
        put_text(cell.text_frame, [{"runs": [(t, sz, b, col)], "align": al}],
                 anchor="m", ml=0.08, mr=0.04, mt=0.0, mb=0.0)
        cx += w

# --- 右：データから分かったこと ---
add_text(s, 7.55, 1.42, 2.65, 0.30,
         [{"runs": [("データから分かったこと", 12, True, NAVY)]}], anchor="m", ml=0.0, mr=0.0)
finds = [
    ("「転職フェア」は当日しか立たない",
     "フェア名の検索は開催当日（+0.003日）。手前で立つのは「転職」「転職エージェント」。フェア名で待たず、転職の悩みで捕まえる。"),
    ("服装の不安は4〜5日前に立つ",
     "リクルートスーツ −4.7日、spi問題 −4.5日。その2日前に先回りして安心Q&Aを送る。"),
    ("事後は「お礼」より「選考対策」",
     "職務経歴書 +4.0日、面接質問 +5.8日、志望動機 +9.3日。来場後は選考準備に移るため、実務資料の配布が刺さる。"),
]
for i, (h, b) in enumerate(finds):
    y = 1.80 + i * 1.32
    sp = add_box(s, 7.55, y, 2.65, 1.18, fill=PALE)
    put_text(sp.text_frame,
             [{"runs": [(h, 10, True, NAVY)], "sa": 3, "ls": 1.05},
              {"runs": [(b, 8.5, None, INK)], "ls": 1.12}],
             anchor="m", ml=0.12, mr=0.10, mt=0.05, mb=0.05)

bar = add_box(s, 0.60, 5.82, 9.60, 0.48, fill=NAVY)
put_text(bar.text_frame,
         [{"runs": [("「送りたい時」ではなく「調べたくなる直前」に送る", 14, True, WHITE)], "align": "c"}],
         anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)
add_text(s, 0.60, 6.38, 9.60, 0.56,
         [{"runs": [("※日数は「転職イベント」「合同説明会」等の検索が発生した日を0日とした相対値。数値は当社調査（検索需要分析）によるもので、傾向を示す参考値です。", 9, None, MUT)], "ls": 1.1}], ml=0.0)

# ================= 後続の施策番号を繰り下げ =================
for idx, old, new in [(19, "具体施策⑦ Meta広告", "具体施策⑧ Meta広告"),
                      (20, "具体施策⑧ LP離脱防止", "具体施策⑨ LP離脱防止")]:
    sl = slides[idx]  # 0-based で1つ後ろ＝新スライド挿入後の位置
    for sh in sl.shapes:
        if sh.has_text_frame and old in sh.text_frame.text:
            p = sh.text_frame.paragraphs[0]
            if p.runs:
                p.runs[0].text = p.runs[0].text.replace(old, new)
                for r in p.runs[1:]:
                    r._r.getparent().remove(r._r)

prs.save(str(TARGET))
print("saved:", TARGET)
