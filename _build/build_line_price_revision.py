# -*- coding: utf-8 -*-
"""LINE料金改定へのご対応（1枚・4ステップ）

ストーリー：
  ① 課題      10月にLINEヤフー社の追加メッセージ単価が改定される
  ② リスク    一斉配信のままだとコストが増大する
  ③ 解決策    セグメント配信への移行＋弊社特別プランで最適化
  ④ スケジュール 適用開始は10月ではなく11月から
     ※10月は配信実績データの計測期間。実績とインセンティブを踏まえ11月に個別調整。

トーンは ネガ（グレー→赤）→ ポジ（青→紺）へ色で着地させる。

使い方:
    python _build/build_line_price_revision.py
        → _templates/DYM_LINEOA_FMT.pptx から 1枚もの を生成

    python _build/build_line_price_revision.py <既存.pptx> [出力.pptx] [ページ番号]
        → 既存デッキの指定ページ（既定9）を、この内容で作り直す
"""
import shutil
import sys
from pathlib import Path
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
FMT = str(ROOT / "_templates" / "DYM_LINEOA_FMT.pptx")
STANDALONE_OUT = str(ROOT / "20260826_LINE料金改定へのご対応.pptx")

TNAVY = "002060"   # タイトル
NAVY = "1F285A"    # 打ち手・着地
BLUE = "1565C0"    # 解決策
RED = "C00000"     # リスク
GREY = "8C8C8C"    # 課題（動かせない外部要因）
INK = "333333"
MUT = "808080"
WHITE = "FFFFFF"
PALE = "F4F7FF"
FADE = "F2F2F2"
BORDER = "D9D9D9"


# ---------- helpers（build_vivical_9p.py と同じ経路） ----------
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


def clear_slide(slide):
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)


# ---------- 中身 ----------
TITLE = "LINE料金改定へのご対応について"

LEAD = [
    {"runs": [("一斉配信のままでは、10月の料金改定で配信コストが増大します。", 14, True, INK)],
     "align": "c"},
    {"runs": [("セグメント配信への移行と特別プランの適用で最適化し、", 14, True, INK),
              ("11月から", 14, True, RED),
              ("ご提供いたします。", 14, True, INK)], "align": "c"},
]

STEPS = [
    # (no, ラベル, ヘッダー色, カード色, 枠色, 見出し, 本文, 補足の一言, 補足色)
    ("STEP 1", "課題", GREY, FADE, BORDER,
     "10月に料金改定",
     "LINEヤフー社の追加メッセージ単価が改定されます。",
     "20万通超は一律 2.5円/通", MUT),
    ("STEP 2", "リスク", RED, FADE, BORDER,
     "コストが増大",
     "一斉配信のままでは、配信コストが大きく膨らみます。",
     "大量配信の割引が縮小", RED),
    ("STEP 3", "解決策", BLUE, PALE, BLUE,
     "配信を最適化",
     "セグメント配信へ移行し、弊社の特別プランを適用します。",
     "ムダな配信を削減", BLUE),
    ("STEP 4", "スケジュール", NAVY, PALE, NAVY,
     "11月から適用",
     "適用開始は10月からではなく、11月からとなります。",
     "10月は実績の計測期間", NAVY),
]

NOTE = ("10月は配信実績データの計測期間とし、その実績と媒体インセンティブを踏まえて、"
        "11月に各社様ごとの特別プランを個別調整いたします。")

CLOSING = "11月から、各社様に合わせた特別プランでご提供いたします。"

# レイアウト定数（スライド 10.83 × 7.5 inch）
X0, TOTAL_W, GAP = 0.50, 9.83, 0.30
CARD_W = (TOTAL_W - GAP * 3) / 4
CARD_Y, CARD_H = 1.95, 2.52
NOTE_Y, NOTE_H = 4.72, 0.80
CLOSE_Y, CLOSE_H = 5.78, 0.92


def draw_body(s):
    """4ステップ＋補足＋クロージング。タイトル／リードは呼び出し側で置く。"""
    for i, (no, label, hcol, fill, edge, head, body, tag, tagcol) in enumerate(STEPS):
        x = X0 + i * (CARD_W + GAP)
        add_box(s, x, CARD_Y, CARD_W, CARD_H, fill=fill, line=edge)

        hd = add_box(s, x, CARD_Y, CARD_W, 0.52, fill=hcol)
        put_text(hd.text_frame,
                 [{"runs": [(no, 9.5, True, WHITE), ("　", 9.5, True, WHITE),
                            (label, 13, True, WHITE)], "align": "c"}],
                 anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)

        add_text(s, x + 0.10, CARD_Y + 0.62, CARD_W - 0.20, 0.42,
                 [{"runs": [(head, 15, True, INK)], "align": "c"}],
                 anchor="m", ml=0.0, mr=0.0)

        add_text(s, x + 0.12, CARD_Y + 1.10, CARD_W - 0.24, 0.86,
                 [{"runs": [(body, 10.5, None, INK)], "ls": 1.3}],
                 anchor="t", ml=0.0, mr=0.0)

        add_text(s, x + 0.12, CARD_Y + 2.00, CARD_W - 0.24, 0.34,
                 [{"runs": [("▶ " + tag, 9.5, True, tagcol)]}],
                 anchor="m", ml=0.0, mr=0.0)

        if i < 3:
            ar = add_box(s, x + CARD_W + 0.04, CARD_Y + CARD_H / 2 - 0.15,
                         0.22, 0.30, fill=BORDER, shape=MSO_SHAPE.RIGHT_ARROW)
            ar.line.fill.background()

    # 補足（なぜ11月からなのか）
    add_box(s, X0, NOTE_Y, TOTAL_W, NOTE_H, fill=WHITE, line=NAVY)
    add_text(s, X0 + 0.28, NOTE_Y, TOTAL_W - 0.56, NOTE_H,
             [{"runs": [("※ ", 11.5, True, NAVY), (NOTE, 11.5, None, INK)], "ls": 1.25}],
             anchor="m", ml=0.0, mr=0.0)

    # クロージング（ポジティブ着地）
    cl = add_box(s, X0, CLOSE_Y, TOTAL_W, CLOSE_H, fill=NAVY)
    put_text(cl.text_frame,
             [{"runs": [(CLOSING, 18, True, WHITE)], "align": "c"}],
             anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)


def build_standalone(out=STANDALONE_OUT):
    """DYM汎用FMTをコピー → 1枚に削減 → 中身を作り直す。"""
    shutil.copyfile(FMT, out)
    prs = Presentation(out)
    slides = list(prs.slides)

    div_el = None
    for sh in slides[1].shapes:
        if sh._element.tag.endswith('}cxnSp') and abs(sh.top - 1390675) < 60000:
            div_el = deepcopy(sh._element)
            break
    assert div_el is not None, "divider not found"

    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    keep = ids[1]
    for sldId in ids:
        if sldId is keep:
            continue
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
    assert len(list(prs.slides)) == 1

    s = list(prs.slides)[0]
    clear_slide(s)
    s.shapes._spTree.append(deepcopy(div_el))
    add_text(s, 0.60, 0.13, 8.10, 0.40,
             [{"runs": [(TITLE, 16, True, TNAVY)]}], anchor="m", ml=0.0, mr=0.0)
    add_text(s, 0.42, 0.62, 10.02, 0.84, LEAD, anchor="m", ml=0.0, mr=0.0)
    draw_body(s)
    prs.save(out)
    return out


def build_into_deck(src, out, page=9):
    """既存デッキの指定ページを、この内容で作り直す（他ページは触らない）。"""
    shutil.copyfile(src, out)
    prs = Presentation(out)
    s = prs.slides[page - 1]
    clear_slide(s)
    add_text(s, 0.47, 0.10, 8.22, 0.40,
             [{"runs": [(TITLE, 16, None, TNAVY)]}], anchor="m", ml=0.1, mr=0.0)
    add_text(s, 0.16, 0.62, 10.58, 0.84, LEAD, anchor="m", ml=0.0, mr=0.0)
    draw_body(s)
    prs.save(out)
    return out


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("saved:", build_standalone())
    else:
        src = Path(sys.argv[1])
        out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_name(src.stem + "_S9修正.pptx")
        page = int(sys.argv[3]) if len(sys.argv) > 3 else 9
        print("saved:", build_into_deck(str(src), str(out), page))
