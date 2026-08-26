# -*- coding: utf-8 -*-
"""LINE公式アカウント 料金改定に伴う運用最適化・特別プランのご案内（1枚完結）

お客様に口頭で説明する言葉のまま、3ステップに落とした1枚。
  10月〜  料金体系が変わります／今の配信のままではコストが上がります
  ご提案  配信を見直します／一斉配信からセグメント配信へ
  11月〜  特別プランを適用／10月の実績をもとに貴社に合わせて調整

「なぜ10月からではないのか」は下の※行で説明する
（10月＝新料金での実績を計測する期間。原則、通常プラン適用）。

営業用トークスクリプトはスピーカーノートに入れている。
弊社側の事情（原価が読めない・利益割れリスク）は資料には一切書かない。

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
STANDALONE_OUT = str(ROOT / "20260826_LINE料金改定に伴う運用最適化・特別プランのご案内.pptx")

TNAVY = "002060"
NAVY = "1F285A"
BLUE = "1565C0"
RED = "C00000"
GREY = "8C8C8C"
INK = "333333"
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
TITLE = "LINE公式アカウント 料金改定に伴う運用最適化・特別プランのご案内"

LEAD = [
    {"runs": [("配信の見直しと特別プランで、料金改定によるコスト増を抑えます。", 15, True, INK)],
     "align": "c"},
    {"runs": [("適用開始は", 15, True, INK), ("11月から", 15, True, RED),
              ("。10月は新料金での実データを計測する期間です。", 15, True, INK)], "align": "c"},
]

# (チップ, 大見出し, 本文, チップ色, カード色, 枠色, 文字色, チップ文字色)
STEPS = [
    ("10月〜", "料金体系が変わります",
     "LINEヤフー社の追加メッセージ単価が改定されます。今の配信のままでは、コストが上がります。",
     GREY, FADE, BORDER, INK, WHITE),
    ("ご提案", "配信を見直します",
     "一斉配信からセグメント配信へ移行し、ムダな配信コストを削減します。",
     BLUE, PALE, BLUE, INK, WHITE),
    ("11月〜", "特別プランを適用",
     "10月の実績をもとに、貴社に合わせた特別プランを適用します。",
     WHITE, NAVY, NAVY, WHITE, NAVY),
]

NOTE = ("10月は、新料金での配信実績を計測する期間です（原則、通常プランの適用となります）。"
        "その実績と媒体インセンティブを踏まえて、11月から貴社専用の特別プランを適用いたします。")

TALK_SCRIPT = """【営業トークスクリプト】LINE公式アカウント 料金改定に伴う運用最適化・特別プランのご案内

■ ひとことで言うと
10月は新料金での実データを取る期間です。
その実績をもとに、11月から貴社に合わせた特別プランを適用します。

■ 導入
2026年10月から、LINEヤフー社の料金体系が改定されます。
追加メッセージの単価が変わり、大量配信に対する割引が縮小します。
貴社のように配信ボリュームが大きいアカウントほど、影響を受けやすい改定です。

■ 課題
対策をしないと、これまでと同じ配信をしているだけでコストが上がります。
特に全員に一律で送る「一斉配信」は、反応の薄い層にも同じ単価がかかるので、
改定後はムダが大きく出てしまいます。

■ ご提案
一斉配信からセグメント配信への切り替えです。
反応が見込める層に絞って配信することで、通数そのものを減らしながら成果は維持します。
あわせて、貴社の配信規模に合わせた弊社の特別プランをご用意します。

■ なぜ11月からなのか（ここが本題）
特別プランの適用は11月からとさせてください。
新料金で実際にどれくらいの通数・費用になるかは、動かしてみないと正確に出ません。
10月は原則、通常プランで運用いただき、実データを取らせてください。
その実績と媒体インセンティブを踏まえて、11月から貴社専用の特別プランを適用します。

■ 着地
10月に一律で値引きをするのではなく、10月の実績を見たうえで、
貴社にとって一番効果の大きい形で11月から適用する、という進め方です。

■ 想定QA
Q. 10月分は高いままですか？
A. 10月は原則、通常プランでの適用となります。ただしその1ヶ月の実績が、
   11月以降の調整の根拠になります。無駄な1ヶ月にはしません。

Q. どれくらい下がりますか？
A. 現時点では確約できません。10月の配信実績と媒体インセンティブ次第です。
   10月下旬に、貴社専用のシミュレーションをお出しします。

Q. 10月から何かできることはありますか？
A. あります。セグメント配信への切り替えは10月から着手できます。
   ここを進めておくほど、11月の調整余地も大きくなります。
"""

# ---------- レイアウト定数（スライド 10.83 × 7.5 inch） ----------
X0, TOTAL_W = 0.50, 9.83
GAP = 0.30
STEP_W = (TOTAL_W - GAP * 2) / 3
STEP_Y, STEP_H = 1.95, 3.00
NOTE_Y, NOTE_H = 5.30, 1.05


def draw_body(s):
    for i, (chip_t, head, body, ccol, fill, edge, bcol, ctxt) in enumerate(STEPS):
        x = X0 + i * (STEP_W + GAP)
        add_box(s, x, STEP_Y, STEP_W, STEP_H, fill=fill, line=edge,
                lw=1.5 if i == 2 else 1.0)

        cp = add_box(s, x + 0.14, STEP_Y + 0.18, STEP_W - 0.28, 0.50, fill=ccol)
        put_text(cp.text_frame, [{"runs": [(chip_t, 14, True, ctxt)], "align": "c"}],
                 anchor="m", ml=0.0, mr=0.0, mt=0.0, mb=0.0)

        add_text(s, x + 0.14, STEP_Y + 0.86, STEP_W - 0.28, 0.62,
                 [{"runs": [(head, 17, True, bcol)], "align": "c"}],
                 anchor="m", ml=0.0, mr=0.0)

        add_text(s, x + 0.18, STEP_Y + 1.62, STEP_W - 0.36, 1.10,
                 [{"runs": [(body, 12, None, bcol)], "ls": 1.35}],
                 anchor="t", ml=0.0, mr=0.0)

        if i < 2:
            ar = add_box(s, x + STEP_W + 0.04, STEP_Y + STEP_H / 2 - 0.16,
                         0.22, 0.32, fill=BORDER, shape=MSO_SHAPE.RIGHT_ARROW)
            ar.line.fill.background()

    add_box(s, X0, NOTE_Y, TOTAL_W, NOTE_H, fill=WHITE, line=NAVY)
    add_text(s, X0 + 0.30, NOTE_Y, TOTAL_W - 0.60, NOTE_H,
             [{"runs": [("※ ", 12, True, NAVY), (NOTE, 12, None, INK)], "ls": 1.35}],
             anchor="m", ml=0.0, mr=0.0)


def set_notes(s, text):
    s.notes_slide.notes_text_frame.text = text


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
    add_text(s, 0.60, 0.13, 8.60, 0.40,
             [{"runs": [(TITLE, 16, True, TNAVY)]}], anchor="m", ml=0.0, mr=0.0)
    add_text(s, 0.42, 0.60, 10.02, 0.84, LEAD, anchor="m", ml=0.0, mr=0.0)
    draw_body(s)
    set_notes(s, TALK_SCRIPT)
    prs.save(out)
    return out


def build_into_deck(src, out, page=9):
    """既存デッキの指定ページを、この内容で作り直す（他ページは触らない）。"""
    shutil.copyfile(src, out)
    prs = Presentation(out)
    s = prs.slides[page - 1]
    clear_slide(s)
    add_text(s, 0.47, 0.10, 8.60, 0.40,
             [{"runs": [(TITLE, 16, None, TNAVY)]}], anchor="m", ml=0.1, mr=0.0)
    add_text(s, 0.16, 0.60, 10.58, 0.84, LEAD, anchor="m", ml=0.0, mr=0.0)
    draw_body(s)
    set_notes(s, TALK_SCRIPT)
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
