# -*- coding: utf-8 -*-
"""教育（塾）業界 LINEOA施策提案（36枚）

_templates/DYM_LINEOA_FMT.pptx（業界汎用FMT・36枚）をコピーし、
全36枚を _drafts/教育（塾）業界_LINEOA提案原稿15枚.md の内容で構築する。

  python _build/build_juku.py
  python _build/qa_render.py 教育（塾）業界_LINEOA施策提案.pptx

【この資料の背骨】
保護者の検討は、塾を探す26日前に始まっている。塾が姿を現せるのは最後の数日だけ。
そして最初の問いは「いくら」ではなく「いつから」。
友だちは広告で買えるが、運用設計がなければ資産にならない（ビザビ371万人の凍結）。

【実測データ（★このデッキの強み。捏造なし）】
- S10 Googleトレンド：塾ピーク2/8・夏期講習7/12（_data/trends/教育塾/）
- S11 前後検索：起点「塾 費用」の26日タイムライン（_data/journey/教育塾/）
- S12 他社分析：友だち数実測＋ビザビ配信停止（_data/lineoa_competitors/）

【未取得データ（★破線の差込枠。数値は捏造していない）】
- S07/S08 塾のCPC・CPA・入会率・退会率＝公的統計が存在しない（社内実績待ち）
- S18 トライ・早稲アカの診断の実物スクショ
- S35 LINEヤフー公式の教育・スクール導入事例

【落とし穴メモ（_build/README.md より）】
- put_text() は必ず reset_tf() を通す（既存テキストへの追記事故を防ぐ）
- スライドの新規追加はしない。ベース36枚を clear_slide() して作り直す
- 一括置換をしない。数値は行・列を特定して書く
"""
import shutil
from pathlib import Path

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
SRC = str(ROOT / "_templates" / "DYM_LINEOA_FMT.pptx")
OUT = str(ROOT / "教育（塾）業界_LINEOA施策提案.pptx")

# ---- 配色（DYM 3色：ネイビー・炭黒・オレンジ。LINEグリーンはUI部分のみ）----
TNAVY = "002060"   # タイトル
NAVY = "1F285A"    # 打ち手・カード見出し
ORANGE = "ED7D31"  # 強調
RED = "C00000"     # 課題・警告
INK = "333333"     # 本文
MUT = "7F7F7F"
WHITE = "FFFFFF"
PALE = "F4F7FF"    # 淡ネイビー
PORANGE = "FCE4D6"  # 淡オレンジ
GREY = "F2F2F2"
BORDER = "D9D9D9"
GREEN = "06C755"   # LINE UI 専用
BEZEL = "2B2B2B"
SCREEN = "EFF2F7"
PRED = "FDF2F2"

# ---- FMTのレイアウト座標（cm）----
SW, SH = 27.52, 19.05
TITLE_XY = (1.52, 0.38, 24.4, 0.94)
LEAD_XY = (1.20, 1.80, 25.1, 1.90)
DIV_Y = 3.86
CX0, CW = 1.20, 25.12
CY0, CY1 = 4.30, 17.00
FOOT_Y = 17.35

shutil.copyfile(SRC, OUT)
prs = Presentation(OUT)
slides = list(prs.slides)
assert len(slides) == 36, len(slides)
assert (prs.slide_width, prs.slide_height) == (9906000, 6858000)


# ================= helpers（build_chumon_jutaku.py と共通の作法）=================
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
    """既存の段落・runを全消去（追記事故の防止）"""
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)


def put_text(tf, paras, anchor="t", ml=0.14, mr=0.14, mt=0.06, mb=0.06, wrap=True):
    reset_tf(tf)
    tf.word_wrap = wrap
    tf.margin_left = Cm(ml)
    tf.margin_right = Cm(mr)
    tf.margin_top = Cm(mt)
    tf.margin_bottom = Cm(mb)
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
        for item in p["runs"]:
            t, sz, b, c = item
            r = para.add_run()
            r.text = t
            set_font(r, sz, b, c)
    return tf


def T(slide, x, y, w, h, paras, anchor="t", **kw):
    box_ = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    put_text(box_.text_frame, paras, anchor=anchor, **kw)
    return box_


def one(text, sz, b=None, c=INK, align="l", sa=None, ls=None):
    d = {"runs": [(text, sz, b, c)], "align": align}
    if sa is not None:
        d["sa"] = sa
    if ls is not None:
        d["ls"] = ls
    return d


def multi(runs, align="l", sa=None, ls=None):
    """1段落に複数run（色・太さを混在させたいとき）"""
    d = {"runs": runs, "align": align}
    if sa is not None:
        d["sa"] = sa
    if ls is not None:
        d["ls"] = ls
    return d


def box(slide, x, y, w, h, fill=None, line=None, lw=1.0,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, dash=None):
    sp = slide.shapes.add_shape(shape, Cm(x), Cm(y), Cm(w), Cm(h))
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
        if dash:
            sp.line.dash_style = dash
    sp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    reset_tf(sp.text_frame)
    return sp


def card(slide, x, y, w, h, head, body, hcol=NAVY, fill=PALE,
         hsz=11, bsz=9, line=None, anchor="t", ls=1.18):
    sp = box(slide, x, y, w, h, fill=fill, line=line)
    paras = [one(head, hsz, True, hcol, sa=3)]
    for b in (body if isinstance(body, list) else [body]):
        if b:
            paras.append(one(b, bsz, None, INK, ls=ls, sa=1))
    put_text(sp.text_frame, paras, anchor=anchor, ml=0.22, mr=0.18, mt=0.14, mb=0.10)
    return sp


def clear_slide(slide):
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)


def frame(slide, title, lead):
    """FMT準拠：タイトル16pt紺（y=0.38）＋リード12pt（y=1.80・2行以内）＋区切り線"""
    clear_slide(slide)
    T(slide, *TITLE_XY, [one(title, 16, True, TNAVY)], anchor="m", ml=0, mr=0)
    T(slide, *LEAD_XY, [one(l, 12, None, INK, ls=1.28) for l in lead],
      anchor="m", ml=0, mr=0)
    ln = slide.shapes.add_connector(1, Cm(0), Cm(DIV_Y), Cm(SW), Cm(DIV_Y))
    ln.line.color.rgb = RGBColor.from_string(BORDER)
    ln.line.width = Pt(1.0)


def foot(slide, text):
    T(slide, CX0, FOOT_Y, CW, 0.9, [one(text, 7.5, None, MUT, ls=1.15)], ml=0, mr=0)


def badge(slide, x, y, w, h, text, fill=PORANGE, col=ORANGE, sz=8.5, dash=None):
    sp = box(slide, x, y, w, h, fill=fill, line=col, lw=1.0, dash=dash)
    put_text(sp.text_frame, [one(text, sz, True, col, align="c")],
             anchor="m", ml=0.06, mr=0.06, mt=0, mb=0)
    return sp


def band(slide, y, text, fill=NAVY, col=WHITE, sz=11.5, h=0.86, x=CX0, w=CW):
    sp = box(slide, x, y, w, h, fill=fill, radius=0.10)
    put_text(sp.text_frame, [one(text, sz, True, col, align="c")],
             anchor="m", ml=0.2, mr=0.2, mt=0, mb=0)
    return sp


def placeholder(slide, x, y, w, h, label, note):
    """★未取得データの差込枠（破線）。捏造しない。"""
    sp = box(slide, x, y, w, h, fill=WHITE, line=MUT, lw=1.25,
              dash=MSO_LINE_DASH_STYLE.DASH)
    put_text(sp.text_frame,
              [one(label, 11, True, MUT, align="c", sa=4),
               one(note, 8.5, None, MUT, align="c", ls=1.25)],
              anchor="m", ml=0.3, mr=0.3, mt=0.1, mb=0.1)
    return sp


def phone(slide, x, y, w, h, header, lines):
    """LINEトーク画面のモック。lines = [(kind, text)]
    kind: 'in'（BOT吹き出し）/ 'chip'（選択肢）/ 'btn'（CTA）/ 'note'（注記）"""
    box(slide, x, y, w, h, fill=BEZEL, radius=0.10)
    box(slide, x + 0.14, y + 0.42, w - 0.28, h - 0.56, fill=SCREEN, radius=0.02,
        shape=MSO_SHAPE.RECTANGLE)
    hd = box(slide, x + 0.14, y + 0.42, w - 0.28, 0.52, fill=GREEN,
             shape=MSO_SHAPE.RECTANGLE)
    put_text(hd.text_frame, [one(header, 7, True, WHITE, align="c")],
             anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)
    cy = y + 1.05
    iw = w - 0.62
    for kind, text in lines:
        nlines = text.count("\n") + 1
        bh = 0.30 + 0.30 * nlines
        if kind == "in":
            sp = box(slide, x + 0.31, cy, iw, bh, fill=WHITE, line=BORDER, lw=0.75)
            put_text(sp.text_frame, [one(l, 6.5, None, INK, ls=1.18)
                                     for l in text.split("\n")],
                     anchor="m", ml=0.14, mr=0.10, mt=0.05, mb=0.05)
        elif kind == "chip":
            sp = box(slide, x + 0.31, cy, iw, bh, fill=WHITE, line=GREEN, lw=0.9)
            put_text(sp.text_frame, [one(l, 6.5, None, "0B7A3B", align="c", ls=1.15)
                                     for l in text.split("\n")],
                     anchor="m", ml=0.08, mr=0.08, mt=0.03, mb=0.03)
        elif kind == "btn":
            sp = box(slide, x + 0.31, cy, iw, bh, fill=GREEN)
            put_text(sp.text_frame, [one(l, 6.8, True, WHITE, align="c", ls=1.15)
                                     for l in text.split("\n")],
                     anchor="m", ml=0.08, mr=0.08, mt=0.03, mb=0.03)
        else:  # note
            sp = slide.shapes.add_textbox(Cm(x + 0.31), Cm(cy), Cm(iw), Cm(bh))
            put_text(sp.text_frame, [one(l, 6.2, None, MUT, ls=1.15)
                                     for l in text.split("\n")],
                     anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)
        cy += bh + 0.12
    return cy


def richmenu(slide, x, y, w, h, tabs):
    """リッチメニュー3タブ×6ボタンのモック。
    tabs = [(タブ名, [ボタン文言6個]), ...]。最初のタブをアクティブ表示。"""
    tab_h = 0.55
    box(slide, x, y, w, h, fill=BEZEL, radius=0.06)
    tw = (w - 0.12) / len(tabs)
    for i, (name, _) in enumerate(tabs):
        active = (i == 0)
        tb = box(slide, x + 0.06 + i * tw, y + 0.06, tw - 0.03, tab_h,
                  fill=GREEN if active else "3F3F3F",
                  shape=MSO_SHAPE.RECTANGLE)
        put_text(tb.text_frame, [one(name, 8, True, WHITE, align="c")],
                 anchor="m", ml=0, mr=0, mt=0, mb=0)
    _, btns = tabs[0]
    gx, gy = x + 0.06, y + 0.06 + tab_h + 0.06
    gw, gh = w - 0.12, h - tab_h - 0.18
    cols, rows = 3, 2
    bw, bh = (gw - 0.06 * (cols - 1)) / cols, (gh - 0.06 * (rows - 1)) / rows
    for i, label in enumerate(btns):
        r, c = divmod(i, cols)
        bb = box(slide, gx + c * (bw + 0.06), gy + r * (bh + 0.06), bw, bh,
                  fill=SCREEN, line=BORDER, lw=0.75, radius=0.04)
        put_text(bb.text_frame, [one(l, 7.2, True, NAVY, align="c", ls=1.1)
                                 for l in label.split("\n")],
                 anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)


def simple_table(slide, x, y, w, h, headers, rows, col_w=None,
                  hsz=9, bsz=8.5, header_fill=NAVY, zebra=PALE,
                  align=None, row_h=None):
    """罫線テーブル（ネイティブpptxテーブル）。headers=[str], rows=[[str]]"""
    n_r, n_c = len(rows) + 1, len(headers)
    gf = slide.shapes.add_table(n_r, n_c, Cm(x), Cm(y), Cm(w), Cm(h))
    tbl = gf.table
    if col_w:
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = Cm(cw)
    if row_h:
        for r in tbl.rows:
            r.height = Cm(row_h)
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor.from_string(header_fill)
        c.margin_left = c.margin_right = Cm(0.12)
        c.margin_top = c.margin_bottom = Cm(0.05)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        put_text(c.text_frame, [one(htext, hsz, True, WHITE,
                                    align=(align[j] if align else "c"))],
                 anchor="m", ml=0.12, mr=0.12, mt=0.02, mb=0.02)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(
                zebra if (zebra and i % 2 == 0) else WHITE)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            put_text(c.text_frame, [one(str(val), bsz, None, INK,
                                        align=(align[j] if align else "l"))],
                     anchor="m", ml=0.12, mr=0.12, mt=0.02, mb=0.02)
    return gf


def find_shape(slide, needle):
    for sh in slide.shapes:
        if sh.has_text_frame and needle in sh.text_frame.text:
            return sh
    return None



# ============================================================
# S01 表紙
# ============================================================
s = slides[0]
sh = find_shape(s, "業界：")
if sh:
    put_text(sh.text_frame, [one("業界：教育（学習塾・業界汎用）", 12, True, TNAVY)], anchor="m")
sh = find_shape(s, "最上位パートナーの知見")
if sh:
    put_text(sh.text_frame,
             [one("保護者が動き出す26日前から伴走し、体験予約・入会率・継続率を同時に引き上げる", 13, True, INK)],
             anchor="m")

# ============================================================
# S02 資料アジェンダ
# ============================================================
s = slides[1]
frame(s, "資料アジェンダ",
      ["この資料が言っていること：保護者の検討は、塾を探す26日前に始まっている。",
       "塾が姿を現せるのは最後の数日だけ。その26日間を取りにいく提案です。"])
AGENDA = [
    ("1", "なぜLINEか", "S03-06"), ("2", "実態調査", "S07-12"),
    ("3", "全体設計", "S13-15"), ("4", "構築", "S16-20"),
    ("5", "配信設計", "S21-26"), ("6", "継続・工数", "S27-28"),
    ("7", "成果と体制", "S29-32"), ("8", "締め", "S33-36"),
]
gx, gy, gw, gh = CX0, CY0 + 0.3, CW, 11.6
cw = gw / 4
ch = gh / 2
for i, (n, t, r) in enumerate(AGENDA):
    row, col = divmod(i, 4)
    x = gx + col * cw
    y = gy + row * ch
    box(s, x + 0.12, y + 0.12, cw - 0.24, ch - 0.24, fill=PALE)
    T(s, x + 0.12, y + 0.22, cw - 0.24, 0.9,
      [one(n, 22, True, ORANGE, align="c")], anchor="t")
    T(s, x + 0.12, y + 0.95, cw - 0.24, 1.3,
      [one(t, 12.5, True, NAVY, align="c")], anchor="t")
    T(s, x + 0.12, y + ch - 0.55, cw - 0.24, 0.5,
      [one(r, 8.5, None, MUT, align="c")], anchor="t")
foot(s, "全8章・36枚｜主語は「LINE公式アカウントの運用」。広告運用の提案書ではない")

# ============================================================
# S03 市場データ｜子どもは減り、市場は伸びている
# ============================================================
s = slides[2]
frame(s, "市場データ｜ニーズ全体図",
      ["出生数は8年で27.7%減った。それなのに学習塾市場は伸びている。",
       "＝より少ない子どもから、より多くの売上を取る競争になっている。"])
nums = [
    ("出生数（2015→2023）", "▲27.7%", "1,005,677人 → 727,277人"),
    ("学習塾 市場規模（2020→2024）", "+13.5%", "4,374億円 → 4,965億円"),
    ("小学校 児童数（2015→2024）", "▲9.7%", "654.3万人 → 591.0万人"),
]
cw3 = (CW - 0.6) / 3
for i, (label, val, sub) in enumerate(nums):
    x = CX0 + i * (cw3 + 0.3)
    box(s, x, CY0 + 0.2, cw3, 4.6, fill=PALE)
    T(s, x, CY0 + 0.5, cw3, 0.9, [one(label, 11, True, NAVY, align="c", ls=1.2)], anchor="m")
    T(s, x, CY0 + 1.5, cw3, 1.6, [one(val, 30, True, ORANGE, align="c")], anchor="m")
    T(s, x, CY0 + 3.2, cw3, 1.3, [one(sub, 9.5, None, INK, align="c", ls=1.2)], anchor="m")
box(s, CX0, CY0 + 5.2, CW, 1.7, fill="FFF2CC")
T(s, CX0 + 0.3, CY0 + 5.2, CW - 0.6, 1.7,
  [one("母数は減り続ける。競争密度が上がるのはこれから。いま必要なのは、取りこぼしを拾う設計。",
       15, True, TNAVY, ls=1.3)], anchor="m")
box(s, CX0, CY0 + 7.3, CW, 2.5, fill=WHITE, line=BORDER, lw=1.0)
T(s, CX0 + 0.3, CY0 + 7.45, CW - 0.6, 2.2,
  [one("⚠ 市場規模は「定義」で2倍変わる。必ず併記すること", 10.5, True, RED, sa=4),
   one("経産省「特定サービス産業動態統計」4,965億円＝受講生500人以上の主要事業者のみ（市場の約4〜5割）。"
       "個人塾・中小は含まない。", 9, None, INK, ls=1.25, sa=2),
   one("矢野経済研究所の全体推計＝9,870億円（2023年度）。個人経営の補習塾・中小予備校まで含むため約2倍。"
       "どちらも正しく、定義が違うだけ。", 9, None, INK, ls=1.25)],
  anchor="t")
foot(s, "出典｜厚生労働省 人口動態統計（確定数）／経済産業省 特定サービス産業動態統計（2025年3月公表・2024年確報）／"
        "文部科学省 学校基本調査（令和6年度確報）／矢野経済研究所「教育産業市場に関する調査」2024年。"
        "※提案前にe-Statで統計表を開いて数値照合すること")

# ============================================================
# S04 塾のWebマーケ構造｜3つの壁
# ============================================================
s = slides[3]
frame(s, "現状の構造｜塾の集客が詰まる3つの壁",
      ["広告費の多くは「予約の手前」と「予約のあと」で溶けている。",
       "①は資金力の勝負。②③は仕組みで直せる＝広告費を増やさずに改善できる区間。"])
WALLS = [
    ("壁①　広告CPAの高騰",
     ["CPC 300〜800円（繁忙期・高単価ジャンルは1,200〜2,500円）",
      "体験申込CPA 25,000〜50,000円",
      "しかも塾のCPC/CPAには公的統計が存在しない",
      "＝適正値が誰にも見えないまま入札している"],
     "直すには「もっと払う」しかない", RED),
    ("壁②　Webでの離脱",
     ["LP CVR 2.0%＝100人来て98人が名前も残さず消える",
      "この98人は興味がないのではない",
      "「合う塾が分からない」63.5%の層（塾ナラ2025）",
      "＝判断材料が足りないだけ"],
     "仕組みで直せる", ORANGE),
    ("壁③　歩留まりの低下",
     ["予約の約20%が体験に来ない（無断キャンセル・直前辞退）",
      "Web広告経由の入会率 40〜50%",
      "一方、紹介・口コミ経由は 70〜80%超",
      "＝同じ体験でも、来た経路で入会率が30pt違う"],
     "仕組みで直せる（本提案の主戦場）", ORANGE),
]
cw3 = (CW - 0.6) / 3
for i, (head, body, tag, col) in enumerate(WALLS):
    x = CX0 + i * (cw3 + 0.3)
    box(s, x, CY0 + 0.2, cw3, 8.2, fill=WHITE, line=BORDER, lw=1.0)
    hb = box(s, x, CY0 + 0.2, cw3, 1.0, fill=col)
    put_text(hb.text_frame, [one(head, 12, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    T(s, x + 0.25, CY0 + 1.5, cw3 - 0.5, 5.2,
      [one("・" + b, 9.5, None, INK, ls=1.35, sa=4) for b in body], anchor="t")
    tb = box(s, x + 0.25, CY0 + 7.1, cw3 - 0.5, 1.0, fill=PALE)
    put_text(tb.text_frame, [one(tag, 9.5, True, NAVY, align="c", ls=1.15)],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
band(s, CY0 + 8.8, "本提案が触るのは②と③だけ。LPも広告も触りません。", fill=NAVY)
foot(s, "CPC/CPA・入会率は業界推計値（出典元の特定不可）。LP CVRはモデル値。"
        "★社内実績（CV地点別CPC/CVR/CPA）が入り次第すべて差し替える")

# ============================================================
# S05 保護者の心理｜検索しているのは40代の母親
# ============================================================
s = slides[4]
frame(s, "エンドユーザー行動心理｜検索しているのは40代の母親",
      ["塾選び最大の苦労は「料金」ではなく「うちの子に合う塾が分からない」63.5%。",
       "費用は9割が負担を感じつつ、6割が「妥当」。払う気はある。決められないだけ。"])
box(s, CX0, CY0 + 0.2, 11.6, 6.4, fill=PALE)
T(s, CX0 + 0.35, CY0 + 0.4, 11.0, 0.8,
  [one("誰が調べているか（LINEヤフー前後検索・実測）", 11.5, True, NAVY)], anchor="m")
FACTS = [
    ("性別", "女性 67.4%", "起点KW「塾」。「塾 費用」も65.5%"),
    ("個別クエリ", "女性 86〜88%", "「塾 いつから通わせる」86.5%／「高校受験 塾 いつから」87.9%"),
    ("年代", "40代 52.8%", "50代23.5%／30代11.1%（同音異義語ノイズ除外後）"),
    ("学年別の違い", "30代 22.1%", "「塾 費用」では30代が上昇＝小学生の保護者は30代"),
]
yy = CY0 + 1.4
for label, val, sub in FACTS:
    box(s, CX0 + 0.35, yy, 2.5, 1.05, fill=NAVY)
    T(s, CX0 + 0.35, yy, 2.5, 1.05, [one(label, 9.5, True, WHITE, align="c")], anchor="m")
    T(s, CX0 + 3.05, yy, 3.0, 1.05, [one(val, 14, True, ORANGE, align="c")], anchor="m")
    T(s, CX0 + 6.2, yy, 5.1, 1.05, [one(sub, 8.5, None, INK, ls=1.2)], anchor="m")
    yy += 1.22
T(s, CX0 + 0.35, CY0 + 6.35, 11.0, 0.5,
  [one("→ ChatGPT調査では「母親か父親かの比率を示す公開調査は無い」だったが、検索実データで裏付けが取れた",
       8.5, None, MUT, ls=1.15)], anchor="t")
VX = CX0 + 12.2
box(s, VX, CY0 + 0.2, 12.92, 6.4, fill=WHITE, line=BORDER, lw=1.0)
T(s, VX + 0.35, CY0 + 0.4, 12.2, 0.8,
  [one("保護者の心の声（データの翻訳）", 11.5, True, NAVY)], anchor="m")
VOICES = [
    ("うちの子に合うのか、通ってみるまで分からない", "塾選び最大の苦労 63.5%（塾ナラ2025・n=200）"),
    ("集団と個別、どっちが向いてるのか判断できない", "「情報が多すぎる」17.0%"),
    ("月謝の他に講習費がいくら乗るのか読めない", "費用に負担を感じる91.2%／一方で「妥当」57.7%（塾シル2026）"),
    ("本人にやる気がないのに入れても無駄では", "決定関与は保護者80.5%・子ども本人57.8%（明光2020）"),
]
yy = CY0 + 1.4
for v, src_ in VOICES:
    bb = box(s, VX + 0.35, yy, 12.2, 1.18, fill=PALE)
    put_text(bb.text_frame,
             [one("「" + v + "」", 10.5, True, INK, sa=2),
              one(src_, 8, None, MUT, ls=1.15)],
             anchor="m", ml=0.28, mr=0.2, mt=0.08, mb=0.08)
    yy += 1.32
band(s, CY0 + 7.0, "払う気はある。合うかどうかが分からないから決められない。"
                   "だから「申し込む前に、合う/合わないを見せてしまう」。", fill=NAVY, sz=13)
box(s, CX0, CY0 + 8.2, CW, 1.5, fill=WHITE, line=BORDER, lw=1.0)
T(s, CX0 + 0.3, CY0 + 8.3, CW - 0.6, 1.3,
  [one("※「勧誘されそうで問い合わせできない」は一般論として語られるが、比率を示す公開調査は存在せず、"
       "前後検索でも該当クエリはヒット0件だった。本資料では論拠として使わない（配信文面の「営業電話はしません」だけ残す）。",
       8.5, None, MUT, ls=1.25)], anchor="m")
foot(s, "出典｜LINEヤフー前後検索（直近1年 2022-05-09〜2023-05-14・実測）／塾ナラ2025（n=200）／"
        "塾シル・ユナイトプロジェクト2026（n=249）／明光ネットワークジャパン2020")

# ============================================================
# S06 できること／できないこと ＋ スコープ宣言
# ============================================================
s = slides[5]
frame(s, "本提案のスコープ｜LINEでできること・できないこと",
      ["LINE公式アカウントの運用提案です。広告運用の提案書ではありません。",
       "「授業の質」も「合格実績」も、LINEでは作れません。作れるのは接点と歩留まりです。"])
CAN = [
    "保護者が動き出す26日間に、接点を持ち続ける",
    "「合う/合わない」を申込前に提示する（診断）",
    "体験の無断キャンセルを減らす（前日リマインド）",
    "在籍生の保護者に成果を毎月見せる（継続率）",
    "欠席・振替の電話を減らす（教室長の工数）",
    "紹介を頼みやすい状態をつくる（口コミ47%）",
]
CANT = [
    "授業の質そのものを上げること",
    "合格実績をつくること",
    "講師の採用・育成",
    "月謝を上げずに利益を増やすこと",
    "友だちを集めるだけで入会を増やすこと",
    "（＝371万人集めても凍結すればゼロ／S12）",
]
hw = (CW - 0.6) / 2
for i, (title_, items, col, fillc) in enumerate(
        [("LINEでできること", CAN, NAVY, PALE), ("LINEではできないこと", CANT, MUT, GREY)]):
    x = CX0 + i * (hw + 0.6)
    box(s, x, CY0 + 0.2, hw, 7.6, fill=fillc)
    hb = box(s, x, CY0 + 0.2, hw, 1.0, fill=col)
    put_text(hb.text_frame, [one(title_, 12.5, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    yy = CY0 + 1.5
    for it in items:
        mark = "✓" if i == 0 else "×"
        T(s, x + 0.35, yy, hw - 0.7, 1.0,
          [multi([(mark + "　", 11, True, col), (it, 10, None, INK)], ls=1.25)], anchor="m")
        yy += 1.02
band(s, CY0 + 8.3, "触るのは「LPから落ちた後」と「予約が入った後」の2箇所だけ。LP CVRも広告も動かしません。",
     fill=NAVY, sz=12.5)
foot(s, "このスコープ宣言は毎回入れる。広告の話に引きずられると、運用の提案が薄まるため")

# ============================================================
# S07 CPC推移｜相場が存在しない市場
# ============================================================
s = slides[6]
frame(s, "広告の実態｜塾のCPC・CPAには「公開された相場」が存在しない",
      ["公的統計も業界統一の公開データも無いことを、再調査で確認した。",
       "＝いくらが適正なのか誰にも見えないまま、全社が入札に参加している。"])
box(s, CX0, CY0 + 0.2, 12.4, 5.4, fill=PALE)
T(s, CX0 + 0.35, CY0 + 0.45, 11.7, 0.8,
  [one("業界推計値（出典元の特定不可・★社内実績で差し替える）", 11.5, True, NAVY)], anchor="m")
EST = [
    ("CPC（一般KW）", "300〜800円", "「学習塾＋地域名」等"),
    ("CPC（繁忙期・高単価）", "1,200〜2,500円", "1〜3月・7月／医学部予備校・中学受験専門"),
    ("体験申込CPA", "25,000〜50,000円", "補習塾・個別指導"),
    ("体験申込CPA（大手）", "60,000〜100,000円", "大手中学受験塾・予備校"),
]
yy = CY0 + 1.4
for label, val, sub in EST:
    T(s, CX0 + 0.35, yy, 4.6, 0.95, [one(label, 10, True, INK)], anchor="m")
    T(s, CX0 + 5.0, yy, 3.6, 0.95, [one(val, 13, True, ORANGE, align="c")], anchor="m")
    T(s, CX0 + 8.8, yy, 3.3, 0.95, [one(sub, 8.5, None, MUT, ls=1.15)], anchor="m")
    yy += 1.02
VX = CX0 + 13.0
placeholder(s, VX, CY0 + 0.2, 12.12, 5.4,
            "★ 検索広告CPCの推移（3年）",
            "社内の塾・スクール系実績から\nCV地点別のCPC/CVR/CPAが入り次第、\nこの枠に折れ線で差し込む")
box(s, CX0, CY0 + 6.0, CW, 3.3, fill="FFF2CC")
T(s, CX0 + 0.4, CY0 + 6.2, CW - 0.8, 3.0,
  [one("相場が見えない市場で、入札を強めるのは最も危ない戦い方", 14, True, TNAVY, sa=6),
   one("・適正CPAが分からないので「いくらまで出すか」を誰も決められない", 10, None, INK, ls=1.3, sa=2),
   one("・繁忙期（1〜3月・7月）はCPCが最大3倍。全社が同じ時期に集中するため", 10, None, INK, ls=1.3, sa=2),
   one("→ だから本提案は入札では戦わない。「取りこぼしを拾う」設計に寄せる（S04の壁②③）",
       10.5, True, NAVY, ls=1.3)], anchor="t")
foot(s, "出典｜Gemini再調査（2026-09-15）で公的統計・業界統一データの不在を確認。"
        "推計値は広告代理店等の運用事例レポートに基づくが、出典元は特定できていない。★社内実績が唯一の一次根拠になる")

# ============================================================
# S08 広告実績CVR（★未取得・差込枠）
# ============================================================
s = slides[7]
frame(s, "広告実績CVR｜CV地点別の実績（DYM社内データ）",
      ["塾・スクール系の社内運用実績を、CV地点別に整理してこの枠に差し込む。",
       "公開データが存在しない領域なので、ここが本提案で唯一の一次根拠になる。"])
placeholder(s, CX0, CY0 + 0.2, CW, 6.2,
            "★ DYM社内 広告実績（塾・スクール系）",
            "CV地点別に「CPC／CVR／CPA」を記載する。\n"
            "CV地点の例：資料請求／体験授業の申込／教室見学の予約／LINE友だち追加\n"
            "※CV地点を明記しない数値はスライドに載せない（比較できないため）")
box(s, CX0, CY0 + 6.8, CW, 3.0, fill=WHITE, line=BORDER, lw=1.0)
T(s, CX0 + 0.4, CY0 + 6.95, CW - 0.8, 2.7,
  [one("この枠が埋まると連動して確定するもの", 11, True, NAVY, sa=5),
   one("・S04 壁①の数値（CPC/CPA）　・S15 施策展開図の前提　・S29 効果測定のKPI基準値",
       9.5, None, INK, ls=1.3, sa=2),
   one("・S30 費用プランの投資回収ライン　・巻末SIMのBefore（現状値）",
       9.5, None, INK, ls=1.3)], anchor="t")
foot(s, "★未取得｜上司確認3点セットの1つ目。CV地点別に揃わない場合は「CV地点＝体験申込」に統一して1行だけでも入れる")

# ============================================================
# S09 広告審査・表現面の懸念
# ============================================================
s = slides[8]
frame(s, "広告審査・表現面の懸念｜教育サービス特有の制約",
      ["塾の訴求は「成果」を約束したくなるが、そこに景品表示法の線がある。",
       "LINE配信の文面にも同じ制約がかかるため、設計段階で型を決めておく。"])
NG = [
    ("成果の断定", "「必ず成績が上がる」「絶対合格」",
     "根拠なく効果を保証する表現は優良誤認のリスク"),
    ("No.1・最上級表記", "「地域No.1」「合格率日本一」",
     "調査主体・範囲・時点の根拠がないと使えない"),
    ("体験談の扱い", "「3ヶ月で偏差値15UP」だけを大きく出す",
     "特異な事例を一般的な効果と誤認させない。条件の併記が要る"),
    ("価格の見せ方", "「月◯◯円〜」だけを強調",
     "講習費・教材費を含めた総額が分かるようにする（有利誤認の回避）"),
]
for i, (cat, ex, why) in enumerate(NG):
    yy = CY0 + 0.3 + i * 1.75
    box(s, CX0, yy, CW, 1.55, fill=WHITE, line=BORDER, lw=1.0)
    hb = box(s, CX0, yy, 4.6, 1.55, fill=PRED, line=None)
    put_text(hb.text_frame, [one(cat, 11, True, RED, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    T(s, CX0 + 4.9, yy, 9.4, 1.55, [one("例：" + ex, 10, True, INK, ls=1.2)], anchor="m")
    T(s, CX0 + 14.6, yy, 10.3, 1.55, [one(why, 9.5, None, MUT, ls=1.25)], anchor="m")
box(s, CX0, CY0 + 7.5, CW, 2.3, fill=PALE)
T(s, CX0 + 0.4, CY0 + 7.65, CW - 0.8, 2.0,
  [one("本提案の対処：成果を約束せず「判断材料を渡す」設計にする", 12.5, True, NAVY, sa=5),
   one("診断は「合う学び方」を示すもので、成績を保証しない。事例は条件（学年・期間・通塾回数）を必ず併記する。"
       "費用は月謝だけでなく年間総額を先に見せる（S18・S22の文面に反映済み）。",
       10, None, INK, ls=1.3)], anchor="t")
foot(s, "景品表示法（優良誤認・有利誤認）の一般的な考え方に基づく整理。"
        "★実際の配信文面は貴社・弊社双方の法務確認を通すこと。塾業界固有の公正競争規約の有無は未確認")

IMG = ROOT / "_images"
SHOT = ROOT / "_data" / "lineoa_competitors"


def pic(slide, path, x, y, w=None, h=None):
    """画像を貼る。無ければ破線の差込枠でしのぐ（捏造しない）"""
    p = Path(path)
    if not p.exists():
        placeholder(slide, x, y, w or 8, h or 5, "★ 画像未生成", str(p.name))
        return None
    kw = {}
    if w:
        kw["width"] = Cm(w)
    if h:
        kw["height"] = Cm(h)
    return slide.shapes.add_picture(str(p), Cm(x), Cm(y), **kw)


# ============================================================
# S10 シーズナリティ｜検索の山と、仕込みを始める日
# ============================================================
s = slides[9]
frame(s, "シーズナリティ｜「入会させたい月」ではなく「その26日前」から動く",
      ["「塾」のピークは2月上旬、「夏期講習」は7月中旬（冬期・春期の3〜5倍）。",
       "検討には26日かかる（S11）。逆算すると、仕込み開始日は1/13と6/16。"])
pic(s, IMG / "juku_trend_juku.png", CX0, CY0 + 0.15, w=12.3)
pic(s, IMG / "juku_trend_koushuu.png", CX0 + 12.8, CY0 + 0.15, w=12.3)
CAL = [
    ("1/13", "「塾」ピーク(2/8)の26日前", "新学年の募集はここから"),
    ("3/22", "春期講習ピーク", "規模は夏の1/3"),
    ("6/16", "夏期講習ピーク(7/12)の26日前", "年間最大の商戦はここから"),
    ("1月中旬〜2月上旬", "受験3種が集中", "大学1/25・高校1/18・中学2/1"),
]
cw4 = (CW - 0.9) / 4
for i, (d, t, sub) in enumerate(CAL):
    x = CX0 + i * (cw4 + 0.3)
    box(s, x, CY0 + 5.6, cw4, 2.5, fill=PALE)
    T(s, x, CY0 + 5.8, cw4, 0.9, [one(d, 15, True, ORANGE, align="c")], anchor="m")
    T(s, x + 0.2, CY0 + 6.7, cw4 - 0.4, 0.7, [one(t, 9.5, True, NAVY, align="c", ls=1.15)], anchor="m")
    T(s, x + 0.2, CY0 + 7.4, cw4 - 0.4, 0.6, [one(sub, 8.5, None, MUT, align="c", ls=1.15)], anchor="m")
band(s, CY0 + 8.4, "多くの塾が販促を始めるのは2月と7月。ピーク当日に走り出しても、保護者の検討はもう終わっている。",
     fill=NAVY, sz=12.5)
foot(s, "出典｜Googleトレンド（日本・2026-01-01〜09-15・週次／_data/trends/教育塾/）。"
        "※取得期間は年初来のため年内の季節性のみ。「◯年で△%上昇」はこのデータでは言えない。"
        "※「塾」と他3語（個別指導/予備校/家庭教師）は約8倍のスケール差があるため別グラフにしている")

# ============================================================
# S11 前後検索｜検討は26日前に始まっている（本提案の背骨）
# ============================================================
s = slides[10]
frame(s, "前後検索｜検討は26日前に始まり、塾が現れるのは最後の数日だけ",
      ["保護者が最初に動くのは「小学生 塾 月謝 相場」。そこから26日かけて塾を探し始める。",
       "塾の名前が検索されるのは、さらにその数日後。手前の26日間は完全な空白。"])
TL_X, TL_W = CX0 + 0.4, CW - 0.8
TL_Y = CY0 + 6.3
D_MIN, D_MAX = -30.0, 13.0


def dx(day):
    return TL_X + (day - D_MIN) / (D_MAX - D_MIN) * TL_W


# --- 区間の帯（軸の上に重ねる）---
gb = box(s, dx(-26), TL_Y - 0.5, dx(0) - dx(-26), 1.0, fill=GREY, line=BORDER, lw=1.0)
put_text(gb.text_frame, [one("接点なし　26日間", 13, True, MUT, align="c")],
         anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
rb = box(s, dx(1), TL_Y - 0.5, dx(10) - dx(1), 1.0, fill=PORANGE, line=ORANGE, lw=1.25)
put_text(rb.text_frame, [one("塾が接触できる\nのはここだけ", 8.5, True, ORANGE, align="c", ls=1.1)],
         anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)
# --- 軸 ---
ln = s.shapes.add_connector(1, Cm(TL_X), Cm(TL_Y), Cm(TL_X + TL_W), Cm(TL_Y))
ln.line.color.rgb = RGBColor.from_string(MUT)
ln.line.width = Pt(1.25)
for d in (-30, -20, -10, 0, 10):
    T(s, dx(d) - 1.3, TL_Y + 0.6, 2.6, 0.5,
      [one(("0日＝塾を探す" if d == 0 else f"{d}日"), 8.5, (d == 0),
           (NAVY if d == 0 else MUT), align="c")], anchor="t")
v = s.shapes.add_connector(1, Cm(dx(0)), Cm(TL_Y - 2.85), Cm(dx(0)), Cm(TL_Y + 0.55))
v.line.color.rgb = RGBColor.from_string(NAVY)
v.line.width = Pt(1.75)


def chip(day, y, h_, w_, kw, rel, fill=WHITE, line=BORDER, col=INK, relcol=MUT, lw=1.0, sz=8.5):
    x_ = min(max(dx(day) - w_ / 2, TL_X), TL_X + TL_W - w_)
    bb = box(s, x_, y, w_, h_, fill=fill, line=line, lw=lw)
    put_text(bb.text_frame,
             [one(kw, sz, (fill != WHITE), col, align="c", ls=1.12, sa=1),
              one(rel, 7.5, True, relcol, align="c")],
             anchor="m", ml=0.08, mr=0.08, mt=0.04, mb=0.04)
    y_mid = y + h_ if y < TL_Y else y
    c = s.shapes.add_connector(1, Cm(dx(day)), Cm(y_mid),
                               Cm(dx(day)), Cm(TL_Y - 0.5 if y < TL_Y else TL_Y + 0.5))
    c.line.color.rgb = RGBColor.from_string(BORDER)
    c.line.width = Pt(0.75)


# --- 検索【前】---
chip(-26.1, TL_Y - 6.15, 1.25, 4.6, "小学生 塾 月謝 相場", "関連度 4.11")
chip(-11.2, TL_Y - 4.65, 1.40, 4.6, "中学生 勉強\n親の関わり方", "関連度 3.24")
chip(-3.9, TL_Y - 3.20, 1.25, 4.6, "高校受験 塾 いつから", "関連度 3.53")
# --- 0日（最重要・オレンジ）---
zb = box(s, dx(0) - 6.3, TL_Y - 1.80, 6.3, 1.22, fill=PORANGE, line=ORANGE, lw=1.5)
put_text(zb.text_frame,
         [one("塾 いつから通わせる", 10.5, True, ORANGE, align="c", sa=1),
          one("関連度1位 4.98　／　塾 成績上がらない 4.67（2位）", 7.8, True, ORANGE, align="c")],
         anchor="m", ml=0.1, mr=0.1, mt=0.04, mb=0.04)
# --- 検索【後】---
chip(1.6, TL_Y + 1.05, 1.25, 5.4, "塾 夏期講習 料金 相場", "関連度 5.09")
chip(6.5, TL_Y + 2.55, 1.35, 7.4,
     "明光義塾 口コミ／森塾 口コミ／各社 料金", "関連度 3.3前後")

READ = [
    ("① 検討は1ヶ月超", "塾が広告・指名検索で拾えるのは最後の数日だけ。手前の3週間は接点ゼロ"),
    ("② 最初の問いは「いつから」", "関連度1位は「塾 いつから通わせる」4.98。料金の確認は+1.6日＝決断の後"),
    ("③ 口コミは最後に立つ", "指名×口コミが出た時点で比較は終盤。その前に友だちになるしかない"),
]
cw3 = (CW - 0.6) / 3
for i, (h_, b_) in enumerate(READ):
    x = CX0 + i * (cw3 + 0.3)
    card(s, x, CY0 + 10.6, cw3, 1.95, h_, [b_], hsz=11, bsz=9, fill=PALE)
foot(s, "出典｜LINEヤフー 前後検索データ（Journey）起点KW「塾 費用」／直近1年 2022-05-09〜2023-05-14・関連度3.0以上を抽出（_data/journey/教育塾/）。"
        "※起点KW「塾」のデータには同音異義語「熟」のノイズが16件混入するため集計から除外している")

# ============================================================
# S12 他社分析｜371万人のリストが、いま止まっている
# ============================================================
s = slides[11]
frame(s, "他社分析｜大手もLINEでは規模を持てていない。そして371万人が凍結している",
      ["友だち数を実測した。KUMON 2,525万人の一方、全国大手の本部は3万人台。",
       "さらにビザビ（371万人）は今年4月で配信を停止していた。"])
ROWS = [
    ("KUMON（公文）", "25,256,318", "A 全国大量", "本部＋教室別", "なし", WHITE),
    ("栄光ゼミナール／ビザビ", "3,713,373", "A 全国大量", "🔴 配信停止中", "なし", PRED),
    ("東進ハイスクール", "136,000", "B 本部小型", "本部＋校舎別", "なし", WHITE),
    ("武田塾", "32,871", "B 本部小型", "本部＋校舎別", "なし", WHITE),
    ("栄光ゼミナール（本部）", "30,000", "B 本部小型", "本部＋教室別", "なし", WHITE),
    ("早稲田アカデミー", "30,000", "B 本部小型", "本部", "◎ 学力診断", PALE),
    ("個別教室のトライ", "計測不能", "C 教室別分散", "教室ごとに分散", "◎ 性格別診断", PALE),
    ("明光義塾", "計測不能", "C 教室別分散", "本部＋FC教室別", "なし", WHITE),
    ("森塾", "未確認", "C 教室別分散", "本部＋教室別", "なし", WHITE),
]
TX, TW = CX0, 15.4
T(s, TX, CY0 + 0.15, TW, 0.6,
  [one("LINE公式アカウント 友だち数（2026-09-15 実機・認証済アカウント）", 10.5, True, NAVY)], anchor="m")
hy = CY0 + 0.85
HD = [("塾", 5.0), ("友だち数", 3.0), ("型", 2.5), ("アカウント構成", 2.6), ("診断", 2.3)]
cx = TX
for label, w_ in HD:
    hb = box(s, cx, hy, w_ - 0.05, 0.75, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    put_text(hb.text_frame, [one(label, 9, True, WHITE, align="c")],
             anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)
    cx += w_
ry = hy + 0.8
for name, n, typ, acc, diag, fillc in ROWS:
    cx = TX
    for val, w_, al, bold in ((name, 5.0, "l", True), (n, 3.0, "c", True),
                              (typ, 2.5, "c", False), (acc, 2.6, "c", False), (diag, 2.3, "c", False)):
        cb = box(s, cx, ry, w_ - 0.05, 0.82, fill=fillc, line=BORDER, lw=0.75,
                 shape=MSO_SHAPE.RECTANGLE)
        col = RED if ("停止" in str(val)) else (ORANGE if (bold and val not in ("計測不能", "未確認")) else INK)
        put_text(cb.text_frame, [one(str(val), 8.5 if not bold else 9.5, bold, col, align=al)],
                 anchor="m", ml=0.12, mr=0.08, mt=0, mb=0)
        cx += w_
    ry += 0.87
# 右：凍結の証跡（スクショは証跡として保全済み。ここは可読性優先で引用にする）
VX = CX0 + 15.8
box(s, VX, CY0 + 0.15, 9.32, 9.6, fill=PRED, line=RED, lw=1.5)
T(s, VX + 0.35, CY0 + 0.3, 8.6, 0.95,
  [one("🔴 371万人が、いま止まっている", 13.5, True, RED)], anchor="m")
qb = box(s, VX + 0.35, CY0 + 1.35, 8.62, 2.6, fill=WHITE, line=RED, lw=1.25)
put_text(qb.text_frame,
         [one("「せっかく登録いただいたのですが、", 9, None, INK, ls=1.25, sa=2),
          one("本アカウントは4月で配信を停止いたしました。", 10.5, True, RED, ls=1.25, sa=2),
          one("今後は『栄光ゼミナール』にて…こちらから\nお友だち追加をお願いいたします」", 9, None, INK, ls=1.25, sa=3),
          one("＝ 実際に友だち追加して返ってきたあいさつメッセージ（2026-09-15・実機）",
              7.5, None, MUT, ls=1.15)],
         anchor="m", ml=0.26, mr=0.22, mt=0.1, mb=0.1)
FACT = [("このアカウントの友だち数", "3,713,373人"),
        ("移行先の栄光ゼミナール本体", "30,000人"),
        ("移行できた割合", "1%未満")]
yy = CY0 + 4.25
for label, val in FACT:
    T(s, VX + 0.45, yy, 5.2, 0.74, [one(label, 9, None, INK)], anchor="m")
    T(s, VX + 5.5, yy, 3.4, 0.74,
      [one(val, 12 if "1%" not in val else 15, True,
           RED if "1%" in val else ORANGE, align="r")], anchor="m")
    yy += 0.80
cb = box(s, VX + 0.35, CY0 + 6.85, 8.62, 2.75, fill=WHITE, line=RED, lw=1.25)
put_text(cb.text_frame,
         [one("友だちは広告で買える。", 12, True, NAVY, align="c", sa=3),
          one("でも運用設計がなければ、資産にならない。", 12, True, NAVY, align="c", sa=6),
          one("370万集めてもゼロになる。", 9.5, None, INK, align="c", ls=1.25, sa=2),
          one("数千人でも、設計があれば戦える。", 9.5, None, INK, align="c", ls=1.25)],
         anchor="m", ml=0.22, mr=0.22, mt=0.12, mb=0.12)
band(s, CY0 + 10.1, "空白は「規模」ではなく「規模 × 診断の両立」。"
                    "規模を持つKUMON・ビザビも、診断で"
                    "「合う/合わない」を見せていない。", fill=NAVY, sz=12)
foot(s, "出典｜LINE公式アカウント（page.line.me・実機／いずれも認証済）2026-09-15取得。証跡は _data/lineoa_competitors/。"
        "第三者サイト（lineoa.jp）の数値は使用していない。"
        "★実名掲載は上司確認事項。★KUMON 2,525万人の内訳（スタンプ/CPF由来か）は未確認のため推定と明記している")


# ============================================================
# S13 カスタマージャーニー｜8フェーズ × CV3段
# ============================================================
s = slides[12]
frame(s, "全体設計｜カスタマージャーニー8フェーズと、CVの3段",
      ["塾は「買う人（保護者）」と「使う人（子ども）」が違う。だからCVを3段に分ける。",
       "軽いCVで先に接点を取り、重いCVへ引き上げる。入会後もLTVのために設計する。"])
PHASES = [
    ("①", "接触\n（認知・流入）", "広告・チラシ・口コミ", GREY),
    ("②", "サイト離脱\n（対策）", "LP CVR 2%＝98%が離脱", PORANGE),
    ("③", "育成\n（興味喚起）", "26日間の空白を埋める", PORANGE),
    ("④", "リード獲得\n(web/LINE)", "CV① 友だち追加＋診断", PALE),
    ("⑤", "リード有効化", "CV② 体験・見学の予約", PALE),
    ("⑥", "有効リードの\n再育成", "体験に来なかった層へ", PORANGE),
    ("⑦", "マネタイズ\n（成約）", "CV③ 入会", PALE),
    ("⑧", "LTV最大化\n・紹介", "継続（退会防止）＋紹介", PALE),
]
cw8 = (CW - 0.35 * 7) / 8
for i, (n, ph, sub, fillc) in enumerate(PHASES):
    x = CX0 + i * (cw8 + 0.35)
    box(s, x, CY0 + 0.2, cw8, 3.5, fill=fillc, line=BORDER, lw=1.0)
    T(s, x, CY0 + 0.35, cw8, 0.7, [one(n, 15, True, NAVY, align="c")], anchor="m")
    T(s, x + 0.1, CY0 + 1.05, cw8 - 0.2, 1.3,
      [one(ph, 9.5, True, INK, align="c", ls=1.15)], anchor="m")
    T(s, x + 0.1, CY0 + 2.4, cw8 - 0.2, 1.2,
      [one(sub, 8, None, MUT, align="c", ls=1.15)], anchor="m")
    if i < 7:
        ar = box(s, x + cw8 + 0.05, CY0 + 1.6, 0.25, 0.5, fill=MUT,
                 shape=MSO_SHAPE.RIGHT_ARROW)
CVS = [
    ("CV①　軽いCV", "LINE友だち追加 ＋ 30秒診断",
     ["氏名・電話番号は不要（匿名で始まる）", "④リード獲得に位置づけ",
      "ここを取らないと26日間の空白が埋まらない"], "④"),
    ("CV②　主CV", "体験授業・教室見学の予約",
     ["診断で「合う/合わない」を確認した後に来る", "⑤リード有効化に位置づけ",
      "予約の質が上がるので入会率も上がる"], "⑤"),
    ("CV③　最終CV／LTV", "入会 → 継続 → 紹介",
     ["入会がゴールではない（在籍年数がLTVを決める）", "⑦マネタイズ＋⑧LTV最大化",
      "紹介経由の入会率は70〜80%超"], "⑦⑧"),
]
cw3 = (CW - 0.6) / 3
for i, (head, sub, items, ph) in enumerate(CVS):
    x = CX0 + i * (cw3 + 0.3)
    box(s, x, CY0 + 4.3, cw3, 5.4, fill=WHITE, line=NAVY, lw=1.5)
    hb = box(s, x, CY0 + 4.3, cw3, 1.0, fill=NAVY)
    put_text(hb.text_frame, [one(head, 12, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    T(s, x + 0.25, CY0 + 5.5, cw3 - 0.5, 1.1,
      [one(sub, 12, True, ORANGE, align="c", ls=1.2)], anchor="m")
    T(s, x + 0.35, CY0 + 6.75, cw3 - 0.7, 2.3,
      [one("・" + it, 9, None, INK, ls=1.3, sa=4) for it in items], anchor="t")
    badge(s, x + cw3 - 2.2, CY0 + 9.05, 1.9, 0.55, "フェーズ " + ph, sz=8)
band(s, CY0 + 10.1, "軽いCVで先に接点を取り、診断で納得させてから重いCVへ。入会後は継続と紹介まで設計する。",
     fill=NAVY, sz=12.5)
foot(s, "CV3段の設計はDYM見解。②③⑥が本提案の主戦場（広告費を増やさずに改善できる区間）")

# ============================================================
# S14 施策全体像｜LINEは新規獲得と退会防止の二毛作
# ============================================================
s = slides[13]
frame(s, "施策全体像｜LINEの後半戦は「辞めさせない」",
      ["入会して終わりではない。塾の売上構造で重いのは在籍年数（LTV）のほう。",
       "退会を1人防ぐことは、新規を1人獲得するのと同じ。しかも獲得コストはゼロ。"])
FLOW = [
    ("広告・チラシ\n口コミ", GREY, "①"),
    ("LP", GREY, "①"),
    ("離脱防止POP\nCPF広告", PORANGE, "②"),
    ("友だち追加\n＋30秒診断", PALE, "④"),
    ("14日\nナーチャリング", PALE, "③⑤"),
    ("体験予約", PALE, "⑤"),
    ("前日\nリマインド", PORANGE, "⑤"),
    ("入会", PALE, "⑦"),
    ("在籍中の保護者接点\n離脱予兆検知", PALE, "⑧"),
    ("継続・紹介", PALE, "⑧"),
]
cwf = (CW - 0.28 * 9) / 10
for i, (t_, fillc, ph) in enumerate(FLOW):
    x = CX0 + i * (cwf + 0.28)
    box(s, x, CY0 + 0.3, cwf, 2.3, fill=fillc, line=BORDER, lw=1.0)
    T(s, x + 0.06, CY0 + 0.45, cwf - 0.12, 1.5,
      [one(t_, 8.5, True, INK, align="c", ls=1.15)], anchor="m")
    T(s, x + 0.06, CY0 + 2.0, cwf - 0.12, 0.5,
      [one(ph, 8, True, MUT, align="c")], anchor="m")
    if i < 9:
        box(s, x + cwf + 0.03, CY0 + 1.25, 0.22, 0.4, fill=MUT, shape=MSO_SHAPE.RIGHT_ARROW)
ub = box(s, CX0, CY0 + 2.75, (cwf + 0.28) * 7 - 0.28, 0.75, fill=NAVY)
put_text(ub.text_frame, [one("前半戦：新規獲得（ここで終わる提案が多い）", 10.5, True, WHITE, align="c")],
         anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
ub2 = box(s, CX0 + (cwf + 0.28) * 7, CY0 + 2.75, (cwf + 0.28) * 3 - 0.28, 0.75, fill=ORANGE)
put_text(ub2.text_frame, [one("後半戦：退会防止・紹介", 10.5, True, WHITE, align="c")],
         anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
MAP = [
    ("離脱防止ポップアップ", "未CVの98%", "CV①（友だち化）", "LP離脱時に「まだ決めなくて大丈夫です」"),
    ("30秒ぴったり塾診断", "友だち全員", "CV②（体験予約）", "「合う/合わない」を先に提示し、回答をタグ化"),
    ("14日ステップ配信", "未予約の友だち", "CV②（体験予約）", "売り込みは2回だけ。残りは家庭で使える情報"),
    ("体験の前日リマインド", "予約者", "歩留まり（実施率）", "持ち物・アクセス・変更1タップ。キャンセル料なし"),
    ("月次の学習レポート", "在籍生の保護者", "継続率（LTV）", "退会検討理由の50.6%はコミュニケーション不足"),
    ("離脱予兆検知", "在籍生", "継続率（LTV）", "欠席2回連続／宿題未提出／未読3週間で自動声かけ"),
    ("紹介キャンペーン", "在籍生の保護者", "CV①〜③（紹介）", "塾を知るきっかけ1位は口コミ47%"),
]
simple_table(s, CX0, CY0 + 3.9, CW, 6.0,
             ["施策", "誰を拾うか", "どのCVに効くか", "中身"],
             MAP, col_w=[5.6, 4.0, 4.6, 10.92],
             align=["l", "c", "c", "l"], hsz=9.5, bsz=8.5)
foot(s, "退会1人の防止＝月謝28,000円×残存18ヶ月≒50万円（モデル値）。新規1人の獲得より安く、確実")

# ============================================================
# S15 施策展開図｜初期構築と月次運用
# ============================================================
s = slides[14]
frame(s, "施策展開図｜初期に何を作り、毎月何を回すか",
      ["初期構築は1〜2ヶ月。作り込むのは診断・リッチメニュー・14日ステップの3つ。",
       "月次は「企画配信・セグメント配信・予兆検知・レポート」の4本を定常で回す。"])
INIT = ["LINE公式アカウント開設・プロフィール設計", "30秒ぴったり塾診断（4問）の設計と実装",
        "リッチメニュー3タブ×6枠の制作", "あいさつメッセージ（営業電話しない宣言を含む）",
        "14日ステップ配信（14通）の原稿・実装", "キーワード自動応答（料金・振替・持ち物）",
        "LP離脱防止ポップアップの設置", "タグ設計（学年／動機／適性／温度）"]
MONTH = ["企画配信（月1〜2本・年間カレンダーに沿って）", "セグメント配信（診断タグ別に出し分け）",
         "体験予約者への前日リマインド（自動）", "在籍生の保護者へ月次学習レポート",
         "離脱予兆検知のトリガー確認と声かけ", "GA・LINE分析のレポーティング",
         "定例会（配信結果の振り返りと翌月の企画）", "文面・リッチメニューの改善"]
hw = (CW - 0.6) / 2
for i, (title_, items, col, sub) in enumerate(
        [("初期（初動1〜2ヶ月）", INIT, NAVY, "ここを作り込まないと、あとの配信が全部弱くなる"),
         ("月次（定常運用）", MONTH, ORANGE, "毎月これを回す。DYMが実務を持つ")]):
    x = CX0 + i * (hw + 0.6)
    box(s, x, CY0 + 0.2, hw, 8.9, fill=WHITE, line=BORDER, lw=1.0)
    hb = box(s, x, CY0 + 0.2, hw, 1.0, fill=col)
    put_text(hb.text_frame, [one(title_, 12.5, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    yy = CY0 + 1.5
    for it in items:
        bb = box(s, x + 0.3, yy, hw - 0.6, 0.78, fill=PALE if i == 0 else PORANGE)
        put_text(bb.text_frame, [one(it, 9.5, None, INK, ls=1.15)],
                 anchor="m", ml=0.25, mr=0.15, mt=0, mb=0)
        yy += 0.86
    T(s, x + 0.3, CY0 + 8.35, hw - 0.6, 0.6, [one(sub, 8.5, None, MUT, align="c")], anchor="m")
band(s, CY0 + 9.5, "アカウント開設・リッチメニュー・診断・ステップ配信・タグ管理は、すべて費用内（無償付帯）。",
     fill=NAVY, sz=12)
foot(s, "別途費用：LP離脱防止（Sitelead 初期10万円＋月5万円）／通知メッセージ／API連携。"
        "配信ツールはLstepまたはHachidoriのどちらか一方を要件定義で選定する")

# ============================================================
# S16 友だち追加動線
# ============================================================
s = slides[15]
frame(s, "構築①｜友だち追加の動線　CTAは「今すぐ体験」ではなく「受け取る」",
      ["26日間の空白を埋めるには、まず匿名で接点を持つこと。",
       "個人情報の要求をゼロにするだけで、拾える層が変わる。"])
RULES = [("個人情報の要求ゼロ", "タップ1回で友だち追加が完了。氏名・電話・成績の入力を求めない"),
         ("「決めなくていい」と明言", "検討初期（26日前）の層は、決断を迫られると離脱する"),
         ("営業しない宣言を入口から", "※比率データは無いが、明記のコストはゼロで効果は大きい")]
for i, (h_, b_) in enumerate(RULES):
    card(s, CX0, CY0 + 0.25 + i * 1.5, 11.6, 1.35, h_, [b_], hsz=10.5, bsz=8.5, fill=PALE)
T(s, CX0, CY0 + 5.0, 11.6, 0.7,
  [one("導線は3つ（LP・広告・オフライン）", 11.5, True, NAVY)], anchor="m")
ROUTES = [("LP離脱防止ポップアップ", "離脱動作を検知して表示。Sitelead（初期10万＋月5万・別途）"),
          ("LINE広告 CPF（友だち追加課金）", "フォーム入力が存在しない。検討初期層を低単価でリスト化"),
          ("オフライン（チラシ・教室掲示）", "既存のチラシにQRを1つ足すだけ。折込の反応率低下を補う")]
for i, (h_, b_) in enumerate(ROUTES):
    card(s, CX0, CY0 + 5.8 + i * 1.5, 11.6, 1.35, h_, [b_], hsz=10.5, bsz=8.5,
         fill=WHITE, line=BORDER)
VX = CX0 + 12.2
T(s, VX, CY0 + 0.25, 12.92, 0.7, [one("クリエイティブ実文（そのまま使えます）", 11.5, True, NAVY)], anchor="m")
box(s, VX, CY0 + 1.0, 12.92, 4.3, fill=PALE)
T(s, VX + 0.35, CY0 + 1.2, 12.2, 3.9,
  [one("▼ LP離脱防止ポップアップ", 10, True, NAVY, sa=5),
   one("まだ、塾を決めなくて大丈夫です。", 13, True, INK, sa=5),
   one("30秒の無料診断で、お子さまに合う学び方（集団／個別／映像）だけ先にわかります。"
       "営業のお電話は一切ありません。", 9.5, None, INK, ls=1.3, sa=5),
   one("［ LINEで30秒診断を受け取る ］", 11, True, WHITE if False else ORANGE, sa=3),
   one("※お名前・電話番号は不要です", 8.5, None, MUT)], anchor="t")
box(s, VX, CY0 + 5.6, 12.92, 4.3, fill=PORANGE)
T(s, VX + 0.35, CY0 + 5.8, 12.2, 3.9,
  [one("▼ LINE広告 CPF", 10, True, ORANGE, sa=5),
   one("うちの子、塾に行くべき？", 13, True, INK, sa=5),
   one("テストの点が下がってきたら読むLINE。学年別「いま家庭ですべきこと」を週1で無料配信中。",
       9.5, None, INK, ls=1.3, sa=5),
   one("［ 友だち追加して受け取る ］", 11, True, ORANGE, sa=3),
   one("※前後検索の実測で、関連度1位は「塾 いつから通わせる」。"
       "料金訴求より「いつから」で刺す", 8.5, None, MUT, ls=1.2)], anchor="t")
foot(s, "CTAの言い換えはDYM見解。訴求の軸（料金ではなく「いつから」）はS11の前後検索実測に基づく")

# ============================================================
# S17 あいさつメッセージ
# ============================================================
s = slides[16]
frame(s, "構築②｜あいさつメッセージ　友だち追加の10秒後に何が届くか",
      ["1通目で「誰が・何を・いつ送るか」と「営業電話はしない」を先に伝える。",
       "アンケート（診断）は1通目に置く。参加率が最も高くなる。"])
phone(s, CX0 + 0.5, CY0 + 0.2, 7.2, 12.0, "◯◯塾 △△教室",
      [("in", "はじめまして！\n◯◯塾 △△教室・教室長の佐藤です😊"),
       ("in", "友だち追加ありがとうございます。\nこのLINEでは、\n📊 学年別「いま家庭ですべきこと」\n"
              "📝 テスト前の勉強のコツ\n💰 塾の費用のリアル\nを週1回だけお届けします。"),
       ("in", "⚠️ こちらから営業のお電話をする\nことは一切ありません。\n"
              "ご質問には私が直接お答えします\n（平日14〜21時）。"),
       ("in", "まずは30秒だけ👇\n【ぴったり塾診断】で、お子さまに\n合う学び方をチェックしてみて\nください。"),
       ("btn", "▶ 診断をはじめる")])
VX = CX0 + 8.4
DESIGN = [
    ("① 担当者名を出す", "「教室長の佐藤です」。BOTではなく人がいると示す。塾は「先生」で選ばれる業界"),
    ("② 有人対応の時間帯を明記", "「平日14〜21時」。いつ返事が来るか分かると質問のハードルが下がる"),
    ("③ 配信頻度を先に約束する", "「週1回だけ」。頻度を明示するとブロック率が下がる"),
    ("④ 最大の不安に先回りする", "「営業のお電話は一切ありません」。※一般論だが明記コストはゼロ"),
    ("⑤ 診断は1通目に置く", "後ろに回すほど参加率が落ちる。ここで属性を取れないと配信が全部一斉配信になる"),
]
T(s, VX, CY0 + 0.25, 16.7, 0.7, [one("この5つを外すと効かない", 12, True, NAVY)], anchor="m")
for i, (h_, b_) in enumerate(DESIGN):
    card(s, VX, CY0 + 1.1 + i * 1.65, 16.72, 1.5, h_, [b_], hsz=11, bsz=9.5, fill=PALE)
box(s, VX, CY0 + 9.5, 16.72, 2.4, fill=WHITE, line=NAVY, lw=1.25)
T(s, VX + 0.35, CY0 + 9.65, 16.0, 2.1,
  [one("なぜ1通目で「営業しない」と言うのか", 11, True, NAVY, sa=4),
   one("保護者の検討は26日前から始まっている（S11）。その段階の人に決断を迫ると離脱する。"
       "先に「決めなくていい・営業しない」と伝えて、26日間ずっと隣にいることのほうが、"
       "最後の数日で指名検索に勝つより確実。", 9.5, None, INK, ls=1.3)], anchor="t")
foot(s, "文面はそのまま使える実文。教室長名・対応時間・配信頻度のみ貴社の運用に合わせて差し替える")

# ============================================================
# S18 診断・アンケート
# ============================================================
s = slides[17]
frame(s, "構築③｜30秒ぴったり塾診断　保護者の63.5%が詰まる「合うか分からない」を解く",
      ["4問すべて選択式・自由入力ゼロ。回答がそのまま配信セグメントのタグになる。",
       "診断は集客ツールであると同時に、早期退塾を防ぐリテンション施策でもある。"])
QS = [
    ("Q1", "お子さまの学年は？", "小1〜4／小5・6／中1・2／中3／高1・2／高3・既卒", "学年タグ"),
    ("Q2", "いちばん近いお悩みは？", "成績が下がってきた／やる気が出ない／受験対策を始めたい／授業についていけない", "動機タグ"),
    ("Q3", "お子さまの性格に近いのは？", "競争があると燃える／マイペースにコツコツ／質問するのが苦手／集中が続かない", "適性タグ（集団/個別の判定）"),
    ("Q4", "塾はいつから？", "今すぐ／今学期中に／次の学年から／情報収集だけ", "温度タグ"),
]
T(s, CX0, CY0 + 0.2, 15.0, 0.7, [one("診断4問（すべて選択式）", 11.5, True, NAVY)], anchor="m")
for i, (n, q, a, tag) in enumerate(QS):
    yy = CY0 + 1.0 + i * 1.85
    box(s, CX0, yy, 15.0, 1.65, fill=WHITE, line=BORDER, lw=1.0)
    nb = box(s, CX0 + 0.2, yy + 0.25, 1.1, 1.15, fill=NAVY)
    put_text(nb.text_frame, [one(n, 11, True, WHITE, align="c")], anchor="m", ml=0, mr=0, mt=0, mb=0)
    T(s, CX0 + 1.5, yy + 0.15, 9.6, 0.65, [one(q, 10.5, True, INK)], anchor="m")
    T(s, CX0 + 1.5, yy + 0.8, 9.6, 0.7, [one(a, 8.5, None, MUT, ls=1.15)], anchor="m")
    tb = box(s, CX0 + 11.4, yy + 0.35, 3.4, 0.95, fill=PORANGE)
    put_text(tb.text_frame, [one("→ " + tag, 8.5, True, ORANGE, align="c", ls=1.1)],
             anchor="m", ml=0.08, mr=0.08, mt=0, mb=0)
VX = CX0 + 15.4
T(s, VX, CY0 + 0.2, 9.72, 0.7, [one("診断結果の型｜共感 → 実例 → 費用 → CTA", 11.5, True, NAVY)], anchor="m")
box(s, VX, CY0 + 1.0, 9.72, 7.4, fill=PALE)
T(s, VX + 0.35, CY0 + 1.2, 9.0, 7.0,
  [one("【診断結果】お子さまは「じっくり個別型」です📘", 11, True, NAVY, sa=6),
   one("質問が苦手なお子さまは、集団授業だと「わからないまま次の単元へ」が起きやすいタイプです。",
       9, None, INK, ls=1.3, sa=6),
   one("同じタイプの中2生は、個別指導で「質問しなくても講師から声をかける」形に変えて、"
       "数学58点→82点（3ヶ月・当教室例※）になりました。", 9, None, INK, ls=1.3, sa=6),
   one("費用は週1コマ 月◯◯円〜。季節講習などの追加費用も、先にすべてお見せします。",
       9, None, INK, ls=1.3, sa=6),
   one("まずは教室の空気だけ、見に来ませんか？\n（体験後の勧誘連絡はしない決まりです）",
       9, None, INK, ls=1.3, sa=6),
   one("▶ 体験の空き日程を見る", 10.5, True, ORANGE, sa=4),
   one("※ 事例数値は貴社の実績に差し替える。学年・期間・通塾回数の条件を必ず併記（S09）",
       8, None, MUT, ls=1.2)], anchor="t")
box(s, VX, CY0 + 8.7, 9.72, 3.2, fill=WHITE, line=NAVY, lw=1.25)
T(s, VX + 0.3, CY0 + 8.85, 9.1, 2.9,
  [one("診断が「継続率」にも効く理由", 10.5, True, NAVY, sa=4),
   one("1年未満で退塾した家庭の43%が「入塾前に十分調べなかった」と回答（1年以上継続は14%）。"
       "＝入口で適合を確認させることが、そのまま早期退塾の予防になる。",
       9, None, INK, ls=1.3, sa=4),
   one("出典：DeltaX「塾選」退塾に関する調査 2024年（n=65）", 7.5, None, MUT)], anchor="t")
foot(s, "費用は結果の後半に置く（先に出すと離脱する）。"
        "★競合ではトライ（性格別学習法診断）・早稲田アカデミー（学力診断シミュレーション）が実装済み。実物スクショは未取得")

# ============================================================
# S19 リッチメニュー
# ============================================================
s = slides[18]
frame(s, "構築④｜リッチメニュー　検討中・予約済み・塾生の3つの顔に出し分ける",
      ["常設の受付窓口。左上は常に「いま一番押してほしい行動」を置く。",
       "タブ③に退会相談を隠さず置く。相談を受け止める塾のほうが、結果として続く。"])
TABS = [
    ("はじめての方", ["🎯 30秒\nぴったり塾診断", "💰 料金まるわかり\n（講習費まで）", "📈 成績アップ事例",
                      "🏫 教室・講師を見る", "📅 無料体験を予約", "❓ よくある質問"]),
    ("体験予約済みの方", ["📋 体験当日の流れ", "🗺️ アクセス・持ち物", "🔁 日程変更（1タップ）",
                          "👨‍🏫 担当講師の紹介", "💬 保護者の声", "✉️ LINEで質問する"]),
    ("塾生・保護者の方", ["🙋 欠席・振替連絡\n（電話不要）", "📆 面談を予約", "📢 お知らせ・時間割",
                          "📔 宿題・学習記録", "🎁 ごきょうだい紹介", "🍀 休会・退会のご相談"]),
]
richmenu(s, CX0, CY0 + 0.3, 8.6, 6.6, TABS)
T(s, CX0, CY0 + 7.1, 8.6, 0.6,
  [one("実装：LINE大サイズ 2500×1686px／2列×3行×3タブ", 8.5, None, MUT, align="c")], anchor="m")
VX = CX0 + 9.2
for i, (name, btns) in enumerate(TABS):
    x = VX + i * ((CW - 9.2) / 3 + 0.0)
    w_ = (CW - 9.2) / 3 - 0.3
    box(s, x, CY0 + 0.3, w_, 7.6, fill=WHITE, line=BORDER, lw=1.0)
    hb = box(s, x, CY0 + 0.3, w_, 0.9, fill=[NAVY, ORANGE, "2E7D32"][i])
    put_text(hb.text_frame, [one(f"タブ{'①②③'[i]}　{name}", 10.5, True, WHITE, align="c")],
             anchor="m", ml=0.08, mr=0.08, mt=0, mb=0)
    yy = CY0 + 1.4
    for j, b in enumerate(btns):
        bb = box(s, x + 0.25, yy, w_ - 0.5, 0.95, fill=PALE if j == 0 else WHITE,
                 line=NAVY if j == 0 else BORDER, lw=1.25 if j == 0 else 0.75)
        put_text(bb.text_frame,
                 [one(b.replace("\n", " "), 8.5, j == 0, NAVY if j == 0 else INK, ls=1.1)],
                 anchor="m", ml=0.15, mr=0.1, mt=0, mb=0)
        yy += 1.02
    T(s, x + 0.25, CY0 + 7.5, w_ - 0.5, 0.5,
      [one(["左上＝最重要CTA", "予約後の不安をつぶす", "電話を減らす＋継続"][i],
           8, True, MUT, align="c")], anchor="m")
box(s, CX0, CY0 + 8.3, CW, 2.5, fill=PALE)
T(s, CX0 + 0.4, CY0 + 8.45, CW - 0.8, 2.2,
  [one("参考｜KUMON（友だち2,525万人）の実機リッチメニュー（2026-09-15取得）", 10.5, True, NAVY, sa=4),
   one("・タブ切替を実装（「今すぐチェック！」／「インフォメーション」）＝本提案の3タブ構成と同じ思想", 9, None, INK, ls=1.25, sa=2),
   one("・左上に「学年別」を置いている＝タブ①を学年起点にする設計の裏付けになる", 9, None, INK, ls=1.25, sa=2),
   one("・ただし診断も体験予約も無く、「詳しくはこちら→」でWebへ送るだけ＝LINE内で完結していない",
       9.5, True, ORANGE, ls=1.25)], anchor="t")
foot(s, "証跡｜_data/lineoa_competitors/KUMON_リッチメニュー_20260915.png。"
        "退会導線を隠す塾より、相談を受け止める塾のほうが継続する（S27の予兆検知と連動）")

# ============================================================
# S20 生徒情報・CRM連携
# ============================================================
s = slides[19]
frame(s, "構築⑤｜生徒情報との連携　配信を「うちの子の話」に変える",
      ["一斉配信のLINEはブロックされる。「うちの子の話」をするLINEはブロックされない。",
       "診断タグ＋在籍データを紐づけて、配信を1対1に近づける。"])
LAYER = [
    ("① 診断タグ（入会前）", ["学年／動機／適性／温度の4軸", "友だち追加直後の4問で取得",
                              "配信の出し分けはここから始まる"], NAVY),
    ("② 行動タグ（入会前）", ["どのリッチメニューを押したか", "体験予約したか／来たか",
                              "配信の未読が続いていないか"], NAVY),
    ("③ 在籍データ（入会後）", ["出席・欠席／振替の記録", "小テスト・定期テストの推移",
                                "宿題の提出状況"], ORANGE),
]
cw3 = (CW - 0.6) / 3
for i, (h_, items, col) in enumerate(LAYER):
    x = CX0 + i * (cw3 + 0.3)
    box(s, x, CY0 + 0.25, cw3, 3.9, fill=WHITE, line=BORDER, lw=1.0)
    hb = box(s, x, CY0 + 0.25, cw3, 0.95, fill=col)
    put_text(hb.text_frame, [one(h_, 11, True, WHITE, align="c")],
             anchor="m", ml=0.08, mr=0.08, mt=0, mb=0)
    T(s, x + 0.3, CY0 + 1.45, cw3 - 0.6, 2.5,
      [one("・" + it, 9.5, None, INK, ls=1.3, sa=5) for it in items], anchor="t")
USE = [
    ("学年タグ × 季節", "小5の保護者だけに「中学準備講座」を配信。小3には出さない"),
    ("適性タグ × 事例", "「質問が苦手」タグの家庭には、個別指導で伸びた事例だけを送る"),
    ("温度タグ × オファー", "「今すぐ」「今学期中」タグだけに体験オファー。「情報収集だけ」には教育コンテンツのみ"),
    ("在籍データ × レポート", "毎月の出席・小テスト推移を保護者へ自動配信（S27の本丸）"),
    ("在籍データ × 予兆検知", "欠席2回連続／宿題未提出／未読3週間で自動で声かけ（S27）"),
]
T(s, CX0, CY0 + 4.6, CW, 0.7, [one("何ができるようになるか", 12, True, NAVY)], anchor="m")
simple_table(s, CX0, CY0 + 5.4, CW, 4.6,
             ["組み合わせ", "配信の具体例"], USE,
             col_w=[7.5, 17.62], align=["l", "l"], hsz=10, bsz=9.5)
box(s, CX0, CY0 + 10.3, CW, 1.6, fill=PALE)
T(s, CX0 + 0.4, CY0 + 10.4, CW - 0.8, 1.4,
  [one("③の在籍データ連携は、貴社の塾システム（Comiru等）との接続方式によって工数が変わります。"
       "初期は手動CSV取込から始め、運用が固まってからAPI連携に移行する進め方を推奨します。",
       9.5, None, INK, ls=1.3)], anchor="m")
foot(s, "API連携・通知メッセージは別途費用。★貴社が使用中の塾管理システムの確認が必要（要件定義で確定）")

# ============================================================
# S21 シナリオ2本の設計表
# ============================================================
s = slides[20]
frame(s, "配信設計｜シナリオは2本だけ。未入会と、在籍中。",
      ["未入会の友だちには14日で体験予約へ。在籍生の保護者には毎月レポートで継続へ。",
       "14日に寄せる理由は2つ：追加15日でアクティブ率が落ちる／後15日で他塾の指名検索が立つ。"])
SC = [
    ("シナリオA｜未入会の友だち", NAVY,
     [("Day0", "あいさつ＋30秒診断", "属性取得・不安の先回り", "全員"),
      ("Day3", "家庭でできる勉強法（塾に入らなくても役立つ）", "信頼の貯金", "全員"),
      ("Day5", "費用のリアル（月謝＋講習費の年間総額）", "費用不安の解消", "全員"),
      ("Day7", "診断タイプ別の成績アップ事例 ＋ 体験オファー①", "CV②へ", "適性タグ別"),
      ("Day10", "よくある不安Q&A（断ったらどうなる／総額は）", "最後の壁を外す", "未予約のみ"),
      ("Day14", "次の定期テストからの逆算 ＋ 体験オファー②", "締切で後押し", "未予約のみ")]),
    ("シナリオB｜在籍生の保護者", ORANGE,
     [("毎月1回", "学習レポート（出席・小テスト推移・講師コメント）", "退会理由1位2位に効く", "在籍生全員"),
      ("学期ごと", "面談のご案内（1タップ予約）", "対面の接点を確保", "在籍生全員"),
      ("随時", "離脱予兆検知トリガー（欠席2連続・宿題未提出・未読3週間）", "退会届の前に会話", "該当者のみ"),
      ("年2回", "ごきょうだい・ご友人紹介のご案内", "口コミ47%を動線化", "継続6ヶ月以上"),
      ("進級期", "次学年のコース案内（2〜3月）", "継続率・単価", "小6・中3以外"),
      ("随時", "欠席・振替のLINE受付（自動応答）", "教室長の工数削減", "在籍生全員")]),
]
hw = (CW - 0.5) / 2
for i, (title_, col, rows) in enumerate(SC):
    x = CX0 + i * (hw + 0.5)
    hb = box(s, x, CY0 + 0.2, hw, 0.95, fill=col)
    put_text(hb.text_frame, [one(title_, 12, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    simple_table(s, x, CY0 + 1.25, hw, 8.0,
                 ["タイミング", "配信内容", "狙い", "対象"],
                 rows, col_w=[2.2, 5.6, 2.5, 2.0],
                 align=["c", "l", "l", "c"], hsz=8.5, bsz=8, header_fill=col)
band(s, CY0 + 9.6, "14日のうち、売り込みは2通だけ（Day7・Day14）。残り12通は家庭で使える情報にする。",
     fill=NAVY, sz=12)
foot(s, "出口は2つ：予約した人は前日リマインドへ分岐、しなかった人はタグ別の定常配信（週1）へ合流。"
        "在籍生は入会と同時にシナリオBへ自動で移す")

# ============================================================
# S22 実文面①｜未入会（14日ステップ）
# ============================================================
s = slides[21]
frame(s, "実文面①｜未入会の友だちに届くLINE（そのまま使えます）",
      ["通知プレビューの冒頭15文字で開封が決まる。各通に必ず設計する。",
       "Day3のように「塾に入らなくても役立つ話」を先に渡すから、Day7とDay14が効く。"])
MSGS = [
    ("Day3", "テスト後、9割のご家庭がしない事",
     "◯◯塾の佐藤です📚\n"
     "塾を探す前に、今日からできることを1つだけ。\n\n"
     "テスト返却後、多くのご家庭は「点数を見て終わり」。\n"
     "でも伸びる子の家庭は、返却当日にこれをやっています👇\n"
     "✅ 間違いを「ケアレスミス」と「わからなかった」に仕分け\n"
     "✅ 「わからなかった」だけ教科書の該当ページに付箋\n"
     "✅ 週末に付箋の問題だけ解き直す（30分でOK）\n\n"
     "これだけで次のテスト+10点は狙えます。\n"
     "塾はその先、「付箋が多すぎて手に負えない」ときの\n選択肢で大丈夫です😊"),
    ("Day7", "数学58点→82点になった話",
     "◯◯塾の佐藤です。\n"
     "去年の今ごろ入塾した中2・Aさんの話をさせてください。\n\n"
     "お母さまの最初のご相談は「本人にやる気がなくて。\n"
     "塾に入れても無駄になりそうで怖い」でした。\n\n"
     "診断は「質問が苦手なマイペース型」。集団ではなく個別で、\n"
     "質問しなくても講師から聞きにいく形にしたところ——\n"
     "📈 数学 58点 → 82点（3ヶ月後の期末）\n\n"
     "やる気は「出させる」ものではなく、\n「できた」の後についてきます。\n\n"
     "🎓 今月の無料体験は残り◯枠。体験後の勧誘連絡は\nしない決まりです。\n"
     "▶ 空き日程を見る"),
    ("Day14", "次のテストまで、あと◯日です",
     "◯◯塾の佐藤です。\n"
     "次の定期テストまで、あと約◯日。\n"
     "塾をご検討中なら、今週がひとつの分かれ目です。\n\n"
     "理由はシンプルで、\n"
     "📅 入塾手続き→座席・講師調整→授業開始まで 約1週間\n"
     "📅 テスト対策には最低3週間\n"
     "かかるからです。\n\n"
     "「まだ迷っている」段階でも大丈夫。体験だけ受けて、\n"
     "テスト後にゆっくり決めたご家庭も多いです😊\n\n"
     "▶ 今週の体験枠（残り◯席）を見る\n"
     "▶ LINEで質問だけしてみる"),
]
cw3 = (CW - 0.7) / 3
for i, (day, prev, body) in enumerate(MSGS):
    x = CX0 + i * (cw3 + 0.35)
    hb = box(s, x, CY0 + 0.2, cw3, 0.85, fill=NAVY)
    put_text(hb.text_frame, [one(day, 12, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    pb = box(s, x, CY0 + 1.15, cw3, 0.8, fill=PORANGE)
    put_text(pb.text_frame,
             [one("通知プレビュー：" + prev, 8.5, True, ORANGE, align="c", ls=1.1)],
             anchor="m", ml=0.12, mr=0.12, mt=0, mb=0)
    bb = box(s, x, CY0 + 2.1, cw3, 8.6, fill=WHITE, line=BORDER, lw=1.0)
    put_text(bb.text_frame,
             [one(l if l else "　", 8.2, None, INK, ls=1.28) for l in body.split("\n")],
             anchor="t", ml=0.3, mr=0.25, mt=0.25, mb=0.15)
foot(s, "Day0（あいさつ＋診断）はS17、Day5・Day10はS21の設計表を参照。"
        "★事例の数値（58点→82点）は貴社の実績に差し替え、学年・期間・通塾回数の条件を併記すること（S09）")

# ============================================================
# S23 実文面②｜在籍生の保護者
# ============================================================
s = slides[22]
frame(s, "実文面②｜在籍生の保護者に届くLINE　退会理由の50.6%はここに効く",
      ["退会を検討した理由の2位3位は、どちらも「コミュニケーション不足」。合計50.6%。",
       "成績が上がらないこと以上に、それについて何も言ってもらえないことで辞めている。"])
M2 = [
    ("月次の学習レポート（毎月1回・自動）", NAVY,
     "【9月の学習レポート】◯◯さん\n\n"
     "📘 出席 8/8回（皆勤です！）\n"
     "✍️ 小テスト平均 72点 → 81点（前月比 +9点）\n"
     "🔍 いま重点的に取り組んでいる単元：一次関数の利用\n\n"
     "💬 教室長より：\n"
     "応用問題で手が止まる場面が減ってきました。\n"
     "10月は文章題の読み取りを重点的に進めます。\n\n"
     "ご不明な点があれば、このトークにそのまま\nご返信ください📩",
     "退会検討理由 1位「成績が上がらない」44.6%／2位「保護者と講師のコミュニケーション不足」27.7% に同時に効く"),
    ("離脱予兆検知トリガー（該当時のみ・自動）", ORANGE,
     "◯◯さんのお母さま、教室長の佐藤です。\n\n"
     "今月お休みが続いていたので、\n少し気になってご連絡しました🍀\n\n"
     "ご家庭やお子さまのご様子で、\n変わったことはありませんか？\n\n"
     "「部活が忙しい」「曜日が合わない」「少し疲れている」\n"
     "——どれもよくあることで、曜日変更やコマ数の調整で\n"
     "解決できる場合が多いです。\n\n"
     "よろしければ15分だけ、お話を聞かせてください。\n\n"
     "▶ 面談の空き時間を見る　▶ このままLINEで相談する",
     "実際に辞めた理由の1位は「子どもが塾に行くことを嫌がった」27%。突然ではなく、必ず欠席という予兆が出る"),
]
hw = (CW - 0.6) / 2
for i, (title_, col, body, why) in enumerate(M2):
    x = CX0 + i * (hw + 0.6)
    hb = box(s, x, CY0 + 0.2, hw, 0.9, fill=col)
    put_text(hb.text_frame, [one(title_, 11.5, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    bb = box(s, x, CY0 + 1.2, hw, 7.0, fill=WHITE, line=BORDER, lw=1.0)
    put_text(bb.text_frame,
             [one(l if l else "　", 9, None, INK, ls=1.3) for l in body.split("\n")],
             anchor="t", ml=0.35, mr=0.3, mt=0.28, mb=0.15)
    wb = box(s, x, CY0 + 8.4, hw, 1.5, fill=PALE)
    put_text(wb.text_frame, [one(why, 9, None, INK, ls=1.3)],
             anchor="m", ml=0.3, mr=0.25, mt=0.1, mb=0.1)
TRIG = [("欠席が2回連続", "曜日・時間が合っていない可能性"),
        ("宿題の未提出が続く", "量が合っていない／つまずいている"),
        ("配信の未読が3週間", "関心が薄れている／保護者が忙しい"),
        ("小テストが2回連続で下降", "単元のつまずきが蓄積している")]
T(s, CX0, CY0 + 10.1, CW, 0.6, [one("予兆検知のトリガー（4つ）", 11, True, NAVY)], anchor="m")
cw4 = (CW - 0.9) / 4
for i, (t_, why) in enumerate(TRIG):
    x = CX0 + i * (cw4 + 0.3)
    bb = box(s, x, CY0 + 10.8, cw4, 1.25, fill=PORANGE)
    put_text(bb.text_frame,
             [one(t_, 9.5, True, ORANGE, align="c", sa=2),
              one(why, 8, None, INK, align="c", ls=1.15)],
             anchor="m", ml=0.15, mr=0.15, mt=0.06, mb=0.06)
foot(s, "出典｜POPER「Comiru」保護者と学習塾の意識調査 2022年（保護者300人）／"
        "DeltaX「塾選」退塾に関する調査 2024年（退塾経験のある保護者65人）")

# ============================================================
# S24 体験予約のリマインド設計
# ============================================================
s = slides[23]
frame(s, "歩留まり①｜体験予約の「来ない」を減らす",
      ["予約の約20%が体験に来ない。ここは仕組みだけで戻せる区間。",
       "前日リマインドで、持ち物・アクセス・変更手段の3つの不安を先につぶす。"])
STEPS = [("予約完了 直後", "予約内容の確認＋カレンダー登録リンク", "その場で不安を残さない"),
         ("前々日", "担当講師の紹介（顔写真・ひとこと）", "誰に会うか分かると欠席が減る"),
         ("前日", "持ち物・アクセス・日程変更の案内", "当日朝の「行くのやめようか」を防ぐ"),
         ("当日 2時間前", "「お待ちしています」の短い一言", "最後のひと押し"),
         ("翌日", "お礼＋診断結果の再送＋相談導線", "断る自由を残したまま次へ")]
for i, (when, what, why) in enumerate(STEPS):
    yy = CY0 + 0.25 + i * 1.5
    box(s, CX0, yy, 14.6, 1.3, fill=WHITE, line=BORDER, lw=1.0)
    wb = box(s, CX0, yy, 3.4, 1.3, fill=NAVY)
    put_text(wb.text_frame, [one(when, 10, True, WHITE, align="c", ls=1.1)],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    T(s, CX0 + 3.7, yy, 5.9, 1.3, [one(what, 9.5, True, INK, ls=1.2)], anchor="m")
    T(s, CX0 + 9.8, yy, 4.6, 1.3, [one(why, 8.5, None, MUT, ls=1.2)], anchor="m")
VX = CX0 + 15.0
T(s, VX, CY0 + 0.25, 10.1, 0.6, [one("前日リマインドの実文", 11.5, True, NAVY)], anchor="m")
box(s, VX, CY0 + 0.95, 10.12, 5.6, fill=WHITE, line=BORDER, lw=1.0)
T(s, VX + 0.35, CY0 + 1.2, 9.4, 5.2,
  [one("【明日◯時〜】体験授業のご確認です😊", 10, True, INK, ls=1.3, sa=6),
   one("◯◯塾 △△教室（地図はこちら▼）", 9, None, INK, ls=1.3, sa=6),
   one("✏️ 持ち物：筆記用具だけでOK（教材はご用意します）", 9, None, INK, ls=1.3, sa=3),
   one("🚲 駐輪場：教室裏にあります", 9, None, INK, ls=1.3, sa=3),
   one("👪 保護者さまは「見学」でも「面談だけ」でも大丈夫です", 9, None, INK, ls=1.3, sa=6),
   one("ご都合が変わったら、このトークに「変更」と送って\nください。1タップで別日程をご案内します\n"
       "（キャンセル料はありません）", 9, None, INK, ls=1.3)], anchor="t")
ef = box(s, VX, CY0 + 6.9, 10.12, 2.2, fill=PALE)
put_text(ef.text_frame,
         [one("期待効果", 10, True, NAVY, sa=4),
          one("予約→体験の実施率　80% → 92%", 14, True, ORANGE, sa=4),
          one("弊社支援実績レンジ（保守置き）。★社内数値で確定させる", 8, None, MUT)],
         anchor="m", ml=0.3, mr=0.25, mt=0.12, mb=0.12)
box(s, VX, CY0 + 9.4, 10.12, 2.4, fill=WHITE, line=NAVY, lw=1.25)
T(s, VX + 0.3, CY0 + 9.55, 9.5, 2.1,
  [one("「キャンセル料はありません」を必ず書く", 10, True, NAVY, sa=4),
   one("保護者が体験をためらう理由に、断りにくさがある（※比率データは無く一般論）。"
       "変更・キャンセルの自由を明示したほうが、結果として実施率は上がる。",
       9, None, INK, ls=1.3)], anchor="t")
foot(s, "無断キャンセルは「行く気がない」のではなく「思い出せない・持ち物が不安・変更方法が分からない」で起きる")

# ============================================================
# S25 年間の企画投稿カレンダー
# ============================================================
s = slides[24]
frame(s, "配信設計｜年間の企画カレンダー　ピークではなく「26日前」に動く",
      ["検索の山（Googleトレンド実測）から逆算して、仕込みを始める日を決める。",
       "企画投稿は月1〜2本。その月に保護者が何を考えているかに合わせる。"])
CAL = [
    ("1月", "新学年の準備が動き出す", "「塾」ピーク2/8の26日前＝1/13から仕込み",
     "小5・中2の保護者へ「次学年で変わること」", ORANGE),
    ("2-3月", "入会の最大期／春期講習", "受験3種のピークも1月中旬〜2月上旬",
     "新学年スタート応援。春期講習は夏の1/3規模", ORANGE),
    ("4-5月", "最初の定期テスト", "中間テストで成績不安が顕在化",
     "「テスト後、9割の家庭がしない事」（S22 Day3）", PALE),
    ("6月", "夏期講習の仕込み開始", "夏期講習ピーク7/12の26日前＝6/16",
     "年間最大の商戦。ここを外すと1年が変わる", ORANGE),
    ("7-8月", "夏期講習ピーク", "冬期・春期の3〜5倍の規模",
     "夏の学習計画・自由研究・受験生の夏", ORANGE),
    ("9-10月", "下半期スタート／退会注意期", "部活引退後の8〜10月は退会ピークの1つ",
     "予兆検知を強める月。面談オファーを厚く", PALE),
    ("11-12月", "冬期講習・受験直前", "冬期講習は12月下旬がピーク",
     "受験生は直前期、非受験生は「来年こそ」", PALE),
]
for i, (m, theme, timing, plan, fillc) in enumerate(CAL):
    yy = CY0 + 0.25 + i * 1.62
    box(s, CX0, yy, CW, 1.45, fill=WHITE, line=BORDER, lw=1.0)
    mb_ = box(s, CX0, yy, 2.6, 1.45, fill=NAVY if fillc == PALE else ORANGE)
    put_text(mb_.text_frame, [one(m, 13, True, WHITE, align="c")],
             anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)
    T(s, CX0 + 2.9, yy, 5.4, 1.45, [one(theme, 10, True, INK, ls=1.2)], anchor="m")
    T(s, CX0 + 8.5, yy, 8.2, 1.45, [one(timing, 9, None, MUT, ls=1.25)], anchor="m")
    T(s, CX0 + 17.0, yy, 8.0, 1.45, [one(plan, 9, None, INK, ls=1.25)], anchor="m")
foot(s, "出典｜Googleトレンド（日本・2026年年初来・週次／_data/trends/教育塾/）＋前後検索の26日（S11）。"
        "退会ピーク（3月・8〜10月）は業界推計値。★貴社の入会・退会の月別実績が入れば、この表を実データで置き換える")

# ============================================================
# S26 通知メッセージ
# ============================================================
s = slides[25]
frame(s, "配信設計｜通知メッセージ　電話番号だけで届く、塾と相性のいい機能",
      ["友だち未追加でも、電話番号があればLINEにメッセージを届けられる（別途費用）。",
       "塾は在籍生の電話番号を持っている。既存リストをLINEへ移す最短の導線になる。"])
USE = [
    ("既存の在籍生リストをLINEへ移行", "紙・電話・メールで連絡している保護者をLINEに集約。"
     "以後の連絡コストが下がる", "在籍生の保護者", NAVY),
    ("休会・退会した家庭への再接触", "「その後いかがですか」の一言。塾は再入会が起きやすい業態",
     "過去の在籍者", ORANGE),
    ("資料請求だけで止まっている層", "体験に進まなかった人へ、季節講習のタイミングで再案内",
     "過去の問い合わせ者", ORANGE),
    ("兄弟姉妹への案内", "上の子が卒塾した家庭に、下の子の学年到達時に案内",
     "卒塾家庭", NAVY),
]
for i, (h_, b_, who, col) in enumerate(USE):
    yy = CY0 + 0.3 + i * 2.1
    box(s, CX0, yy, 16.4, 1.85, fill=WHITE, line=BORDER, lw=1.0)
    hb = box(s, CX0, yy, 5.6, 1.85, fill=col)
    put_text(hb.text_frame, [one(h_, 10.5, True, WHITE, align="c", ls=1.2)],
             anchor="m", ml=0.15, mr=0.15, mt=0, mb=0)
    T(s, CX0 + 5.9, yy + 0.15, 7.4, 1.55, [one(b_, 9.5, None, INK, ls=1.25)], anchor="m")
    T(s, CX0 + 13.5, yy + 0.15, 2.7, 1.55, [one(who, 9, True, MUT, align="c", ls=1.2)], anchor="m")
VX = CX0 + 16.8
box(s, VX, CY0 + 0.3, 8.32, 4.0, fill=PALE)
T(s, VX + 0.35, CY0 + 0.5, 7.6, 3.6,
  [one("塾にとっての価値", 11, True, NAVY, sa=5),
   one("塾は在籍生・卒塾生の電話番号を必ず持っている。"
       "この資産をLINEに変えられるのが通知メッセージ。", 9.5, None, INK, ls=1.3, sa=5),
   one("特に「卒塾した家庭の下の子」は、獲得コストゼロで"
       "入会率が最も高い層（紹介経由70〜80%と同じ構造）。",
       9.5, None, INK, ls=1.3)], anchor="t")
box(s, VX, CY0 + 4.6, 8.32, 4.0, fill=WHITE, line=RED, lw=1.25)
T(s, VX + 0.35, CY0 + 4.8, 7.6, 3.6,
  [one("⚠ 注意点", 11, True, RED, sa=5),
   one("・別途費用（アカウント種別・配信数による）", 9.5, None, INK, ls=1.3, sa=3),
   one("・「知らないのに届いた」と受け取られない文面設計が要る", 9.5, None, INK, ls=1.3, sa=3),
   one("・在籍生・過去の問い合わせ者など、関係性が明確な相手に限定する",
       9.5, None, INK, ls=1.3)], anchor="t")
box(s, VX, CY0 + 8.9, 8.32, 2.5, fill=PORANGE)
T(s, VX + 0.35, CY0 + 9.05, 7.6, 2.2,
  [one("導入の優先度", 10.5, True, ORANGE, sa=4),
   one("初期は不要。まず在籍生をLINEに集めて運用を固め、"
       "リストが溜まってから検討するのが安全。", 9.5, None, INK, ls=1.3)], anchor="t")
foot(s, "通知メッセージは別途費用。利用可否・単価はLINEヤフーの規定により変動するため、要件定義時に最新条件を確認する")

# ============================================================
# S27 退会防止（本資料の山）
# ============================================================
s = slides[26]
frame(s, "歩留まり②｜退会防止　退会は「突然」ではない。必ず予兆がある。",
      ["退会検討理由の1位は「成績が上がらない」44.6%。2位3位のコミュニケーション不足が合計50.6%。",
       "成績が上がらないこと以上に、それについて説明がないことで辞めている。"])
BARS = [("成績が上がらない", 44.6, ORANGE),
        ("保護者と講師のコミュニケーション不足", 27.7, RED),
        ("子どもと講師のコミュニケーション不足", 22.9, RED)]
T(s, CX0, CY0 + 0.25, 15.5, 0.65,
  [one("退会を検討した理由（POPER 2022年・保護者300人）", 11, True, NAVY)], anchor="m")
BW = 8.4
for i, (label, v, col) in enumerate(BARS):
    yy = CY0 + 1.1 + i * 1.35
    T(s, CX0, yy, 6.6, 1.1, [one(label, 9.5, None, INK, ls=1.2)], anchor="m")
    box(s, CX0 + 6.8, yy + 0.2, BW, 0.7, fill=GREY, shape=MSO_SHAPE.RECTANGLE)
    box(s, CX0 + 6.8, yy + 0.2, BW * v / 50.0, 0.7, fill=col, shape=MSO_SHAPE.RECTANGLE)
    T(s, CX0 + 6.8 + BW * v / 50.0 + 0.15, yy + 0.2, 2.0, 0.7,
      [one(f"{v}%", 11, True, col)], anchor="m")
bb = box(s, CX0 + 6.8, CY0 + 5.3, BW + 1.0, 1.15, fill=PRED, line=RED, lw=1.25)
put_text(bb.text_frame,
         [one("2位＋3位 ＝ 50.6%　コミュニケーション起因", 10.5, True, RED, align="c")],
         anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
T(s, CX0, CY0 + 6.8, 16.5, 0.65,
  [one("実際に辞めた理由（DeltaX「塾選」2024年・退塾経験のある保護者65人）", 11, True, NAVY)], anchor="m")
REAL = [("子どもが塾に行くことを嫌がった", "27%"), ("部活・習い事と両立できなかった", "14%"),
        ("進路を変更した", "10%"), ("テストの点数が上がらなかった", "10%")]
cw4 = (16.5 - 0.9) / 4
for i, (t_, v) in enumerate(REAL):
    x = CX0 + i * (cw4 + 0.3)
    bb = box(s, x, CY0 + 7.5, cw4, 1.5, fill=PALE)
    put_text(bb.text_frame,
             [one(v, 15, True, ORANGE, align="c", sa=2),
              one(t_, 8.5, None, INK, align="c", ls=1.15)],
             anchor="m", ml=0.12, mr=0.12, mt=0.06, mb=0.06)
VX = CX0 + 17.0
box(s, VX, CY0 + 0.25, 8.12, 5.9, fill=WHITE, line=NAVY, lw=1.5)
T(s, VX + 0.3, CY0 + 0.45, 7.5, 5.5,
  [one("早期退塾には予測因子がある", 11, True, NAVY, sa=5),
   one("1年未満で退塾した家庭", 9.5, True, INK, sa=3),
   one("・入塾理由が「友達が通っていた」26%（1年以上継続は3%）", 9, None, INK, ls=1.3, sa=3),
   one("・「入塾前に十分調べなかった」43%（1年以上継続は14%）", 9, None, INK, ls=1.3, sa=6),
   one("→ 入口の適合確認（S18の診断）が、そのまま継続率の施策になる",
       9.5, True, ORANGE, ls=1.3)], anchor="t")
ACT = [("① 成果の可視化", "月1で学習レポートを自動配信。退会理由1位と2位に同時に効く"),
       ("② 予兆検知", "欠席2連続・宿題未提出・未読3週間で自動声かけ→面談へ"),
       ("③ 入口での適合確認", "診断で「合う/合わない」を先に見せる（＝43%の予防）")]
for i, (h_, b_) in enumerate(ACT):
    card(s, VX, CY0 + 6.5 + i * 1.75, 8.12, 1.6, h_, [b_], hsz=10.5, bsz=8.5, fill=PALE)
band(s, CY0 + 11.9, "退会を年5名防ぐ＝新規を5名多く獲得するのと同じ。しかも獲得コストはゼロ。", fill=NAVY, sz=12.5)
foot(s, "出典｜POPER「Comiru」保護者と学習塾の意識調査 2022年（n=300）／DeltaX「塾選」退塾に関する調査 2024年（n=65）。"
        "年間中途退会率10〜20%は業界推計値（公的統計なし）")

# ============================================================
# S28 工数削減 ＋ 改善モデル
# ============================================================
s = slides[27]
frame(s, "現場の工数｜教室長を「夕方の電話番」から解放する",
      ["塾の現場は夕方に全部が重なる。欠席連絡・授業準備・面談が同じ時間帯に集中する。",
       "電話をLINEに逃がすだけで、浮いた時間を予兆の出た生徒との面談に回せる。"])
BEFORE = [("16:00-17:00", "欠席連絡の電話が集中", "授業準備が止まる"),
          ("17:00-19:00", "授業開始。電話に出られない", "折り返しが溜まる"),
          ("19:00-21:00", "授業と面談が重なる", "保護者対応が後回しに"),
          ("21:00-", "折り返し電話・翌日準備", "残業が常態化")]
AFTER = [("終日", "欠席・振替はLINEで自動受付", "電話が鳴らない"),
         ("終日", "料金・持ち物・振替はKW自動応答", "24時間365日対応"),
         ("授業時間", "授業に集中できる", "折り返しが発生しない"),
         ("浮いた時間", "予兆の出た生徒の保護者と面談", "退会防止に直結")]
hw = (CW - 0.8) / 2
for i, (title_, rows, col) in enumerate([("Before｜いまの夕方", BEFORE, MUT),
                                          ("After｜LINE導入後", AFTER, NAVY)]):
    x = CX0 + i * (hw + 0.8)
    hb = box(s, x, CY0 + 0.25, hw, 0.95, fill=col)
    put_text(hb.text_frame, [one(title_, 12, True, WHITE, align="c")],
             anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    for j, (t_, what, res) in enumerate(rows):
        yy = CY0 + 1.4 + j * 1.55
        box(s, x, yy, hw, 1.4, fill=WHITE if i else GREY, line=BORDER, lw=1.0)
        tb = box(s, x + 0.2, yy + 0.2, 2.9, 1.0, fill=col)
        put_text(tb.text_frame, [one(t_, 8.5, True, WHITE, align="c", ls=1.1)],
                 anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)
        T(s, x + 3.3, yy + 0.12, 5.6, 0.62, [one(what, 9.5, True, INK, ls=1.15)], anchor="m")
        T(s, x + 3.3, yy + 0.7, 5.6, 0.58, [one(res, 8.5, None, MUT, ls=1.15)], anchor="m")
if True:
    x = CX0 + hw + 0.15
    box(s, x, CY0 + 3.4, 0.5, 1.0, fill=ORANGE, shape=MSO_SHAPE.RIGHT_ARROW)
box(s, CX0, CY0 + 8.0, CW, 2.2, fill=PALE)
T(s, CX0 + 0.4, CY0 + 8.15, CW - 0.8, 1.9,
  [one("工数削減を「時給換算」で語らない", 11, True, NAVY, sa=4),
   one("削減額を人件費で計算すると額が小さく見え、論点もズレる。決裁者に効くのは"
       "「ピーク時に手が止まる」という現場の実感と、"
       "浮いた時間を退会防止の面談に回せるという接続。",
       9.5, None, INK, ls=1.3)], anchor="t")
band(s, CY0 + 10.5, "電話を減らすこと自体が目的ではない。空いた時間を、辞めそうな生徒との面談に使えることが価値。",
     fill=NAVY, sz=12)
foot(s, "欠席・振替のLINE受付、キーワード自動応答は費用内（無償付帯）。日程調整の自動化は要件により別途")

# ============================================================
# S29 効果測定の設計
# ============================================================
s = slides[28]
frame(s, "成果の見方｜友だち数をKPIにしない",
      ["371万人集めても配信を止めればゼロ（S12）。友だち数は成果ではない。",
       "見るのは「体験予約」「入会率」「継続率」の3つと、その先の粗利だけ。"])
KPI = [
    ("主KPI", "体験予約数／体験予約CPA", "LINE経由と広告経由を分けて見る"),
    ("主KPI", "体験→入会率", "診断を通った人と通らない人で比較する"),
    ("主KPI", "年間の中途退会率", "予兆検知の声かけ→面談→継続の歩留まりも見る"),
    ("副KPI", "予約→体験の実施率", "前日リマインドの効果はここに出る"),
    ("副KPI", "診断の完了率", "低ければ設問が重い。4問を超えていないか"),
    ("副KPI", "配信の開封率・ブロック率", "ブロック率が上がったら配信頻度か内容の問題"),
    ("見ない", "友だち数（単体）", "増やすだけなら広告で買える。成果と相関しない"),
]
simple_table(s, CX0, CY0 + 0.3, CW, 6.4,
             ["区分", "指標", "見方"], KPI,
             col_w=[3.2, 8.2, 13.72], align=["c", "l", "l"], hsz=10, bsz=9.5)
FRAME = [("月次で見る", "体験予約数・予約CPA・実施率・診断完了率・ブロック率"),
         ("四半期で見る", "体験→入会率・退会率・紹介経由の入会数"),
         ("年次で見る", "生徒1名あたりの生涯粗利（LTV×粗利率）と、投下費用の回収")]
T(s, CX0, CY0 + 7.1, CW, 0.65, [one("見る周期を分ける（短期の数字で施策を殺さない）", 12, True, NAVY)], anchor="m")
cw3 = (CW - 0.6) / 3
for i, (h_, b_) in enumerate(FRAME):
    x = CX0 + i * (cw3 + 0.3)
    card(s, x, CY0 + 7.9, cw3, 2.0, h_, [b_], hsz=11, bsz=9, fill=PALE)
box(s, CX0, CY0 + 10.2, CW, 1.7, fill=WHITE, line=NAVY, lw=1.25)
T(s, CX0 + 0.4, CY0 + 10.3, CW - 0.8, 1.5,
  [one("継続率は四半期で見る。塾の退会は年度末（3月）と部活引退後（8〜10月）に集中するため、"
       "月次の増減だけを見ると施策の効果を読み違える。",
       9.5, None, INK, ls=1.3)], anchor="m")
foot(s, "レポートは月次の定例会で共有（費用内）。GA連携・LINE分析の数値は弊社側で取得して提出する")

# ============================================================
# S30 費用プラン
# ============================================================
s = slides[29]
frame(s, "費用プラン｜6ヶ月〜・税抜",
      ["アカウント構築から配信の実務まで、下記はすべて費用内（無償付帯）。",
       "貴社の運用体制に合わせて、①〜④のいずれかを選んでいただきます。"])
PLANS = [
    ("① コンサル基本", "初期 10万円〜", "月 20万円（3投稿）\n30万円（5投稿）\n50万円（9投稿）",
     "企画から配信まで全部任せたい", ORANGE),
    ("② 初動設計＋運用", "初期 10万円〜", "月 5万円〜",
     "構築は任せて、配信は自社で回したい", NAVY),
    ("③ 運用代行・効率改善", "初期 10万円〜", "月 0万円〜\n（アカウント費のみ）",
     "既にアカウントがあり、改善だけ頼みたい", NAVY),
    ("④ 成果報酬型", "初期 5万円〜", "単価 × 成果数\n＋ 固定費",
     "成果に連動させたい", NAVY),
]
cw4 = (CW - 0.9) / 4
for i, (name, init, monthly, who, col) in enumerate(PLANS):
    x = CX0 + i * (cw4 + 0.3)
    box(s, x, CY0 + 0.25, cw4, 5.4, fill=WHITE, line=col, lw=1.5 if i == 1 else 1.0)
    hb = box(s, x, CY0 + 0.25, cw4, 0.95, fill=col)
    put_text(hb.text_frame, [one(name, 11, True, WHITE, align="c")],
             anchor="m", ml=0.08, mr=0.08, mt=0, mb=0)
    T(s, x + 0.2, CY0 + 1.4, cw4 - 0.4, 0.6, [one(init, 10, None, MUT, align="c")], anchor="m")
    T(s, x + 0.2, CY0 + 2.1, cw4 - 0.4, 1.9,
      [one(l, 13, True, col, align="c", ls=1.2) for l in monthly.split("\n")], anchor="m")
    T(s, x + 0.2, CY0 + 4.2, cw4 - 0.4, 1.2, [one(who, 9, None, INK, align="c", ls=1.2)], anchor="m")
    if i == 1:
        badge(s, x + cw4 - 2.4, CY0 + 0.05, 2.3, 0.5, "推奨", sz=9)
FREE = ["アカウント開設", "プロフィール設計", "リッチメニュー制作", "あいさつメッセージ",
        "初期アンケート（診断）", "KW自動応答", "ステップ配信", "セグメント配信",
        "タグ管理", "GAレポート連携", "クリエイティブ制作", "月次定例会"]
T(s, CX0, CY0 + 6.0, CW, 0.65, [one("無償付帯（費用内で実施）", 12, True, NAVY)], anchor="m")
cw6 = (CW - 0.5 * 5) / 6
for i, f in enumerate(FREE):
    r, c = divmod(i, 6)
    bb = box(s, CX0 + c * (cw6 + 0.5), CY0 + 6.75 + r * 1.0, cw6, 0.85, fill=PALE)
    put_text(bb.text_frame, [one("✓ " + f, 8.5, None, NAVY, align="c", ls=1.1)],
             anchor="m", ml=0.06, mr=0.06, mt=0, mb=0)
box(s, CX0, CY0 + 9.1, CW, 2.5, fill=WHITE, line=RED, lw=1.25)
T(s, CX0 + 0.4, CY0 + 9.25, CW - 0.8, 2.2,
  [one("別途費用（誤解が起きやすいので明記します）", 11, True, RED, sa=4),
   one("・LP離脱防止（Sitelead）：初期 10万円 ＋ 月 5万円", 9.5, None, INK, ls=1.3, sa=2),
   one("・通知メッセージ／API連携（塾管理システムとの接続）：要件により個別見積", 9.5, None, INK, ls=1.3, sa=2),
   one("・配信ツール（Lstep または Hachidori のどちらか一方を要件定義で選定）", 9.5, None, INK, ls=1.3)],
  anchor="t")
foot(s, "契約は6ヶ月〜。塾の商戦（2〜3月／7〜8月）を1サイクル含む期間で効果を見ていただくため。"
        "4ヶ月目以降の月額は上記のとおり継続します")

# ============================================================
# S31 運用スケジュール
# ============================================================
s = slides[30]
frame(s, "導入スケジュール｜次の商戦から逆算して着手する",
      ["構築に約1.5ヶ月。そこから14日ステップが一周して、ようやく数字が動き出す。",
       "「塾」ピーク2/8なら12月着手、夏期講習7/12なら5月着手が目安。"])
PH = [
    ("Phase 1", "要件定義・設計", "2週間",
     ["現状ヒアリング（在籍数・月謝・退会率・広告実績）", "CV3段とタグ設計の確定",
      "塾管理システムとの連携方式の決定", "配信ツールの選定（Lstep / Hachidori）"], NAVY),
    ("Phase 2", "構築・制作", "3〜4週間",
     ["アカウント開設・プロフィール", "30秒診断（4問）の設計と実装",
      "リッチメニュー3タブ×6枠の制作", "14日ステップ（14通）の原稿・実装", "LP離脱防止の設置"], NAVY),
    ("Phase 3", "テスト配信・調整", "1〜2週間",
     ["社内テスト配信（文面・分岐の確認）", "法務確認（S09の表現チェック）",
      "既存の在籍生への案内・友だち化"], ORANGE),
    ("Phase 4", "本格運用・改善", "継続",
     ["企画配信（月1〜2本）", "セグメント配信・予兆検知の稼働",
      "月次レポート＋定例会", "四半期で入会率・退会率を検証"], ORANGE),
]
cw4 = (CW - 0.9) / 4
for i, (ph, name, dur, items, col) in enumerate(PH):
    x = CX0 + i * (cw4 + 0.3)
    cv = box(s, x, CY0 + 0.3, cw4, 1.5, fill=col, shape=MSO_SHAPE.CHEVRON)
    put_text(cv.text_frame,
             [one(ph, 10, True, WHITE, align="c", sa=2),
              one(name, 11.5, True, WHITE, align="c", ls=1.15)],
             anchor="m", ml=0.5, mr=0.2, mt=0, mb=0)
    T(s, x, CY0 + 1.95, cw4, 0.6, [one(dur, 10, True, ORANGE, align="c")], anchor="m")
    box(s, x, CY0 + 2.65, cw4, 5.2, fill=WHITE, line=BORDER, lw=1.0)
    T(s, x + 0.3, CY0 + 2.85, cw4 - 0.6, 4.8,
      [one("・" + it, 8.5, None, INK, ls=1.3, sa=5) for it in items], anchor="t")
T(s, CX0, CY0 + 8.3, CW, 0.65, [one("逆算のしかた（S10の検索ピークから）", 12, True, NAVY)], anchor="m")
BACK = [("2〜3月の入会期を取りにいく", "12月着手 → 1月中旬に配信開始（「塾」ピーク2/8の26日前）"),
        ("7〜8月の夏期講習を取りにいく", "5月着手 → 6月中旬に配信開始（夏期講習ピーク7/12の26日前）")]
for i, (h_, b_) in enumerate(BACK):
    bb = box(s, CX0, CY0 + 9.1 + i * 1.35, CW, 1.2, fill=PORANGE)
    put_text(bb.text_frame,
             [multi([(h_ + "　→　", 11, True, ORANGE), (b_, 10, None, INK)], ls=1.2)],
             anchor="m", ml=0.4, mr=0.3, mt=0, mb=0)
foot(s, "在籍生の友だち化（Phase 3）は初月から効く。新規獲得の数字が動くのはPhase 4に入ってから")

# ============================================================
# S32 サポート体制
# ============================================================
s = slides[31]
frame(s, "体制｜誰が何をやるか",
      ["配信の実務（原稿・制作・設定・分析）は弊社が持ちます。",
       "貴社にお願いするのは、教室の情報と、月1回の定例会だけです。"])
ROLE = [
    ("DYM｜運用統括", ["全体設計・KPI管理", "月次レポートと改善提案", "定例会の進行"], NAVY),
    ("DYM｜配信担当", ["配信原稿の作成", "セグメント・ステップの設定", "予兆検知トリガーの管理"], NAVY),
    ("DYM｜クリエイティブ", ["リッチメニュー制作", "配信画像・バナー", "LP離脱防止の設置"], NAVY),
    ("貴社｜教室", ["教室の情報提供（実績・講師）", "体験・面談の実施", "月1回の定例会にご参加"], ORANGE),
]
cw4 = (CW - 0.9) / 4
for i, (name, items, col) in enumerate(ROLE):
    x = CX0 + i * (cw4 + 0.3)
    box(s, x, CY0 + 0.3, cw4, 4.6, fill=WHITE, line=col, lw=1.25)
    hb = box(s, x, CY0 + 0.3, cw4, 1.0, fill=col)
    put_text(hb.text_frame, [one(name, 11, True, WHITE, align="c")],
             anchor="m", ml=0.08, mr=0.08, mt=0, mb=0)
    T(s, x + 0.3, CY0 + 1.55, cw4 - 0.6, 3.1,
      [one("・" + it, 9.5, None, INK, ls=1.3, sa=6) for it in items], anchor="t")
FLOW2 = [("毎月", "配信企画の提案（弊社）→ 貴社confirm → 配信 → レポート"),
         ("毎月", "定例会（30〜60分）：先月の数字の振り返りと翌月の企画"),
         ("四半期", "入会率・退会率の検証と、施策の組み替え"),
         ("随時", "LINEトークで質問・相談（専用グループ）")]
T(s, CX0, CY0 + 5.3, CW, 0.65, [one("進め方のリズム", 12, True, NAVY)], anchor="m")
for i, (freq, what) in enumerate(FLOW2):
    yy = CY0 + 6.1 + i * 1.35
    box(s, CX0, yy, CW, 1.2, fill=PALE if i % 2 == 0 else WHITE, line=BORDER, lw=1.0)
    fb = box(s, CX0 + 0.2, yy + 0.2, 2.6, 0.8, fill=NAVY)
    put_text(fb.text_frame, [one(freq, 10, True, WHITE, align="c")],
             anchor="m", ml=0.05, mr=0.05, mt=0, mb=0)
    T(s, CX0 + 3.2, yy, 21.6, 1.2, [one(what, 10, None, INK, ls=1.2)], anchor="m")
band(s, CY0 + 11.6, "貴社の負担は「教室の情報」と「月1回の定例会」だけ。配信の実務は弊社が持ちます。",
     fill=NAVY, sz=12.5)
foot(s, "体制の人数・担当者名は契約時に確定します。定例会はオンライン可")

# ============================================================
# S33 飛び道具
# ============================================================
s = slides[32]
frame(s, "飛び道具｜他塾がやっていない企画（ご参考）",
      ["競合8社のLINEはすべて「予約の受け皿」で止まっている（S12）。",
       "以下は地域で最初の1社になれる企画案です。優先度はDYM見解。"])
IDEA = [
    ("◎", "学年別・逆算カレンダー配信", "小3〜中2の保護者へ、学年ごとに「この時期にやるべきこと」を毎月配信。"
     "入塾学年の山（中受小4／他小5-6）の1学年前から接点を持つ", ORANGE),
    ("◎", "テスト返却週のピンポイント配信", "定期テストの返却時期に合わせて「点数の見方」を配信。"
     "中学生の入塾きっかけ1位は「成績が不安」52.4%", ORANGE),
    ("○", "きょうだい・卒塾生への通知メッセージ", "卒塾家庭の下の子が該当学年に達したタイミングで案内。"
     "紹介経由の入会率は70〜80%超", NAVY),
    ("○", "保護者向け「進路のきほん」ミニ講座", "LINEで完結する3分×5回の配信講座。"
     "決定関与は保護者80.5%。保護者自身の知識不足を埋める", NAVY),
    ("△", "AIによる学習相談ボット", "24時間の質問対応。ただし塾は「先生」で選ばれる業界なので、"
     "有人対応の価値を下げないよう設計が要る", MUT),
]
for i, (mark, name, body, col) in enumerate(IDEA):
    yy = CY0 + 0.3 + i * 2.15
    box(s, CX0, yy, CW, 1.9, fill=WHITE, line=BORDER, lw=1.0)
    mb_ = box(s, CX0 + 0.2, yy + 0.35, 1.2, 1.2, fill=col)
    put_text(mb_.text_frame, [one(mark, 15, True, WHITE, align="c")],
             anchor="m", ml=0, mr=0, mt=0, mb=0)
    T(s, CX0 + 1.7, yy + 0.2, 7.4, 1.5, [one(name, 11, True, INK, ls=1.2)], anchor="m")
    T(s, CX0 + 9.3, yy + 0.2, 15.6, 1.5, [one(body, 9.5, None, INK, ls=1.3)], anchor="m")
box(s, CX0, CY0 + 11.2, CW, 1.0, fill=PALE)
T(s, CX0 + 0.4, CY0 + 11.25, CW - 0.8, 0.9,
  [one("◎＝根拠データあり・すぐ実行できる　／　○＝条件付きで有効　／　△＝この業界では優先度低（DYM見解）",
       9.5, None, INK, align="c")], anchor="m")
foot(s, "優先度の判定はDYM見解。根拠データはS03〜S12および巻末の出典一覧を参照")

# ============================================================
# S34 第2の提案軸｜紹介・口コミの動線化
# ============================================================
s = slides[33]
frame(s, "第2の提案軸｜塾は広告より口コミで選ばれる。そこを動線にする。",
      ["塾を知ったきっかけ1位は「知人・友人の口コミ」47.0%。塾の公式サイトより強い。",
       "しかも紹介経由の入会率は70〜80%超。Web広告経由の40〜50%と30pt以上違う。"])
SRC2 = [("知人・友人の口コミ", 47.0, ORANGE), ("インターネット上の口コミ", 28.3, NAVY)]
T(s, CX0, CY0 + 0.25, 13.5, 0.65,
  [one("塾を知ったきっかけ（POPER「Comiru」2022年・保護者300人）", 11, True, NAVY)], anchor="m")
for i, (label, v, col) in enumerate(SRC2):
    yy = CY0 + 1.05 + i * 1.4
    T(s, CX0, yy, 5.6, 1.1, [one(label, 9.5, None, INK, ls=1.2)], anchor="m")
    box(s, CX0 + 5.8, yy + 0.2, 6.4, 0.7, fill=GREY, shape=MSO_SHAPE.RECTANGLE)
    box(s, CX0 + 5.8, yy + 0.2, 6.4 * v / 50.0, 0.7, fill=col, shape=MSO_SHAPE.RECTANGLE)
    T(s, CX0 + 5.8 + 6.4 * v / 50.0 + 0.15, yy + 0.2, 2.0, 0.7,
      [one(f"{v}%", 11, True, col)], anchor="m")
cb2 = box(s, CX0, CY0 + 4.0, 13.5, 1.9, fill=PORANGE, line=ORANGE, lw=1.25)
put_text(cb2.text_frame,
         [one("入会率の差（業界推計値）", 9.5, True, ORANGE, align="c", sa=3),
          multi([("紹介・口コミ経由 70〜80%超", 13, True, ORANGE),
                 ("　vs　", 10, None, MUT),
                 ("Web広告経由 40〜50%", 13, True, MUT)], align="c")],
         anchor="m", ml=0.2, mr=0.2, mt=0.1, mb=0.1)
VX = CX0 + 14.0
T(s, VX, CY0 + 0.25, 11.1, 0.65, [one("紹介を「頼む」のではなく「頼みやすくする」", 11.5, True, NAVY)], anchor="m")
REF = [
    ("① 紹介したくなる材料を毎月渡す", "月次の学習レポートは、そのまま「うちの子が伸びた証拠」になる。"
     "ママ友に見せられる形にしておく"),
    ("② リッチメニューに常設する", "タブ③に「ごきょうだい・ご友人紹介」を常時置く。"
     "思い立った瞬間に押せる場所が要る"),
    ("③ タイミングを狙って声をかける", "成績が上がった月・皆勤だった月に自動で紹介案内。"
     "満足度が高い瞬間に頼むと通る"),
    ("④ 卒塾家庭の下の子を拾う", "通知メッセージで、該当学年に達したタイミングで案内（S26）"),
]
for i, (h_, b_) in enumerate(REF):
    card(s, VX, CY0 + 1.05 + i * 2.35, 11.12, 2.15, h_, [b_], hsz=10.5, bsz=9, fill=PALE)
box(s, CX0, CY0 + 6.3, 13.5, 3.5, fill=WHITE, line=NAVY, lw=1.25)
T(s, CX0 + 0.35, CY0 + 6.5, 12.8, 3.1,
  [one("ただし、口コミだけで選ばれると続かない", 11, True, RED, sa=5),
   one("1年未満で退塾した家庭の26%が、入塾理由を「仲の良い友達が通っていたから」と回答"
       "（1年以上継続した家庭は3%）。", 9.5, None, INK, ls=1.3, sa=5),
   one("→ 紹介で来た人にも、必ず診断を通す。「友達が行っているから」を"
       "「うちの子に合っているから」に変えてから入会してもらう。",
       9.5, True, NAVY, ls=1.3)], anchor="t")
band(s, CY0 + 10.3, "広告より安く、入会率が高く、続きやすい。紹介は「おまけ」ではなく主要な獲得チャネル。",
     fill=NAVY, sz=12.5)
foot(s, "出典｜POPER「Comiru」2022年（n=300）／DeltaX「塾選」退塾に関する調査 2024年（n=65）／"
        "入会率の差は業界推計値（出典元の特定不可）")

# ============================================================
# S35 LINEOA実績・出典一覧
# ============================================================
s = slides[34]
frame(s, "LINE公式アカウントの実績と、この資料の出典",
      ["LINEヤフー公式の「教育・スクール」導入事例は未取得のため差込枠にしています。",
       "捏造はしません。この資料で使った数値の出所を、すべて下に並べます。"])
placeholder(s, CX0, CY0 + 0.25, 12.3, 4.4,
            "★ LINEヤフー公式 導入事例（教育・スクール）",
            "lycbiz.com/jp/case-study から該当業種の事例を取得して差し込む。\n"
            "該当なしの場合は「公式事例は存在しない＝先行者になれる」と正直に書く\n"
            "（注文住宅案件では実際に該当なしだった）")
box(s, CX0, CY0 + 5.0, 12.3, 4.9, fill=PALE)
T(s, CX0 + 0.35, CY0 + 5.2, 11.6, 4.5,
  [one("LINEヤフー公式データ（倍率表現は使っていません）", 11, True, NAVY, sa=5),
   one("・LINE 国内月間利用者数：1億人突破（2025年12月末時点）", 9.5, None, INK, ls=1.3, sa=4),
   one("・メッセージの開封：受信直後 約2割／3〜6時間で約5割／当日中に約8割",
       9.5, None, INK, ls=1.3, sa=6),
   one("※「開封率60〜80%＝メルマガの3〜4倍」という表現は公式値ではないため使用していません。",
       8.5, None, MUT, ls=1.25, sa=3),
   one("※「LINE経由で来店率40〜50%改善」はDYM支援実績であり公式値ではありません。出典を混ぜません。",
       8.5, None, MUT, ls=1.25)], anchor="t")
VX = CX0 + 12.9
T(s, VX, CY0 + 0.25, 12.22, 0.65, [one("この資料の出典一覧", 11.5, True, NAVY)], anchor="m")
SRC_LIST = [
    ("S03", "厚労省 人口動態統計／経産省 特定サービス産業動態統計／文科省 学校基本調査／矢野経済研究所2024"),
    ("S05", "LINEヤフー前後検索（実測）／塾ナラ2025 n=200／塾シル2026 n=249／明光ネットワークジャパン2020"),
    ("S07", "業界推計値（出典元の特定不可）★社内実績で差し替え"),
    ("S10", "Googleトレンド 日本 2026年年初来（実測）"),
    ("S11", "LINEヤフー前後検索 起点KW「塾 費用」2022-05〜2023-05（実測）"),
    ("S12", "LINE公式アカウント page.line.me・実機 2026-09-15取得（実測・証跡あり）"),
    ("S18", "DeltaX「塾選」退塾に関する調査 2024年 n=65"),
    ("S23/S27", "POPER「Comiru」2022年 n=300／DeltaX「塾選」2024年 n=65"),
    ("S25", "Googleトレンド（実測）＋前後検索の26日"),
    ("S34", "POPER「Comiru」2022年 n=300／ネオマーケティング2019"),
    ("入塾学年", "オリコンME 2025年 学習塾 利用実態データ"),
    ("入塾きっかけ", "インタースペース「ママスタ」2021年 n=1,037"),
]
simple_table(s, VX, CY0 + 0.95, 12.22, 8.95,
             ["スライド", "出典"], SRC_LIST,
             col_w=[2.6, 9.62], align=["c", "l"], hsz=9, bsz=8)
box(s, CX0, CY0 + 10.3, CW, 1.6, fill=WHITE, line=RED, lw=1.25)
T(s, CX0 + 0.4, CY0 + 10.4, CW - 0.8, 1.4,
  [one("★ 提出前の必須作業：公表統計（S03）はe-Statで統計表を開いて数値を照合すること。"
       "★ 競合の実名掲載と、民間調査の引用可否は上司確認事項。"
       "特に明光義塾は競合であり、その調査を引用している点に注意。",
       9.5, None, INK, ls=1.3)], anchor="m")
foot(s, "出典のない数値は書かない。業界指標／自社実績／業界推計値／モデル値は、この資料ではラベルを分けて表記している")


# ============================================================
# 業種残骸の最終チェック（正規表現スキャン）
# ============================================================
import re

RESIDUE_PATTERNS = [
    r"○○業界", r"〇〇業界", r"定期購入", r"転職", r"求人", r"カゴ落ち",
    r"利用状況タグ", r"人材\(転職\)業界", r"あああ", r"すうじ", r"なまえ",
]
residue_hits = []
for i, sl in enumerate(slides, 1):
    for sh in sl.shapes:
        if sh.has_text_frame:
            txt = sh.text_frame.text
            for pat in RESIDUE_PATTERNS:
                if re.search(pat, txt):
                    residue_hits.append((i, pat, txt[:44]))
if residue_hits:
    print("★業種残骸が残っています：")
    for i, pat, txt in residue_hits:
        print(f"  slide {i}: /{pat}/ -> {txt!r}")
else:
    print("業種残骸チェック：残存0")

prs.save(OUT)
print("saved:", OUT)
print("slides:", len(list(Presentation(OUT).slides)))
