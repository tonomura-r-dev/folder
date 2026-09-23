# -*- coding: utf-8 -*-
"""株式会社チェングロウス御中｜LINE公式アカウント運用のご提案（v1）

ベース：_templates/DYM_LINEOA_提案FMT_ver2_3.pptx（やる気スイッチ59枚・殿村さん指定）
方針：汎用ページは触らず、クライアント固有ページのみ差し替え。情報を入れ過ぎない。

  python3 _build/build_chengrowth.py

構成操作：
- 文言差し替え（in-place）：P1表紙／P5考え方／P19要件定義／P37プラン
- 作り直し（clear→再構築）：P2現状／P4動線／P7前後検索(★枠)／P10トレンド実測
  ／P11大手LINE実測／P20費用対効果(SIM実数)／P23 14日ステップ／P25年間企画
- 削除：P3・P8・P9・P12・P24・P26・P33・P34（ブランド2つ目＋通知メッセ）→ 51枚

★のまま残した箇所：P7 前後検索（未取得・差込枠）
"""
import re
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
SRC = str(ROOT / "_templates" / "DYM_LINEOA_提案FMT_ver2_3.pptx")
OUT = str(ROOT / "20260922_株式会社チェングロウス御中_LINE公式アカウント運用のご提案.pptx")

TNAVY = "002060"; NAVY = "1F285A"; ORANGE = "ED7D31"; RED = "C00000"
INK = "333333"; MUT = "7F7F7F"; WHITE = "FFFFFF"; PALE = "F4F7FF"
PORANGE = "FCE4D6"; BORDER = "D9D9D9"; GREEN = "06C755"; PRED = "FDF2F2"

SW, SH = 27.52, 19.05
TITLE_XY = (1.52, 0.38, 24.4, 0.94)
LEAD_XY = (1.20, 1.80, 25.1, 1.90)
DIV_Y = 3.86
CX0, CW = 1.20, 25.12
CY0 = 4.30
FOOT_Y = 17.35

shutil.copyfile(SRC, OUT)
prs = Presentation(OUT)
slides = list(prs.slides)
assert len(slides) == 59, len(slides)


# ================= helpers =================
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
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)


def put_text(tf, paras, anchor="t", ml=0.14, mr=0.14, mt=0.06, mb=0.06, wrap=True):
    reset_tf(tf)
    tf.word_wrap = wrap
    tf.margin_left = Cm(ml); tf.margin_right = Cm(mr)
    tf.margin_top = Cm(mt); tf.margin_bottom = Cm(mb)
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
        for t, sz, b, c in p["runs"]:
            r = para.add_run(); r.text = t
            set_font(r, sz, b, c)
    return tf


def T(slide, x, y, w, h, paras, anchor="t", **kw):
    b = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    put_text(b.text_frame, paras, anchor=anchor, **kw)
    return b


def one(text, sz, b=None, c=INK, align="l", sa=None, ls=None):
    d = {"runs": [(text, sz, b, c)], "align": align}
    if sa is not None: d["sa"] = sa
    if ls is not None: d["ls"] = ls
    return d


def multi(runs, align="l", sa=None, ls=None):
    d = {"runs": runs, "align": align}
    if sa is not None: d["sa"] = sa
    if ls is not None: d["ls"] = ls
    return d


def box(slide, x, y, w, h, fill=None, line=None, lw=1.0,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, dash=None):
    sp = slide.shapes.add_shape(shape, Cm(x), Cm(y), Cm(w), Cm(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = RGBColor.from_string(line)
        sp.line.width = Pt(lw)
        if dash: sp.line.dash_style = dash
    sp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try: sp.adjustments[0] = radius
        except Exception: pass
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
    clear_slide(slide)
    T(slide, *TITLE_XY, [one(title, 16, True, TNAVY)], anchor="m", ml=0, mr=0)
    T(slide, *LEAD_XY, [one(l, 12, None, INK, ls=1.28) for l in lead], anchor="m", ml=0, mr=0)
    ln = slide.shapes.add_connector(1, Cm(0), Cm(DIV_Y), Cm(SW), Cm(DIV_Y))
    ln.line.color.rgb = RGBColor.from_string(BORDER)
    ln.line.width = Pt(1.0)


def foot(slide, text):
    T(slide, CX0, FOOT_Y, CW, 0.9, [one(text, 7.5, None, MUT, ls=1.15)], ml=0, mr=0)


def band(slide, y, text, fill=NAVY, col=WHITE, sz=12, h=1.0, x=CX0, w=CW):
    sp = box(slide, x, y, w, h, fill=fill, radius=0.10)
    put_text(sp.text_frame, [one(text, sz, True, col, align="c", ls=1.25)],
             anchor="m", ml=0.3, mr=0.3, mt=0, mb=0)
    return sp


def placeholder(slide, x, y, w, h, label, note):
    sp = box(slide, x, y, w, h, fill=WHITE, line=MUT, lw=1.25,
             dash=MSO_LINE_DASH_STYLE.DASH)
    put_text(sp.text_frame,
             [one(label, 11, True, MUT, align="c", sa=4),
              one(note, 8.5, None, MUT, align="c", ls=1.25)],
             anchor="m", ml=0.3, mr=0.3, mt=0.1, mb=0.1)
    return sp


def simple_table(slide, x, y, w, h, headers, rows, col_w=None,
                 hsz=9.5, bsz=9, header_fill=NAVY, zebra=PALE, align=None, row_h=None):
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
        c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string(header_fill)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        put_text(c.text_frame, [one(htext, hsz, True, WHITE, align=(align[j] if align else "c"))],
                 anchor="m", ml=0.12, mr=0.12, mt=0.02, mb=0.02)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(zebra if (zebra and i % 2 == 0) else WHITE)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            put_text(c.text_frame, [one(str(val), bsz, None, INK, align=(align[j] if align else "l"))],
                     anchor="m", ml=0.12, mr=0.12, mt=0.02, mb=0.02)
    return gf


def walk_replace(shape, mapping):
    """グループ・テーブルを含めてrun単位で文字列置換"""
    if shape.shape_type == 6:
        for c in shape.shapes:
            walk_replace(c, mapping)
        return
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    for r in para.runs:
                        for k, v in mapping.items():
                            if k in r.text:
                                r.text = r.text.replace(k, v)
        return
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for r in para.runs:
                for k, v in mapping.items():
                    if k in r.text:
                        r.text = r.text.replace(k, v)


def replace_on(slide, mapping):
    for sh in slide.shapes:
        walk_replace(sh, mapping)


def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    rId = slides_list[index].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides_list[index])


# ============================================================
# 文言差し替え（in-place）
# ============================================================
replace_on(slides[0], {  # P1 表紙
    "株式会社やる気スイッチグループ 御中": "株式会社チェングロウス 御中",
})
replace_on(slides[4], {  # P5 本提案の考え方
    "体験予約で終わらず、来場 → 入会 → 継続まで並走できる": "応募で終わらず、面談 → 就業 → 定着まで並走できる",
    "体験予約（CV）": "応募・相談（CV）",
    "来場": "面談",
    "入会": "就業",
    "継続・紹介": "定着・紹介",
})
replace_on(slides[18], {  # P19 要件定義
    "①離脱ユーザーの再学習によるCV数増加": "①離脱ユーザーの再訴求によるCV数増加",
    "②体験予約のみユーザーの削減 → 来店・契約サポート": "②応募に至らない検討ユーザーの面談化 → 就業サポート",
    "学習コンテンツの展開でサービス理解度を引": "職種・年収コンテンツの展開でサービス理解度を引",
    "体験予約": "応募・面談",
})
replace_on(slides[26], {  # P27 その他（コンサルに含む）
    "・暮らしスマイルアンケートをLINEリサーチ機能で実施": "・求職者アンケートをLINEリサーチ機能で実施",
    "複数エリア（17アカウント）共通のFAQ対応を効率化": "よくある質問（応募方法・年収・勤務地 等）の対応を効率化",
    "（診断コンテンツ・複数サービスを前提に整理）": "（診断コンテンツを前提に整理）",
})
replace_on(slides[36], {  # P37 運用プラン
    "施策成果件数に伴う費用発生でリスクを抑えた運用スタートを実現。": "月額固定のコンサル運営で、費用を抑えた運用スタートを実現。",
    "保護者ユーザーのカスタマージャーニー（8フェーズ）": "求職者（整備士）のカスタマージャーニー（8フェーズ）",
    "接触〜入会・継続までの動線設計": "接触〜面談・就業までの動線設計",
    "入会": "就業",
})

# ============================================================
# P2 現状把握（作り直し）
# ============================================================
s = slides[1]
frame(s, "現状把握【自動車 求人Navi】",
      ["広告は動いている。しかし、応募につながっていない。"])
nums = [
    ("Meta広告の応募単価", "約70万円/件", "年250万円の投下で\n応募は年3〜4件"),
    ("整備士の応募単価（Google）", "6万円", "目標は2〜2.5万円\n＝2.4〜3倍の乖離"),
    ("月間サイト訪問", "3,215 UU", "SimilarWeb推計\n応募まで進むのはごく一部"),
]
cw3 = (CW - 0.6) / 3
for i, (label, val, sub) in enumerate(nums):
    x = CX0 + i * (cw3 + 0.3)
    box(s, x, CY0 + 0.2, cw3, 4.4, fill=PALE)
    T(s, x, CY0 + 0.5, cw3, 0.7, [one(label, 11.5, True, NAVY, align="c")], anchor="m")
    T(s, x, CY0 + 1.4, cw3, 1.5, [one(val, 26, True, ORANGE, align="c")], anchor="m")
    T(s, x, CY0 + 3.0, cw3, 1.4,
      [one(l, 9.5, None, INK, align="c", ls=1.3) for l in sub.split("\n")], anchor="m")
box(s, CX0, CY0 + 5.0, CW, 2.2, fill=PRED, line=RED, lw=1.25)
T(s, CX0 + 0.5, CY0 + 5.15, CW - 1.0, 1.9,
  [one("あわせて2点、すぐに直せる改善点があります。", 12, True, RED, sa=4),
   one("① トップページの「転職支援サービス申込」ボタンが現在リンク切れ（押した先が非公開の求人ページ）", 10.5, None, INK, ls=1.3),
   one("② 電話相談の受付が平日10:00〜18:00のみ ＝ 整備士の方が働いている時間帯", 10.5, None, INK, ls=1.3)],
  anchor="m")
foot(s, "出典：与件ヒアリング（2026-09）／サイト実査 2026-09-21。UUはSimilarWeb推計のため、GA実測値をいただき次第置き換え")

# ============================================================
# P4 想定動線（作り直し）
# ============================================================
s = slides[3]
frame(s, "想定動線【自動車 求人Navi】",
      ["現状は「今すぐ応募する」の一段のみ。迷っている段階の受け皿をLINEでつくる。"])
T(s, CX0, CY0 + 0.15, 4.0, 0.6, [one("現状", 11, True, MUT)], anchor="m")
steps_now = [("流入", NAVY), ("求人を見る", NAVY), ("今すぐ応募する", ORANGE), ("迷った人は離脱", RED)]
cwn = (CW - 1.8) / 4
for i, (t, c) in enumerate(steps_now):
    x = CX0 + i * (cwn + 0.6)
    bb = box(s, x, CY0 + 0.8, cwn, 1.1, fill=c if c != RED else PRED,
             line=RED if c == RED else None, lw=1.25)
    put_text(bb.text_frame, [one(t, 11, True, WHITE if c != RED else RED, align="c")], anchor="m")
    if i < 3:
        T(s, x + cwn + 0.05, CY0 + 0.95, 0.5, 0.8, [one("→", 14, True, MUT, align="c")], anchor="m")
T(s, CX0, CY0 + 2.3, 6.0, 0.6, [one("LINE導入後", 11, True, GREEN)], anchor="m")
steps_new = ["流入", "LINE友だち化\n（離脱防止・CPF）", "30秒診断\n（職種・資格・温度感）", "14日間の情報配信\n（年収相場・体験談）", "面談・応募", "就業"]
cwm = (CW - 2.5) / 6
for i, t in enumerate(steps_new):
    x = CX0 + i * (cwm + 0.5)
    fill = GREEN if i in (1, 2, 3) else NAVY
    bb = box(s, x, CY0 + 2.95, cwm, 1.5, fill=fill)
    put_text(bb.text_frame, [one(l, 9, True, WHITE, align="c", ls=1.15) for l in t.split("\n")],
             anchor="m", ml=0.08, mr=0.08)
    if i < 5:
        T(s, x + cwm + 0.02, CY0 + 3.25, 0.46, 0.8, [one("→", 12, True, MUT, align="c")], anchor="m")
band(s, CY0 + 5.1, "広告・サイトは変えない。「応募の手前」に、相談の階段を1段足すだけ。", fill=NAVY, sz=13, h=1.1)
foot(s, "意見｜DYM提案。CV地点＝求人応募・転職支援サービス登録（面談）")

# ============================================================
# P7 市場分析：前後検索（★差込枠）
# ============================================================
s = slides[6]
frame(s, "市場分析（前後検索クエリ）",
      ["★LINEヤフー社提供の前後検索クエリを取得後、検討の流れを反映する。"])
bands3 = [("検索“前”", MUT), ("起点KW（整備士 ほか）", NAVY), ("検索“後”", RED)]
cw3b = (CW - 0.6) / 3
for i, (h, c) in enumerate(bands3):
    x = CX0 + i * (cw3b + 0.3)
    bb = box(s, x, CY0 + 0.2, cw3b, 1.0, fill=c)
    put_text(bb.text_frame, [one(h, 11.5, True, WHITE, align="c")], anchor="m")
placeholder(s, CX0, CY0 + 1.5, CW, 5.3, "★差込枠｜前後検索データ（取得中）",
            "対象KWを一語で選定して取得後、検索前後のクエリ変化と\n「LINE追加を仕掛けるタイミング」をこのページに反映する")
foot(s, "★取得中｜出典：LINEヤフー社 管理ツール")

# ============================================================
# P10 市場分析：Googleトレンド実測（作り直し）
# ============================================================
s = slides[9]
frame(s, "市場分析（検索トレンド）",
      ["「整備士」の検索は一年中ほぼフラット。転職検討は季節を選ばず続いている。"])
img = str(ROOT / "_images" / "chengrowth_trend_seibishi_5y.png")
iw = 19.0  # 高さ≒6.4cmに収める
s.shapes.add_picture(img, Cm((SW - iw) / 2), Cm(CY0 + 0.2), width=Cm(iw))
pts = [
    ("山", "3月", "年度替わり・転職の動きが最大"),
    ("谷", "12月", "それでも山との差は約1.4倍"),
    ("含意", "通年", "検討は一年中。接点は3月に集中"),
]
cw3 = (CW - 0.6) / 3
for i, (tag, val, sub) in enumerate(pts):
    x = CX0 + i * (cw3 + 0.3)
    box(s, x, CY0 + 7.0, cw3, 2.1, fill=PALE)
    T(s, x, CY0 + 7.15, cw3, 0.5, [one(tag, 9.5, True, MUT, align="c")], anchor="m")
    T(s, x, CY0 + 7.6, cw3, 0.8, [one(val, 16, True, ORANGE, align="c")], anchor="m")
    T(s, x, CY0 + 8.4, cw3, 0.6, [one(sub, 8.5, None, INK, align="c")], anchor="m")
band(s, CY0 + 9.4, "業界の接点は3月に集中。通年で検討層を受け止める先は、空いている。",
     fill=NAVY, sz=12.5, h=1.0)
foot(s, "出典：Googleトレンド（日本・過去5年・2026-09-21取得）。相対指標のため他KWとの数値比較は不可")

# ============================================================
# P11 競合分析：大手求人サービスのLINE活用（作り直し）
# ============================================================
s = slides[10]
frame(s, "競合分析（求人サービスのLINE活用状況）",
      ["総合大手はLINEが標準装備。一方、自動車・整備士特化は空白地帯。"])
simple_table(s, CX0, CY0 + 0.2, CW, 5.0,
             ["サービス", "種別", "LINE友だち数"],
             [["リクナビNEXT", "総合・転職サイト", "3,235,678人"],
              ["求人ボックス", "総合・アグリゲーション", "1,862,866人"],
              ["マイナビ転職", "総合・転職サイト", "754,302人"],
              ["doda", "総合・エージェント", "公開アカウントなし（会員登録後に1:1連携）"],
              ["プレックスジョブ【整備士】", "整備士特化", "6,288人"],
              ["整備士の求人・転職｜カンパネル", "整備士特化", "196人"]],
             col_w=[CW * 0.36, CW * 0.28, CW * 0.36], row_h=0.75, bsz=9.5)
band(s, CY0 + 5.6, "「整備士 × LINE」はまだ誰も取っていない。総合大手の型を持ち込むだけで先行者になれる。",
     fill=NAVY, sz=12.5, h=1.1)
foot(s, "出典：page.line.me（公式）2026-09-21実測。友だち数は取得日時点の公表値")

# ============================================================
# P20 想定の費用対効果（作り直し・SIM実数）
# ============================================================
s = slides[19]
frame(s, "想定の費用対効果",
      ["Meta広告の目的を「応募獲得」から「友だち獲得」へ切り替える（広告費は据え置き）。"])
box(s, CX0, CY0 + 0.2, CW * 0.42, 5.6, fill=PALE)
T(s, CX0 + 0.4, CY0 + 0.45, CW * 0.42 - 0.8, 0.6, [one("費用（税抜）", 12, True, NAVY)], anchor="m")
T(s, CX0 + 0.4, CY0 + 1.15, CW * 0.42 - 0.8, 4.4,
  [multi([("初期　", 11, True, INK), ("24.5万円", 16, True, ORANGE)], sa=2),
   one("　構築20万＋離脱防止1.5万＋アカウント3万", 9, None, MUT, sa=8),
   multi([("月次　", 11, True, INK), ("10.5万円", 16, True, ORANGE)], sa=2),
   one("　コンサル7万＋離脱防止3万＋アカウント費5千", 9, None, MUT, sa=8),
   multi([("CPF広告費　", 11, True, INK), ("月20.8万円", 13, True, INK)], sa=2),
   one("　既存Meta予算（年250万）の転用＝追加出費ゼロ", 9, None, MUT)],
  anchor="t")
rx = CX0 + CW * 0.42 + 0.4
rw = CW - CW * 0.42 - 0.4
box(s, rx, CY0 + 0.2, rw, 5.6, fill=WHITE, line=NAVY, lw=1.25)
T(s, rx + 0.4, CY0 + 0.45, rw - 0.8, 0.6, [one("効果（6ヶ月・シミュレーション）", 12, True, NAVY)], anchor="m")
half = (rw - 1.2) / 2
T(s, rx + 0.4, CY0 + 1.3, half, 0.6, [one("面談・応募", 10.5, True, INK, align="c")], anchor="m")
T(s, rx + 0.4, CY0 + 1.9, half, 1.3, [one("196件", 30, True, ORANGE, align="c")], anchor="m")
T(s, rx + 0.4 + half + 0.4, CY0 + 1.3, half, 0.6, [one("獲得単価（累計CPA）", 10.5, True, INK, align="c")], anchor="m")
T(s, rx + 0.4 + half + 0.4, CY0 + 1.9, half, 1.3, [one("11,036円", 26, True, ORANGE, align="c")], anchor="m")
T(s, rx + 0.4, CY0 + 3.5, rw - 0.8, 2.1,
  [one("現状の整備士CPA 6万円・目標2〜2.5万円に対し、大幅に下回る水準。", 10.5, True, INK, ls=1.3, sa=3),
   one("6ヶ月目の単月CPAは8,730円まで低下（友だち資産の積み上がりによる）。", 9.5, None, INK, ls=1.3, sa=3),
   one("参考：CPFを使わない場合＝6ヶ月42件・累計CPA 20,833円", 8.5, None, MUT, ls=1.3)],
  anchor="t")
foot(s, "SIM①（CPFあり・コンサル運営）より。前提：UU3,215（SimilarWeb推計）・CPF単価280円ほか仮置き値を含む。実績値をいただき次第、再試算")

# ============================================================
# P23 実施施策：14日間ステップ配信（作り直し）
# ============================================================
s = slides[22]
frame(s, "実施施策【14日間ステップ配信】",
      ["友だち追加後14日間で、診断結果をもとに内容を出し分けながら面談へご案内する。",
       "売り込みは14日間で2回だけ。残りは「不安に先回りして答える」情報配信。"])
days = [
    ("Day 0", "あいさつ＋30秒診断", "「こちらから電話をかけることは\nありません」を明言。\n職種・資格・温度感を4問で取得"),
    ("Day 3", "年収相場コンテンツ", "資格タグ別に出し分け。\n「2級整備士の年収には\n100万円以上の幅がある」"),
    ("Day 7", "面談のご案内 ①", "「履歴書、まだ要りません」\n在職中は夜の面談OK・\n紹介は自動車業界のみを明記"),
    ("Day 14", "面談のご案内 ②", "毎日の配信はここで終了。\n以後は月数回、新着求人と\n相場情報だけに切り替え"),
]
cw4 = (CW - 0.9) / 4
for i, (d, h, b) in enumerate(days):
    x = CX0 + i * (cw4 + 0.3)
    tb = box(s, x, CY0 + 0.2, cw4, 0.7, fill=ORANGE if "7" in d or "14" in d else NAVY)
    put_text(tb.text_frame, [one(f"{d}｜{h}", 9.5, True, WHITE, align="c")], anchor="m", ml=0.05, mr=0.05)
    card(s, x, CY0 + 1.0, cw4, 3.2, "", b.split("\n"), fill=PALE, hsz=1, bsz=9, ls=1.35, anchor="m")
    if i < 3:
        T(s, x + cw4 + 0.02, CY0 + 1.8, 0.3, 0.8, [one("→", 11, True, MUT, align="c")], anchor="m")
band(s, CY0 + 4.6, "面談・応募 196件（6ヶ月累計・SIM①）。未予約の方は月数回の定常配信で通年フォロー。",
     fill=NAVY, sz=12, h=1.0)
foot(s, "意見｜配信文面の実物はコンテンツ案v1（別紙）に全文記載。Day7・Day14がCVオファー")

# ============================================================
# P25 想定企画（作り直し・年間）
# ============================================================
s = slides[24]
frame(s, "想定企画【年間】",
      ["検索の山（3月）に向けて1〜2月から前倒しで仕込む。谷（12月）も検討は続くため配信は止めない。"])
plans = [
    ("1月", "新生活前の転職相談"), ("2月", "3月転職の準備特集"), ("3月", "★年度替わり・求人最大化"),
    ("4月", "新年度スタート応援"), ("5月", "連休明けの見直し訴求"), ("6月", "夏賞与前の情報提供"),
    ("7月", "賞与後の転職検討期"), ("8月", "資格取得・キャリア特集"), ("9月", "下期スタート求人"),
    ("10月", "年内転職ラストコール"), ("11月", "年末年始の働き方特集"), ("12月", "冬も相談だけOK訴求"),
]
cw4 = (CW - 0.9) / 4
for i, (m, t) in enumerate(plans):
    r, c = divmod(i, 4)
    x = CX0 + c * (cw4 + 0.3)
    y = CY0 + 0.2 + r * 1.75
    bb = box(s, x, y, cw4, 1.55, fill=PORANGE if m == "3月" else PALE)
    put_text(bb.text_frame,
             [one(m, 10, True, ORANGE if m == "3月" else NAVY, sa=2),
              one(t, 9, None, INK, ls=1.15)],
             anchor="m", ml=0.2, mr=0.15, mt=0.08, mb=0.06)
band(s, CY0 + 5.7, "配信は月2〜4本。クリックした人でオーディエンスを育て、配信のたびに精度が上がる設計。",
     fill=NAVY, sz=12, h=1.0)
foot(s, "出典：GoogleトレンドKW「整備士」（山3月・谷12月）／企画の詳細文面は個別に作成")

# ============================================================
# P33（旧・通知メッセ）→ LY公式実績ページに作り替え（2026-09-23追加）
# ============================================================
s = slides[32]
frame(s, "公式実績【求人・人材 × LINE】",
      ["LINEヤフー社の公式導入事例より（いずれも実名・公開情報）。"])
cw2 = (CW - 0.5) / 2
box(s, CX0, CY0 + 0.2, cw2, 5.2, fill=PALE)
T(s, CX0 + 0.4, CY0 + 0.5, cw2 - 0.8, 0.6, [one("UZUZ（20代の就職・転職サポート）", 12, True, NAVY)], anchor="m")
T(s, CX0 + 0.4, CY0 + 1.2, cw2 - 0.8, 0.6, [one("LINE経由の面談予約率", 10.5, None, INK, align="c")], anchor="m")
T(s, CX0 + 0.4, CY0 + 1.8, cw2 - 0.8, 1.4, [one("55%", 34, True, ORANGE, align="c")], anchor="m")
T(s, CX0 + 0.4, CY0 + 3.3, cw2 - 0.8, 1.8,
  [one("・問い合わせの約3割がLINE経由。他チャネル比で高水準", 9.5, None, INK, ls=1.35, sa=2),
   one("・友だち追加広告（CPF）の獲得単価 約300円", 9.5, None, INK, ls=1.35, sa=2),
   one("・チャットボットで求人紹介〜面談予約を自動化", 9.5, None, INK, ls=1.35)],
  anchor="t")
x2 = CX0 + cw2 + 0.5
box(s, x2, CY0 + 0.2, cw2, 5.2, fill=PALE)
T(s, x2 + 0.4, CY0 + 0.5, cw2 - 0.8, 0.6, [one("タウンワーク（求人メディア）", 12, True, NAVY)], anchor="m")
T(s, x2 + 0.4, CY0 + 1.2, cw2 - 0.8, 0.6, [one("リッチメニュー経由の応募数", 10.5, None, INK, align="c")], anchor="m")
T(s, x2 + 0.4, CY0 + 1.8, cw2 - 0.8, 1.4, [one("+79%", 34, True, ORANGE, align="c")], anchor="m")
T(s, x2 + 0.4, CY0 + 3.3, cw2 - 0.8, 1.8,
  [one("・リッチメニュー改善でタップ数+20%", 9.5, None, INK, ls=1.35, sa=2),
   one("・その結果、メニュー経由の応募が79%向上", 9.5, None, INK, ls=1.35, sa=2),
   one("・入口設計への投資が応募数に直結した事例", 9.5, None, INK, ls=1.35)],
  anchor="t")
band(s, CY0 + 5.8, "「面談への近さ」と「入口設計の効果」。本提案の2本柱には、どちらも公式実績がある。",
     fill=NAVY, sz=12.5, h=1.1)
foot(s, "出典：LINEヤフー for Business 導入事例（UZUZ／タウンワーク）。数値は各事例記事の公表値")

# ============================================================
# 削除（末尾から）：P34,P26,P24,P12,P9,P8,P3（P33は実績ページとして残す）
# ============================================================
for idx in sorted([33, 25, 23, 11, 8, 7, 2], reverse=True):
    delete_slide(prs, idx)
assert len(prs.slides) == 52, len(prs.slides)

# 実績ページを「想定の費用対効果」の直後へ移動
sld_ids = prs.slides._sldIdLst
titles = []
for i, sl in enumerate(prs.slides):
    t = ""
    for sh in sl.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            t = sh.text_frame.text.strip().split("\n")[0]
            break
    titles.append(t)
src = next(i for i, t in enumerate(titles) if "公式実績" in t)
dst = next(i for i, t in enumerate(titles) if "費用対効果" in t)
ids = list(sld_ids)
el = ids[src]
sld_ids.remove(el)
sld_ids.insert(dst + 1 if src > dst else dst, el)

# ============================================================
# 残骸チェック（最終・全枚）
# ============================================================
PAT = re.compile(r"やる気スイッチ|忍者ナイン|チャイルドアイズ|習い事|理英会|どんちゃか|"
                 r"セントラルスポーツ|biima|知育|幼児|運動会|体験予約|入会|保護者|子ども")
hits = []


def scan(shape, i):
    if shape.shape_type == 6:
        for c in shape.shapes:
            scan(c, i)
        return
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for m in PAT.findall(cell.text):
                    hits.append((i, m))
        return
    if shape.has_text_frame:
        for m in PAT.findall(shape.text_frame.text):
            hits.append((i, m))


for i, sl in enumerate(prs.slides, 1):
    for sh in sl.shapes:
        scan(sh, i)
if hits:
    print("★残骸あり：", sorted(set(hits)))
else:
    print("残骸チェック：残存0")

prs.save(OUT)
print("saved:", OUT, f"（{len(Presentation(OUT).slides)}枚）")
