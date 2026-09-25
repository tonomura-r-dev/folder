# -*- coding: utf-8 -*-
"""教育（塾）業界 LINEOA施策提案 v2（2026-09-25）
正＝殿村さんPC保存版（35枚・9/17）をベースに、構成原稿v2（_drafts/教育（塾）業界_構成原稿v2.md）を反映して32枚にする。
- スライドの新規追加はしない。既存スライドを作り替えて、sldIdLst で並べ替える
- 8フェーズの帯は「検索データ」以降に入れる（各社LINEのページには入れない）
  python3 _build/build_juku_v2.py <入力pptx(35枚版)> <出力pptx>
"""
import copy
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
OLD = list(prs.slides)          # 旧版の並び（0始まり）
S = lambda n: OLD[n - 1]        # 旧ページ番号で引く

NAVY, TNAVY, INK, GRAY, LGRAY = "1F285A", "002060", "333333", "7F7F7F", "D9D9D9"
PALE, ORANGE, RED, GREEN, LINE_GREEN = "F4F7FF", "ED7D31", "C00000", "00897B", "06C755"
BLUE, PURPLE, BEIGE, LBLUE = "3467B2", "8E4EC6", "FFF2CC", "DDEBF7"
PHASES = ["①接触", "②離脱", "③育成", "④リード獲得", "⑤リード有効化", "⑥再育成", "⑦入会（契約）", "⑧紹介"]


# ---------------------------------------------------------------- 基本部品
def meiryo(run):
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", "メイリオ")


def box(s, x, y, w, h, fill, lines, color=INK, size=10, bold=False, anchor=MSO_ANCHOR.MIDDLE,
        align=PP_ALIGN.CENTER, line=None, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, ml=0.2, dash=False, adj=0.08):
    sp = s.shapes.add_shape(shape, Cm(x), Cm(y), Cm(w), Cm(h))
    if fill:
        sp.fill.solid()
        sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sp.fill.background()
    if line:
        sp.line.color.rgb = RGBColor.from_string(line)
        sp.line.width = Pt(lw)
        if dash:
            ln = sp.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sp.adjustments[0] = adj
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(ml)
    tf.margin_top = tf.margin_bottom = Cm(0.08)
    for i, ln in enumerate(lines):
        t, sz, b, c = (ln, size, bold, color) if isinstance(ln, str) else (list(ln) + [None] * 4)[:4]
        sz, b, c = sz or size, bold if b is None else b, c or color
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = t
        r.font.size = Pt(sz)
        r.font.bold = b
        r.font.color.rgb = RGBColor.from_string(c)
        meiryo(r)
    return sp


def text(s, x, y, w, h, lines, size=10, color=INK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    return box(s, x, y, w, h, None, lines, color=color, size=size, bold=bold, align=align, anchor=anchor,
               shape=MSO_SHAPE.RECTANGLE, ml=0.05)


def arrow(s, x1, y1, x2, y2, color="8C8C8C", w=1.75):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    c.line.color.rgb = RGBColor.from_string(color)
    c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle"}))


def chevron(s, x, y, w, h, fill, t, color="FFFFFF", size=9.5, first=False, ml=0.3):
    return box(s, x, y, w, h, fill, [(t, size, True, color)], shape=MSO_SHAPE.PENTAGON if first else MSO_SHAPE.CHEVRON, ml=ml)


def picture(s, rel, x, y, w=None, h=None):
    return s.shapes.add_picture(str(ROOT / rel), Cm(x), Cm(y), Cm(w) if w else None, Cm(h) if h else None)


def set_para_text(shape, lines):
    """書式（最初のrunのrPr・pPr）を保ったまま、段落ごとに文を入れ替える。runは endParaRPr より前に入れる"""
    tf = shape.text_frame
    tx = tf._txBody
    ps = tx.findall(qn("a:p"))
    tmpl = copy.deepcopy(ps[0])
    r0 = tmpl.find(qn("a:r"))
    rpr = copy.deepcopy(r0.find(qn("a:rPr"))) if r0 is not None and r0.find(qn("a:rPr")) is not None else None
    for p in ps:
        tx.remove(p)
    for ln in lines:
        p = copy.deepcopy(tmpl)
        for ch in list(p):
            if ch.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
                p.remove(ch)
        r = p.makeelement(qn("a:r"), {})
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = r.makeelement(qn("a:t"), {})
        t.text = ln
        r.append(t)
        end = p.find(qn("a:endParaRPr"))
        if end is not None:
            end.addprevious(r)
        else:
            p.append(r)
        tx.append(p)


def parts(s):
    """タイトル・リード・区切り線・脚注を探す"""
    title = lead = div = foot = None
    for sh in s.shapes:
        top = Emu(sh.top).cm if sh.top is not None else 99
        if sh.has_text_frame and top < 1.0 and Emu(sh.left).cm > 1.3 and title is None:
            title = sh
        elif sh.has_text_frame and 1.5 < top < 2.0 and Emu(sh.width).cm > 20 and lead is None:
            lead = sh
        elif sh.shape_type == 9 and abs(top - 3.86) < 0.1:
            div = sh
        elif sh.has_text_frame and 17.2 < top < 17.6:
            foot = sh
    return title, lead, div, foot


def rebuild(s, title_t, lead_lines, foot_t=None, phases=None):
    """タイトル・リード・区切り線・脚注だけ残して中身を消す。phases があれば8フェーズの帯を入れる"""
    title, lead, div, foot = parts(s)
    keep = {id(x._element) for x in (title, lead, div, foot) if x is not None}
    tree = s.shapes._spTree
    for sh in list(s.shapes):
        if id(sh._element) not in keep:
            tree.remove(sh._element)
    set_para_text(title, [title_t])
    set_para_text(lead, lead_lines)
    if foot is not None:
        if foot_t:
            set_para_text(foot, [foot_t])
        else:
            tree.remove(foot._element)
    if phases is not None:
        band(s, phases)
        lead.top, lead.height = Cm(2.2), Cm(1.6)
    else:
        lead.top, lead.height = Cm(1.75), Cm(2.0)
    return s


def band(s, active):
    """8フェーズの帯。active＝塗るフェーズ番号のリスト（'all'＝全部を薄く）"""
    x0, y, W, h, g = 1.2, 1.64, 25.12, 0.5, 0.06
    w = (W - g * 7) / 8
    for i, t in enumerate(PHASES):
        n = i + 1
        if active == "all":
            fill, col = "C9D3E6", NAVY
        elif n in active:
            fill, col = NAVY, "FFFFFF"
        else:
            fill, col = "EEEEEE", "A6A6A6"
        chevron(s, x0 + i * (w + g), y, w, h, fill, t, color=col, size=7, first=(i == 0), ml=0.08)


def logic3(s, data, should, action, y=4.15, h=2.75):
    """施策ページの型：データ → すべきこと → LINE施策"""
    W, g = 25.12, 0.55
    w = (W - g * 2) / 3
    heads = [("データ（事実）", "EEEEEE", INK, "FFFFFF", data),
             ("すべきこと", BEIGE, INK, "FFFFFF", should),
             ("LINE施策", NAVY, "FFFFFF", NAVY, action)]
    for i, (hd, hf, hc, bf, body) in enumerate(heads):
        x = 1.2 + i * (w + g)
        box(s, x, y, w, h, bf, [], line=LGRAY if i < 2 else None)
        box(s, x, y, w, 0.62, hf, [(hd, 10, True, hc)], adj=0.2)
        text(s, x + 0.25, y + 0.72, w - 0.5, h - 0.8,
             [(t, 10, False, "FFFFFF" if i == 2 else INK) for t in body], anchor=MSO_ANCHOR.TOP)
        if i < 2:
            arrow(s, x + w + 0.08, y + h / 2, x + w + g - 0.08, y + h / 2, color=NAVY, w=2.25)


def pending(s, x, y, w, h, t, sub=None):
    lines = [(t, 10.5, True, "C55A11")]
    if sub:
        lines.append((sub, 9, False, GRAY))
    return box(s, x, y, w, h, "FFFBF5", lines, line="ED7D31", dash=True, lw=1.25)


def phone(s, x, y, w, h, name="◯◯塾 △△教室", bg="C9D6E8"):
    """スマホ（LINEトーク画面）の模型。画面の内側の座標を返す"""
    box(s, x, y, w, h, "2B2B2B", [], adj=0.1)
    sx, sy, sw, sh_ = x + 0.25, y + 0.55, w - 0.5, h - 1.0
    box(s, sx, sy, sw, sh_, bg, [], shape=MSO_SHAPE.RECTANGLE)
    box(s, sx, sy, sw, 0.6, "FFFFFF", [("‹  " + name, 8, True, INK)], shape=MSO_SHAPE.RECTANGLE, align=PP_ALIGN.LEFT)
    return sx, sy + 0.7, sw, sh_ - 0.7


def bubble(s, x, y, w, h, lines, size=8, fill="FFFFFF"):
    return box(s, x, y, w, h, fill, [(t, size, False, INK) for t in lines], align=PP_ALIGN.LEFT,
               anchor=MSO_ANCHOR.TOP, adj=0.12, ml=0.2)


def clear_all(s):
    tree = s.shapes._spTree
    for sh in list(s.shapes):
        tree.remove(sh._element)


def replace_in(s, pairs):
    n = 0
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                for a, b in pairs:
                    if a in r.text:
                        r.text = r.text.replace(a, b)
                        n += 1
    return n


def shift_lead_for_band(s, phases):
    """中身を残すページ用：リードを下げて帯を入れる"""
    title, lead, div, foot = parts(s)
    band(s, phases)
    lead.top, lead.height = Cm(2.2), Cm(1.6)


# ================================================================ 各ページ
# ---- P1 表紙（旧1）
s = S(1)
for sh in s.shapes:
    if sh.has_text_frame and "ターゲットが動き出す" in sh.text_frame.text:
        set_para_text(sh, ["保護者の「迷っている期間」に寄り添い、", "資料請求から体験予約・入会までをつなぐ"])
    if sh.has_text_frame and sh.text_frame.text.startswith("業界"):
        set_para_text(sh, ["業界：教育（学習塾）"])

# ---- P2 アジェンダ（旧27を作り替え）
s = rebuild(S(27), "アジェンダ", ["市場と広告の課題から入り、保護者の検索データで「すべきこと」を決め、LINE施策に落とします。"])
AG = [("1", "市場の課題", "子どもは減り、大手・中堅で奪い合う", "P3"),
      ("2", "広告の課題", "広告で取れているのは指名検索だけ", "P4"),
      ("3", "検索データ", "保護者はどう塾を選んでいるか", "P5〜8"),
      ("4", "各社LINEの今", "大手2社・中堅2社のLINE", "P9"),
      ("5", "考え方", "LINEでできること・全体設計・導線", "P10〜12"),
      ("6", "施策", "友だち追加から入会まで・施策の設計図", "P13〜29"),
      ("7", "運用・費用・スケジュール", "成果の見方・費用・体制", "P30〜35")]
for i, (n, t, sub, pg) in enumerate(AG):
    y = 4.3 + i * 1.78
    box(s, 3.0, y, 1.4, 1.4, NAVY, [(n, 16, True, "FFFFFF")], shape=MSO_SHAPE.OVAL)
    text(s, 4.9, y + 0.05, 8.5, 0.8, [(t, 15, True, TNAVY)])
    text(s, 4.9, y + 0.8, 14.5, 0.6, [(sub, 11, False, INK)])
    text(s, 21.5, y + 0.35, 3.2, 0.7, [(pg, 11, False, GRAY)], align=PP_ALIGN.RIGHT)
    if i < len(AG) - 1:
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(4.9), Cm(y + 1.6), Cm(24.7), Cm(y + 1.6))
        c.line.color.rgb = RGBColor.from_string("E7E6E6")

# ---- P3 市場データ（旧2）はそのまま

# ---- P4 広告の実態（旧3を作り替え）
s = rebuild(S(3), "広告の実態",
            ["塾の広告で取れているのは、ほとんどが塾名での検索（指名）。「個別指導 ◯◯駅」のような",
             "一般のキーワードは、1件あたりの広告費が高い。"],
            "※広告の実績は取得中（差し替え予定）。LPの離脱率は旧版のモデル値（LP CVR 2.0%）のため、実績に差し替えるか「モデル値」と明記する")
box(s, 1.2, 4.2, 12.3, 0.7, "EEEEEE", [("獲得件数の割合", 11, True, INK)], adj=0.2)
box(s, 14.02, 4.2, 12.3, 0.7, "EEEEEE", [("1件あたりの広告費（CPA）", 11, True, INK)], adj=0.2)
for i, (lab, col) in enumerate([("指名キーワード（塾名）", NAVY), ("一般キーワード（個別指導 ◯◯駅 など）", RED)]):
    for j, x in enumerate([1.2, 14.02]):
        y = 5.3 + i * 2.6
        text(s, x, y, 12.3, 0.6, [(lab, 10.5, True, col)])
        pending(s, x, y + 0.65, 12.3, 1.55, "【データ待ち】広告の実績", "棒グラフで入れる（殿村さん取得中）")
box(s, 1.2, 10.75, 25.12, 2.6, PALE, [], line=LGRAY)
text(s, 1.6, 10.9, 24.3, 0.7, [("LPに来た保護者の多くは、名前を残さず帰っている", 12, True, TNAVY)])
text(s, 1.6, 11.65, 24.3, 1.6, [("広告で連れてきた保護者も、その場で申し込まなければ接点は切れる。", 10.5, False, INK),
                                 ("LPの離脱率：【データ待ち】（実績があれば差し替え。旧版の「100人中98人」はモデル値）", 10.5, False, GRAY)])
box(s, 1.2, 13.8, 25.12, 1.5, NAVY, [("広告費を増やして一般キーワードを取りにいくのは高くつく。", 13, True, "FFFFFF"),
                                      ("サイトに来て、まだ決めかねている保護者を、LINEで逃がさない仕組みが要る。", 13, True, "FFFFFF")])

# ---- P5 誰が調べているか（旧4：中身を残して手直し）
s = S(4)
tree = s.shapes._spTree
for sh in list(s.shapes):
    if sh.shape_id in (23, 31, 32):
        tree.remove(sh._element)
    elif sh.has_text_frame and "本人にやる気" in sh.text_frame.text:
        for p in sh.text_frame.paragraphs:
            rs = p.runs
            if not rs:
                continue
            full = "".join(r.text for r in rs)
            if "本人にやる気" in full:
                rs[0].text = "「集団と個別、どちらに入れたらいいのか」"
            elif "明光" in full:
                rs[0].text = "塾を決めるのに関わる人：保護者80.5%（明光2020）"
            else:
                continue
            for r in rs[1:]:
                r.text = ""
replace_in(s, [("決定関与は保護者80.5%・子ども本人57.8%（明光2020）", "塾を決めるのに関わる人：保護者80.5%（明光2020）"),
               ("「本人にやる気がないのに入れても無駄では」", "「集団と個別、どちらに入れたらいいのか」")])
title, lead, div, foot = parts(s)
set_para_text(title, ["誰が調べているか"])
set_para_text(lead, ["塾を調べているのは40代の母親。選ぶときに一番困っているのは、料金ではなく",
                     "「うちの子に合う塾が分からない」（63.5%）。払う気はある。合うか分からないから決められない。"])
shift_lead_for_band(s, [3])
if foot is not None:
    set_para_text(foot, ["出典｜LINEヤフー前後検索（2022-05〜2023-05・実測）※最新の前後検索で更新予定／塾ナラ2025（n=200）／塾シル・ユナイトプロジェクト2026（n=249）／明光ネットワークジャパン2020"])

# ---- P6 検索の流れ（旧10を作り替え）
s = rebuild(S(10), "検索の流れ",
            ["保護者は、塾の名前を調べる前に何を調べているか。", "最初に調べること → 迷っているときに調べること → 塾の名前、の順に並べる。"],
            "出典｜LINEヤフー 前後検索（直近1年）起点KW：塾／個別指導／体験授業／高校受験 ※「塾」の結果から「熟」の語（熟年・熟語など）は除外",
            phases=[3, 4])
cols = [("最初に調べること", "EEEEEE"), ("迷っているときに調べること", BEIGE), ("塾の名前を調べる", LBLUE)]
for i, (t, f) in enumerate(cols):
    x = 1.2 + i * 8.52
    box(s, x, 4.15, 8.08, 0.7, f, [(t, 11, True, INK)], adj=0.2)
    pending(s, x, 5.0, 8.08, 7.3, "【データ待ち】前後検索", "クエリを時系列で並べ、4色で色分けする")
    if i < 2:
        arrow(s, x + 8.12, 8.6, x + 8.48, 8.6, color=NAVY, w=2.25)
LEG = [("塾選び・形態", "D6E2F3", "2B5797"), ("費用", "E8DAF5", "7440A8"), ("悩み・時期", "FBE0CF", "C0561A"), ("塾名・口コミ", "CCE7E4", "00796B")]
for i, (t, f, c) in enumerate(LEG):
    box(s, 1.2 + i * 3.2, 12.55, 3.0, 0.55, f, [(t, 9, True, c)], adj=0.2)
text(s, 14.2, 12.55, 12.1, 0.6, [("色分けは仮。データが届いたらクエリに合わせて決める", 9, False, GRAY)])
box(s, 1.2, 13.5, 25.12, 1.6, NAVY, [("塾の名前が出てくる前に、保護者はもう調べ始めている。", 13, True, "FFFFFF"),
                                      ("その間、塾と保護者の接点はない。（データが届いたら確定）", 11, False, "FFFFFF")])

# ---- P7 季節の動き（旧9を作り替え）
s = rebuild(S(9), "季節の動き",
            ["「塾」の検索が一番多いのは2月上旬。「夏期講習」は7月中旬で、冬期・春期の3〜5倍。",
             "検索の山の前から、LINEでつながっておく。"],
            "出典｜Googleトレンド（日本・2026年年初来・週次）_data/trends/教育塾/ ※相対値。スケールの違うKWは別のグラフにしている",
            phases=[1, 3])
picture(s, "_images/juku_v2_trend_juku.png", 1.2, 4.2, w=12.4)
picture(s, "_images/juku_v2_trend_koushuu.png", 13.92, 4.2, w=12.4)
picture(s, "_images/juku_v2_trend_juken.png", 1.2, 9.2, w=12.4)
box(s, 13.92, 9.35, 12.4, 3.9, PALE, [], line=LGRAY)
text(s, 14.3, 9.5, 11.7, 3.7, [("読み取れること", 11.5, True, TNAVY),
                                ("・「塾」は2月上旬が山。ただし年間を通して大きくは落ちない", 10.5, False, INK),
                                ("・講習は夏が圧倒的。冬期・春期の3〜5倍", 10.5, False, INK),
                                ("・受験の検索は1月中旬〜2月上旬に集中", 10.5, False, INK),
                                ("→ 山の時期だけ広告を増やすより、", 10.5, True, NAVY),
                                ("　 山の前からLINEで接点を持っておく", 10.5, True, NAVY)])
box(s, 1.2, 14.1, 25.12, 1.1, NAVY, [("検索の山の前から、LINEでつながっておく。", 13, True, "FFFFFF")])

# ---- P8 検索データからわかる「すべきこと」（旧14を作り替え）
s = rebuild(S(14), "検索データからわかる「すべきこと」",
            ["P4〜P7のデータを、すべきことに置き換える。", "塾名を調べる前の保護者と先につながり、申し込む前に「合う・合わない」を見せる。"],
            None, phases=[1, 2, 3, 4])
ROWS = [("広告で取れているのは指名検索だけ（P4）", "塾名を調べる前の保護者と、先につながる", "サイトの離脱防止からLINEへ"),
        ("塾名を調べる前から、保護者は調べ始めている（P6）", "比べている間、役に立つ情報を届け続ける", "14日間の配信"),
        ("一番の悩みは「うちの子に合うか分からない」63.5%（P5）", "申し込む前に「合う・合わない」を見せる", "30秒ぴったり診断"),
        ("検索の山は2月と7月（P7）", "山の前から接点を持っておく", "時期に合わせた企画配信")]
hx = [1.2, 10.2, 18.9]
hw = [8.5, 8.2, 7.42]
for x, w_, (t, f, c) in zip(hx, hw, [("データ（事実）", "EEEEEE", INK), ("すべきこと", BEIGE, INK), ("LINE施策", NAVY, "FFFFFF")]):
    box(s, x, 4.2, w_, 0.7, f, [(t, 11, True, c)], adj=0.2)
for i, (a, b, c) in enumerate(ROWS):
    y = 5.15 + i * 2.35
    box(s, hx[0], y, hw[0], 2.0, "FFFFFF", [(a, 10.5, False, INK)], line=LGRAY, align=PP_ALIGN.LEFT)
    box(s, hx[1], y, hw[1], 2.0, "FFFFFF", [(b, 10.5, True, INK)], line="E0C36A", align=PP_ALIGN.LEFT)
    box(s, hx[2], y, hw[2], 2.0, NAVY, [(c, 11, True, "FFFFFF")])
    arrow(s, hx[0] + hw[0] + 0.05, y + 1.0, hx[1] - 0.05, y + 1.0, color=NAVY, w=2)
    arrow(s, hx[1] + hw[1] + 0.05, y + 1.0, hx[2] - 0.05, y + 1.0, color=NAVY, w=2)
box(s, 1.2, 14.7, 25.12, 1.3, BEIGE, [("診断と資料をきっかけにLINEの友だちになってもらい、比べている間ずっと情報を届ける。", 12.5, True, INK)])

# ---- P9 各社LINEの今（旧11を作り替え：中身は調査結果で埋める）
s = rebuild(S(11), "各社LINEの今",
            ["大手・中堅の塾は、LINEで保護者の迷いにどこまで応えているか。",
             "大手2社（個別教室のトライ・明光義塾）と中堅2社（森塾・ナビ個別指導学院）を実際に確認した。"],
            "出典｜各社LINE公式アカウント（page.line.me・実測）・公式サイト・LINEヤフー for Business 導入事例（明光ネットワークジャパン 2024）2026-09-25取得。「確認できず」＝公開情報で確認できなかったもの。★競合の実名掲載は上司確認")
P9 = s
# 調査結果：_data/lineoa_competitors/教育塾_4社LINE_20260925.md
COLS = [("", 3.9), ("本部アカウント\n（友だち数）", 4.3), ("教室別の\nアカウント", 2.9), ("誰向け", 2.3),
        ("診断", 3.9), ("LINEで資料請求\n・体験予約", 4.1), ("チャット相談", 3.72)]
ROWS9 = [("明光義塾", "大手", "明光義塾\n308,344人", "あり（多数）\n例：623人", "保護者",
          "あり\nタイプ別の勉強法・相性チェック", "確認できず\n（教室は電話・Web・LINEで予約）", "あり\nAIが回答"),
         ("個別教室のトライ", "大手", "トライさん\n120,131人", "確認できず", "生徒\n（中高生）",
          "Webの性格診断あり\nLINEでは確認できず", "確認できず\n（Webフォーム）", "あり\nLINE相談窓口"),
         ("森塾", "中堅", "運営会社名義\n（スプリックス）\n2,043,233人", "見つからず", "確認できず",
          "確認できず", "サイトに記載なし\n（Webフォームのみ）", "確認できず"),
         ("ナビ個別指導学院", "中堅", "見つからず", "見つからず", "―", "―", "サイトに記載なし\n（Webフォームのみ）", "―")]
x = 1.2
xs = []
for t, w_ in COLS:
    xs.append((x, w_))
    if t:
        box(s, x, 4.15, w_ - 0.06, 1.15, NAVY, [(ln, 9.5, True, "FFFFFF") for ln in t.split("\n")], shape=MSO_SHAPE.RECTANGLE, ml=0.05)
    x += w_
for r, row in enumerate(ROWS9):
    y = 5.38 + r * 1.95
    f = "FFFFFF" if r % 2 == 0 else "F7F9FC"
    name, tier = row[0], row[1]
    box(s, xs[0][0], y, xs[0][1] - 0.06, 1.87, PALE, [(tier, 8.5, True, GRAY), (name, 10.5, True, NAVY)], shape=MSO_SHAPE.RECTANGLE, line="E7E6E6", ml=0.05)
    for c, val in enumerate(row[2:]):
        cx_, cw_ = xs[c + 1]
        miss = val.startswith("確認できず") or val.startswith("見つからず") or val.startswith("サイトに記載なし") or val == "―"
        good = val.startswith("あり") or val.startswith("保護者")
        col = GRAY if miss else (GREEN if good else INK)
        lines = val.split("\n")
        box(s, cx_, y, cw_ - 0.06, 1.87, f, [(lines[0], 9.5, True, col)] + [(ln, 8.5, False, GRAY if miss else INK) for ln in lines[1:]],
            shape=MSO_SHAPE.RECTANGLE, line="E7E6E6", ml=0.05)
box(s, 1.2, 13.3, 12.3, 2.75, PALE, [("明光義塾はLINEで成果を出している（LINEヤフー公式事例）", 10.5, True, NAVY),
                                      ("・サイトの離脱時ポップアップから友だち追加", 9.5, False, INK),
                                      ("・約5カ月で約20万人の友だち（うち約15万人はスタンプの二次拡散）", 9.5, False, INK),
                                      ("・資料請求が前年比120％ほど", 9.5, False, INK)],
    line="8EA9DB", align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
box(s, 14.02, 13.3, 12.3, 2.75, BEIGE, [("空いているところ", 10.5, True, INK),
                                         ("・保護者向けに「診断 → 資料請求 → 体験予約」を", 9.5, False, INK),
                                         ("　LINEの中でつないでいる塾は、確認できなかった", 9.5, False, INK),
                                         ("・トライは生徒向け。森塾・ナビは公式サイトにLINEの導線がない", 9.5, False, INK)],
    line="E0C36A", align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
box(s, 1.2, 16.25, 25.12, 0.95, NAVY, [("保護者向けに、診断から資料請求・体験予約までをLINEでつなぐ。ここはまだ空いている。", 11.5, True, "FFFFFF")])

# ---- P10 LINEでできること（旧5を作り替え）
s = rebuild(S(5), "LINEでできること",
            ["保護者の「迷っている期間」に、LINEでできることは4つ。", "授業の質や合格実績はLINEではつくれない。つくれるのは、接点と、体験・入会までの流れ。"],
            None, phases="all")
cx, cy = 13.76, 10.1
px_, py_, pw_, ph_ = phone(s, cx - 2.3, 5.0, 4.6, 8.4)
bubble(s, px_ + 0.15, py_ + 0.25, pw_ - 0.4, 1.3, ["30秒で、お子さまに", "合う学び方が分かります"], size=7.5)
box(s, px_ + 0.2, py_ + 1.8, pw_ - 0.6, 0.6, LINE_GREEN, [("診断をはじめる", 8, True, "FFFFFF")], adj=0.3)
bubble(s, px_ + 0.15, py_ + 2.7, pw_ - 0.4, 1.1, ["コースと料金の資料を", "お送りしました"], size=7.5)
bubble(s, px_ + 0.15, py_ + 4.0, pw_ - 0.4, 1.1, ["明日17時、体験授業で", "お待ちしています"], size=7.5)
FOUR = [(1.2, 4.3, "① タイプ別の診断", "「うちの子に合うのはどれ？」に、申し込む前に答える", "④リード獲得"),
        (18.1, 4.3, "② LINEで資料請求", "電話も来店もなしで、コースと料金の資料が届く", "④リード獲得（CV①）"),
        (1.2, 10.3, "③ 14日間の配信", "迷っている間に役立つ情報を届けて、体験へ", "③育成・⑤リード有効化"),
        (18.1, 10.3, "④ 体験前のリマインド", "日時・持ち物・変更方法を前日に届けて、来てもらう", "⑤リード有効化 → ⑦入会")]
for x, y, t, sub, ph in FOUR:
    box(s, x, y, 8.2, 4.6, "FFFFFF", [], line="8EA9DB", lw=1.5, adj=0.1)
    box(s, x, y, 8.2, 1.0, NAVY, [(t, 12.5, True, "FFFFFF")], adj=0.15)
    text(s, x + 0.35, y + 1.25, 7.5, 2.0, [(sub, 11.5, False, INK)])
    box(s, x + 0.35, y + 3.5, 7.5, 0.7, PALE, [(ph, 9.5, True, NAVY)], adj=0.2)
arrow(s, 9.45, 6.6, cx - 2.35, 7.6, color="8EA9DB")
arrow(s, 18.05, 6.6, cx + 2.35, 7.6, color="8EA9DB")
arrow(s, 9.45, 12.6, cx - 2.35, 11.6, color="8EA9DB")
arrow(s, 18.05, 12.6, cx + 2.35, 11.6, color="8EA9DB")
box(s, 1.2, 15.35, 25.12, 1.0, "EEEEEE", [("LINEではできないこと：授業の質を上げる／合格実績をつくる／講師を採用・育成する", 10.5, False, GRAY)])

# ---- P11 全体設計（旧12を作り替え）
s = rebuild(S(12), "全体設計",
            ["軽い申し込み（資料請求）で先につながり、体験予約、入会へ進んでもらう。",
             "入会は教室で行うため、LINEで計測するのは資料請求（CV①）と体験予約（CV②）の2つ。"],
            None, phases="all")
DESC = ["広告・チラシ\n・口コミ", "帰る保護者を\n離脱防止でLINEへ", "比べている間\n情報を届ける", "30秒診断 →\n資料をLINEで", "資料を読む人へ\n追いの配信→体験予約",
        "体験に来なかった・\n入会しなかった人へ", "教室で入会\n（LINEでは計測しない）"]
w = (25.12 - 0.12 * 6) / 7
for i, (ph, d) in enumerate(zip(PHASES[:7], DESC)):
    x = 1.2 + i * (w + 0.12)
    f, c = ("EEEEEE", GRAY) if i == 7 else (NAVY, "FFFFFF")
    chevron(s, x, 4.4, w, 1.0, f, ph, color=c, size=9.5, first=(i == 0))
    box(s, x + 0.05, 5.6, w - 0.1, 2.4, "FFFFFF",
        [(t, 11, False, INK) for t in d.split("\n")], line=LGRAY, ml=0.08)
CV = [(3, "CV① 資料請求", GREEN), (4, "CV② 体験予約", GREEN), (6, "入会（契約）", NAVY)]
for i, t, col in CV:
    x = 1.2 + i * (w + 0.12)
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x + w / 2), Cm(7.95), Cm(x + w / 2), Cm(8.9))
    c.line.color.rgb = RGBColor.from_string(col)
    c.line.width = Pt(2)
    a_, b_ = (t.split(" ", 1) + [""])[:2] if " " in t else ("入会", "（契約）")
    box(s, x + 0.05, 8.9, w - 0.1, 1.2, col, [(a_, 10, True, "FFFFFF"), (b_, 10, True, "FFFFFF")], adj=0.2, ml=0.05)
x4 = 1.2 + 3 * (w + 0.12)
text(s, x4 - 0.3, 10.12, 2 * w + 0.72, 0.45, [("↑ この間に「読んで・比べて・迷う」期間がある ↑", 8.5, True, ORANGE)], align=PP_ALIGN.CENTER)
box(s, 1.2, 10.6, 25.12, 4.1, PALE, [], line=LGRAY)
text(s, 1.6, 10.75, 24.3, 0.8, [("CVを2段に分ける理由", 14, True, TNAVY)])
text(s, 1.6, 11.7, 11.8, 3.0, [("CV① 資料請求（軽い）", 13, True, GREEN),
                                ("・まだ決めていない保護者でも、気軽に申し込める", 12, False, INK),
                                ("・学年や悩みが分かり、その後の配信を合わせられる", 12, False, INK)])
text(s, 13.9, 11.7, 12.0, 3.0, [("CV② 体験予約（重い）", 13, True, GREEN),
                                 ("・資料を読んで「合いそう」と思った保護者が進む", 12, False, INK),
                                 ("・体験に来てもらえれば、入会は教室で決まる", 12, False, INK)])
box(s, 1.2, 15.1, 25.12, 1.3, NAVY, [("いきなり体験を迫らず、資料請求 → 体験予約 → 入会の順に進んでもらう。", 14, True, "FFFFFF")])

# ---- P12 導線（旧26を作り替え）
s = rebuild(S(26), "導線",
            ["電話や資料請求フォームに「ハードルが高い」と感じる保護者がいる。",
             "LINEなら、名前も電話番号も入れずに始められる。"],
            "※離脱防止ポップアップはSitelead等（別途費用）", phases=[2, 4, 5, 7])
FLOW = [("サイトから\n帰ろうとする", "EEEEEE", INK), ("離脱防止\nポップアップ", "FFFFFF", NAVY), ("LINE\n友だち追加", "FFFFFF", NAVY),
        ("30秒\n診断", "FFFFFF", NAVY), ("資料請求\n（CV①）", GREEN, "FFFFFF"), ("体験予約\n（CV②）", GREEN, "FFFFFF"), ("入会\n（契約）", NAVY, "FFFFFF")]
w, g = 3.2, 0.45
for i, (t, f, c) in enumerate(FLOW):
    x = 1.2 + i * (w + g)
    box(s, x, 5.0, w, 2.4, f, [(ln, 11.5, True, c) for ln in t.split("\n")], line=None if f != "FFFFFF" else "8EA9DB", lw=1.5)
    if i < len(FLOW) - 1:
        arrow(s, x + w + 0.03, 6.2, x + w + g - 0.03, 6.2, color=NAVY, w=2.25)
SUB = ["", "「まだ決めなくて\n大丈夫です」", "名前・電話番号\nは不要", "学年・悩み・性格\nいつから", "コース・料金の\n資料が届く", "LINEで日時を\n選ぶだけ", "教室で\n入会手続き"]
for i, t in enumerate(SUB):
    if t:
        text(s, 1.2 + i * (w + g), 7.6, w, 1.4, [(ln, 9.5, False, GRAY) for ln in t.split("\n")], align=PP_ALIGN.CENTER)
box(s, 1.2, 9.6, 12.3, 4.6, "FFFFFF", [], line=LGRAY)
text(s, 1.6, 9.75, 11.6, 4.4, [("いま", 12, True, GRAY),
                                ("サイト → 資料請求フォーム／電話 → 体験", 10.5, False, INK),
                                ("・フォームに名前・住所・電話番号を入れる", 10.5, False, INK),
                                ("・電話がかかってくるかもしれない", 10.5, False, INK),
                                ("→ 決めていない保護者は、ここで帰る", 10.5, True, RED)])
box(s, 14.02, 9.6, 12.3, 4.6, PALE, [], line="8EA9DB")
text(s, 14.4, 9.75, 11.6, 4.4, [("LINEを入れると", 12, True, NAVY),
                                 ("サイト → LINE → 診断 → 資料 → 体験", 10.5, False, INK),
                                 ("・タップ1回で友だち追加。入力はいらない", 10.5, False, INK),
                                 ("・資料を読んでから、体験を決められる", 10.5, False, INK),
                                 ("→ 決めていない保護者とも、つながっておける", 10.5, True, NAVY)])
box(s, 1.2, 14.6, 25.12, 1.1, NAVY, [("入口の敷居を下げて、決めていない保護者ともつながる。", 12.5, True, "FFFFFF")])

# ---- P13 施策全体像（旧13を作り替え）
s = rebuild(S(13), "施策全体像",
            ["初めの1〜2ヶ月でつくるのは、診断・リッチメニュー・14日間の配信。", "あとは毎月、企画配信と出し分け配信を回す。"],
            None, phases="all")
box(s, 1.2, 4.2, 12.3, 0.8, NAVY, [("初めにつくるもの（1〜2ヶ月）", 12, True, "FFFFFF")], adj=0.15)
INIT = [("LINE公式アカウントの設定（未開設の場合）", "①"), ("サイトの離脱防止ポップアップ", "②"), ("あいさつメッセージ", "④"),
        ("30秒ぴったり診断（4問）", "④"), ("LINEで資料請求の仕組み", "④"), ("リッチメニュー（2タブ）", "全体"),
        ("14日間の配信（未入会の保護者向け）", "③⑤"), ("キーワード自動応答（料金・体験・持ち物）", "⑤")]
for i, (t, ph) in enumerate(INIT):
    y = 5.2 + i * 1.18
    box(s, 1.2, y, 12.3, 1.0, "FFFFFF", [(t, 10.5, False, INK)], line=LGRAY, align=PP_ALIGN.LEFT)
    box(s, 11.3, y + 0.2, 2.0, 0.6, PALE, [(ph, 9, True, NAVY)], adj=0.3)
box(s, 14.02, 4.2, 12.3, 0.8, GREEN, [("毎月まわすもの", 12, True, "FFFFFF")], adj=0.15)
MON = [("企画配信（月1〜2本・年間カレンダーに沿って）", "③"), ("診断結果に合わせた出し分け配信", "⑤⑥"),
       ("体験予約者への前日リマインド（自動）", "⑤⑦"), ("過去に問い合わせた保護者への再案内", "⑥"), ("月次レポートと定例会", "全体")]
for i, (t, ph) in enumerate(MON):
    y = 5.2 + i * 1.18
    box(s, 14.02, y, 12.3, 1.0, "FFFFFF", [(t, 10.5, False, INK)], line=LGRAY, align=PP_ALIGN.LEFT)
    box(s, 24.1, y + 0.2, 2.0, 0.6, PALE, [(ph, 9, True, NAVY)], adj=0.3)
box(s, 14.02, 11.3, 12.3, 3.3, BEIGE, [("作り込むのは、診断・リッチメニュー・14日間の配信の3つ。", 11, True, INK),
                                         ("ここが弱いと、あとの配信が全部弱くなる。", 11, False, INK)])
text(s, 1.2, 14.8, 25.12, 0.6, [("右の丸は、8フェーズのどこに効くか（番号）", 9, False, GRAY)])

# ---- P14 構築① 友だち追加の入口（旧15を作り替え）
s = rebuild(S(15), "構築① 友だち追加の入口",
            ["決断を迫らず、「資料と診断を受け取るだけ」でつながる入口をつくる。", "ボタンは「今すぐ体験」ではなく「受け取る」。"],
            "※ボタン文言の言い換えはDYM見解。離脱防止ポップアップはSitelead等（別途費用）", phases=[1, 2, 4])
logic3(s, ["サイトに来た保護者の多くは、名前を残さず帰る（P4）", "一番の悩みは「合うか分からない」63.5%（P5）"],
       ["決めていない保護者でも、", "気軽につながれる入口をつくる"],
       ["入口は3つ：", "離脱防止ポップアップ／チラシ・教室のQR／友だちからの紹介"])
text(s, 1.2, 7.25, 12.5, 0.6, [("入口は3つ", 11.5, True, TNAVY)])
ENT = [("サイトの離脱防止ポップアップ", "帰ろうとした瞬間に表示。広告費を足さずにつながれる"),
       ("チラシ・教室掲示のQR", "今あるチラシにQRを1つ足すだけ"),
       ("友だちからの紹介（入会前）", "「一緒に体験しよう」をLINEで送ってもらう（P15）")]
for i, (t, sub) in enumerate(ENT):
    y = 7.95 + i * 2.2
    box(s, 1.2, y, 12.5, 1.95, "FFFFFF", [(t, 11, True, NAVY), (sub, 10, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
text(s, 14.3, 7.25, 12, 0.6, [("離脱防止ポップアップの実文", 11.5, True, TNAVY)])
box(s, 14.3, 7.95, 12.0, 6.35, "FFFFFF", [], line=LGRAY, lw=1.25)
text(s, 14.8, 8.2, 11.0, 4.3, [("まだ、塾を決めなくて大丈夫です。", 14, True, TNAVY),
                                ("30秒の診断で、お子さまに合う学び方（集団／個別／映像）と、", 10.5, False, INK),
                                ("コース・料金の資料をLINEでお届けします。", 10.5, False, INK)])
box(s, 15.8, 11.6, 9.0, 1.1, LINE_GREEN, [("LINEで資料と診断を受け取る", 12, True, "FFFFFF")], adj=0.4)
text(s, 14.8, 12.9, 11.0, 0.6, [("※お名前・お電話番号は不要です", 9.5, False, GRAY)], align=PP_ALIGN.CENTER)

# ---- P15 友だちと一緒に通うメリット（旧22を作り替え）
s = rebuild(S(22), "構築① つづき：友だちと一緒に",
            ["「友だちと一緒なら行ってみたい」を、体験予約のきっかけにする。", "保護者どうしの「どこの塾に行ってる？」を、LINEでつなぐ。"],
            "出典｜POPER「Comiru」保護者と学習塾の意識調査 2022年（保護者300人）。特典の中身は塾ごとに決める", phases=[1, 4, 5])
logic3(s, ["塾を知ったきっかけの1位は", "「知人・友人の口コミ」47.0%"],
       ["友だちとの会話を、", "体験予約のきっかけにする"],
       ["友だちと一緒の体験枠をLINEで予約", "LINEの転送だけで紹介できる"])
text(s, 1.2, 7.25, 12.5, 0.6, [("友だちと一緒に通うメリット（保護者に伝えること）", 11.5, True, TNAVY)])
MER = [("📚", "通う習慣がつきやすい", "一緒に行く友だちがいると、休みにくくなる"),
       ("💪", "励まし合える", "テスト前も「一緒に頑張ろう」がある"),
       ("🚗", "送り迎えを分担できる", "保護者どうしで助け合える")]
for i, (ic, t, sub) in enumerate(MER):
    y = 7.95 + i * 2.2
    box(s, 1.2, y, 1.9, 1.95, PALE, [(ic, 18, False, INK)], line="8EA9DB")
    box(s, 3.3, y, 10.4, 1.95, "FFFFFF", [(t, 11.5, True, NAVY), (sub, 10, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
text(s, 14.3, 7.25, 12, 0.6, [("LINEでの仕組み", 11.5, True, TNAVY)])
px_, py_, pw_, ph_ = phone(s, 14.3, 7.9, 4.8, 8.5)
bubble(s, px_ + 0.2, py_ + 0.3, pw_ - 0.5, 2.0, ["お友だちと一緒の", "体験授業も受け付けて", "います😊 2人とも", "特典があります"])
box(s, px_ + 0.2, py_ + 2.5, pw_ - 0.5, 0.65, LINE_GREEN, [("一緒に体験を予約", 8, True, "FFFFFF")], adj=0.3)
box(s, px_ + 0.2, py_ + 3.3, pw_ - 0.5, 0.65, "FFFFFF", [("友だちに送る", 8, True, LINE_GREEN)], line=LINE_GREEN, adj=0.3)
STEP = [("1", "友だちと一緒の体験枠を用意", "LINEで2人分をまとめて予約できる"),
        ("2", "2人とも特典", "例：体験後の学習相談、教材プレゼント"),
        ("3", "LINEの転送で紹介", "トークを友だちの保護者に送るだけ")]
for i, (n, t, sub) in enumerate(STEP):
    y = 7.95 + i * 2.2
    box(s, 19.6, y + 0.45, 1.0, 1.0, NAVY, [(n, 12, True, "FFFFFF")], shape=MSO_SHAPE.OVAL)
    text(s, 20.8, y + 0.2, 5.6, 1.7, [(t, 11, True, NAVY), (sub, 9.5, False, INK)])

# ---- P16 構築② あいさつメッセージ（旧16を作り替え）
s = rebuild(S(16), "構築② あいさつメッセージ",
            ["友だち追加の直後は、一番メッセージが読まれるタイミング。", "最初の1通で安心できることを先に伝え、診断へ進んでもらう。"],
            "文面の教室長名・対応時間・配信頻度は貴社の運用に合わせて差し替える", phases=[4])
logic3(s, ["友だち追加の直後は、", "一番メッセージが読まれる"],
       ["安心できることを先に伝えて、", "診断へ進んでもらう"],
       ["あいさつメッセージの", "組み立てを決めておく"])
text(s, 1.2, 7.25, 13, 0.6, [("メッセージの組み立て", 11.5, True, TNAVY)])
GRT = [("教室長の名前を出す", "人の顔が見えると安心できる"), ("返信できる時間を書く", "例：平日14〜21時"),
       ("届く頻度を書く", "例：週1回"), ("営業のお電話はしないと書く", "連絡先を渡す不安をなくす"), ("30秒診断へのボタン", "そのまま診断へ進める")]
for i, (t, sub) in enumerate(GRT):
    y = 7.95 + i * 1.62
    box(s, 1.2, y + 0.25, 0.9, 0.9, NAVY, [(str(i + 1), 11, True, "FFFFFF")], shape=MSO_SHAPE.OVAL)
    box(s, 2.3, y, 11.4, 1.42, "FFFFFF", [(t, 11, True, NAVY), (sub, 9.5, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
px_, py_, pw_, ph_ = phone(s, 16.8, 7.25, 7.2, 9.6)
bubble(s, px_ + 0.25, py_ + 0.25, pw_ - 0.9, 6.3,
       ["はじめまして、◯◯塾 △△教室の", "教室長、佐藤です😊", "友だち追加ありがとうございます。", "",
        "このLINEでは、お子さまの勉強に", "役立つ情報を週1回お届けします。", "営業のお電話は一切しません。",
        "ご質問は、平日14〜21時に", "このトークでお返事します。", "", "まずは30秒の診断で、お子さまに", "合う学び方とコース・料金の資料を", "お受け取りください👇"], size=8)
box(s, px_ + 0.25, py_ + 6.75, pw_ - 0.9, 0.75, LINE_GREEN, [("30秒診断をはじめる", 9, True, "FFFFFF")], adj=0.3)

# ---- P17 構築③ 30秒ぴったり診断（旧17を作り替え）
s = rebuild(S(17), "構築③ 30秒ぴったり診断",
            ["申し込む前に「集団・個別・映像のどれが合うか」を見せて、納得してから体験へ。", "4問すべて選択式。回答はそのまま配信の出し分けに使う。"],
            "※競合の診断（トライの性格別学習法診断など）はP9の調査結果を参照", phases=[4, 5])
logic3(s, ["塾選びで一番困ったこと", "「うちの子に合う塾が分からない」63.5%"],
       ["申し込む前に、", "「合う・合わない」を見せる"],
       ["30秒診断 → 資料請求（CV①）", "→ 体験予約（CV②）"])
Q = [("Q1 お子さまの学年は？", "小1〜4／小5・6／中1・2／中3／高1・2／高3", "学年"),
     ("Q2 いちばん近いお悩みは？", "成績が下がってきた／やる気が出ない／受験対策を始めたい／授業についていけない", "悩み"),
     ("Q3 お子さまの性格に近いのは？", "競争があると燃える／マイペースにコツコツ／質問するのが苦手／集中が続かない", "性格"),
     ("Q4 塾はいつから？", "今すぐ／今学期中に／次の学年から／情報収集だけ", "時期")]
for i, (q, a, tag) in enumerate(Q):
    y = 7.3 + i * 2.1
    box(s, 1.2, y, 13.2, 1.85, "FFFFFF", [(q, 10.5, True, NAVY), (a, 9, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
    box(s, 14.6, y + 0.55, 1.8, 0.75, PALE, [(tag, 9, True, NAVY)], adj=0.3)
text(s, 17.2, 7.25, 9, 0.6, [("診断結果の流れ", 11.5, True, TNAVY)])
RES = [("共感", "「◯◯タイプのお子さまは、こんな所でつまずきやすい」", "EEEEEE", INK),
       ("合う学び方", "集団／個別／映像のどれが合うか", "EEEEEE", INK),
       ("同じタイプの事例", "同じ悩みだった子の変化", "EEEEEE", INK),
       ("資料を受け取る", "CV① 資料請求", GREEN, "FFFFFF"),
       ("無料体験を予約する", "CV② 体験予約", GREEN, "FFFFFF")]
for i, (t, sub, f, c) in enumerate(RES):
    y = 7.95 + i * 1.62
    box(s, 17.2, y, 9.1, 1.35, f, [(t, 10.5, True, c), (sub, 9, False, c)])
    if i < len(RES) - 1:
        arrow(s, 21.75, y + 1.36, 21.75, y + 1.6, color=NAVY)

# ---- P18 構築④ リッチメニュー（旧18を作り替え）
s = rebuild(S(18), "構築④ リッチメニュー",
            ["スマホの親指が一番届きやすいのは、画面の右下。", "一番押してほしい行動（無料体験・教室見学の予約）を右下に置く。"],
            "※体験予約済みの方には、タブ②（体験当日の流れ・アクセス・日程変更など）を出し分ける", phases="all")
logic3(s, ["スマホを片手で持つと、", "親指が一番届くのは右下"],
       ["一番押してほしい行動を、", "右下に置く"],
       ["右下＝無料体験・教室見学の予約", "（体験予約済みの方は日程の確認・変更）"])
px_, py_, pw_, ph_ = phone(s, 1.2, 7.2, 8.0, 9.9)
bubble(s, px_ + 0.25, py_ + 0.3, pw_ - 1.2, 1.2, ["診断結果とコース・料金の", "資料をお送りしました📄"])
MENU = [("🎯", "30秒\nぴったり診断"), ("💰", "コース・料金"), ("🏫", "教室・講師\n紹介"),
        ("📄", "資料を\n受け取る"), ("👫", "友だちと\n一緒に体験"), ("📅", "無料体験・\n教室見学の予約")]
mx, my, mw = px_, py_ + ph_ - 4.4, pw_
cw_, ch_ = mw / 3, 2.2
for i, (ic, t) in enumerate(MENU):
    x = mx + (i % 3) * cw_
    y = my + (i // 3) * ch_
    last = i == 5
    box(s, x, y, cw_, ch_, RED if last else "FFFFFF", [(ic, 13, False, INK)] + [(ln, 8, True, "FFFFFF" if last else NAVY) for ln in t.split("\n")],
        line="BFBFBF", shape=MSO_SHAPE.RECTANGLE, ml=0.05)
box(s, mx + cw_ * 1.2, my - 0.75, cw_ * 1.8, 0.6, ORANGE, [("右下＝一番押しやすい ↘", 8, True, "FFFFFF")], adj=0.3, ml=0.05)
text(s, 10.2, 7.25, 16, 0.6, [("ボタンの並び（タブ① はじめての方）", 11.5, True, TNAVY)])
GRID = [["🎯 30秒ぴったり診断", "💰 コース・料金", "🏫 教室・講師紹介"],
        ["📄 資料を受け取る（CV①）", "👫 友だちと一緒に体験", "📅 無料体験・教室見学の予約（CV②）"]]
for r, row in enumerate(GRID):
    for c, t in enumerate(row):
        last = r == 1 and c == 2
        box(s, 10.2 + c * 5.4, 7.95 + r * 1.55, 5.25, 1.4, RED if last else "FFFFFF", [(t, 10, True, "FFFFFF" if last else NAVY)],
            line="8EA9DB", shape=MSO_SHAPE.RECTANGLE)
text(s, 10.2, 11.3, 16, 0.6, [("置き方の考え方", 11.5, True, TNAVY)])
text(s, 10.2, 11.95, 16.1, 3.4, [("・右下：一番押してほしい行動。体験・見学の予約（CV②）", 10.5, False, INK),
                                  ("・左下：資料を受け取る（CV①）。体験はまだ早い保護者の受け皿", 10.5, False, INK),
                                  ("・上段：比べるときに見たい情報（診断・料金・教室）", 10.5, False, INK),
                                  ("・体験予約済みの方には、タブ②に切り替えて右下を「日程の確認・変更」にする", 10.5, False, INK)])
box(s, 10.2, 15.5, 16.1, 1.2, PALE, [("タブ②（体験予約済みの方）：体験当日の流れ／アクセス・持ち物／担当講師の紹介／保護者の声／LINEで質問／日程の確認・変更（右下）", 9.5, False, NAVY)], align=PP_ALIGN.LEFT)

# ---- P19 構築⑤ 配信の出し分け（旧19を作り替え）
s = rebuild(S(19), "構築⑤ 配信の出し分け",
            ["同じ内容を全員に送ると、合わない人にはブロックされる。", "診断の回答と、LINEでの行動で、保護者ごとに「うちの子の話」になる配信にする。"],
            None, phases=[5, 6])
logic3(s, ["同じ内容を全員に送ると、", "合わない人にはブロックされる"],
       ["保護者ごとに、", "「うちの子の話」になる配信にする"],
       ["診断タグ＋行動タグで", "配信を出し分ける"])
TAGS = [("① 診断タグ（友だち追加の直後）", ["学年", "悩み", "性格（集団／個別の判定）", "いつから（温度）"]),
        ("② 行動タグ（LINEでの行動）", ["どのボタンを押したか", "資料請求したか", "体験予約したか・来たか", "配信を開いているか"])]
for i, (t, items) in enumerate(TAGS):
    x = 1.2 + i * 8.52
    box(s, x, 7.3, 8.08, 0.8, NAVY, [(t, 11, True, "FFFFFF")], adj=0.15)
    for j, it in enumerate(items):
        box(s, x, 8.3 + j * 1.25, 8.08, 1.05, "FFFFFF", [(it, 10.5, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
box(s, 18.24, 7.3, 8.08, 0.8, GREEN, [("何ができるようになるか", 11, True, "FFFFFF")], adj=0.15)
EX = ["中3で「受験対策」→ 高校受験の情報を中心に", "「質問するのが苦手」→ 個別指導の事例を中心に", "資料請求したが予約なし → 体験の案内を追加で",
      "「情報収集だけ」→ 売り込みは控えて週1の情報だけ"]
for j, it in enumerate(EX):
    box(s, 18.24, 8.3 + j * 1.25, 8.08, 1.05, PALE, [(it, 9.5, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)

# ---- P20 14日間の配信（旧20を作り替え）
s = rebuild(S(20), "配信設計：14日間の配信",
            ["比べている間は、売り込みではなく役に立つ情報を届け、体験予約に進んでもらう。", "14日のうち、体験の案内は2通だけ（7日後・14日後）。"],
            "※保護者が比べている間に何を調べているかは、前後検索（取得中）で確定させ、配信の中身を合わせる", phases=[3, 5, 6])
logic3(s, ["保護者は比べている間に、", "費用・時期・口コミを調べている", "（前後検索で確定させる）"],
       ["売り込みより先に、", "役に立つ情報を届ける"],
       ["未入会の保護者向けに", "14日間の配信を自動で送る"])
TL = [("登録した日", "友だち追加の直後", "あいさつ＋30秒診断", False),
      ("3日後", "比べ始めたころ", "家でできる勉強のコツ", False),
      ("5日後", "費用が気になるころ", "月謝＋講習費の年間の目安", False),
      ("7日後", "候補を絞るころ", "同じタイプの子の事例＋体験のご案内①", True),
      ("10日後", "まだ予約していない人", "よくある質問（断ってもいい？総額は？）", False),
      ("14日後", "決めるころ", "次の定期テストからの逆算＋体験のご案内②", True)]
c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(1.6), Cm(8.2), Cm(25.9), Cm(8.2))
c.line.color.rgb = RGBColor.from_string(NAVY)
c.line.width = Pt(3)
w = 25.12 / 6
for i, (d, scene, body, cv) in enumerate(TL):
    x = 1.2 + i * w
    col = GREEN if cv else NAVY
    box(s, x + w / 2 - 0.35, 7.85, 0.7, 0.7, col, [], shape=MSO_SHAPE.OVAL)
    text(s, x, 7.1, w, 0.65, [(d, 11, True, col)], align=PP_ALIGN.CENTER)
    box(s, x + 0.1, 8.85, w - 0.2, 1.0, "EEEEEE", [(scene, 9.5, True, INK)], adj=0.15)
    box(s, x + 0.1, 10.0, w - 0.2, 3.0, col if cv else "FFFFFF", [(body, 10, True, "FFFFFF" if cv else INK)],
        line=None if cv else "8EA9DB")
box(s, 1.2, 13.5, 12.3, 2.3, PALE, [("体験予約した人", 11, True, GREEN), ("→ 前日リマインドへ（P22）", 10.5, False, INK)], line="8EA9DB")
box(s, 14.02, 13.5, 12.3, 2.3, PALE, [("予約しなかった人", 11, True, NAVY), ("→ 週1回の配信へ合流し、講習の時期に再案内（⑥再育成）", 10.5, False, INK)], line="8EA9DB")

# ---- P21 実文面（旧21を作り替え）
s = rebuild(S(21), "実文面：場面ごとのメッセージ",
            ["メッセージごとに「いつ・どの場面で送るか」を決めておく。", "通知の冒頭（最初の15文字ほど）で開かれるかが決まるので、冒頭も1通ずつ決める。"],
            "※事例の数字は貴社の実績に差し替え、学年・期間・通った回数を書く。体験予約の前日のメッセージはP22", phases=[3, 5])
MSG = [("登録から3日後", "比べ始めたころ", "テスト後、9割のご家庭がしない事",
        ["◯◯塾の佐藤です📚", "塾を探す前に、今日からできることを1つ", "だけ。", "", "伸びる子の家庭は、テスト返却の日に", "これをやっています👇",
         "✅ 間違いを「ケアレスミス」と", "　「わからなかった」に仕分け", "✅ 「わからなかった」だけ付箋", "✅ 週末に付箋の問題だけ解き直す", "",
         "塾はその先、「付箋が多すぎて手に負えな", "い」ときの選択肢で大丈夫です😊"], None),
       ("登録から7日後", "候補を絞るころ", "同じ悩みだった◯年生の話",
        ["◯◯塾の佐藤です。", "診断で「◯◯タイプ」だったお子さまと", "同じタイプの、中2のAさんのお話です。", "",
         "入塾前は「授業についていけない」が", "悩みでしたが、分からない所まで戻った", "ことで、3ヶ月後の定期テストで", "数学が◯点→◯点に。", "",
         "お子さまに合うかどうかは、", "体験で確かめられます。"], "無料体験の空き枠を見る"),
       ("登録から14日後", "決めるころ", "次のテストまで、あと◯日です",
        ["◯◯塾の佐藤です。", "次の定期テストまで、あと約◯日。", "", "入塾手続きから授業開始まで約1週間、", "テスト対策には3週間ほどかかるため、",
         "ご検討中なら今週がひとつの分かれ目", "です。", "", "「まだ迷っている」段階でも大丈夫。", "体験だけ受けて、テスト後にゆっくり", "決めたご家庭も多いです😊"], "今週の体験枠を見る")]
for i, (when, scene, pre, body, btn) in enumerate(MSG):
    x = 1.2 + i * 8.52
    box(s, x, 4.15, 8.08, 1.2, NAVY, [(f"【{when}】", 11, True, "FFFFFF"), (scene, 10, False, "FFFFFF")], adj=0.12)
    box(s, x, 5.5, 8.08, 0.75, BEIGE, [(f"通知の冒頭：{pre}", 9.5, True, INK)], adj=0.15, align=PP_ALIGN.LEFT)
    box(s, x, 6.4, 8.08, 10.7, "C9D6E8", [], shape=MSO_SHAPE.RECTANGLE)
    bh = 0.4 * len(body) + 0.3
    bubble(s, x + 0.3, 6.65, 7.4, bh, body, size=8.5)
    if btn:
        box(s, x + 0.3, 6.65 + bh + 0.2, 7.4, 0.75, LINE_GREEN, [(f"▶ {btn}", 9, True, "FFFFFF")], adj=0.3)

# ---- P22 体験予約から入会まで（旧23を作り替え）
s = rebuild(S(23), "体験予約から入会まで",
            ["体験を予約しても、当日来ない人がいる。理由の多くは「忘れた」「持ち物が不安」「変更の仕方が分からない」。",
             "前日までに先回りして、体験に来てもらい、入会につなげる。"],
            "※体験に来ない割合・実施率の改善幅は、貴社の実績で確定させる（旧版の「約20%」「80%→92%」は出典未確認のため載せていない）",
            phases=[5, 7])
logic3(s, ["体験を予約しても、", "当日来ない人がいる", "（割合は貴社の実績で確定）"],
       ["来ない理由を、", "前日までに先回りしてなくす"],
       ["予約の直後から翌日まで、", "LINEで5回自動で届ける"])
TL2 = [("予約した直後", "予約内容の確認", "カレンダー登録"), ("2日前", "担当講師の紹介", "顔写真とひとこと"),
       ("前日", "持ち物・アクセス", "日程の変更方法"), ("当日2時間前", "「お待ちしています」", "短い一言"), ("翌日", "お礼＋診断結果", "入会のご相談はこちら")]
w = 25.12 / 5
c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(1.6), Cm(8.0), Cm(25.9), Cm(8.0))
c.line.color.rgb = RGBColor.from_string(NAVY)
c.line.width = Pt(3)
for i, (d, t, sub) in enumerate(TL2):
    x = 1.2 + i * w
    col = GREEN if i == 4 else NAVY
    box(s, x + w / 2 - 0.35, 7.65, 0.7, 0.7, col, [], shape=MSO_SHAPE.OVAL)
    text(s, x, 7.0, w, 0.6, [(d, 10.5, True, col)], align=PP_ALIGN.CENTER)
    box(s, x + 0.15, 8.6, w - 0.3, 1.5, "FFFFFF", [(t, 10.5, True, NAVY), (sub, 9.5, False, INK)], line="8EA9DB")
text(s, 1.2, 10.45, 12, 0.6, [("前日のメッセージ（実文）", 11.5, True, TNAVY)])
box(s, 1.2, 11.1, 14.0, 5.9, "C9D6E8", [], shape=MSO_SHAPE.RECTANGLE)
bubble(s, 1.5, 11.3, 13.4, 5.5, ["【明日◯時〜】体験授業のご確認です😊", "◯◯塾 △△教室（地図はこちら▼）", "✏️ 持ち物：筆記用具だけでOK（教材はご用意します）",
                                  "🚲 駐輪場：教室裏にあります", "👪 保護者さまは「見学」でも「面談だけ」でも大丈夫です", "",
                                  "ご都合が変わったら、このトークに「変更」と送って", "ください。別の日程をご案内します。", "（キャンセル料はありません）"], size=9)
box(s, 15.7, 11.1, 10.62, 2.8, BEIGE, [("「キャンセル料はありません」を必ず書く", 11, True, INK),
                                         ("変更・キャンセルの自由を先に伝えたほうが、", 10, False, INK), ("体験に来てもらいやすい", 10, False, INK)], align=PP_ALIGN.LEFT)
box(s, 15.7, 14.2, 10.62, 2.8, PALE, [("翌日のメッセージで入会へ", 11, True, NAVY),
                                        ("お礼と診断結果を送り、入会の相談は", 10, False, INK), ("LINEで受け付けて教室へつなぐ", 10, False, INK)], align=PP_ALIGN.LEFT)

# ---- P23 過去に問い合わせた保護者へ（旧25を作り替え）
s = rebuild(S(23 + 2), "過去に問い合わせた保護者へ",
            ["塾には、資料請求や体験だけで止まった保護者の連絡先が残っている。", "講習の時期に合わせて、もう一度声をかける。"],
            "通知メッセージは別途費用。利用条件・単価はLINEヤフーの規定で変わるため、要件定義で最新の条件を確認する", phases=[6])
logic3(s, ["資料請求や体験だけで止まった", "保護者の電話番号が残っている"],
       ["講習の時期に合わせて、", "もう一度声をかける"],
       ["通知メッセージ（電話番号で", "LINEに届く機能）で再案内"])
TGT = [("資料請求だけで止まっている保護者", "資料は読んだが、体験に進まなかった"), ("体験に来たが、入会しなかった保護者", "時期が合わなかった・迷っていた")]
for i, (t, sub) in enumerate(TGT):
    box(s, 1.2, 7.4 + i * 2.3, 12.3, 2.0, "FFFFFF", [(t, 11, True, NAVY), (sub, 10, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
text(s, 14.02, 7.3, 12.3, 0.6, [("送るタイミング", 11.5, True, TNAVY)])
for i, t in enumerate(["夏期講習の前", "冬期講習の前", "新学年の前"]):
    box(s, 14.02 + i * 4.2, 8.0, 3.9, 1.2, GREEN, [(t, 11, True, "FFFFFF")])
box(s, 14.02, 9.6, 12.3, 2.1, PALE, [("例：「夏期講習のご案内です。以前いただいた資料請求の", 9.5, False, INK),
                                       ("内容をもとに、お子さまの学年に合うコースをお送りします」", 9.5, False, INK)], align=PP_ALIGN.LEFT)
box(s, 1.2, 12.3, 25.12, 2.4, "FFF7F0", [("注意すること", 11, True, "C55A11"),
                                          ("・関係がはっきりしている相手だけに送る（資料請求・体験をした保護者）", 10, False, INK),
                                          ("・「知らないのに届いた」と思われない文面にする（いつ・何で接点があったかを書く）", 10, False, INK)],
    line=ORANGE, align=PP_ALIGN.LEFT)

# ---- P24 年間の配信カレンダー（旧24：中身を残して手直し）
s = S(24)
title, lead, div, foot = parts(s)
set_para_text(title, ["年間の配信カレンダー"])
set_para_text(lead, ["検索の山（2月・7月）の前から、その時期に保護者が考えていることを届ける。", "企画配信は月1〜2本。"])
shift_lead_for_band(s, [1, 3])
replace_in(s, [("「塾」ピーク2/8の26日前＝1/13から仕込み", "「塾」の検索は2月上旬が山"),
               ("夏期講習ピーク7/12の26日前＝6/16", "夏期講習の検索は7月中旬が山"),
               ("（S22 Day3）", "（P21）"), ("（S21 Day3）", "（P21）")])
for sh in s.shapes:
    if sh.has_text_frame and ("26日" in sh.text_frame.text or "仕込み" in sh.text_frame.text):
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.text = r.text.replace("26日前", "山の前").replace("仕込み開始", "配信の準備").replace("仕込み", "準備")

for sh in s.shapes:
    if sh.shape_id == 32:
        set_para_text(sh, ["2学期の定期テスト"])
    elif sh.shape_id == 33:
        set_para_text(sh, ["中間・期末テストで成績の不安が出やすい"])
    elif sh.shape_id == 34:
        set_para_text(sh, ["テスト返却週に「点数の見方」を配信"])
    elif sh.shape_id == 40:
        set_para_text(sh, ["出典｜Googleトレンド（日本・2026年年初来・週次／_data/trends/教育塾/）。各月の内容はDYM見解。★貴社の月別の入会実績があれば、この表を合わせる"])

# ---- P25 第2の提案軸：友だちと一緒に（旧33を作り替え）
s = rebuild(S(33), "第2の提案軸：友だちと一緒に",
            ["塾を知ったきっかけの1位は「知人・友人の口コミ」47.0%。塾の公式サイトより多い。",
             "保護者どうしの「どこの塾に行ってる？」を、LINEの友だち追加につなげる。"],
            "出典｜POPER「Comiru」保護者と学習塾の意識調査 2022年（保護者300人）", phases=[1, 4])
logic3(s, ["塾を知ったきっかけ", "1位「知人・友人の口コミ」47.0%", "2位「インターネット上の口コミ」28.3%"],
       ["保護者どうしの会話を、", "LINEの友だち追加につなげる"],
       ["友だちと一緒の体験枠／", "診断結果を友だちに送れる"])
text(s, 1.2, 7.3, 12.3, 0.6, [("塾を知ったきっかけ（上位2つ）", 11.5, True, TNAVY)])
for i, (t, v) in enumerate([("知人・友人の口コミ", 47.0), ("インターネット上の口コミ", 28.3)]):
    y = 8.1 + i * 1.6
    text(s, 1.2, y + 0.2, 4.8, 0.8, [(t, 10, True, INK)])
    box(s, 6.1, y + 0.1, 7.0 * v / 50, 1.0, NAVY if i == 0 else "8EA9DB", [], shape=MSO_SHAPE.RECTANGLE)
    text(s, 6.2 + 7.0 * v / 50, y + 0.2, 1.8, 0.8, [(f"{v}%", 11, True, NAVY)])
text(s, 14.02, 7.3, 12.3, 0.6, [("LINEでの仕掛け（入会前）", 11.5, True, TNAVY)])
IDEA = [("友だちと一緒の体験枠", "2人まとめてLINEで予約（P15）"), ("診断結果を友だちに送れる", "「うちはこのタイプだった」を共有しやすく"),
        ("体験に来た保護者へ案内", "「お友達も一緒にどうぞ」を翌日のメッセージで")]
for i, (t, sub) in enumerate(IDEA):
    box(s, 14.02, 8.0 + i * 1.85, 12.3, 1.6, "FFFFFF", [(t, 11, True, NAVY), (sub, 9.5, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
box(s, 1.2, 13.9, 25.12, 1.2, NAVY, [("口コミで選ばれる業界だから、口コミが生まれる場所にLINEの入口を置く。", 12.5, True, "FFFFFF")])


# ---- P26〜P29 施策の設計図（①②③④で1ページずつ。旧8・旧7・旧6＋1枚追加）
def rebuild2(s, title_t, lead_lines, foot_t, phases):
    """2_レイアウトのスライド用：上部の紺ブロック・線・タイトルを残して作り直す"""
    tree = s.shapes._spTree
    title = None
    for sh in list(s.shapes):
        top = Emu(sh.top).cm
        if top < 1.6 and Emu(sh.height).cm < 1.7:
            if sh.has_text_frame and sh.text_frame.text.strip():
                title = sh
            continue
        tree.remove(sh._element)
    set_para_text(title, [title_t])
    band(s, phases)
    text(s, 1.2, 2.2, 25.1, 1.6, [(t, 12, False, INK) for t in lead_lines], anchor=MSO_ANCHOR.MIDDLE)
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, Cm(3.86), Cm(27.52), Cm(3.86))
    c.line.color.rgb = RGBColor.from_string(LGRAY)
    text(s, 1.2, 17.35, 25.12, 0.9, [(foot_t, 7.5, False, GRAY)])
    return s


COLW = [2.8, 4.5, 7.2, 5.4, 3.3, 1.92]    # タイミング／対象（想定数）／訴求／開封×クリック×CVR／件数／費用
HEADS = ["タイミング", "対象（想定数）", "訴求", "開封率×クリック率×CVR", "件数", "費用"]


def flow_head(s, y, heads=HEADS, widths=COLW):
    x = 1.2
    for h, w_ in zip(heads, widths):
        box(s, x, y, w_ - 0.12, 0.6, "EEEEEE", [(h, 10.5, True, INK)], shape=MSO_SHAPE.RECTANGLE, ml=0.03)
        x += w_


def flow_row(s, y, vals, h=0.92, widths=COLW, cv=None):
    x = 1.2
    for i, (v, w_) in enumerate(zip(vals, widths)):
        lines = v.split("\n")
        if i == 4 and cv:
            f, c = (GREEN, "FFFFFF")
        elif i == 4:
            f, c = ("F2F2F2", GRAY)
        else:
            f, c = ("FFFFFF", INK)
        box(s, x, y, w_ - 0.12, h, f, [(lines[0], 11.5, i in (0, 4), c)] + [(ln, 9.5, False, c if i == 4 else GRAY) for ln in lines[1:]],
            line="C9D3E6", shape=MSO_SHAPE.RECTANGLE, ml=0.08, align=PP_ALIGN.CENTER if i in (0, 3, 4, 5) else PP_ALIGN.LEFT)
        if i in (1, 2, 3):
            arrow(s, x + w_ - 0.13, y + h / 2, x + w_ + 0.01, y + h / 2, color=NAVY, w=1.5)
        x += w_


def sec(s, y, t):
    box(s, 1.2, y, 25.12, 0.62, NAVY, [(t, 12, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE, align=PP_ALIGN.LEFT, ml=0.3)



def new_slide_like(src):
    """4_レイアウトで1枚追加し、タイトル・リード・区切り線・脚注を src から写す"""
    lay = [l for l in prs.slide_layouts if l.name == "4_タイトルとコンテンツ"][0]
    ns = prs.slides.add_slide(lay)
    for ph in list(ns.placeholders):
        ph._element.getparent().remove(ph._element)
    title, lead, div, foot = parts(src)
    for x in (title, lead, div, foot):
        if x is not None:
            ns.shapes._spTree.append(copy.deepcopy(x._element))
    return ns


def page_no(s, n):
    text(s, 25.6, 18.28, 1.6, 0.7, [(str(n), 12, True, "FFFFFF")], align=PP_ALIGN.RIGHT)


FOOT_MODEL = "※1教室あたりのモデル値（例）。開封率・クリック率＝DYM SIMの係数（ステップ配信 72.5%×12%／企画配信 78%×10%）。CVR＝0.3〜1.0%。件数は小数点第1位まで"

# ======== ① 友だち追加の動線 ========
s = rebuild(S(8), "施策の設計図①　友だち追加の動線",
            ["どこで、何を訴求して、LINEの友だちになってもらうか。1教室あたり月250人を置いた。",
             "一番大きいのは、サイトから帰ろうとする保護者を拾う離脱防止ポップアップ。"],
            "※友だち追加数は1教室あたりのモデル値（例）。費用は月額・初期とも税抜。離脱防止ポップアップはSitelead等", phases=[1, 2, 4])
W1 = [4.6, 10.0, 3.8, 3.4, 3.32]
flow_head(s, 4.15, ["箇所", "訴求", "LINE追加", "費用（月）", "初期費用"], W1)
ROUTES = [("00 離脱防止\nポップアップ", "「まだ決めなくて大丈夫です」\n30秒診断と資料をLINEで。営業電話なし", "200人/月", "3万円", "1.5万円"),
          ("01 完了画面\nからの誘導", "資料請求・体験予約の完了画面で\n「日程の確認・変更はLINEで」", "30人/月", "5万円", "10万円"),
          ("02 チラシ・\n教室のQR", "LINE限定\n「学年別テスト対策プリント」プレゼント", "20人/月", "費用内", "費用内")]
for i, r in enumerate(ROUTES):
    y = 4.85 + i * 1.72
    x = 1.2
    for j, (v, w_) in enumerate(zip(r, W1)):
        lines = v.split("\n")
        f, c = (LINE_GREEN, "FFFFFF") if j == 2 else ("FFFFFF", INK)
        box(s, x, y, w_ - 0.15, 1.6, f, [(lines[0], 14 if j == 2 else 12.5, j in (0, 2), c)] + [(ln, 11, j == 0, INK if j == 0 else GRAY) for ln in lines[1:]],
            line="C9D3E6", shape=MSO_SHAPE.RECTANGLE, ml=0.12, align=PP_ALIGN.LEFT if j in (0, 1) else PP_ALIGN.CENTER)
        if j in (0, 1):
            arrow(s, x + w_ - 0.16, y + 0.8, x + w_ + 0.01, y + 0.8, color=NAVY, w=2)
        x += w_
box(s, 1.2, 10.1, 25.12, 1.0, LINE_GREEN, [("新しい友だち　合計 月250人（1教室あたり）", 15, True, "FFFFFF")], adj=0.2)
text(s, 1.2, 11.3, 25, 0.7, [("それぞれの訴求（実文）", 13, True, TNAVY)])
MOCK = [("00 離脱防止ポップアップ", ["まだ、塾を決めなくて大丈夫です。", "30秒の診断で、お子さまに合う学び方と", "コース・料金の資料をLINEでお届けします。"], "LINEで資料と診断を受け取る"),
        ("01 完了画面", ["資料請求ありがとうございます。", "体験の日程の確認・変更は、", "LINEからいつでもできます。"], "LINEで友だち追加"),
        ("02 チラシ・教室のQR", ["LINEの友だち限定", "学年別「テスト対策プリント」を", "プレゼントしています。"], "QRから受け取る")]
for i, (t, body, btn) in enumerate(MOCK):
    x = 1.2 + i * 8.52
    box(s, x, 12.05, 8.08, 5.1, "FFFFFF", [], line=LGRAY, lw=1.25)
    text(s, x + 0.3, 12.15, 7.5, 0.6, [(t, 11, True, GRAY)])
    text(s, x + 0.3, 12.8, 7.5, 2.6, [(ln, 12 if k == 0 else 11.5, k == 0, TNAVY if k == 0 else INK) for k, ln in enumerate(body)])
    box(s, x + 0.6, 15.85, 6.9, 1.0, LINE_GREEN, [(btn, 12, True, "FFFFFF")], adj=0.4)

# ======== ② 効率改善（ステップ配信・企画配信） ========
s = rebuild2(S(7), "施策の設計図②　効率改善（ステップ配信・企画配信）",
             ["友だち追加から14日間はステップ配信で、その後は時期に合わせた企画配信で、",
              "資料請求（CV①）と体験予約（CV②）を取る。"],
             FOOT_MODEL, phases=[3, 4, 5, 6])
sec(s, 4.1, "[ステップ配信]　友だち追加から14日間・自動（新しい友だち 月250人）")
flow_head(s, 4.78)
STEPS = [("0日後", "友だち追加者全員\n250人", "あいさつ＋診断 → 資料を受け取る", "100% × 12% × 1.0%", "0.3件\nCV① 資料請求", True),
         ("3日後", "全員\n250人", "家でできる勉強のコツ", "72.5% × 12%", "CVは狙わない\n信頼づくり", False),
         ("5日後", "資料未請求の人\n約250人", "月謝＋講習費の年間の目安", "72.5% × 12% × 1.0%", "0.2件\nCV① 資料請求", True),
         ("7日後", "全員\n250人", "事例＋体験のご案内①", "72.5% × 12% × 0.8%", "0.2件\nCV② 体験予約", True),
         ("14日後", "体験未予約の人\n約250人", "テストから逆算＋体験の案内②", "72.5% × 12% × 1.0%", "0.2件\nCV② 体験予約", True)]
for i, r in enumerate(STEPS):
    flow_row(s, 5.43 + i * 1.03, list(r[:5]) + ["費用内"], h=0.97, cv=r[5])
sec(s, 10.65, "[企画配信]　月1〜2本（友だち 累計約1,000人・ブロックを除いて約680人）")
flow_head(s, 11.33)
PLANS = [("講習の前", "友だち全体\n約680人", "講習・新学年の先行案内", "78% × 10% × 1.0%", "0.5件\nCV② 体験・講習", True),
         ("テストの前", "中学生の保護者\n約270人", "テスト対策講座・見直しポイント", "78% × 10% × 0.5%", "0.1件\nCV② 体験予約", True),
         ("資料請求後", "資料を請求した人\n（件数は実績で）", "事例・費用のQ&A・教室紹介", "78% × 10% × 1.0%", "実績で計算\nCV② 体験予約", False),
         ("講習の時期", "体験・入会に\n至らなかった人", "講習のご案内（通知メッセージ）", "通知メッセージ", "実績で計算\n別途費用", False)]
for i, r in enumerate(PLANS):
    flow_row(s, 11.98 + i * 1.03, list(r[:5]) + (["費用内"] if i < 3 else ["別途"]), h=0.97, cv=r[5])
box(s, 1.2, 16.15, 25.12, 1.1, BEIGE, [("1教室あたり月に 資料請求 約0.5件・体験予約 約1.0件　→　100教室なら 資料請求 約50件・体験予約 約100件", 12.5, True, INK)])
page_no(s, 27)

# ======== ③ 満足度改善（ユーザー） ========
s = rebuild2(S(6), "施策の設計図③　満足度改善（保護者）",
             ["保護者の疑問に、待たせずに答える。", "すぐ答えられることは自動で、相談ごとは人が答える。押してほしいボタンは右下に置く。"],
             "※有人対応の時間帯（平日14〜21時）は例。教室の運用に合わせて決める", phases=[3, 5])
U3 = [("QA自動化", "料金・コース・対象学年・駐車場などは、キーワード自動応答ですぐ答える", "費用内"),
      ("個別チャット", "志望校の相談や日程の調整は、平日14〜21時に教室スタッフが1対1で対応", "費用内"),
      ("その他", "リッチメニューの右下に「無料体験・教室見学の予約」を常に置く", "費用内")]
for i, (t, d, fee) in enumerate(U3):
    y = 4.2 + i * 2.0
    box(s, 1.2, y, 3.4, 1.8, NAVY, [(f"[{t}]", 12, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE, ml=0.03)
    box(s, 4.7, y, 11.3, 1.8, "FFFFFF", [(d, 12.5, False, INK)], shape=MSO_SHAPE.RECTANGLE, line="C9D3E6", align=PP_ALIGN.LEFT)
    box(s, 16.1, y, 2.3, 1.8, PALE, [(fee, 11, True, NAVY)], shape=MSO_SHAPE.RECTANGLE, line="C9D3E6", ml=0.05)
px_, py_, pw_, ph_ = phone(s, 19.0, 4.1, 7.3, 13.1)
box(s, px_ + pw_ - 3.6, py_ + 0.3, 3.3, 0.8, "A6E3A1", [("料金はいくら？", 9, False, INK)], adj=0.3)
bubble(s, px_ + 0.25, py_ + 1.3, pw_ - 1.2, 2.6, ["（自動応答）", "コース別の月謝の目安です👇", "・小学生　◯◯円〜", "・中学生　◯◯円〜", "講習費を含めた年間の目安は", "こちら▶"], size=9.5)
box(s, px_ + pw_ - 3.9, py_ + 4.1, 3.6, 0.8, "A6E3A1", [("志望校の相談をしたい", 9, False, INK)], adj=0.3)
bubble(s, px_ + 0.25, py_ + 5.1, pw_ - 1.2, 1.9, ["（教室長の佐藤です）", "ご相談ありがとうございます。", "お子さまの学年を教えて", "いただけますか？"], size=9.5)
mx, my, mw = px_, py_ + ph_ - 3.4, pw_
cw_, ch_ = mw / 3, 1.7
for k, t in enumerate(["30秒診断", "コース・料金", "教室・講師", "資料を受け取る", "友だちと体験", "無料体験・\n見学の予約"]):
    x = mx + (k % 3) * cw_
    y = my + (k // 3) * ch_
    last = k == 5
    box(s, x, y, cw_, ch_, RED if last else "FFFFFF", [(ln, 8, True, "FFFFFF" if last else NAVY) for ln in t.split("\n")],
        line="BFBFBF", shape=MSO_SHAPE.RECTANGLE, ml=0.03)
box(s, 1.2, 10.4, 17.2, 4.3, PALE, [], line="8EA9DB")
text(s, 1.6, 10.6, 16.5, 4.0, [("保護者にとって良くなること", 14.5, True, NAVY),
                                ("・夜や休みの日でも、料金やコースがすぐ分かる", 13.5, False, INK),
                                ("・電話をしなくても、LINEで気軽に相談できる", 13.5, False, INK),
                                ("・体験の予約が、画面の右下からすぐできる", 13.5, False, INK)])
box(s, 1.2, 15.1, 17.2, 2.0, BEIGE, [("すぐ答える仕組みが、", 13, True, INK), ("資料請求・体験予約の取りこぼしを減らす。", 13, True, INK)])
page_no(s, 28)

# ======== ④ 効率改善（管理側） ========
s = new_slide_like(S(31))
s = rebuild(s, "施策の設計図④　効率改善（教室・運用側）",
            ["教室の手間を増やさずに回す。予約の受付・日程変更・持ち物の案内は、LINEで自動にする。",
             "LINE経由の資料請求・体験予約の数も、計測できるようにする。"],
            "※GA4のパラメータ（utm）でLINE経由の流入と資料請求・体験予約を計測。費用は月額費用内", phases=[5, 7])
U4 = [("自動応答", "体験予約の受付・日程の変更・持ち物の案内を、24時間自動で受け付ける", "費用内"),
      ("その他", "GA4のパラメータで、LINE経由の資料請求・体験予約の数を計測する", "費用内")]
for i, (t, d, fee) in enumerate(U4):
    y = 4.2 + i * 1.75
    box(s, 1.2, y, 3.4, 1.55, NAVY, [(f"[{t}]", 13, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE, ml=0.05)
    box(s, 4.7, y, 19.3, 1.55, "FFFFFF", [(d, 12.5, False, INK)], shape=MSO_SHAPE.RECTANGLE, line="C9D3E6", align=PP_ALIGN.LEFT)
    box(s, 24.1, y, 2.22, 1.55, PALE, [(fee, 11, True, NAVY)], shape=MSO_SHAPE.RECTANGLE, line="C9D3E6", ml=0.05)
text(s, 1.2, 7.85, 25, 0.7, [("[自動応答]　体験予約の流れ（教室は当日迎えるだけ）", 13, True, TNAVY)])
FL = [("LINEで\n「体験予約」", "FFFFFF", NAVY), ("日時を選ぶ\n（空き枠）", "FFFFFF", NAVY), ("予約完了を\n自動で送る", "FFFFFF", NAVY),
      ("前日に\n持ち物・地図", "FFFFFF", NAVY), ("変更は\n「変更」と送る", "FFFFFF", NAVY), ("教室は\n迎えるだけ", GREEN, "FFFFFF")]
w, g = 3.7, 0.584
for i, (t, f, c) in enumerate(FL):
    x = 1.2 + i * (w + g)
    box(s, x, 8.6, w, 2.3, f, [(ln, 12, True, c) for ln in t.split("\n")], line=None if f == GREEN else "8EA9DB", lw=1.5)
    if i < len(FL) - 1:
        arrow(s, x + w + 0.04, 9.75, x + w + g - 0.04, 9.75, color=NAVY, w=2)
text(s, 1.2, 11.35, 25, 0.7, [("[その他]　LINE経由の成果を数える", 13, True, TNAVY)])
FL2 = [("LINEの配信", "リンクに目印を付ける"), ("サイト", "GA4で見分ける"), ("資料請求（CV①）", "件数を計測"), ("体験予約（CV②）", "件数を計測"), ("月次レポート", "定例会で共有")]
w, g = 4.5, 0.655
for i, (t, sub) in enumerate(FL2):
    x = 1.2 + i * (w + g)
    col = GREEN if "CV" in t else NAVY
    box(s, x, 12.1, w, 2.3, "FFFFFF", [(t, 12.5, True, col), (sub, 11, False, GRAY)], line="8EA9DB", lw=1.5)
    if i < len(FL2) - 1:
        arrow(s, x + w + 0.04, 13.25, x + w + g - 0.04, 13.25, color=NAVY, w=2)
box(s, 1.2, 15.1, 25.12, 1.9, BEIGE, [("教室の手間を増やさずに、LINE経由の成果が毎月数字で見える。", 15, True, INK)])
PAGE4 = s

# ---- P26 成果の見方（旧28を作り替え）
s = rebuild(S(28), "成果の見方",
            ["友だちの数は成果にしない。見るのは、資料請求・体験予約・入会の3つ。", "入会は教室で決まるので、3ヶ月ごとに教室から報告をもらって確認する。"],
            "レポートは月次の定例会で共有（費用内）", phases="all")
KPI = [("毎月見る", NAVY, ["友だち追加数", "診断の回答数", "資料請求数（CV①）", "体験予約数（CV②）", "体験に来た数", "ブロック率"]),
       ("3ヶ月ごとに見る", GREEN, ["入会数（教室から報告をもらう）", "体験 → 入会の割合", "資料請求 → 体験予約の割合"])]
for i, (t, col, items) in enumerate(KPI):
    x = 1.2 + i * 12.82
    box(s, x, 4.3, 12.3, 0.9, col, [(t, 12, True, "FFFFFF")], adj=0.15)
    for j, it in enumerate(items):
        box(s, x, 5.4 + j * 1.3, 12.3, 1.1, "FFFFFF", [(it, 11, False, INK)], line="8EA9DB", align=PP_ALIGN.LEFT)
box(s, 1.2, 13.6, 25.12, 1.3, BEIGE, [("友だちの数ではなく、資料請求・体験予約・入会で見る。", 12.5, True, INK)])

# ---- P27 費用プラン（旧29：文言だけ直す）
s = S(29)
replace_in(s, [("例：6通×2セット（未入会向け・在籍生向け）＝12通", "例：6通×1セット（未入会の保護者向け）"),
               ("＝通知メッセ・保有リスト活用", "＝通知メッセージ・過去に問い合わせた保護者への再案内")])
shift_lead_for_band(s, "all")

# ---- P28 導入スケジュール（旧30：文言を直す）
s = S(30)
replace_in(s, [("現状ヒアリング（在籍数・月謝・退会率・広告実績）", "現状ヒアリング（月謝・広告実績・資料請求と体験予約の件数）"),
               ("CV3段とタグ設計の確定", "CV（資料請求・体験予約）とタグ設計の確定"),
               ("塾管理システムとの連携方式の決定", "LINEで資料請求の仕組みの決定"),
               ("法務確認（S09の表現チェック）", "表現の確認（成果の断定・最上級表現がないか）"),
               ("既存の在籍生への案内・友だち化", "教室スタッフへの運用説明"),
               ("セグメント配信・予兆検知の稼働", "出し分け配信・前日リマインドの稼働"),
               ("四半期で入会率・退会率を検証", "3ヶ月ごとに体験→入会の割合を確認"),
               ("（S10の検索ピークから）", "（検索の山から）"),
               ("（「塾」ピーク2/8の26日前）", ""), ("（夏期講習ピーク7/12の26日前）", ""),
               ("「塾」ピーク2/8なら12月着手、夏期講習7/12なら5月着手が目安。", "2月の入会期なら12月着手、夏期講習なら5月着手が目安。")])
title, lead, div, foot = parts(s)
if foot is not None:
    set_para_text(foot, ["構築に約1.5ヶ月。14日間の配信が一周してから数字が動き出す"])
shift_lead_for_band(s, "all")

# ---- P29 体制（旧31：文言を直す）
s = S(31)
replace_in(s, [("予兆検知トリガーの管理", "前日リマインドの管理"), ("入会率・退会率の検証と、施策の組み替え", "体験→入会の割合の確認と、施策の組み替え")])
shift_lead_for_band(s, "all")

# ---- P30 企画のご参考（旧32：文言を直す）
s = S(32)
tree = s.shapes._spTree
for sh in list(s.shapes):
    if sh.shape_id in (13, 14, 15, 16):
        tree.remove(sh._element)
    elif 17 <= sh.shape_id <= 24:
        sh.top = sh.top - Cm(2.15)
replace_in(s, [("飛び道具", "企画のご参考"), ("競合8社のLINEはすべて「予約の受け皿」で止まっている（S12）。", "他塾がまだやっていない企画の例。"),
               ("根拠データはS03〜S12および巻末の出典一覧を参照", "根拠は巻末の出典一覧を参照")])
shift_lead_for_band(s, "all")

# ---- P31 出典（旧34を作り替え）
s = rebuild(S(34), "この資料の出典",
            ["この資料で使った数値の出所を、すべて並べます。", "出典のない数値は書きません。取得中のデータは、届き次第差し替えます。"],
            "業界指標／自社実績／業界推計値／モデル値は、ラベルを分けて表記している")
SRCS = [("P3", "厚労省 人口動態統計／経産省 特定サービス産業動態統計／文科省 学校基本調査／矢野経済研究所 2024"),
        ("P4", "広告の実績（取得中）"),
        ("P5", "LINEヤフー前後検索（実測）／塾ナラ 2025 n=200／塾シル・ユナイトプロジェクト 2026 n=249／明光ネットワークジャパン 2020"),
        ("P6・P8・P20", "LINEヤフー前後検索 直近1年（取得中）KW：塾／個別指導／体験授業／高校受験"),
        ("P7・P24", "Googleトレンド 日本 2026年年初来（実測）"),
        ("P9", "各社LINE公式アカウント（page.line.me）・公式サイト 2026-09-25取得"),
        ("P15・P25", "POPER「Comiru」保護者と学習塾の意識調査 2022年 n=300"),
        ("P34", "オリコンME 2025 学習塾 利用実態データ／インタースペース「ママスタ」2021年 n=1,037")]
box(s, 1.2, 4.2, 3.6, 0.7, NAVY, [("ページ", 10.5, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE)
box(s, 4.9, 4.2, 21.42, 0.7, NAVY, [("出典", 10.5, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE)
for i, (pg, src) in enumerate(SRCS):
    y = 5.0 + i * 0.95
    f = "FFFFFF" if i % 2 == 0 else PALE
    box(s, 1.2, y, 3.6, 0.88, f, [(pg, 10, True, NAVY)], shape=MSO_SHAPE.RECTANGLE, line="E7E6E6")
    box(s, 4.9, y, 21.42, 0.88, f, [(src, 9.5, False, INK)], shape=MSO_SHAPE.RECTANGLE, line="E7E6E6", align=PP_ALIGN.LEFT)
box(s, 1.2, 12.9, 12.3, 3.0, PALE, [("LINEヤフー公式データ", 11, True, NAVY),
                                     ("・LINE 国内月間利用者数：1億人突破（2025年12月末時点）", 9.5, False, INK),
                                     ("・メッセージの開封：受信直後 約2割／3〜6時間で約5割／当日中に約8割", 9.5, False, INK)],
    line="8EA9DB", align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
box(s, 14.02, 12.9, 12.3, 3.0, "FFF7F7", [("提出前に確認すること", 11, True, RED),
                                          ("・公表統計（P3）はe-Statで統計表を開いて数値を照合する", 9.5, False, INK),
                                          ("・競合の実名掲載（P9）と、民間調査の引用可否は上司に確認する", 9.5, False, INK)],
    line=RED, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)


# ================================================================ 並べ替え・削除
ORDER = [1, 27, 2, 3, 4, 10, 9, 14, 11, 5, 12, 26, 13, 15, 22, 16, 17, 18, 19, 20, 21, 23, 25, 24, 33, 8, 7, 6, 36, 28, 29, 30, 31, 32, 34, 35]
DROP = []
lst = prs.slides._sldIdLst
ids = list(lst)
for n in DROP:
    rId = ids[n - 1].get(qn("r:id"))
    prs.part.drop_rel(rId)
for e in ids:
    lst.remove(e)
for n in ORDER:
    lst.append(ids[n - 1])

# 図形IDの重複を解消
for sl in prs.slides:
    tree = sl.shapes._spTree
    seen, els = set(), [e for e in tree.iter() if e.tag == qn("p:cNvPr")]
    mx = max([int(e.get("id")) for e in els] + [1])
    for e in els:
        if int(e.get("id")) in seen:
            mx += 1
            e.set("id", str(mx))
        seen.add(int(e.get("id")))

prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
