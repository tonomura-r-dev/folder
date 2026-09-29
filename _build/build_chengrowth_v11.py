# -*- coding: utf-8 -*-
"""チェングロウス アジェンダv11（21枚＋裏表紙）の提案資料を作る。

ベース：20260922_..ご提案ver1.2.pptx（殿村さんPC保存版・DYM FMT）。デザインと一部ページを流用し、残りは作り直す。
画像（画面イメージ）はClaudeが作らない。差し込み枠だけ置き、画像生成プロンプトは別途チャットで渡す。

  python3 _build/build_chengrowth_v11.py
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
import io

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "20260922_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver1.2.pptx"
OUT = ROOT / "20260929_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver2.0.pptx"
TW = ROOT / "_data/参考資料/タウンワークLINE_20260928"
TREND = ROOT / "_images/chengrowth_trend_seibishi_5y_L.png"

TNAVY = "002060"; NAVY = "1F285A"; ORANGE = "ED7D31"; RED = "C00000"
INK = "333333"; MUT = "7F7F7F"; WHITE = "FFFFFF"; PALE = "F4F7FF"
PORANGE = "FCE4D6"; BORDER = "D9D9D9"; GREEN = "06C755"; PGREEN = "E8F8EE"

SW, SH = 27.52, 19.05
TITLE_XY = (1.52, 0.38, 24.4, 0.94)
LEAD_XY = (1.20, 1.80, 25.1, 1.90)
DIV_Y = 3.86
CX0, CW = 1.20, 25.12
CY0 = 4.30
FOOT_Y = 17.35

shutil.copyfile(SRC, OUT)
prs = Presentation(str(OUT))
S = list(prs.slides)
assert len(S) == 52, len(S)


# ================= helpers（build_chengrowth.py と同じ作法） =================
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


def card(slide, x, y, w, h, head, body, hcol=NAVY, fill=PALE, hsz=12.5, bsz=10.5,
         line=None, anchor="t", ls=1.25):
    sp = box(slide, x, y, w, h, fill=fill, line=line)
    paras = [one(head, hsz, True, hcol, sa=5)]
    for b in (body if isinstance(body, list) else [body]):
        if b:
            paras.append(one(b, bsz, None, INK, ls=ls, sa=2))
    put_text(sp.text_frame, paras, anchor=anchor, ml=0.35, mr=0.3, mt=0.25, mb=0.15)
    return sp


def chip(slide, x, y, w, h, text, fill=NAVY, col=WHITE, sz=11):
    sp = box(slide, x, y, w, h, fill=fill, radius=0.18)
    put_text(sp.text_frame, [one(text, sz, True, col, align="c")], anchor="m", ml=0.1, mr=0.1, mt=0, mb=0)
    return sp


def arrow(slide, x, y, w=0.9, h=0.9, fill=NAVY):
    sp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(x), Cm(y), Cm(w), Cm(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp


def down(slide, x, y, w=1.4, h=0.8, fill=NAVY):
    sp = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Cm(x), Cm(y), Cm(w), Cm(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp


def clear_slide(slide):
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)


def frame(slide, title, lead):
    clear_slide(slide)
    T(slide, *TITLE_XY, [one(title, 16, True, TNAVY)], anchor="m", ml=0, mr=0)
    T(slide, *LEAD_XY, [one(l, 12.5, None, INK, ls=1.3) for l in lead], anchor="m", ml=0, mr=0)
    ln = slide.shapes.add_connector(1, Cm(0), Cm(DIV_Y), Cm(SW), Cm(DIV_Y))
    ln.line.color.rgb = RGBColor.from_string(BORDER)
    ln.line.width = Pt(1.0)


def foot(slide, text):
    T(slide, CX0, FOOT_Y, CW, 0.9, [one(text, 7.5, None, MUT, ls=1.15)], ml=0, mr=0)


def band(slide, y, text, fill=NAVY, col=WHITE, sz=13, h=1.2, x=CX0, w=CW):
    sp = box(slide, x, y, w, h, fill=fill, radius=0.10)
    put_text(sp.text_frame, [one(text, sz, True, col, align="c", ls=1.25)],
             anchor="m", ml=0.3, mr=0.3, mt=0, mb=0)
    return sp


def imgslot(slide, x, y, w, h, label, note):
    """画像の差し込み枠（画像はClaudeが作らない）"""
    sp = box(slide, x, y, w, h, fill=WHITE, line=MUT, lw=1.25, dash=MSO_LINE_DASH_STYLE.DASH)
    put_text(sp.text_frame,
             [one("［画像］" + label, 11, True, MUT, align="c", sa=4),
              one(note, 9, None, MUT, align="c", ls=1.25)],
             anchor="m", ml=0.3, mr=0.3, mt=0.1, mb=0.1)
    return sp


IMG = ROOT / "_images"


def fitpic(slide, name, x, y, w, h):
    """殿村さんが画像生成した画面イメージを、枠(x,y,w,h)に縦横比を保って中央に置く"""
    from PIL import Image as _I
    iw, ih = _I.open(IMG / name).size
    sc = min(w / iw, h / ih)
    pw, ph = iw * sc, ih * sc
    pic = slide.shapes.add_picture(str(IMG / name), Cm(x + (w - pw) / 2), Cm(y + (h - ph) / 2), Cm(pw), Cm(ph))
    pic.line.color.rgb = RGBColor.from_string(BORDER)
    return pic


def table(slide, x, y, w, h, headers, rows, col_w=None, hsz=10.5, bsz=10.5,
          header_fill=NAVY, zebra=PALE, align=None, bold_rows=(), red_cells=()):
    gf = slide.shapes.add_table(len(rows) + 1, len(headers), Cm(x), Cm(y), Cm(w), Cm(h))
    tbl = gf.table
    if col_w:
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = Cm(cw)
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string(header_fill)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        put_text(c.text_frame, [one(htext, hsz, True, WHITE, align="c")], anchor="m",
                 ml=0.12, mr=0.12, mt=0.02, mb=0.02)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(zebra if (zebra and i % 2 == 0) else WHITE)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            col = RED if (i, j) in red_cells else INK
            put_text(c.text_frame, [one(str(val), bsz, (i in bold_rows) or (i, j) in red_cells, col,
                                        align=(align[j] if align else "l"))],
                     anchor="m", ml=0.12, mr=0.12, mt=0.02, mb=0.02)
    return gf


def walk_replace(shape, mapping):
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


def set_shape_text(slide, contains, paras, **kw):
    for sh in slide.shapes:
        if sh.has_text_frame and contains in sh.text_frame.text:
            put_text(sh.text_frame, paras, **kw)
            return sh
    raise KeyError(contains)


# ============================================================
# 流用するページの素材を先に抜いておく
# ============================================================
login_pic = [sh for sh in S[48].shapes if sh.shape_type == 13][0]
LOGIN_BLOB = login_pic.image.blob

# 使い回し先（中身を作り直す）のスライド：レイアウトが「4_タイトルとコンテンツ」の不要ページ
spare = [S[i] for i in (1, 2, 3, 4, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20, 21, 22, 24, 25)]
spare_it = iter(spare)


def new(title, lead):
    s = next(spare_it)
    frame(s, title, lead)
    return s


ORDER = []

# ============================================================
# 1 表紙（流用）
# ============================================================
s = S[0]
replace_on(s, {"最上位パートナーの知見を活用し、獲得に強い動線を構築する": "―サイトリニューアルに伴うご提案―"})
ORDER.append(s)

# ============================================================
# 2 タウンワークの事例
# ============================================================
s = new("求人メディアのLINE活用事例｜タウンワーク",
        ["リッチメニューに人気の検索条件を並べ、LINEからワンタップで求人を探せるようにしました。",
         "その結果、LINE経由の応募は、配信よりもリッチメニュー経由が最も多くなりました。"])
kx, kw, gap = CX0, 7.9, 0.71
for i, (big, lab, sub) in enumerate([
        ("＋79%", "応募数", "LINEからの応募が大きく増加"),
        ("＋20%", "リッチメニューのタップ数", "探す入口として使われるように"),
        ("最多", "リッチメニュー経由の応募", "LINE経由の応募の中で、配信より多い")]):
    x = kx + i * (kw + gap)
    sp = box(s, x, 4.5, kw, 4.6, fill=PALE, line=BORDER)
    put_text(sp.text_frame, [one(lab, 13, True, NAVY, align="c", sa=6),
                             one(big, 40, True, ORANGE, align="c", sa=6),
                             one(sub, 11, None, INK, align="c")], anchor="m")
T(s, CX0, 9.6, CW, 0.8, [one("何をしたか", 13, True, NAVY)], ml=0)
steps = [("1", "リッチメニューに人気の検索条件を並べる", "単発・日払い・在宅など、よく探される条件をボタンに"),
         ("2", "トークで条件を選ぶと、求人が表示される", "選んだ条件に合う求人がカード形式で並ぶ"),
         ("3", "そのまま求人ページへ", "条件を入れた状態の検索結果に移動し、サイトで応募")]
for i, (n, h, b) in enumerate(steps):
    x = kx + i * (kw + gap)
    card(s, x, 10.5, kw, 3.9, f"{n}　{h}", b)
    if i < 2:
        arrow(s, x + kw + 0.02, 12.0, 0.66, 0.8)
band(s, 15.0, "LINEを「求人を探す入口」にしたことが、応募の増加につながりました。")
foot(s, "出典：LINEヤフー for Business 導入事例（タウンワーク・2019年公開）")
ORDER.append(s)

# ============================================================
# 3 タウンワークの実際の画面
# ============================================================
s = new("タウンワークの実際の画面",
        ["LINEは「条件を選ぶ入口」で、検索と応募はサイト側で行う作りです。",
         "サイトの会員とLINEを連携し、条件別の検索結果ページへそのまま移動させています。"])
caps = [("1_あいさつ.png", "あいさつ", "条件に合う求人を受け取るための\n会員連携を案内"),
        ("2_リッチメニュー.png", "リッチメニュー", "時間帯・学生歓迎などの\n条件がワンタップ"),
        ("3_遷移先サイト.png", "移動先のサイト", "条件を入れた状態の\n検索結果ページ"),
        ("4_連携リンク先_リクルートIDログイン.png", "会員連携の画面", "リクルートIDでログイン\n（メール＋パスワード）")]
iw, ih = 4.9, 10.6
gx = (CW - 4 * iw) / 3
for i, (fn, h, b) in enumerate(caps):
    x = CX0 + i * (iw + gx)
    chip(s, x, 4.3, iw, 0.8, f"{i + 1}　{h}", sz=11.5)
    pic = s.shapes.add_picture(str(TW / fn), Cm(x + (iw - ih * 1170 / 2532) / 2), Cm(5.25), height=Cm(ih))
    pic.line.color.rgb = RGBColor.from_string(BORDER)
    T(s, x - 0.3, 15.95, iw + 0.6, 1.3, [one(l, 10, None, INK, align="c") for l in b.split("\n")], ml=0, mr=0)
foot(s, "画面：タウンワーク LINE公式アカウント（2026年9月 弊社確認）")
ORDER.append(s)

# ============================================================
# 4 LINEのメリット① 興味別に届けられる
# ============================================================
s = new("LINEのメリット①｜興味別に届けられる",
        ["配信の中でクリックされた内容をもとに、友だちを自動で分けられます（LINE公式アカウントの標準機能）。",
         "配信するたびに分け方の精度が上がり、全員に同じ内容を送らずに済みます。"])
chip(s, CX0, 4.5, 6.2, 0.9, "STEP1　配信する", sz=12)
card(s, CX0, 5.5, 6.2, 7.2, "新着求人のお知らせ", ["・整備士の求人", "・営業の求人", "・未経験OK・資格取得支援の求人",
                                                "", "3つの求人を1通で届ける"], bsz=11.5)
arrow(s, 7.7, 8.6)
chip(s, 9.0, 4.5, 8.3, 0.9, "STEP2　クリックで自動に分かれる", sz=12)
for i, (h, b) in enumerate([("整備士の求人をクリック", "→「整備士に興味」のグループ"),
                            ("営業の求人をクリック", "→「営業に興味」のグループ"),
                            ("資格取得支援をクリック", "→「未経験・これから資格」のグループ")]):
    card(s, 9.0, 5.5 + i * 2.45, 8.3, 2.25, h, b, hsz=11.5, bsz=11)
arrow(s, 17.75, 8.6)
chip(s, 19.0, 4.5, 7.32, 0.9, "STEP3　合う人にだけ届ける", sz=12)
card(s, 19.0, 5.5, 7.32, 7.2, "次の配信から", ["整備士グループには整備士の求人を、", "未経験グループには資格取得支援を。",
                                            "", "興味のない案内が減るので、", "ブロックされにくくなります。"],
     fill=PORANGE, hcol=ORANGE, bsz=11.5)
band(s, 13.6, "登録時に細かく聞かなくても、クリックの履歴で興味が分かっていきます。")
foot(s, "※LINE公式アカウント管理画面の「オーディエンス」機能（クリックリターゲティング）を使用")
ORDER.append(s)

# ============================================================
# 5 LINEのメリット② 何度でも接点を持てる
# ============================================================
s = new("LINEのメリット②｜何度でも接点を持てる",
        ["広告は、クリックされた一度きりで接点が終わります。",
         "LINEは、一度友だちになれば、広告費をかけ直さずに何度でも求人をご案内できます。"])
T(s, CX0, 4.6, 2.6, 1.2, [one("広告", 14, True, NAVY)], anchor="m", ml=0)
flow1 = ["広告を見る", "求人を見る", "応募しない", "接点が終わる"]
for i, t in enumerate(flow1):
    x = 4.0 + i * 5.7
    chip(s, x, 4.6, 4.6, 1.2, t, fill=(RED if i == 3 else "E6E9F0"), col=(WHITE if i == 3 else INK), sz=12)
    if i < 3: arrow(s, x + 4.7, 4.85, 0.8, 0.7, fill=MUT)
T(s, CX0, 7.0, 2.6, 1.2, [one("LINE", 14, True, GREEN)], anchor="m", ml=0)
flow2 = ["求人を見る", "LINEで友だちに", "新着求人・\n条件の合う求人", "応募"]
for i, t in enumerate(flow2):
    x = 4.0 + i * 5.7
    sp = chip(s, x, 7.0, 4.6, 1.2, t.replace("\n", ""), fill=(GREEN if i in (1, 3) else PGREEN),
              col=(WHITE if i in (1, 3) else INK), sz=12)
    if i < 3: arrow(s, x + 4.7, 7.25, 0.8, 0.7, fill=GREEN)
T(s, 15.4, 8.4, 5.0, 0.8, [one("↻ 何度でもご案内", 11, True, GREEN, align="c")], ml=0)
for i, (h, b) in enumerate([("応募しなかった人も残る", "今すぐ応募しない人も、友だちとして残り続けます。"),
                            ("タイミングが合ったときに届く", "転職を考え始めた時期に、新着求人が届きます。"),
                            ("連絡がLINEで完結する", "電話やメールより気軽にやりとりできます。")]):
    card(s, CX0 + i * 8.61, 10.0, 7.9, 3.6, h, b, bsz=11.5)
band(s, 14.3, "一度の広告費で、何度もご案内できる「見込みの人の名簿」ができます。")
foot(s, "※配信通数に応じたLINE公式アカウントの利用料は別途かかります")
ORDER.append(s)

# ============================================================
# 6 LINEのメリット③ 登録の手間を下げられる（Profile+）
# ============================================================
s = new("LINEのメリット③｜登録の手間を下げられる",
        ["「LINEでログイン」を使うと、LINEのアカウントでサイトの会員登録ができます。",
         "さらにLINE Profile+を使うと、氏名・電話番号などが自動で入り、確認して押すだけで登録が終わります。"])
pic = s.shapes.add_picture(io.BytesIO(LOGIN_BLOB), Cm(CX0), Cm(4.4), width=Cm(15.6))
T(s, CX0, 4.4 + 15.6 * 12.64 / 25.22 + 0.1, 15.6, 0.6, [one("「LINEでログイン」の画面の流れ（イメージ）", 9, None, MUT)], ml=0)
card(s, 17.3, 4.4, 9.02, 4.4, "LINE Profile+で自動入力できる項目",
     ["氏名（カナを含む）", "性別・生年月日", "電話番号・住所"], bsz=11.5)
card(s, 17.3, 9.0, 9.02, 4.2, "利用の条件",
     ["日本の法人のみ", "使う項目ごとに申請し、LINEヤフー社の審査がある", "ユーザーの同意が必要"], bsz=11,
     fill=PORANGE, hcol=ORANGE)
band(s, 14.4, "入力が「確認して押すだけ」になり、会員登録のハードルを大きく下げられます。")
foot(s, "出典：LINE Developers（LINE Profile+・2025年11月更新）／画面の流れの引用：https://prtimes.jp/main/html/rd/p/000003920.000001594.html")
ORDER.append(s)

# ============================================================
# 7 他社のLINE活用（流用）
# ============================================================
s = S[8]
replace_on(s, {
    "競合分析（求人サービスのLINE活用状況）": "他社のLINE活用｜求人サービスの状況",
    "総合大手はLINEが標準装備。一方、自動車・整備士特化は空白地帯。": "総合大手は、数十万〜数百万人規模でLINEを活用しています。一方、整備士特化のサービスでは、まだ少ない状況です。",
})
set_shape_text(s, "まだ誰も取っていない",
               [one("整備士特化で、LINEを本格的に活用しているサービスはまだ少なく、先に始めた方が有利です。", 13, True, WHITE, align="c")],
               anchor="m")
ORDER.append(s)

# ============================================================
# 8 タウンワークの型を分解
# ============================================================
s = new("タウンワークの型を分解すると、3つの仕組み",
        ["応募が増えた理由は、LINEを「求人を探す入口」にしたことです。",
         "この3つは、サイト側に「条件別の検索結果ページ」と「LINEとの会員連携」があれば再現できます。"])
for i, (h, b) in enumerate([
        ("① リッチメニューで条件検索", ["よく探される条件をボタンに並べる。", "トーク画面の下に常にあるので、", "探したいときにいつでも使える。"]),
        ("② 全員共通のあいさつ＋タップで分岐", ["あいさつで会員連携を案内する。", "条件のカードをタップしてもらい、", "興味を聞く。"]),
        ("③ 検索結果へそのまま移動", ["条件を入れた状態の検索結果ページへ。", "探す手間がなく、", "応募はサイト側で行う。"])]):
    x = CX0 + i * 8.61
    chip(s, x, 4.5, 7.9, 1.0, h, sz=12)
    card(s, x, 5.6, 7.9, 4.6, "", b, bsz=12)
T(s, CX0, 10.8, CW, 0.8, [one("貴社で再現するときに必要なもの", 13, True, NAVY)], ml=0)
card(s, CX0, 11.7, 12.2, 2.6, "サイト側", ["条件別の検索結果ページ（職種×地域×資格）", "LINEで会員登録・ログインできる仕組み"],
     fill=PORANGE, hcol=ORANGE, bsz=11.5)
card(s, 14.1, 11.7, 12.22, 2.6, "LINE側", ["リッチメニュー・あいさつ・配信の設計", "条件で選んで送る仕組み"],
     fill=PGREEN, hcol="0B7A3B", bsz=11.5)
band(s, 14.9, "サイトリニューアルのタイミングなら、この仕組みを最初から入れられます。")
ORDER.append(s)

# ============================================================
# 9 貴社に当てはめると
# ============================================================
s = new("貴社に当てはめると｜登録はLINEで、課題は応募",
        ["サイトリニューアルに合わせて、会員登録をLINEで完結できる形にすることをご提案します。",
         "入力の手間がほぼなくなり、登録のハードルを大きく下げられます。"])
chip(s, CX0, 4.4, 11.9, 0.95, "現在", fill=MUT, sz=12.5)
card(s, CX0, 5.45, 11.9, 5.0, "入口は「応募」か「転職支援の申込」",
     ["・会員登録は、ほとんど使われていない", "・応募しなかった人とは、接点が残らない", "・比べている途中の人が残る場所がない"],
     fill="F2F2F2", hcol=INK, bsz=11.5)
arrow(s, 13.35, 7.5, 0.9, 0.9)
chip(s, 14.42, 4.4, 11.9, 0.95, "リニューアル後", fill=GREEN, sz=12.5)
card(s, 14.42, 5.45, 11.9, 5.0, "会員登録＝LINEでワンタップ",
     ["・登録と同時にLINEの友だちになる", "・応募しなかった人も、登録者として残る", "・登録した人に、合う求人を届け続けられる"],
     fill=PGREEN, hcol="0B7A3B", bsz=11.5)
down(s, 12.96, 10.75)
sp = box(s, CX0, 11.8, CW, 2.3, fill=PORANGE, line=ORANGE, lw=1.5)
put_text(sp.text_frame, [one("登録（CV①）はワンタップで取れる。課題は、登録した人を応募（CV②）までどう運ぶか。", 14, True, ORANGE, align="c", sa=4),
                         one("→ リッチメニューと配信で、LINEから応募まで運びます（次ページ以降）", 12, None, INK, align="c")],
         anchor="m")
foot(s, "CV①＝会員登録（LINEで登録）／CV②＝求人への応募")
ORDER.append(s)

# ============================================================
# 10 貴社向けの工夫
# ============================================================
s = new("貴社向けの工夫｜タウンワークより一段軽く、整備士向けに",
        ["タウンワークの型をそのまま使うのではなく、貴社のサイトとお仕事に合わせて2点を変えます。"])
chip(s, CX0, 4.4, 12.2, 1.0, "工夫①　登録がタウンワークより一段軽い", sz=12.5)
table(s, CX0, 5.6, 12.2, 4.2, ["", "タウンワーク", "貴社（ご提案）"],
      [["LINEとつなぐ前", "リクルートIDで\nログインが必要", "不要"],
       ["入力するもの", "メール・パスワード", "ほぼなし\n（Profile+で自動入力）"],
       ["つながった状態", "既存の会員と連携", "登録と友だち追加が\n同時に完了"]],
      col_w=[3.4, 4.2, 4.6], align=["c", "c", "c"], bsz=10.5, red_cells=((1, 2), (2, 2), (3, 2)))
chip(s, 14.1, 4.4, 12.22, 1.0, "工夫②　条件を整備士向けにする", sz=12.5)
card(s, 14.1, 5.6, 12.22, 4.2, "リッチメニュー・配信で使う条件",
     ["・職種：整備士／営業／受付・事務", "・地域：都道府県・市区町村", "・資格：1級・2級・3級・無資格",
      "・未経験OK・資格取得支援あり"], bsz=11.5)
card(s, CX0, 10.4, CW, 2.6, "掲載中の求人を、条件で探せるようにします",
     ["全国2,134件（うち正社員・整備士946件、正社員・営業180件）の求人を、LINEから条件別に探せる入口をつくります。"],
     bsz=11.5)
band(s, 13.6, "「軽い登録」と「整備士に合わせた条件」で、タウンワークの型を貴社向けに最適化します。")
foot(s, "出典：自動車求人Navi 検索画面（2026年9月28日時点の掲載件数）")
ORDER.append(s)

# ============================================================
# 11 友だち追加の動線
# ============================================================
s = new("友だち追加の動線｜サイトに来た人をLINEに残す",
        ["サイトに来た人を、2つの入口でLINEの友だち（登録者）にします。"])
for i, (tag, head, flow, cost, slot, note) in enumerate([
        ("00", "離脱防止ポップアップ", "求人ページで帰ろうとした人 → 「条件に合う新着求人をLINEでお届け」 → LINE追加",
         "※月3万円（初期1.5万円）", "chengrowth_v11_popup.png", ""),
        ("01", "「LINEで登録」ボタン", "求人詳細・会員登録ページ → 「LINEで登録（入力ほぼ不要）」 → LINE追加＋会員登録",
         "※サイト側で実装（リニューアルの要件）", "chengrowth_v11_register.png", "")]):
    x = CX0 + i * 12.92
    chip(s, x, 4.3, 12.2, 1.0, f"動線{tag}　{head}", sz=12.5)
    T(s, x, 5.4, 12.2, 1.9, [one(flow, 11, None, INK, ls=1.25), one(cost, 10, True, ORANGE)], ml=0.1)
    fitpic(s, slot, x, 7.6, 12.2, 7.1)
band(s, 15.2, "応募しなかった人も、帰ろうとした人も、LINEの友だちとして残ります。")
foot(s, "※画面はイメージです")
ORDER.append(s)

# ============================================================
# 12 リッチメニューと配信で応募まで運ぶ
# ============================================================
s = new("リッチメニューと配信で、応募まで運ぶ",
        ["登録した人が「探したいとき」はリッチメニュー、「まだ迷っているとき」は配信で、応募まで運びます。"])
chip(s, CX0, 4.3, 8.2, 0.95, "リッチメニュー（いつでも探せる入口）", sz=11.5)
fitpic(s, "chengrowth_v11_richmenu.png", CX0, 5.35, 8.2, 5.6)
T(s, CX0, 11.05, 8.2, 1.6, [one("整備士・営業・近くの求人・資格・未経験OK・新着を、ワンタップで検索結果へ", 10.5, None, INK, ls=1.25)], ml=0)
chip(s, 9.8, 4.3, 16.52, 0.95, "配信（迷っている人を応募へ）", sz=11.5)
table(s, 9.8, 5.35, 12.4, 6.6, ["配信", "対象", "内容", "開封×クリック"],
      [["ステップ配信", "登録した人\n（0〜14日）", "職種を聞く→合う求人\n→応募のご案内", "72.5%×12%"],
       ["新着求人", "条件が合う人", "条件が合う新着だけ", "78%×10%"],
       ["年間の企画", "友だち全員", "3月の山に向けて前倒し\n（月4本）", "78%×10%"]],
      col_w=[2.6, 2.6, 4.4, 2.8], align=["c", "c", "l", "c"], bsz=10.5, hsz=10.5)
fitpic(s, "chengrowth_v11_talk_step.png", 22.5, 5.35, 3.82, 6.6)
band(s, 13.3, "応募（CV②）は、リッチメニューと配信の両方から生まれます。")
foot(s, "※画面はイメージです／開封×クリックは弊社シミュレーションの前提値")
ORDER.append(s)

# ============================================================
# 13 【推し施策】応募を採用までつなげる
# ============================================================
s = new("【推し施策】応募を採用までつなげる｜サンクスLINE",
        ["応募の完了画面でLINEの友だち追加を案内し、応募した人とLINEでつながります。",
         "面談日程のリマインドで取りこぼしを防ぎ、応募の先の「採用」を増やす施策です。"])
chip(s, CX0, 4.3, 16.0, 1.0, "動線02　応募の完了画面（サンクスLINE）", sz=12.5)
T(s, CX0, 5.4, 16.0, 1.5, [one("応募の完了画面 → 「今後のご連絡はLINEでお送りします」 → LINE追加", 11.5, None, INK),
                         one("※オプション：月3万円〜（初期10万円）", 10, True, ORANGE)], ml=0.1)
for i, (h, b) in enumerate([("面談日程のリマインド", ["前日・当日にLINEでご案内し、", "面談の無断キャンセルを防ぐ"]),
                            ("よくある質問の自動応答", ["資格・受験料・勤務地などに", "LINEで24時間すぐ回答"]),
                            ("個別のやりとり", ["日程変更や相談も", "電話なしでLINEで完結"])]):
    card(s, CX0 + i * 5.43, 7.2, 5.0, 4.4, h, b, bsz=10.5, hsz=10.5)
fitpic(s, "chengrowth_v11_talk_thanks.png", 17.7, 4.3, 8.62, 9.2)
sp = box(s, CX0, 12.0, 16.0, 1.5, fill=PORANGE, line=ORANGE)
put_text(sp.text_frame, [one("応募（CV②）の数は増やさず、応募から採用までの取りこぼしを減らす施策です。", 11.5, True, ORANGE, align="c")], anchor="m")
band(s, 14.1, "応募して終わりにせず、面談・採用まで同じLINEで伴走します。")
foot(s, "※画面はイメージです／シミュレーションの応募数・費用には含めていません（応募の先の採用に効く施策のため）")
ORDER.append(s)

# ============================================================
# 14 最初から入れた方が得
# ============================================================
s = new("最初から入れた方が得｜後から入れると勿体ない",
        ["LINEの友だち・会員の情報は、始めた日からしか貯まりません。",
         "サイトと同時に始めれば、翌3月の転職の山を、貯まった友だちで迎えられます。"])
for i, (h, b) in enumerate([
        ("① データの蓄積が遅れない", ["友だち・会員・クリックの履歴は", "始めた日から貯まる。", "半年後には約700人の友だちに", "（弊社シミュレーション）"]),
        ("② 開発のやり直しを避けられる", ["会員登録の仕組みを", "後からLINE対応に作り替えると、", "サイトの改修が二度手間になる。"]),
        ("③ 3月の山に間に合う", ["整備士の検索は一年中あり、", "山は3月。", "秋までに友だちを貯めておけば、", "山の時期に案内できる。"])]):
    card(s, CX0 + i * 8.61, 4.4, 7.9, 6.2, h, b, bsz=12, hsz=13)
chip(s, CX0, 11.1, 3.6, 1.0, "同時に開始", fill=GREEN, sz=11.5)
box(s, 5.0, 11.3, 21.3, 0.6, fill=GREEN)
T(s, 5.0, 11.95, 21.3, 0.7, [one("4月 ────── 友だちが貯まり続ける ────── 翌3月（転職の山）", 10.5, True, "0B7A3B", align="c")], ml=0)
chip(s, CX0, 12.9, 3.6, 1.0, "後から開始", fill=MUT, sz=11.5)
box(s, 5.0, 13.1, 10.0, 0.6, fill="E6E9F0")
box(s, 15.0, 13.1, 11.3, 0.6, fill=GREEN)
T(s, 5.0, 13.75, 10.0, 0.7, [one("この間の登録者はLINEに残らない", 10.5, True, MUT, align="c")], ml=0)
band(s, 15.0, "サイトリニューアルと同時に、LINEも最初から入れることをご提案します。")
ORDER.append(s)

# ============================================================
# 15 サイトリニューアルに入れる要件3点
# ============================================================
s = new("サイトリニューアルに入れる要件3点",
        ["LINEで応募まで運ぶために、リニューアルの要件として次の3点を入れていただくことをご提案します。"])
for i, (h, b) in enumerate([
        ("① 会員登録をLINEで完結させる", ["「LINEでログイン」とLINE Profile+で、", "入力ほぼなしで会員登録。", "登録と同時に友だち追加も完了。"]),
        ("② 会員の情報とLINEをひもづける", ["会員名簿にLINEの情報を保存し、", "職種・地域・資格で選んで", "配信できるようにする。"]),
        ("③ 条件別の検索結果ページ", ["職種×地域×資格ごとのURLを用意し、", "リッチメニューや配信から", "そのまま開けるようにする。"])]):
    x = CX0 + i * 8.61
    chip(s, x, 4.4, 7.9, 1.0, h, sz=12)
    card(s, x, 5.5, 7.9, 5.4, "", b, bsz=12)
card(s, CX0, 11.5, CW, 2.4, "LINE連携の開発について",
     ["LINE連携の開発は、弊社でも対応が可能です（内容に応じて都度お見積もり）。サイト制作会社様と仕様をすり合わせて進めます。"],
     fill=PORANGE, hcol=ORANGE, bsz=11.5)
band(s, 14.5, "この3点が入っていれば、タウンワークの型を貴社サイトで再現できます。")
ORDER.append(s)

# ============================================================
# 16 スケジュール
# ============================================================
s = new("スケジュール｜4月にサイトと同時に開設",
        ["LINE Profile+は申請とLINEヤフー社の審査があるため、予算が決まり次第、早めに着手します。"])
sched = [("11〜12月", "予算のご決定", "来期の予算にLINEを組み込む"),
         ("1〜2月", "お申込み・申請", "アカウント開設、LINE Profile+の申請、サイト制作会社様と仕様のすり合わせ"),
         ("2〜3月", "初期構築", "あいさつ・リッチメニュー・ステップ配信の作成、サイトとの連携テスト"),
         ("4月", "サイトと同時に開設", "運用開始。会員登録・友だち追加を開始"),
         ("〜翌3月", "友だちを貯めて山に備える", "毎月の配信・改善を続け、3月の転職の山に備える")]
for i, (m, h, b) in enumerate(sched):
    y = 4.4 + i * 2.25
    chip(s, CX0, y, 3.6, 1.9, m, fill=(GREEN if m == "4月" else NAVY), sz=13)
    card(s, 5.1, y, 21.22, 1.9, h, b, hsz=12.5, bsz=11, anchor="m", fill=(PGREEN if m == "4月" else PALE))
foot(s, "※LINE Profile+の審査期間は、申請内容により変わります")
ORDER.append(s)

# ============================================================
# 17 シミュレーション（SIM ver2.6）
# ============================================================
s = new("成果シミュレーション（4〜9月）",
        ["会員登録（CV①）はLINEでワンタップ。登録した人を、リッチメニューと配信で応募（CV②）まで運びます。",
         "友だちが貯まるほど応募が増え、7月に応募単価が2.5万円を下回る見込みです。"])
months = ["4月", "5月", "6月", "7月", "8月", "9月", "合計"]
table(s, CX0, 4.4, CW, 6.2, ["", *months],
      [["友だち（累計）", "82人", "187人", "307人", "438人", "570人", "709人", "―"],
       ["会員登録（CV①）", "58件", "74件", "85件", "93件", "93件", "98件", "501件"],
       ["応募（CV②）", "3件", "5件", "8件", "11件", "13件", "15件", "55件"],
       ["応募単価（CPA）", "7.8万円", "4.7万円", "2.9万円", "2.1万円", "1.8万円", "1.6万円", "―"]],
      col_w=[4.72] + [2.9] * 6 + [3.0], align=["l"] + ["c"] * 7, bsz=11.5, hsz=11.5,
      bold_rows=(3,), red_cells=((4, 4), (4, 5), (4, 6)))
for i, (lab, big, sub) in enumerate([("半年の会員登録（CV①）", "501件", "LINEでワンタップ登録"),
                                     ("半年の応募（CV②）", "55件", "リッチメニュー・配信から"),
                                     ("9月の応募単価", "1.6万円", "目標2〜2.5万円を下回る")]):
    x = CX0 + i * 8.61
    sp = box(s, x, 11.1, 7.9, 3.4, fill=(PORANGE if i == 2 else PALE), line=BORDER)
    put_text(sp.text_frame, [one(lab, 12, True, NAVY, align="c", sa=3),
                             one(big, 28, True, (ORANGE if i == 2 else NAVY), align="c", sa=3),
                             one(sub, 10.5, None, INK, align="c")], anchor="m")
foot(s, "前提：サイト来訪 月約6,100人（広告のシミュレーション確定後に差し替え）／LINEで登録＝来訪者の1.5%／費用 月23.5万円／"
        "応募単価＝月の費用÷その月の応募数（初期費用は含まない）")
ORDER.append(s)

# ============================================================
# 18 費用と体制
# ============================================================
s = new("費用と体制",
        ["LINE単独のお見積もりです（サイト制作費・広告運用費とは別）。運用は弊社が担当します。"])
table(s, CX0, 4.4, 14.6, 7.6, ["項目", "初期", "月額"],
      [["LINE運用コンサル\n（設計・構築・配信・改善）", "20万円", "20万円"],
       ["離脱防止ポップアップ", "1.5万円", "3万円"],
       ["LINE公式アカウント（開設／利用料）", "3万円", "0.5万円〜"],
       ["合計", "24.5万円", "23.5万円"],
       ["サンクスLINE（オプション）", "10万円", "3万円〜"],
       ["LINE Profile+／LINE連携の開発", "都度お見積もり", "―"]],
      col_w=[7.4, 3.6, 3.6], align=["l", "c", "c"], bsz=11, hsz=11.5, bold_rows=(4,),
      red_cells=((4, 1), (4, 2)))
T(s, 16.4, 4.4, 9.92, 0.8, [one("体制", 13, True, NAVY)], ml=0)
for i, (h, b) in enumerate([("弊社：戦略・設計", "アカウント設計、動線・配信の設計、効果の分析と改善のご提案"),
                            ("弊社：運用・制作", "あいさつ・リッチメニュー・配信の作成と設定、毎月のレポート"),
                            ("貴社にお願いすること", "求人情報・面談体制の共有、サイト制作会社様との連携")]):
    card(s, 16.4, 5.3 + i * 2.25, 9.92, 2.05, h, b, bsz=10.5, hsz=11.5,
         fill=(PORANGE if i == 2 else PALE), hcol=(ORANGE if i == 2 else NAVY))
band(s, 12.6, "月23.5万円で、会員登録から応募までをLINEで運ぶ仕組みを構築・運用します。")
foot(s, "※LINE公式アカウントの利用料は配信通数により変動します")
ORDER.append(s)

# ============================================================
# 19 参考：業界の人手不足
# ============================================================
s = new("参考：業界の人手不足",
        ["整備士・営業は、求人に対して人が大きく足りない「超売り手市場」です。担い手も増えていません。"])
for i, (lab, big, sub) in enumerate([("整備士の有効求人倍率", "5.46倍", "2025年度"),
                                     ("自動車営業の有効求人倍率", "12.55倍", "2025年度（受付事務は0.96倍）"),
                                     ("人手不足を感じる整備事業者", "約6割", "国土交通省の資料")]):
    x = CX0 + i * 8.61
    sp = box(s, x, 4.4, 7.9, 4.0, fill=PALE, line=BORDER)
    put_text(sp.text_frame, [one(lab, 12, True, NAVY, align="c", sa=4), one(big, 34, True, ORANGE, align="c", sa=4),
                             one(sub, 10.5, None, INK, align="c")], anchor="m")
table(s, CX0, 9.0, CW, 3.3, ["", "以前", "最近", "変化"],
      [["整備学校の入学者", "12,394人（2003年度）", "7,068人（2024年度）", "約4割減"],
       ["整備要員の平均年齢", "39.7歳", "47.2歳", "高齢化"]],
      col_w=[6.5, 6.5, 6.5, 5.62], align=["l", "c", "c", "c"], bsz=11.5, hsz=11.5, red_cells=((1, 3), (2, 3)))
band(s, 13.0, "人が集まりにくい市場だからこそ、来てくれた人を逃さない仕組みが効きます。")
foot(s, "出典：厚生労働省 job tag（有効求人倍率・2025年度）／国土交通省資料（出典：自動車整備白書）")
ORDER.append(s)

# ============================================================
# 20 参考：求職者の動き方
# ============================================================
s = new("参考：求職者の動き方",
        ["整備士は、複数の経路を併用して比べながら仕事を探しています。検索は一年中あり、山は3月です。"])
T(s, CX0, 4.3, 11.0, 0.8, [one("仕事を探した経路（複数回答）", 12.5, True, NAVY)], ml=0)
for i, (lab, v) in enumerate([("転職エージェント", 35.3), ("総合型の求人サイト", 34.1), ("業界特化型の求人サイト", 33.0)]):
    y = 5.3 + i * 1.5
    T(s, CX0, y, 4.4, 1.1, [one(lab, 10.5, None, INK)], anchor="m", ml=0)
    box(s, 5.7, y + 0.2, 5.0 * v / 35.3, 0.7, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    T(s, 5.8 + 5.0 * v / 35.3, y, 1.8, 1.1, [one(f"{v}%", 11, True, NAVY)], anchor="m", ml=0)
card(s, CX0, 10.0, 11.0, 2.4, "特化型サイトの登録者は、84%が有資格者", "整備士に特化したサイトには、資格を持つ人が集まる", bsz=10.5)
T(s, 13.0, 4.3, 13.3, 0.8, [one("「整備士」の検索（過去5年）", 12.5, True, NAVY)], ml=0)
s.shapes.add_picture(str(TREND), Cm(13.0), Cm(5.2), width=Cm(13.3))
card(s, 13.0, 10.0, 13.3, 2.4, "山は3月、谷は12月。差は約1.4倍", "一年中、転職を考える人がいる。3月に向けて前倒しで接点を持つ", bsz=10.5)
band(s, 13.0, "比べている途中の人を、LINEで登録者として残しておくことが大切です。")
foot(s, "出典：モビリア総研調査（直近3年に転職した整備士1,249人・Response 2026年9月23日）／Googleトレンド「整備士」（日本・過去5年）")
ORDER.append(s)

# ============================================================
# 21 参考：前後検索（流用）
# ============================================================
s = S[6]
replace_on(s, {"市場分析（前後検索クエリ）": "参考：前後検索（カーディーラー）",
               "面談のご案内①": "求人のご案内①", "面談のご案内②": "求人のご案内②"})
ORDER.append(s)

# 裏表紙（流用）
ORDER.append(S[51])

# ============================================================
# 並べ替えと不要ページの削除
# ============================================================
keep_ids = {id(x) for x in ORDER}
sldIdLst = prs.slides._sldIdLst
items = list(sldIdLst)
by_slide = {id(sl): el for sl, el in zip(S, items)}
for sl, el in zip(S, items):
    if id(sl) not in keep_ids:
        prs.part.drop_rel(el.rId)
        sldIdLst.remove(el)
for el in list(sldIdLst):
    sldIdLst.remove(el)
for sl in ORDER:
    sldIdLst.append(by_slide[id(sl)])

prs.save(str(OUT))
print("saved:", OUT.name, len(Presentation(str(OUT)).slides), "枚")
