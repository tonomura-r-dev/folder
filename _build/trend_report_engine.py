"""
LINE公式アカウント 媒体最新情報 トレンドレポート 共通エンジン（v2）。

トンマナは実際に配布された過去号そのもの（2026年6月号・7月号）から抽出した
_templates/DYM_LINEOA_TREND_FMT.pptx（＝2026年7月号の実ファイル）を土台にする。
以前使っていた _templates/DYM_LINEOA_BUFFF_FMT.pptx（営業提案資料のFMT）とは別物。
2026-09-16、殿村さんから過去号2本の提供を受けて全面差し替えた。

やり方：
  1. FMTをコピーし、表紙(0)・目次(1)・トピックのひな形(2＝LINE VOOMのページ)の
     3枚だけ残す
  2. ひな形スライドを duplicate_slide() でトピック数ぶん複製する
     （python-pptxの add_slide() は使うが、レイアウトは複製元と同じものを使うだけで
     新しいレイアウトは作らない。CLAUDE.mdの「スライド新規追加はしない」は
     FMTコピー後に空のスライドを新規に足すケースの話で、これは既存スライドの複製）
  3. 各スライドの中身（対象業界／分類タグ／バナー見出し／3行サマリー／表／出典）を
     テキストだけ書き換える。ヘッダーの「今後のアップデート情報」・DYMロゴ・
     区切り線・フッターはひな形をそのまま複製しているので触らない
  4. トピック本文は「概要・変更点／料金体系・出稿条件／スケジュール・導入フロー／
     注意点・影響」の4行×2列表に統一（ひな形の3列×7行の実施日テーブルは
     差し替える。理由：過去号は毎回テーブル構成がバラバラで再現しきれないため、
     既存のTOPICSデータ構造（overview/pricing/schedule/caution）をそのまま使える
     4行固定の表に寄せた）

月次スクリプトは build() を呼ぶだけでよい。作業手順は
.claude/skills/trend-report/SKILL.md を参照。
"""

import copy
import shutil
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR, MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
SRC = str(ROOT / "_templates" / "DYM_LINEOA_TREND_FMT.pptx")

# ひな形（FMT内の7月号）のスライド番号（0始まり）
COVER_IDX = 0
AGENDA_IDX = 1
TEMPLATE_IDX = 2  # 【LINE VOOM】ページ。分類タグは「LINE公式アカウント」が選択済み

CATEGORIES = ["LINE公式アカウント", "開発\nLINE API", "オプション\n商材", "その他"]

TAG_BORDER_BLUE = RGBColor(0x4F, 0x81, 0xBD)

# ネイティブ図形での本文表現（2026-09-17〜、「アップデート内容」「インパクト」の
# 2ページ構成デモで確定した配色）
NAVY = RGBColor(0x15, 0x13, 0x3D)
ACCENT = RGBColor(0x34, 0x67, 0xB2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BODY_GRAY = RGBColor(0x33, 0x33, 0x33)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
LIGHT_BLUE = RGBColor(0xEA, 0xF0, 0xFA)
BORDER_GRAY = RGBColor(0xD9, 0xD9, 0xD9)


# ---------- スライド複製 ----------

def _rewrite_and_copy_shape(source_part, dest_part, shape_el):
    """シェイプXMLをdeepcopyし、画像参照(r:embed)をdestパート側の新しい関係に張り替える。"""
    el = copy.deepcopy(shape_el)
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    for blip in el.iter(qn("a:blip")):
        old_rid = blip.get(qn("r:embed"))
        if not old_rid:
            continue
        rel = source_part.rels[old_rid]
        new_rid = dest_part.rels._add_relationship(rel.reltype, rel._target)
        blip.set(qn("r:embed"), new_rid)
    return el


def duplicate_slide(prs, index):
    """既存スライド(index)を複製して末尾に追加する。同じレイアウトを使うだけで
    新規レイアウトは作らない。戻り値は新しいSlideオブジェクト。"""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)
    # add_slide がレイアウト側プレースホルダを自動で入れてくることがあるので一旦全消し
    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)
    for shape in source.shapes:
        new_el = _rewrite_and_copy_shape(source.part, dest.part, shape._element)
        dest.shapes._spTree.append(new_el)
    return dest


def load_base(out_path, n_topics):
    """FMTをコピーし、表紙・目次・トピック用スライド(n_topics枚)を用意して返す。"""
    shutil.copyfile(SRC, out_path)
    prs = Presentation(out_path)

    # 表紙・目次・ひな形(1枚)以外は削除
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    keep_ids = {ids[COVER_IDX], ids[AGENDA_IDX], ids[TEMPLATE_IDX]}
    for sldId in ids:
        if sldId in keep_ids:
            continue
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)

    slides = list(prs.slides)
    assert len(slides) == 3, len(slides)
    cover, agenda, template = slides

    topic_slides = [template]
    for _ in range(n_topics - 1):
        topic_slides.append(duplicate_slide(prs, 2))

    return prs, cover, agenda, topic_slides


# ---------- テキスト編集ヘルパー ----------

def set_single_run_text(shape, new_text):
    """1パラグラフ1ランを前提に、フォーマットを保ったままテキストだけ差し替える。"""
    p = shape.text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = new_text
    for r in list(p.runs[1:]):
        r._r.getparent().remove(r._r)


def set_multiline_text(shape, lines):
    """既存の段落数 <= len(lines) を前提に、各段落の先頭ランへ1行ずつ入れる。
    段落が足りない場合は最後の段落の書式を複製して追加する。"""
    tf = shape.text_frame
    paras = tf.paragraphs
    while len(paras) < len(lines):
        new_p_el = copy.deepcopy(paras[-1]._p)
        paras[-1]._p.addnext(new_p_el)
        paras = tf.paragraphs
    for i, line in enumerate(lines):
        p = paras[i]
        if not p.runs:
            p.add_run()
        p.runs[0].text = line
        for r in list(p.runs[1:]):
            r._r.getparent().remove(r._r)
        # 元の段落に改行(Shift+Enter)で複数行が押し込まれているケースがあり、
        # runs[1:] の削除だけでは <a:br/> が残って空行になる。ここで一緒に消す。
        for br in list(p._p.findall(qn("a:br"))):
            br.getparent().remove(br)
    for extra in list(paras[len(lines):]):
        extra._p.getparent().remove(extra._p)


def find_shape(slide, predicate):
    for sh in slide.shapes:
        if sh.has_text_frame and predicate(sh.text_frame.text):
            return sh
    return None


# ---------- コンテンツ編集 ----------

def edit_cover(cover, month_label):
    """month_label 例: '2026年9月'"""
    edit_cover_custom(cover, f"{month_label}　LINEOAトレンドレポート　")


def edit_cover_custom(cover, full_text):
    """表紙タイトルを丸ごと差し替える（トレンドレポート以外の単発デッキ用）。"""
    sh = find_shape(cover, lambda t: "LINEOAトレンドレポート" in t)
    p = sh.text_frame.paragraphs[0]
    p.runs[0].text = full_text
    for r in list(p.runs[1:]):
        r._r.getparent().remove(r._r)


def edit_agenda(agenda, topics):
    """topics: [{"title": "..."}] の順で P.3, P.4, ... を振る。ページ番号は
    表紙1・目次1・本編なので3から開始。"""
    sh = find_shape(agenda, lambda t: t.strip() != "" and "目次" not in t)
    lines = [f"P.{i + 3}　　{t['title']}" for i, t in enumerate(topics)]
    set_multiline_text(sh, lines)


def set_tag_highlight(slide, category_label):
    """4つの分類タグのうち category_label に一致するものだけハイライト
    （塗り=TEXT_2テーマ色・枠=実線）にし、他は非選択（白塗り・点線）に戻す。"""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        text = sh.text_frame.text.replace("\n", "").replace("\x0b", "")
        if text not in [c.replace("\n", "") for c in CATEGORIES]:
            continue
        selected = (text == category_label.replace("\n", ""))
        sh.fill.solid()
        if selected:
            sh.fill.fore_color.theme_color = MSO_THEME_COLOR.TEXT_2
            sh.line.color.rgb = TAG_BORDER_BLUE
            sh.line.dash_style = MSO_LINE_DASH_STYLE.SOLID
        else:
            sh.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
            sh.line.color.rgb = TAG_BORDER_BLUE
            sh.line.dash_style = MSO_LINE_DASH_STYLE.SQUARE_DOT


def find_body_area(slide):
    """ひな形の実施日テーブルの位置を、本文エリア（左・上・全幅）として使う。
    テーブル自体はもう使わないのでここで消す。"""
    old_table_shape = None
    for sh in slide.shapes:
        if getattr(sh, "has_table", False):
            old_table_shape = sh
            break
    left, top = old_table_shape.left, old_table_shape.top
    old_table_shape._element.getparent().remove(old_table_shape._element)
    width, height = Inches(10.2), Inches(3.75)
    return left, top, width, height


def add_image(slide, left, top, box_w, box_h, image_path):
    """指定エリアに収まるよう縦横比を保ってスケールし、中央寄せで画像を配置する。"""
    from PIL import Image
    with Image.open(image_path) as im:
        img_w, img_h = im.size
    box_ratio = box_w / box_h
    img_ratio = img_w / img_h
    if img_ratio > box_ratio:
        w = box_w
        h = int(box_w / img_ratio)
    else:
        h = box_h
        w = int(box_h * img_ratio)
    left_off = left + (box_w - w) // 2
    top_off = top + (box_h - h) // 2
    slide.shapes.add_picture(image_path, left_off, top_off, width=w, height=h)


def _label_chip(slide, x, y, w, h, text, fill, text_color):
    chip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    chip.fill.solid()
    chip.fill.fore_color.rgb = fill
    chip.line.fill.background()
    chip.shadow.inherit = False
    tf = chip.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = text_color
    r.font.name = "メイリオ"


def _content_box(slide, x, y, w, h, lines, fill, border, big_bold_first=False):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    box.line.color.rgb = border
    box.line.width = Pt(1)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.15)
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.1
        if i > 0:
            p.space_before = Pt(5)
        r = p.add_run()
        r.text = line
        if i == 0 and big_bold_first:
            r.font.size = Pt(16) if len(lines) > 1 else Pt(18)
            r.font.bold = True
            r.font.color.rgb = NAVY
        else:
            r.font.size = Pt(10.5)
            r.font.color.rgb = BODY_GRAY
        r.font.name = "メイリオ"


def _merit_card(slide, x, y, w, h, headline, desc):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.adjustments[0] = 0.06
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_BLUE
    card.line.color.rgb = ACCENT
    card.line.width = Pt(1)
    card.shadow.inherit = False
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.2)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    r0 = p0.add_run()
    r0.text = headline
    r0.font.size = Pt(17)
    r0.font.bold = True
    r0.font.color.rgb = NAVY
    r0.font.name = "メイリオ"
    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.LEFT
    p1.space_before = Pt(6)
    p1.line_spacing = 1.15
    r1 = p1.add_run()
    r1.text = desc
    r1.font.size = Pt(11)
    r1.font.color.rgb = BODY_GRAY
    r1.font.name = "メイリオ"


def render_before_after(slide, x, y, w, h, data):
    """AS-IS/TO-BEの比較（上段・コンパクト）＋メリット2枚（下段・主役）。
    2026-09-17、③プロモーションスタンプを「アップデート内容」「インパクト」の
    2ページに分割した際に確定したレイアウト。data のキー：
    as_is（1行）、to_be_lines（1〜2行）、merits（[(headline, desc), ...] 2件）。"""
    top_h = Inches(1.0)
    label_h = Inches(0.3)
    arrow_w = Inches(0.5)
    box_w = (w - arrow_w) // 2

    _label_chip(slide, x, y, box_w, label_h, "AS-IS", LIGHT_GRAY, BODY_GRAY)
    _content_box(slide, x, y + label_h, box_w, top_h - label_h,
                 [data["as_is"]], LIGHT_GRAY, BORDER_GRAY, big_bold_first=True)

    atb = slide.shapes.add_textbox(x + box_w, y, arrow_w, top_h)
    atf = atb.text_frame
    atf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ap = atf.paragraphs[0]
    ap.alignment = PP_ALIGN.CENTER
    ap.text = "▶"
    for r in ap.runs:
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = ACCENT

    x2 = x + box_w + arrow_w
    _label_chip(slide, x2, y, box_w, label_h, "TO-BE", NAVY, WHITE)
    _content_box(slide, x2, y + label_h, box_w, top_h - label_h,
                 data["to_be_lines"], LIGHT_BLUE, ACCENT, big_bold_first=True)

    merit_y = y + top_h + Inches(0.25)
    merit_h = Inches(1.5)
    gap = Inches(0.2)
    card_w = (w - gap) // 2
    for i, (headline, desc) in enumerate(data["merits"][:2]):
        _merit_card(slide, x + i * (card_w + gap), merit_y, card_w, merit_h, headline, desc)


def render_impact(slide, x, y, w, h, data):
    """数字カード（4枚）＋内訳テーブル＋活用メリット。data のキー：
    stats（[(value, label, note), ...] 4件）、
    breakdown_headers・breakdown_values（同じ長さの配列、無ければ省略可）、
    merit_headline・merit_desc。"""
    gap = Inches(0.15)
    card_h = Inches(1.5)
    stats = data["stats"]
    card_w = (w - gap * (len(stats) - 1)) // len(stats)
    cx = x
    for value, label, note in stats:
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, y, card_w, card_h)
        card.adjustments[0] = 0.06
        card.fill.solid()
        card.fill.fore_color.rgb = LIGHT_BLUE
        card.line.color.rgb = ACCENT
        card.line.width = Pt(1)
        card.shadow.inherit = False
        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.1)
        tf.margin_right = Inches(0.1)
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        r0 = p0.add_run()
        r0.text = value
        r0.font.size = Pt(26)
        r0.font.bold = True
        r0.font.color.rgb = NAVY
        r0.font.name = "メイリオ"
        p1 = tf.add_paragraph()
        p1.alignment = PP_ALIGN.CENTER
        p1.space_before = Pt(2)
        r1 = p1.add_run()
        r1.text = label
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = ACCENT
        r1.font.name = "メイリオ"
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(2)
        r2 = p2.add_run()
        r2.text = note
        r2.font.size = Pt(8)
        r2.font.color.rgb = BODY_GRAY
        r2.font.name = "メイリオ"
        cx += card_w + gap

    y_after = y + card_h
    headers = data.get("breakdown_headers")
    values = data.get("breakdown_values")
    if headers and values:
        ty = y_after + Inches(0.2)
        th = Inches(0.9)
        gframe = slide.shapes.add_table(2, len(headers), x, ty, w, th)
        table = gframe.table
        for c, htext in enumerate(headers):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = htext
            r.font.size = Pt(9)
            r.font.bold = True
            r.font.color.rgb = WHITE
            r.font.name = "メイリオ"
        for c, vtext in enumerate(values):
            cell = table.cell(1, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_BLUE
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = vtext
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = NAVY
            r.font.name = "メイリオ"
        y_after = ty + th

    y3 = y_after + Inches(0.2)
    _merit_card(slide, x, y3, w, Inches(0.85), data["merit_headline"], data["merit_desc"])


def render_flow(slide, x, y, w, h, data):
    """Nステップの横矢印フロー（申込み〜掲載までの流れ等）＋補足メモ1枚。data のキー：
    steps（[(day_label, desc), ...]）、note_headline・note_desc（任意）。
    2026-09-17、⑥LINEオープンキャンペーンのスケジュール専用スライドで追加。"""
    steps = data["steps"]
    n = len(steps)
    arrow_w = Inches(0.35)
    step_h = Inches(1.5)
    label_h = Inches(0.4)
    box_w = (w - arrow_w * (n - 1)) // n

    cx = x
    for i, (day_label, desc) in enumerate(steps):
        _label_chip(slide, cx, y, box_w, label_h, day_label, NAVY, WHITE)
        _content_box(slide, cx, y + label_h, box_w, step_h - label_h,
                     [desc], LIGHT_BLUE, ACCENT, big_bold_first=False)
        cx += box_w
        if i < n - 1:
            atb = slide.shapes.add_textbox(cx, y, arrow_w, step_h)
            atf = atb.text_frame
            atf.vertical_anchor = MSO_ANCHOR.MIDDLE
            ap = atf.paragraphs[0]
            ap.alignment = PP_ALIGN.CENTER
            ap.text = "▶"
            for r in ap.runs:
                r.font.size = Pt(16)
                r.font.bold = True
                r.font.color.rgb = ACCENT
            cx += arrow_w

    if data.get("note_headline"):
        note_y = y + step_h + Inches(0.25)
        _merit_card(slide, x, note_y, w, Inches(1.0), data["note_headline"], data["note_desc"])


def render_grid_notes(slide, x, y, w, h, data):
    """注意点・条件を2×2のカードで見せる。data のキー：items（[(headline, desc), ...] 最大4件）。
    2026-09-17、⑥LINEオープンキャンペーンの注意点専用スライドで追加。"""
    items = data["items"][:4]
    gap_x = Inches(0.2)
    gap_y = Inches(0.2)
    card_w = (w - gap_x) // 2
    card_h = min((h - gap_y) // 2, Inches(1.6))
    for i, (headline, desc) in enumerate(items):
        col, row = i % 2, i // 2
        cx = x + col * (card_w + gap_x)
        cy = y + row * (card_h + gap_y)
        _merit_card(slide, cx, cy, card_w, card_h, headline, desc)


def add_image_placeholder(slide, left, top, width, height, caption):
    """画像を後から差し込むための空きスペース（破線の枠＋説明キャプション）を置く。
    画像が届いたらこの枠を picture に差し替える。"""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
    box.line.color.rgb = RGBColor(0xA6, 0xA6, 0xA6)
    box.line.width = Pt(1.25)
    box.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run()
    r1.text = "画像挿入予定"
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    r1.font.name = "メイリオ"
    if caption:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(6)
        r2 = p2.add_run()
        r2.text = caption
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        r2.font.name = "メイリオ"


def edit_topic_slide(slide, data):
    banner = find_shape(slide, lambda t: t.startswith("【"))
    set_single_run_text(banner, f"【{data['bracket']}】{data['headline']}")

    # 説明バンド＝3段落・中央揃えのシェイプ（ひな形固有の構造で特定）
    desc = None
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        paras = sh.text_frame.paragraphs
        if len(paras) == 3 and paras[0].alignment == PP_ALIGN.CENTER:
            desc = sh
            break
    set_multiline_text(desc, data["description_lines"])

    # 対象業界の値ボックス＝ひな形の元テキストで特定
    target = find_shape(slide, lambda t: t.strip() == "友だち獲得を強化したい全業種")
    if target is not None:
        set_single_run_text(target, data["industry"])

    src = find_shape(slide, lambda t: t.startswith("出典："))
    if src is not None:
        set_single_run_text(src, f"出典：{data['source']}")

    sec_header = find_shape(slide, lambda t: t.strip() == "■ 施策スケジュール")
    if sec_header is not None:
        set_single_run_text(sec_header, "■ 詳細")

    # ヘッダー見出し「今後のアップデート情報」＝左上の紺色の角アイコンに寄りすぎているので右へ寄せる
    header = find_shape(slide, lambda t: "今後のアップデート情報" in t)
    if header is not None:
        header.left = header.left + Inches(0.2)

    set_tag_highlight(slide, data["category"])
    body_left, body_top, body_width, body_height = find_body_area(slide)
    render = data.get("render")
    if render == "before_after":
        render_before_after(slide, body_left, body_top, body_width, body_height, data)
    elif render == "impact":
        render_impact(slide, body_left, body_top, body_width, body_height, data)
    elif render == "flow":
        render_flow(slide, body_left, body_top, body_width, body_height, data)
    elif render == "grid_notes":
        render_grid_notes(slide, body_left, body_top, body_width, body_height, data)
    else:
        image_path = data.get("image_path")
        if image_path and Path(image_path).exists():
            add_image(slide, body_left, body_top, body_width, body_height, image_path)
        else:
            add_image_placeholder(
                slide, body_left, body_top, body_width, body_height, data.get("image_caption", "")
            )


def force_font(prs, font_name="メイリオ"):
    """全スライドの全テキスト（表のセルも含む）のフォントを font_name に統一する。
    ひな形（FMT）由来のランには "Meiryo"（英語表記）や未指定（テーマ既定）が
    混ざっているため、2026-09-17、殿村さんの指示で明示的に揃えることにした。"""
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.name = font_name
            if getattr(sh, "has_table", False):
                for row in sh.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                run.font.name = font_name


def build(out_path, month_label, topics):
    """topics: dict のリスト。各要素は
    {title, bracket, headline, industry, category, description_lines(3行),
     overview, pricing, schedule, caution, source}
    title は目次に出すページタイトル（通常は【bracket】headline と同じでよい）。
    """
    prs, cover, agenda, topic_slides = load_base(out_path, len(topics))

    edit_cover(cover, month_label)
    edit_agenda(agenda, topics)
    for slide, data in zip(topic_slides, topics):
        edit_topic_slide(slide, data)

    force_font(prs)

    prs.save(out_path)
    print(f"saved: {out_path}")


def build_standalone(out_path, cover_title, topics):
    """トレンドレポートと同じトンマナ（DYM_LINEOA_TREND_FMT.pptx）で、
    特定の案件・キャンペーン専用の単発デッキを作る。topics の形式は build() と同じ。
    表紙タイトルだけ自由文言にできる（月次ラベル形式を強制しない）。"""
    prs, cover, agenda, topic_slides = load_base(out_path, len(topics))

    edit_cover_custom(cover, cover_title)
    edit_agenda(agenda, topics)
    for slide, data in zip(topic_slides, topics):
        edit_topic_slide(slide, data)

    force_font(prs)

    prs.save(out_path)
    print(f"saved: {out_path}")
