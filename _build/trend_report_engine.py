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

TABLE_LABEL_NAVY = RGBColor(0x15, 0x13, 0x3D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BODY_GRAY = RGBColor(0x33, 0x33, 0x33)
BORDER_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
TAG_BORDER_BLUE = RGBColor(0x4F, 0x81, 0xBD)

SECTIONS = ["概要・変更点", "料金体系・出稿条件", "スケジュール・導入フロー", "注意点・影響"]


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
    sh = find_shape(cover, lambda t: "LINEOAトレンドレポート" in t)
    p = sh.text_frame.paragraphs[0]
    p.runs[0].text = f"{month_label}　LINEOAトレンドレポート　"
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


IMAGE_GAP = Inches(0.15)
IMAGE_WIDTH = Inches(2.7)


def rebuild_body_table(slide, data):
    """ひな形の実施日テーブル(3列×7行)を消して、4行×2列の
    区分/内容 表に差し替える。右端は画像プレースホルダー分だけ幅を空けておく。"""
    old_table_shape = None
    for sh in slide.shapes:
        if getattr(sh, "has_table", False):
            old_table_shape = sh
            break
    left, top = old_table_shape.left, old_table_shape.top
    full_width, height = Inches(10.2), Inches(3.5)
    width = full_width - IMAGE_GAP - IMAGE_WIDTH
    old_table_shape._element.getparent().remove(old_table_shape._element)

    rows, cols = 4, 2
    gframe = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = gframe.table
    table.columns[0].width = Inches(1.9)
    table.columns[1].width = width - Inches(1.9)

    sections = [
        (SECTIONS[0], data.get("overview", [])),
        (SECTIONS[1], data.get("pricing", [])),
        (SECTIONS[2], data.get("schedule", [])),
        (SECTIONS[3], data.get("caution", [])),
    ]

    # 行の高さを内容量（文字数）に比例配分する。「料金体系・出稿条件」「スケジュール・
    # 導入フロー」は「－」等の空回答が多く、他の区分と同じ高さだと余白ばかりになるため、
    # 情報が少ない行は縮め、浮いた分を情報が多い行（概要・注意点など）に回す。
    # フォントサイズは変えない（行の高さだけの調整なので、はみ出しの心配がない）。
    lens = [max(sum(len(x) for x in items), 15) for _, items in sections]
    total_w = sum(lens)
    min_h = Inches(0.5)
    raw_heights = [height * (w / total_w) for w in lens]
    raw_heights = [max(h, min_h) for h in raw_heights]
    overflow = sum(raw_heights) - height
    if overflow > 0:
        above_min = [i for i, h in enumerate(raw_heights) if h > min_h]
        above_sum = sum(raw_heights[i] - min_h for i in above_min)
        if above_sum > 0:
            for i in above_min:
                share = (raw_heights[i] - min_h) / above_sum
                raw_heights[i] -= int(overflow * share)
    for r, h in enumerate(raw_heights):
        table.rows[r].height = int(h)

    for r, (label, items) in enumerate(sections):
        label_cell = table.cell(r, 0)
        label_cell.fill.solid()
        label_cell.fill.fore_color.rgb = TABLE_LABEL_NAVY
        label_cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        label_cell.margin_left = Inches(0.08)
        label_cell.margin_right = Inches(0.05)
        tf = label_cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = label
        for run in p.runs:
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = WHITE
            run.font.name = "メイリオ"

        content_cell = table.cell(r, 1)
        content_cell.fill.solid()
        content_cell.fill.fore_color.rgb = WHITE
        content_cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        content_cell.margin_left = Inches(0.1)
        content_cell.margin_right = Inches(0.1)
        content_cell.margin_top = Inches(0.04)
        content_cell.margin_bottom = Inches(0.04)
        ctf = content_cell.text_frame
        ctf.word_wrap = True
        items = items or ["－"]
        total_len = sum(len(x) for x in items)
        # 画像プレースホルダー分だけ列幅が狭くなった（同じ文字数でも折り返しが増える）ので、
        # 閾値を厳しめにしてある。
        if total_len > 150:
            size, spacing, space_after = 8.5, 1.0, 1
        elif total_len > 90:
            size, spacing, space_after = 9.5, 1.05, 1.5
        else:
            size, spacing, space_after = 10.5, 1.1, 2
        for i, item in enumerate(items):
            p = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
            p.line_spacing = spacing
            p.space_after = Pt(space_after)
            r_ = p.add_run()
            r_.text = f"・{item}"
            r_.font.size = Pt(size)
            r_.font.color.rgb = BODY_GRAY
            r_.font.name = "メイリオ"

    return left, top, width, height


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
    t_left, t_top, t_width, t_height = rebuild_body_table(slide, data)
    add_image_placeholder(
        slide,
        t_left + t_width + IMAGE_GAP,
        t_top,
        IMAGE_WIDTH,
        t_height,
        data.get("image_caption", ""),
    )


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

    prs.save(out_path)
    print(f"saved: {out_path}")
