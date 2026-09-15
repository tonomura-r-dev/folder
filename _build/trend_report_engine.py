"""
LINE公式アカウント 媒体最新情報 トレンドレポート 共通エンジン。

トンマナは実際のDYM営業資料FMT（_templates/DYM_LINEOA_BUFFF_FMT.pptx）から抽出。
このFMTをコピーし、必要な枚数だけ残して clear_slide() → 作り直す経路
（build_special_plan.py と同じ。スライドの新規追加はしない＝CLAUDE.md参照）。

ヘッダーのネイビー角アイコン・区切り線・DYMロゴ・フッターの著作権表記／ページ番号は
スライドレイアウト側に定義されており、clear_slide() で個々のシェイプを消しても
自動的に継承表示される（要検証済み：2026-09、test_inherit.jpg）。
なので自前では描画しない。

月ごとのビルドスクリプトはこのモジュールを import し、TOPICS_LIST / TOPICS /
OUT / MONTH_LABEL だけを定義して build() を呼ぶ（_build/build_trend_report_202609.py 参照）。
毎月の作業手順・データの集め方は .claude/skills/trend-report/SKILL.md を参照。
"""

import shutil
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
SRC = str(ROOT / "_templates" / "DYM_LINEOA_BUFFF_FMT.pptx")

FONT = "メイリオ"

# 実FMTから抽出した配色（2026-09、LINEOA_BUFFF_3.pptxを解析）
TITLE_NAVY = RGBColor(0x00, 0x20, 0x60)   # ヘッダー見出し・本文タイトル色
CHIP_NAVY = RGBColor(0x10, 0x25, 0x3F)    # セクション見出しチップ・フラグチップ
BODY_GRAY = RGBColor(0x34, 0x34, 0x34)    # 本文
META_GRAY = RGBColor(0x49, 0x48, 0x48)    # メタ情報
RED = RGBColor(0xFF, 0x00, 0x00)          # 強調・注意（FMT実測）
LIGHT_BOX = RGBColor(0xF2, 0xF2, 0xF2)    # メインメッセージ等のボックス塗り
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
SOURCE_GRAY = RGBColor(0x80, 0x80, 0x80)

SLIDE_W = Inches(10.83)
SLIDE_H = Inches(7.5)

# ヘッダーの角アイコン（レイアウト側で幅0.49in・高さ0.6in）と被らない開始位置
CONTENT_LEFT = Inches(0.62)
CONTENT_RIGHT = Inches(10.33)
CONTENT_W = CONTENT_RIGHT - CONTENT_LEFT


def set_font_all(text_frame, name=FONT):
    for para in text_frame.paragraphs:
        for run in para.runs:
            run.font.name = name
            rPr = run._r.get_or_add_rPr()
            for tag in ("a:latin", "a:ea", "a:cs"):
                el = rPr.find(qn(tag))
                if el is None:
                    el = rPr.makeelement(qn(tag), {})
                    rPr.append(el)
                el.set("typeface", name)


def clear_slide(slide):
    """スライド自身が持つシェイプだけを消す。レイアウト/マスター側の
    ヘッダーアイコン・区切り線・DYMロゴ・フッターは継承表示のため残る。"""
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)


def load_trimmed_fmt(out_path, n_slides):
    """FMTをコピーし、先頭n_slides枚だけ残す（追加はしない）。"""
    shutil.copyfile(SRC, out_path)
    prs = Presentation(out_path)
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    keep_ids = ids[:n_slides]
    for sldId in ids:
        if sldId in keep_ids:
            continue
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
    slides = list(prs.slides)
    assert len(slides) == n_slides, len(slides)
    for slide in slides:
        clear_slide(slide)
    return prs, slides


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = line_w or Pt(0.75)
    return shp


def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=None, word_wrap=True):
    """runs: list of paragraphs; each paragraph is a list of (text, size, bold, color) tuples."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = word_wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, para_runs in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        for text, size, bold, color in para_runs:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
    set_font_all(tf)
    return box


def add_bullets(slide, x, y, w, h, items, size=10, color=BODY_GRAY, bullet="・",
                 line_spacing=1.2):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line_spacing
        p.space_after = Pt(3)
        r = p.add_run()
        r.text = f"{bullet}{item}"
        r.font.size = Pt(size)
        r.font.color.rgb = color
    set_font_all(tf)
    return box


def chip(slide, x, y, w, h, text, fill=CHIP_NAVY, color=WHITE, size=9.5, bold=True):
    c = add_rect(slide, x, y, w, h, fill=fill)
    add_text(slide, x, y, w, h, [[(text, size, bold, color)]],
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return c


def page_title(slide, text, size=17):
    """ヘッダー見出し（レイアウト側の角アイコンの右、区切り線の上）。"""
    add_text(slide, CONTENT_LEFT, Inches(0.1), Inches(8.1), Inches(0.42),
              [[(text, size, True, TITLE_NAVY)]], anchor=MSO_ANCHOR.MIDDLE)


def source_line(slide, url):
    if not url:
        return
    add_text(slide, CONTENT_LEFT, Inches(6.8), Inches(9.5), Inches(0.22),
              [[(f"出典：{url}", 7.5, False, SOURCE_GRAY)]], anchor=MSO_ANCHOR.MIDDLE)


def cover_slide(slide, month_label):
    """month_label 例: '2026年9月'（"年"区切りで年と月を分けてタイトルに使う）"""
    _, month_only = month_label.split("年")
    add_text(slide, Inches(0.7), Inches(2.5), Inches(9), Inches(0.4),
              [[("LINE公式アカウント", 14, False, META_GRAY)]])
    add_text(slide, Inches(0.7), Inches(2.85), Inches(9.3), Inches(1.0),
              [[(f"媒体最新情報　{month_only} トレンドレポート", 30, True, TITLE_NAVY)]],
              line_spacing=1.15)
    add_rect(slide, Inches(0.7), Inches(3.95), Inches(9.0), Pt(1.5), fill=TITLE_NAVY)
    chip(slide, Inches(0.7), Inches(4.15), Inches(2.4), Inches(0.4), f"{month_label}号",
         fill=LIGHT_BOX, color=META_GRAY, size=11, bold=False)
    add_text(slide, Inches(0.7), Inches(6.5), Inches(6), Inches(0.4),
              [[("株式会社DYM（DYM × LINEOA）", 12, True, TITLE_NAVY)]])


def agenda_slide(slide, topics):
    """topics: 表示用に番号込みで整形済みの文字列リスト（例 '1. ○○'）"""
    page_title(slide, "目次")
    y = Inches(1.0)
    row_h = Inches(0.62)
    for i, t in enumerate(topics, start=1):
        chip(slide, CONTENT_LEFT, y, Inches(0.44), Inches(0.42), str(i), size=12)
        add_text(slide, CONTENT_LEFT + Inches(0.62), y, Inches(9.0), Inches(0.42),
                  [[(t, 13, False, BODY_GRAY)]], anchor=MSO_ANCHOR.MIDDLE)
        y += row_h


def topic_slide(slide, data):
    """data のキー：title, message, overview[], pricing[], schedule[], caution[], source"""
    page_title(slide, data["title"])

    # メインメッセージ
    msg_y = Inches(0.75)
    add_rect(slide, CONTENT_LEFT, msg_y, CONTENT_W, Inches(0.55), fill=LIGHT_BOX)
    add_text(slide, CONTENT_LEFT + Inches(0.2), msg_y, CONTENT_W - Inches(0.4), Inches(0.55),
              [[(data["message"], 12, True, TITLE_NAVY)]],
              anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)

    # 4区分ボックス（2x2）
    box_top = Inches(1.45)
    box_h = Inches(2.55)
    gap = Inches(0.15)
    hgap = Inches(0.2)
    box_w = (CONTENT_W - hgap) / 2
    positions = [
        (CONTENT_LEFT, box_top),
        (CONTENT_LEFT + box_w + hgap, box_top),
        (CONTENT_LEFT, box_top + box_h + gap),
        (CONTENT_LEFT + box_w + hgap, box_top + box_h + gap),
    ]
    sections = [
        ("概要・変更点", data.get("overview", [])),
        ("料金体系・出稿条件", data.get("pricing", [])),
        ("スケジュール・導入フロー", data.get("schedule", [])),
        ("注意点・影響", data.get("caution", [])),
    ]
    chip_h = Inches(0.34)
    for (x, y), (label, items) in zip(positions, sections):
        add_rect(slide, x, y, box_w, box_h, fill=WHITE, line=BORDER_GRAY, line_w=Pt(0.75))
        chip(slide, x, y, box_w, chip_h, label, size=10.5)
        items = items or ["－"]
        # 項目数が多い箱は自動でフォントを詰めて収める
        size, spacing = (8.5, 1.1) if len(items) >= 6 else (9, 1.15) if len(items) >= 5 else (10, 1.2)
        add_bullets(slide, x + Inches(0.18), y + chip_h + Inches(0.1),
                    box_w - Inches(0.36), box_h - chip_h - Inches(0.2), items,
                    size=size, line_spacing=spacing)

    source_line(slide, data.get("source"))


def build(out_path, month_label, topics_list, topics):
    """月次スクリプトから呼ぶエントリーポイント。
    topics_list: 目次に出す番号付き文字列のリスト
    topics: topic_slide() に渡す dict のリスト（topics_listと同じ順・同じ件数）
    """
    n_slides = 2 + len(topics)  # 表紙1 + 目次1 + トピックN
    prs, slides = load_trimmed_fmt(out_path, n_slides)

    cover_slide(slides[0], month_label)
    agenda_slide(slides[1], topics_list)

    for slide, data in zip(slides[2:], topics):
        topic_slide(slide, data)

    prs.save(out_path)
    print(f"saved: {out_path}")
