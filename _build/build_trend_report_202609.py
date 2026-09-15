"""
LINE公式アカウント 媒体最新情報 2026年9月 トレンドレポート ビルドスクリプト。

_drafts/トレンドレポート_2026年9月_構成案.md の内容をPPTX化する。
対象業界・フェーズ・フラグは元データに記載が無いため空欄のまま出力する。

DYM汎用FMT（_templates/DYM_LINEOA_FMT.pptx）をコピーし、必要な11枚だけ残して
clear_slide() → 作り直す経路（build_special_plan.py と同じ）。
スライドの新規追加はしない（python-pptxで壊れるため。CLAUDE.md参照）。

使い方: python3 _build/build_trend_report_202609.py
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
SRC = str(ROOT / "_templates" / "DYM_LINEOA_FMT.pptx")
OUT = str(ROOT / "LINEOA媒体最新情報_2026年9月トレンドレポート.pptx")

FONT = "メイリオ"

NAVY = RGBColor(0x15, 0x13, 0x3D)
SUB_BLUE = RGBColor(0xC5, 0xD8, 0xF1)
SUB_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
RED = RGBColor(0xC0, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_GRAY = RGBColor(0x33, 0x33, 0x33)
BORDER_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
PLACEHOLDER_GRAY = RGBColor(0x99, 0x99, 0x99)

SLIDE_W = Inches(10.83)
SLIDE_H = Inches(7.5)

N_SLIDES = 11  # 表紙1 + 目次1 + トピック9


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
    """FMT側の残骸シェイプを全部消して、まっさらな状態にする。"""
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split("}")[-1] in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            spTree.remove(el)


def load_trimmed_fmt(n_slides):
    """FMTをコピーし、先頭n_slides枚だけ残す（追加はしない）。"""
    shutil.copyfile(SRC, OUT)
    prs = Presentation(OUT)
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


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None, shadow=False):
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


def add_bullets(slide, x, y, w, h, items, size=11, color=TEXT_GRAY, bullet="・",
                 line_spacing=1.15, bold_first=False):
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


def header(slide, flag_agency=None, flag_noshare=None):
    """共通ヘッダー：ネイビー帯 + 「9月 媒体最新情報」+ 属性フラグ（空欄可）。"""
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.62), fill=NAVY)
    add_text(slide, Inches(0.4), Inches(0.06), Inches(5), Inches(0.5),
              [[("9月 媒体最新情報", 15, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)

    chip_x = Inches(10.75)
    for label, flag in [("展開禁止", flag_noshare), ("代理店限定", flag_agency)]:
        text = flag if flag else "－"
        chip_w = Inches(1.55)
        chip_x -= chip_w + Inches(0.12)
        add_rect(slide, chip_x, Inches(0.13), chip_w, Inches(0.36),
                 fill=SUB_GRAY, line=BORDER_GRAY, line_w=Pt(0.75))
        add_text(slide, chip_x, Inches(0.13), chip_w, Inches(0.36),
                  [[(f"{label}：{text}", 9, True, NAVY)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def footer(slide, source_url):
    add_rect(slide, 0, Inches(7.16), SLIDE_W, Inches(0.34), fill=SUB_GRAY)
    add_text(slide, Inches(0.4), Inches(7.18), Inches(8.5), Inches(0.3),
              [[(f"出典：{source_url}" if source_url else "", 8, False, RGBColor(0x66, 0x66, 0x66))]],
              anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(6.5), Inches(7.18), Inches(4.13), Inches(0.3),
              [[("Copyright(c) DYM Co., Ltd. All Rights Reserved.", 7.5, False, RGBColor(0x66, 0x66, 0x66))]],
              align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def subheader(slide, industry=None, phase=None):
    """対象業界／フェーズ帯。値が無ければ空欄（－）のまま表示。"""
    y = Inches(0.68)
    add_rect(slide, 0, y, SLIDE_W, Inches(0.32), fill=SUB_BLUE)
    ind = industry or "－"
    ph = phase or "－"
    add_text(slide, Inches(0.4), y, Inches(5.5), Inches(0.32),
              [[("対象業界：", 9.5, True, NAVY), (ind, 9.5, False, NAVY)]],
              anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(6.5), y, Inches(3.9), Inches(0.32),
              [[("マーケティングフェーズ：", 9.5, True, NAVY), (ph, 9.5, False, NAVY)]],
              anchor=MSO_ANCHOR.MIDDLE)


def cover_slide(slide, theme):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=NAVY)
    add_rect(slide, 0, Inches(4.55), SLIDE_W, Inches(0.03), fill=SUB_BLUE)

    add_text(slide, Inches(0.9), Inches(2.55), Inches(9), Inches(1.6),
              [[("LINE公式アカウント", 24, True, WHITE)],
               [("媒体最新情報　9月 トレンドレポート", 30, True, WHITE)]],
              line_spacing=1.15)

    add_text(slide, Inches(0.9), Inches(4.75), Inches(9), Inches(0.4),
              [[(theme or "－", 13, False, SUB_BLUE)]])

    add_text(slide, Inches(0.9), Inches(6.7), Inches(6), Inches(0.5),
              [[("株式会社DYM（DYM × LINEOA）", 12, True, WHITE)]])
    return slide


def agenda_slide(slide, topics):
    header(slide)
    add_text(slide, Inches(0.4), Inches(1.05), Inches(8), Inches(0.5),
              [[("目次", 20, True, NAVY)]])
    add_rect(slide, Inches(0.4), Inches(1.62), Inches(1.2), Pt(2.5), fill=RED)

    y = Inches(2.0)
    row_h = Inches(0.52)
    for i, t in enumerate(topics, start=1):
        add_rect(slide, Inches(0.4), y, Inches(0.5), Inches(0.4), fill=SUB_BLUE)
        add_text(slide, Inches(0.4), y, Inches(0.5), Inches(0.4),
                  [[(str(i), 13, True, NAVY)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.1), y, Inches(9.3), Inches(0.4),
                  [[(t, 13, False, TEXT_GRAY)]], anchor=MSO_ANCHOR.MIDDLE)
        y += row_h
    footer(slide, "")
    return slide


def topic_slide(slide, data):
    header(slide, flag_agency=data.get("flag_agency"), flag_noshare=data.get("flag_noshare"))
    subheader(slide, industry=data.get("industry"), phase=data.get("phase"))

    # タイトル
    add_text(slide, Inches(0.4), Inches(1.1), Inches(10.0), Inches(0.6),
              [[(data["title"], 16, True, NAVY)]], line_spacing=1.05)

    # メインメッセージ
    msg_y = Inches(1.75)
    add_rect(slide, Inches(0.4), msg_y, Inches(10.03), Inches(0.75),
             fill=SUB_BLUE)
    add_text(slide, Inches(0.65), msg_y, Inches(9.6), Inches(0.75),
              [[(data["message"], 11.5, True, NAVY)]],
              anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)

    # 4区分ボックス（2x2）
    box_top = Inches(2.72)
    box_h = Inches(2.1)
    gap = Inches(0.2)
    box_w = (SLIDE_W - Inches(0.8) - gap) / 2
    positions = [
        (Inches(0.4), box_top),
        (Inches(0.4) + box_w + gap, box_top),
        (Inches(0.4), box_top + box_h + gap),
        (Inches(0.4) + box_w + gap, box_top + box_h + gap),
    ]
    sections = [
        ("概要・変更点", data.get("overview", [])),
        ("料金体系・出稿条件", data.get("pricing", [])),
        ("スケジュール・導入フロー", data.get("schedule", [])),
        ("注意点・影響", data.get("caution", [])),
    ]
    for (x, y), (label, items) in zip(positions, sections):
        add_rect(slide, x, y, box_w, box_h, fill=WHITE, line=BORDER_GRAY, line_w=Pt(0.75))
        add_rect(slide, x, y, box_w, Inches(0.36), fill=NAVY)
        add_text(slide, x + Inches(0.15), y, box_w - Inches(0.3), Inches(0.36),
                  [[(f"■ {label}", 11, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
        items = items or ["－"]
        add_bullets(slide, x + Inches(0.2), y + Inches(0.5), box_w - Inches(0.4),
                    box_h - Inches(0.65), items, size=10.5)

    footer(slide, data.get("source", ""))
    return slide


TOPICS_LIST = [
    "1. ［提供終了］LINE公式アカウントの「リサーチ」機能終了",
    "2. 既読API（旧）新規利用受付終了（リマインド）",
    "3. LINEミニアプリ お気に入り追加促進機能",
    "4. プロモーションスタンプ・絵文字：早期発注特典①の締切",
    "5. LINE公式アカウント「分析」データ参照期間 過去36カ月に変更（告知延期中）",
    "6. プロモーションスタンプ 効果改善アップデート（サジェスト面 無料スタンプ表示条件緩和）",
    "7. 友だち追加オプション 複数アカウント対応",
    "8. LINE Sales Promotion Manager 利用停止期間",
    "9. LINEオープンキャンペーン（アンケート型）26年10-12月 期間限定キャンペーン",
]

TOPICS = [
    dict(
        title="［提供終了］LINE公式アカウントの「リサーチ」機能終了",
        message="2026年12月2日をもって「リサーチ」「アカウント満足度調査」機能が終了。年内に必要なデータの取得・代替運用への切替検討が必要。",
        overview=["LINE公式アカウントの「リサーチ」機能を終了",
                   "「リサーチ」「アカウント満足度調査」の機能が段階的に利用不可に"],
        pricing=["該当なし"],
        schedule=["社外告知：2026年8月17日（月）CMSお知らせ掲載済",
                   "LPP掲載：2026年8月19日（水）／8月26日（水）",
                   "終了予定：2026年12月2日（水）"],
        caution=["リサーチを使ったアンケート施策・満足度調査を運用中のクライアントは代替手段への移行が必要"],
        source="https://workers-hub.box.com/s/mtoua30hxf68prdf5x5xof495fc3e6pj",
    ),
    dict(
        title="既読API（旧）新規利用受付終了（リマインド）",
        message="事前申請不要の新「既読API」が一般提供済み。旧既読API（LBPM経由）の新規申請受付は終了。",
        overview=["新しい「既読API」が一般提供開始（事前の利用申請不要）",
                   "従来のLBPM経由「既読API（旧）」の新規利用申請窓口を停止"],
        pricing=["－"],
        schedule=["新既読API仕様：developers.line.biz/ja/docs/messaging-api/mark-as-read/"],
        caution=["既読API（旧）を利用予定だった案件は新API仕様への切替が必要"],
        source="https://workers-hub.box.com/s/nxgvv2ejax65j6cadz7gdsdalxd8zioa",
    ),
    dict(
        title="LINEミニアプリ お気に入り追加促進機能",
        message="特定条件を満たしたユーザーに、ミニアプリのお気に入り追加を促すポップアップが自動表示される新機能。",
        overview=["トリガー①：マイミニアプリの「履歴」からアクセス",
                   "トリガー②：7日以内に2日以上アクセス",
                   "条件を満たすとお気に入り追加を促すポップアップが表示"],
        pricing=["－"],
        schedule=["－（リリース日は要確認）"],
        caution=["リピート利用ユーザーの再訪導線が強化される。ミニアプリ運用中クライアントへの案内材料になる"],
        source="https://workers-hub.box.com/s/9uo4m0fiqyxmddz5br3xe9691w0q4ht7",
    ),
    dict(
        title="プロモーションスタンプ・絵文字：早期発注特典①の締切",
        message="3月早期発注特典①を適用するには9月30日17時までの発注が必須。中間締切はクリエイターコラボ9/10、企業制作9/24。",
        overview=["3月早期発注特典 特典①適用は9月1日〜9月30日17時までの発注が条件"],
        pricing=["特典②適用の場合は10月30日17時まで"],
        schedule=["クリエイターコラボ：「依頼申請／バリエーション依頼シート提出」を9/10(木)17時までにLBPMで申請",
                   "企業制作：「キャラクター可否審査申請」を9/24(木)17時までにLBPMで申請"],
        caution=["コラボクリエイター候補が未定の企業は締切厳守で要進行",
                   "人気クリエイターは出稿がカニバる可能性があり早めの進行を推奨",
                   "内容によっては上記日程で申請しても発注期日までに間に合わない場合あり"],
        source="https://workers-hub.box.com/s/uw4owpzcmvhcf84ijd66s0kc78izwx0n",
        accent_deadlines=True,
    ),
    dict(
        title="LINE公式アカウント「分析」データ参照期間 過去36カ月に変更",
        message="「分析」タブの参照期間が過去36カ月までに変更予定。※社外告知は延期中、日程未確定につき最新情報を要確認。",
        overview=["Web版管理画面／管理アプリの「分析」タブで確認できる実績値の参照期間を過去36カ月までに変更"],
        pricing=["該当なし"],
        schedule=["リリース日：2026年11月頃（予定）",
                   "社外告知：当初9/2予定だったが内容変更の可能性ありいったん延期。改めて案内予定"],
        caution=["36カ月より前の実績データが必要なクライアントには変更前の確認・保存を案内",
                   "告知日程が流動的なため最新情報を要確認"],
        source="https://workers-hub.box.com/s/ms8iztyci2dn5z2qna88rklxwxp7o5ja",
    ),
    dict(
        title="プロモーションスタンプ 効果改善アップデート（サジェスト面 無料スタンプ表示条件緩和）",
        message="LINEキーボードのサジェスト機能で無料スタンプが表示される条件が緩和され、全ユーザーに表示可能に。DL数は平均+5.63%増加。",
        overview=["サジェスト面での無料スタンプ表示ユーザー条件を緩和し全ユーザーに表示されるよう変更",
                   "配信中の案件を含む全案件に自動適用（広告主・案件側の対応は不要）"],
        pricing=["追加費用なし（自動適用）"],
        schedule=["リリース日：2026年8月28日（金）"],
        caution=["A/Bテスト結果：1案件あたり平均DL数+5.63%増",
                   "直近18ヶ月DL履歴のないユーザーでも約2倍のDL、普段DLしない層への到達も改善",
                   "運用中クライアントへの実績報告・追加提案の材料に使える"],
        source="https://workers-hub.box.com/s/rgfhp6hxyusj7domi1irjx68jld23k9r",
    ),
    dict(
        title="友だち追加オプション 複数アカウント対応",
        message="1つのLINEミニアプリに連携するLINE公式アカウントを複数設定できるようになる。",
        overview=["LINEミニアプリと連携するLINE公式アカウントを複数設定可能に"],
        pricing=["－"],
        schedule=["リリース予定日：9月2日"],
        caution=["複数ブランド・複数拠点アカウントを運用するクライアントの友だち獲得導線が広がる"],
        source="https://workers-hub.box.com/s/ko2jjwgxk5fs18ejrs97fvfqv3bm3089",
    ),
    dict(
        title="LINE Sales Promotion Manager 利用停止期間",
        message="データベース管理システムのスケール最適化のため、LINEで応募の管理画面が一時利用停止となる。",
        overview=["データベース管理システムのスケール最適化実施のため対象期間中は管理画面が利用停止"],
        pricing=["該当なし"],
        schedule=["社外への案内開始日・リリース日：2026年9月2日",
                   "－（停止期間の具体的な日時は要確認）"],
        caution=["停止期間中は応募管理・当選者確認等の作業ができない",
                   "稼働中キャンペーンがあるクライアントには事前周知が必要"],
        source="https://workers-hub.box.com/s/mo3y54qld9vj4kung0c4040z41mfhoip",
    ),
    dict(
        title="LINEオープンキャンペーン（アンケート型）26年10-12月 期間限定キャンペーン",
        message="2026年10月20日〜12月25日に終了する案件限定で、特別価格・特別期間・特別通数の3特典が適用される。",
        overview=["対象：2026年10月20日開始〜12月25日までに終了する案件限定",
                   "景品はLINEポイント"],
        pricing=["特典①特別価格：基本費用（キャンペーン価格）1,500万円",
                   "特典②特別掲載期間：キャンペーン掲載期間15日間",
                   "特典③特別通数：告知メッセージ（1通）の通数150万通増量（最大1,050万通）"],
        schedule=["キャンペーン期間：2026年10月20日〜12月25日終了の案件が対象"],
        caution=["期間限定・条件限定のため、対象クライアントへは早めの提案・申込誘導が必要"],
        source="https://workers-hub.box.com/s/odilqpe44tp398zlfko148yggcgj68ny",
    ),
]


def apply_deadline_accent(slide):
    """スライド4（早期発注特典）だけ、締切日付を赤字強調する。"""
    keywords = ["9月30日17時", "9/10", "9/24", "10月30日17時"]
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if any(k in run.text for k in keywords):
                    run.font.color.rgb = RED
                    run.font.bold = True


def main():
    prs, slides = load_trimmed_fmt(N_SLIDES)

    cover_slide(slides[0], theme=None)
    agenda_slide(slides[1], TOPICS_LIST)

    for slide, data in zip(slides[2:], TOPICS):
        topic_slide(slide, data)
        if data.get("accent_deadlines"):
            apply_deadline_accent(slide)

    prs.save(OUT)
    print(f"saved: {OUT}")


if __name__ == "__main__":
    main()
