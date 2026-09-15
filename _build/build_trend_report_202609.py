"""
LINE公式アカウント 媒体最新情報 2026年9月 トレンドレポート ビルドスクリプト。

トンマナは実際のDYM営業資料FMT（_templates/DYM_LINEOA_BUFFF_FMT.pptx）から抽出。
このFMTをコピーし、必要な8枚だけ残して clear_slide() → 作り直す経路
（build_special_plan.py と同じ。スライドの新規追加はしない＝CLAUDE.md参照）。

ヘッダーのネイビー角アイコン・区切り線・DYMロゴ・フッターの著作権表記／ページ番号は
スライドレイアウト側に定義されており、clear_slide() で個々のシェイプを消しても
自動的に継承表示される（要検証済み：2026-09、test_inherit.jpg）。
なので自前では描画しない。

採否は上長確認済み（2026-09）：
  ・［提供終了］リサーチ機能終了 → 先月既出のため除外
  ・既読API（旧）新規受付終了 → 一旦は「新API中心に書き直し」で残したが、
    その後「入れるには渋い」との判断で除外（最終6アジェンダには含めない）
  ・早期発注特典①の締切 → 期限切れのため除外
  ・LINE Sales Promotion Manager → 優先度低だが一旦は残す。掲載順は元の号数どおり
    （オープンキャンペーンより前）

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
SRC = str(ROOT / "_templates" / "DYM_LINEOA_BUFFF_FMT.pptx")
OUT = str(ROOT / "LINEOA媒体最新情報_2026年9月トレンドレポート.pptx")

FONT = "メイリオ"

# 実FMTから抽出した配色（2026-09、LINEOA_BUFFF_3.pptxを解析）
TITLE_NAVY = RGBColor(0x00, 0x20, 0x60)   # ヘッダー見出し・本文タイトル色
CHIP_NAVY = RGBColor(0x10, 0x25, 0x3F)    # セクション見出しチップ・フラグチップ
BODY_GRAY = RGBColor(0x34, 0x34, 0x34)    # 本文
META_GRAY = RGBColor(0x49, 0x48, 0x48)    # メタ情報（対象業界・フェーズ等）
RED = RGBColor(0xFF, 0x00, 0x00)          # 強調・注意（FMT実測）
LIGHT_BOX = RGBColor(0xF2, 0xF2, 0xF2)    # メインメッセージ等のボックス塗り
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
SOURCE_GRAY = RGBColor(0x80, 0x80, 0x80)

SLIDE_W = Inches(10.83)
SLIDE_H = Inches(7.5)

N_SLIDES = 8  # 表紙1 + 目次1 + トピック6

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


def cover_slide(slide):
    add_text(slide, Inches(0.7), Inches(2.5), Inches(9), Inches(0.4),
              [[("LINE公式アカウント", 14, False, META_GRAY)]])
    add_text(slide, Inches(0.7), Inches(2.85), Inches(9.3), Inches(1.0),
              [[("媒体最新情報　9月 トレンドレポート", 30, True, TITLE_NAVY)]],
              line_spacing=1.15)
    add_rect(slide, Inches(0.7), Inches(3.95), Inches(9.0), Pt(1.5), fill=TITLE_NAVY)
    chip(slide, Inches(0.7), Inches(4.15), Inches(2.4), Inches(0.4), "2026年9月号",
         fill=LIGHT_BOX, color=META_GRAY, size=11, bold=False)
    add_text(slide, Inches(0.7), Inches(6.5), Inches(6), Inches(0.4),
              [[("株式会社DYM（DYM × LINEOA）", 12, True, TITLE_NAVY)]])


def agenda_slide(slide, topics):
    page_title(slide, "目次")
    y = Inches(1.0)
    row_h = Inches(0.62)
    for i, t in enumerate(topics, start=1):
        chip(slide, CONTENT_LEFT, y, Inches(0.44), Inches(0.42), str(i), size=12)
        add_text(slide, CONTENT_LEFT + Inches(0.62), y, Inches(9.0), Inches(0.42),
                  [[(t, 13, False, BODY_GRAY)]], anchor=MSO_ANCHOR.MIDDLE)
        y += row_h


def topic_slide(slide, data):
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
        size, spacing = (8.5, 1.1) if len(items) >= 6 else (9, 1.15) if len(items) >= 5 else (10, 1.2)
        add_bullets(slide, x + Inches(0.18), y + chip_h + Inches(0.1),
                    box_w - Inches(0.36), box_h - chip_h - Inches(0.2), items,
                    size=size, line_spacing=spacing)

    source_line(slide, data.get("source"))


TOPICS_LIST = [
    "1. LINEミニアプリ お気に入り追加促進機能",
    "2. LINE公式アカウント「分析」データ参照期間 過去36カ月に変更",
    "3. プロモーションスタンプ 効果改善アップデート",
    "4. 友だち追加オプション 複数アカウント対応",
    "5. LINE Sales Promotion Manager 利用停止期間",
    "6. LINEオープンキャンペーン 26年10-12月 期間限定キャンペーン",
]

TOPICS = [
    dict(
        title="LINEミニアプリ お気に入り追加促進機能",
        message="特定条件を満たしたユーザーに、ミニアプリのお気に入り追加を促すポップアップが自動表示される新機能。",
        overview=["トリガー①：マイミニアプリの「履歴」からアクセス",
                   "トリガー②：7日以内に2日以上アクセス",
                   "条件を満たすとミニアプリを閉じたタイミングでポップアップが表示",
                   "「お気に入りに追加」を押すとアプリタブの「お気に入り」に追加され、同時にアプリタブへ自動遷移"],
        pricing=["－"],
        schedule=["リリース予定日：8月31日（アプリのバージョンアップデートに伴って変更の可能性あり）",
                   "表示トリガーは数週間のABテストを経て、より効果の高い条件を100%反映予定（現状は暫定条件で運用中）"],
        caution=["リピート利用ユーザーの再訪導線が強化される。ミニアプリ運用中クライアントへの案内材料になる",
                   "既にお気に入り登録済みのミニアプリにはポップアップは表示されない",
                   "追加後に削除した場合、同じミニアプリのポップアップは720時間（30日間）再表示されない",
                   "「Not now」で見送った場合は14日間再表示されない",
                   "同一ミニアプリへの表示上限は2回。2回とも閉じると永久に対象外になる",
                   "同一ユーザーへの表示は1日最大1回（直近表示から24時間以内は非表示）"],
        source="https://workers-hub.box.com/s/9uo4m0fiqyxmddz5br3xe9691w0q4ht7",
    ),
    dict(
        title="LINE公式アカウント「分析」データ参照期間 過去36カ月に変更",
        message="「分析」タブの参照期間が過去36カ月までに変更予定。※社外告知は延期中、日程未確定につき最新情報を要確認。",
        overview=["Web版管理画面／管理アプリの「分析」タブで確認できる実績値の参照期間を過去36カ月までに変更",
                   "対象は「分析」タブの実績値のみ。それ以外のデータや作成済みメッセージには影響なし"],
        pricing=["該当なし"],
        schedule=["リリース日：2026年11月頃（予定・日程は決まり次第、管理画面のお知らせで案内）",
                   "社外告知：当初9/2予定だったが内容変更の可能性ありいったん延期。改めて案内予定"],
        caution=["36カ月より前の実績データが必要なクライアントには変更前の確認・保存を案内",
                   "「分析」タブの一部データはダウンロード可能（詳細は分析マニュアル参照）",
                   "告知日程が流動的なため最新情報を要確認"],
        source="https://workers-hub.box.com/s/ms8iztyci2dn5z2qna88rklxwxp7o5ja",
    ),
    dict(
        title="プロモーションスタンプ 効果改善アップデート",
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
        overview=["従来は1つのミニアプリに連携できる公式アカウントは1つのみだったが、複数設定可能に",
                   "流入経路（店舗・ブランド等）ごとに異なるLINE公式アカウントへ友だち追加を誘導できる"],
        pricing=["－"],
        schedule=["リリース予定日：9月2日"],
        caution=["全店舗共通の一斉配信から、店舗ごとのセグメント配信に切り替えられる",
                   "設定はLINE Developersの「友だち追加オプション」から。デフォルト設定／許可リスト設定の組み合わせで3パターン運用可能",
                   "紐づけ条件：ミニアプリと同一プロバイダーの公式アカウントであること／操作者がMessaging APIのadmin権限を保有していること。最大1,000件まで紐づけ可"],
        source="https://workers-hub.box.com/s/ko2jjwgxk5fs18ejrs97fvfqv3bm3089",
    ),
    dict(
        title="LINE Sales Promotion Manager 利用停止期間",
        message="データベース管理システムのスケール最適化のため、LINEで応募の管理画面が一時利用停止となる。",
        overview=["データベース管理システムのスケール最適化実施のため対象期間中は管理画面が利用停止",
                   "対象メニュー：LINEオープンキャンペーン（抽選型）。アンケート型・季節商品の過去案件も対象"],
        pricing=["該当なし"],
        schedule=["社外への案内開始日・リリース日：2026年9月2日",
                   "対象期間：9月8日（火）〜9月10日（木） 10:00〜14:30の間"],
        caution=["影響範囲：サマリーレポートの作成・ダウンロードが利用不可",
                   "LBPMでの案件申請・入稿など、上記以外の機能は通常どおり利用可能",
                   "稼働中キャンペーンがあるクライアントには事前周知が必要"],
        source="https://workers-hub.box.com/s/mo3y54qld9vj4kung0c4040z41mfhoip",
    ),
    dict(
        title="LINEオープンキャンペーン 26年10-12月 期間限定キャンペーン",
        message="2026年10月20日〜12月25日に終了する案件限定で、特別価格・特別期間・特別通数の3特典が適用される。",
        overview=["対象：2026年10月20日開始〜12月25日までに終了する案件限定",
                   "景品はLINEポイント（2ポイント固定・応募者全員に総付）",
                   "対象アカウント条件：LINE公式アカウントかつMessaging API対応／認証プロバイダーであることが必須"],
        pricing=["基本費用：通常1,800万円→キャンペーン価格1,500万円",
                   "掲載期間：通常7日間→15日間",
                   "通数：通常最大900万通→最大1,050万通（150万通増量・ノンセグメント1配信のみ、セグメント／オーディエンス配信は不可）",
                   "基本費用に友だち追加・uid提供・アンケート・景品・メッセージ配信1回分を含む",
                   "オプション：応募後の強制視聴ビデオ300万円（トライアル価格・変更の可能性あり）"],
        schedule=["①開始30営業日前：企業・商材審査／企画審査",
                   "②開始25営業日前17時：メッセージ配信枠おさえ・キャンペーン申込み",
                   "③開始20営業日前17時：各種入稿物の提出",
                   "④開始日当日：キャンペーン掲載開始",
                   "掲載開始は平日（LINEヤフー社営業日）のみ、土日祝日不可"],
        caution=["期間限定・条件限定のため、対象クライアントへは早めの提案・申込誘導が必要",
                   "キャンセル規定：発注受領後、または発注期日（開始25営業日前）超過後のキャンセルは最低発注金額1,500万円を請求",
                   "他割引との併用不可。景品がLINEポイントから変更になる可能性あり（変更不可の申し出は受付不可）"],
        source="https://workers-hub.box.com/s/odilqpe44tp398zlfko148yggcgj68ny",
    ),
]


def main():
    prs, slides = load_trimmed_fmt(N_SLIDES)

    cover_slide(slides[0])
    agenda_slide(slides[1], TOPICS_LIST)

    for slide, data in zip(slides[2:], TOPICS):
        topic_slide(slide, data)

    prs.save(OUT)
    print(f"saved: {OUT}")


if __name__ == "__main__":
    main()
