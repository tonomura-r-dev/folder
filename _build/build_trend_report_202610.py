"""
LINE公式アカウント 媒体最新情報 2026年10月 トレンドレポート ビルドスクリプト。

トンマナ・レイアウトは trend_report_engine.py に集約。このファイルは10月号のデータだけを持つ。
作業手順は .claude/skills/trend-report/SKILL.md を参照。

ネタは殿村さんから2026-10-21に受領した9本（掲載順も受領順のまま）。
資料URL（workers-hub.box.com）はこの環境から開けないため、中身は受領した概要文のみで作成。
出典は9月号の最終版（殿村さん手直し版）に合わせ、公開URLが概要文にあるものだけURL、他は「代理店情報」。

使い方: python3 _build/build_trend_report_202610.py
"""

from pathlib import Path
from pptx.util import Pt
import trend_report_engine as engine

ROOT = Path(__file__).resolve().parent.parent
OUT = str(ROOT / "LINEOA媒体最新情報_2026年10月トレンドレポート.pptx")
MONTH_LABEL = "2026年10月"

TOPICS = [
    dict(
        bracket="LINE通知メッセージ",
        headline="ネイティブアプリ遷移の条件付き解禁",
        industry="アプリを運営している全業種",
        category="開発\nLINE API",
        description_lines=[
            "ベースプラン限定で、一定の条件を満たした場合に限り、",
            "通知メッセージのリンクからネイティブアプリへ遷移できるようガイドラインが緩和された。",
            "2026年9月9日（水）からLBPMでネイティブアプリ遷移の審査受付を開始。",
        ],
        render="before_after",
        as_is="通知メッセージのリンクからネイティブアプリへは遷移できない",
        as_is_ratio=0.4,
        as_is_headline_size=Pt(13),
        to_be_headline_size=Pt(18),
        to_be_lines=["ベースプラン限定・条件付きで遷移可能に",
                     "テンプレート／フレキシブルの両方が対象。審査は10万円／1UX"],
        merits=[
            ("通知からアプリへ直接誘導できる",
             "通知メッセージのリンクからアプリを直接開けるため、Webを経由しない導線を設計できる。"),
            ("遷移には事前審査が必要",
             "LBPMでの審査（10万円／1UX）が前提。社外告知は2026年9月8日（火）、審査受付は9月9日（水）から。"),
        ],
        source="代理店情報",
    ),
    dict(
        bracket="LINE通知メッセージ",
        headline="UX追加申請の請求書発行をLBPMへ移行",
        industry="通知メッセージをご利用中の全業種",
        category="開発\nLINE API",
        description_lines=[
            "「UX追加申請（LINE通知メッセージ）」の請求書の発行元が、",
            "これまでのAGPからLBPMに変更される。発注画面・申請方法も合わせて変更。",
            "2026年9月受領分（2026年10月月初の請求書発行分）から適用。",
        ],
        render="before_after",
        as_is="請求書はAGPから発行",
        as_is_ratio=0.4,
        as_is_headline_size=Pt(14),
        to_be_headline_size=Pt(20),
        to_be_lines=["請求書はLBPMから発行", "2026年9月受領分（10月月初発行分）から"],
        merits=[
            ("発注画面・申請方法も変更",
             "LBPM上の「UX追加申請（LINE通知メッセージ）」の発注画面と申請方法が新しくなる。"),
            ("請求書の発行先が変わる",
             "2026年10月月初発行分からは、LBPMから請求書を発行する。"),
        ],
        source="代理店情報",
    ),
    dict(
        bracket="LINE公式アカウント",
        headline="「分析」データ参照期間 リリース日の変更",
        industry="全業種",
        category="LINE公式アカウント",
        description_lines=[
            "9月号でご案内した「分析」タブの参照期間を過去36カ月に変更する件について、",
            "リリース日が2026年11月頃から2027年2月頃に変更された。",
            "36カ月より前の実績値が必要な場合は、仕様変更前に確認しておく。",
        ],
        render="before_after",
        as_is="リリース日：2026年11月頃",
        as_is_ratio=0.4,
        as_is_headline_size=Pt(13),
        to_be_headline_size=Pt(18),
        to_be_lines=["リリース日：2027年2月頃に変更", "変更内容（参照期間を過去36カ月まで）は同じ"],
        merits=[
            ("対象はWeb版・管理アプリの両方",
             "Web版管理画面と管理アプリの「分析」タブで確認できる実績値が、過去36カ月までになる。"),
            ("過去データの確認は変更前に",
             "36カ月より前の実績値が必要なクライアントには、リリース前の確認・保存を案内する。"),
        ],
        source="代理店情報",
    ),
    dict(
        bracket="LINEシリアルコードスタンプ",
        headline="媒体資料の更新・進行ガイドのご案内",
        industry="友だち獲得を強化したい全業種",
        category="オプション\n商材",
        description_lines=[
            "2026年10月1日（木）リリースの新商品「LINEシリアルコードスタンプ」について、",
            "媒体資料（更新版）と進行ガイドの2点が公開された。",
            "既存のプロモーションスタンプと体裁を揃え、スペックの一部も更新されている。",
        ],
        render="grid_notes",
        items=[
            ("媒体資料：体裁を統一", "既存のプロモーションスタンプの媒体資料と体裁を統一。"),
            ("媒体資料：スペックの一部更新", "期間の延長・短縮化など、スペックの一部を更新。"),
            ("進行ガイド：独自フロー", "シリアルコードスタンプ独自の進行フローを詳しく記載。"),
            ("進行ガイド：設計・制作ルール", "キャンペーン設計ルールと制作ガイドを詳しく記載。"),
        ],
        source="代理店情報",
    ),
    dict(
        bracket="LINE公式アカウント",
        headline="「利用と請求」のメンテナンス",
        industry="各種オプションをご利用中の全業種",
        category="LINE公式アカウント",
        description_lines=[
            "「利用と請求」の一部申し込み・契約画面のメンテナンスが実施される。",
            "日程は2026年9月14日（月）14:00〜14:30頃の予定。",
            "期間中は、CRM・レストラン・Beautyオプションの新規申込・契約変更ができない。",
        ],
        render="grid_notes",
        items=[
            ("日程", "2026年9月14日（月）14:00〜14:30頃。日程・時間は前後する場合あり。"),
            ("利用できない操作", "新規申込・契約変更（解約予約・解約予約キャンセル）。"),
            ("対象オプション", "CRMオプション／レストランオプション／Beautyオプション。"),
            ("対応", "メンテナンス終了後に、あらためて手続きを行う。"),
        ],
        source="代理店情報",
    ),
    dict(
        bracket="LINE通知メッセージ",
        headline="ユニット単位の統計情報取得に対応",
        industry="通知メッセージをご利用中の全業種",
        category="開発\nLINE API",
        description_lines=[
            "2026年9月15日（火）から、ユニットごとの統計情報を取得するAPIが",
            "LINE通知メッセージでも利用できるようになった。",
            "ユニットごとに、開封数・URLタップ数などの統計情報を確認できる。",
        ],
        render="before_after",
        as_is="LINE通知メッセージではユニット単位の統計を取得できない",
        as_is_ratio=0.4,
        as_is_headline_size=Pt(13),
        to_be_headline_size=Pt(18),
        to_be_lines=["ユニット単位の統計情報をAPIで取得可能に",
                     "開封数・URLタップ数などをユニットごとに確認できる"],
        merits=[
            ("要望の多かった機能",
             "2025年10月・2026年6月の「LINE通知メッセージ」アンケートで、特に要望が多かった機能。"),
            ("通知ごとの反応を比較できる",
             "ユニットごとの開封数・URLタップ数を見て、通知の内容を見直せる。"),
        ],
        source="https://developers.line.biz/ja/docs/messaging-api/unit-based-statistics-aggregation/",
    ),
    dict(
        bracket="LINEポイント Friends AD",
        headline="FY26 冬季キャンペーン",
        industry="友だち獲得を強化したい全業種",
        category="オプション\n商材",
        description_lines=[
            "Friends Swipe AD／Friends Video ADのALL配信を対象に、",
            "通常配信6円・リーチ最適化配信9円、付与ポイント2Pの特別単価で提供される。",
            "2026年12月1日〜2027年3月31日までに掲載終了する案件が対象。",
        ],
        render="impact",
        stats=[
            ("6円", "通常配信", "キャンペーン特別単価"),
            ("9円", "リーチ最適化配信", "キャンペーン特別単価"),
            ("2P", "付与ポイント", "ALL配信限定"),
            ("200万円〜", "最低発注金額", "最新の媒体資料に準拠"),
        ],
        breakdown_headers=["対象期間", "対象メニュー", "対象セグメント", "対象クライアント"],
        breakdown_values=["26/12/1〜27/3/31", "Swipe AD／Video AD", "ALL配信のみ", "全クライアント"],
        merit_headline="発注時の注意",
        merit_desc="備考欄に「FY26 Winter Campaign適用」と記載。配信中・発注受領済みの案件、他割引との併用、デリタゲ配信は対象外。",
        source="代理店情報",
    ),
    dict(
        bracket="レストランオプション",
        headline="一部機材の価格改定",
        industry="飲食業",
        category="オプション\n商材",
        description_lines=[
            "仕入れ先であるエプソン販売株式会社の価格改定に伴い、",
            "レストランオプションで提供している一部機材の販売価格が改定（値上げ）される。",
            "2027年2月上旬以降のお申し込み分から新価格を適用。",
        ],
        render="grid_notes",
        items=[
            ("改定時期", "2027年2月上旬予定。同日以降のお申し込みページアクセス分から新価格。"),
            ("対象機材", "EPSON製のプリンター・キャッシュドロアー。"),
            ("価格", "新価格は代理店資料に掲載。"),
            ("告知", "OAMの申し込み導線・AGPお知らせ・飲食CMSお知らせにも掲載予定。"),
        ],
        source="代理店情報",
    ),
    dict(
        bracket="CRMオプション",
        headline="導入事例の追加",
        industry="全業種",
        category="オプション\n商材",
        description_lines=[
            "CRMオプションの利用企業で、新たな導入事例が共有された。",
            "フォームを使ったアンケートで回答率が上がり、",
            "回答内容に沿った配信でメッセージのCTRが改善した。",
        ],
        render="grid_notes",
        items=[
            ("① アンケート回答率UP", "リサーチ機能ではなくフォームを使うことで、アンケートの回答率が上がった。"),
            ("② 配信メッセージのCTR改善", "ユーザーの回答内容に沿った配信を行い、メッセージのCTRが改善した。"),
        ],
        source="代理店情報",
    ),
]

for t in TOPICS:
    t["title"] = f"【{t['bracket']}】{t['headline']}"


if __name__ == "__main__":
    engine.build(OUT, MONTH_LABEL, TOPICS)
