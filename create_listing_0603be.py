import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03be.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 56, 16, 16, 22, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

# 全社 WebSearch実確認済みLP（企業名カッコなし正式名称のみ）
DATA = [
    # ── プログラミング動画学習 ──
    ("株式会社Techpit",               "https://enterprise.techpit.jp/company",                   "リスティング広告", "", "エンジニア向けリスキリング・法人研修",       "エンジニア 研修 SaaS, プログラミング 学習 法人, TechPit 料金 評判",     "EdTech"),
    # ── フォーム作成 ──
    ("株式会社ベーシック",            "https://form.run/ja",                                      "リスティング広告", "", "ノーコードフォーム作成・問い合わせ管理",       "フォーム 作成 ツール 無料, 問い合わせ フォーム 比較, formrun 評判",     "ノーコードSaaS"),
    # ── 経営管理クラウド ──
    ("株式会社ログラス",              "https://www.loglass.jp/lp/y01-test0201",                   "リスティング広告", "", "予算管理・予実管理・経営管理クラウド",        "予算 管理 SaaS 比較, 予実 管理 ツール, ログラス 料金 評判",            "経営管理SaaS"),
    # ── チャットボット ──
    ("チャットプラス株式会社",        "https://chatplus.jp/lp/",                                  "リスティング広告", "", "AI搭載チャットボット・問い合わせ自動化",       "チャットボット 導入 実績 No1, AI 自動応答 ツール, チャットプラス 料金", "カスタマーサポートSaaS"),
    # ── 製造業調達 ──
    ("株式会社アペルザ",              "https://dx.aperza.com/lp/desk/",                           "リスティング広告", "", "製造業向け受発注DX・1日250円で効率化",       "製造業 受発注 DX, 工場 調達 デジタル化, アペルザDESK 料金",           "製造業SaaS"),
    # ── 社内wiki ──
    ("株式会社プロダクトライフサイクル", "https://notepm.jp/register",                            "リスティング広告", "", "社内wiki・マニュアル作成・ナレッジ管理",      "社内 wiki ツール 比較, マニュアル 作成 クラウド, NotePM 料金 評判",   "ナレッジ管理SaaS"),
    # ── 広告運用代行 ──
    ("株式会社カルテットコミュニケーションズ", "https://quartet-communications.com/",             "リスティング広告", "", "リスティング広告運用代行・4,200社実績",        "リスティング 広告 代行 中小企業, Google 広告 運用 代行 費用, カルテット 評判", "広告代理店"),
    # ── SFA/CRM AI ──
    ("株式会社マツリカ",              "https://lp-senses.mazrica.com/glistings-mazrica/",         "リスティング広告", "", "Mazrica Sales・AI搭載SFA/CRM・案件管理",     "SFA AI 搭載 比較, マジカ 料金 評判, 営業 案件 管理 ツール",           "SFA SaaS"),
    # ── SNS運用代行 ──
    ("株式会社コムニコ",              "https://www.comnico.jp/service-doc-form",                  "リスティング広告", "", "SNS運用代行・Instagram・X・TikTok",          "SNS 運用代行 費用 比較, Instagram 運用 代行, コムニコ 評判 料金",     "SNSマーケ"),
    # ── マルチチャネルMA ──
    ("Repro株式会社",                 "https://repro.io/lp/download/",                            "リスティング広告", "", "マルチチャネルMA・プッシュ通知・LPO",         "MA ツール プッシュ通知 比較, アプリ MA, Repro 料金 評判",              "MAツールSaaS"),
    # ── 広告運用代行 ──
    ("リードプラス株式会社",          "https://lp.leadplus.co.jp/ads/service",                    "リスティング広告", "", "インターネット広告運用代行・AI自動最適化",     "広告 運用 代行 中小 企業, リスティング 広告 代行 費用, リードプラス 評判", "広告代理店"),
    # ── オンボードAI ──
    ("株式会社ビズリーチ",            "https://onboard-ai.jp/lp/",                                "リスティング広告", "", "オンボードAI・入社定着・マネジメント改革",   "新入社員 定着 ツール, オンボード AI SaaS, 入社 研修 効率化",          "HR SaaS"),
    # ── テスト自動化 ──
    ("株式会社Autify",               "https://autify.com/ja/demo",                               "リスティング広告", "", "E2Eテスト自動化・AIノーコード・QA効率化",    "テスト 自動化 SaaS 比較, E2E ツール ノーコード, Autify 料金",          "DevOps SaaS"),
    # ── 物流DX ──
    ("ハコベル株式会社",              "https://lp.hacobell.com/",                                 "リスティング広告", "", "物流DX・配送マッチング・即日手配",            "物流 DX プラットフォーム, ハコベル 評判 料金, 配送 マッチング サービス", "物流SaaS"),
    # ── 医師向けメディア ──
    ("メドピア株式会社",              "https://medpeer.co.jp/service/company.html",               "リスティング広告", "", "医師18万人プラットフォーム広告・製薬支援",    "医師 広告 媒体, 製薬 マーケティング 費用, メドピア 掲載 申し込み",    "医療メディア"),
    # ── オンライン商談 ──
    ("ベルフェイス株式会社",          "https://bell-face.com/lp/",                                "リスティング広告", "", "電話×オンライン商談・金融業界シェアNo.1",    "オンライン 商談 電話 ツール, ベルフェイス 料金 評判, Web 商談 比較",  "セールスSaaS"),
    # ── プロシェアリング ──
    ("株式会社サーキュレーション",    "https://circu.co.jp/lp_hojin_01/",                         "リスティング広告", "", "プロシェアリング・副業専門家・経営課題解決",   "プロシェアリング 費用 比較, 外部 専門家 活用, サーキュレーション 評判", "人材マッチング"),
    # ── 採用求人 ──
    ("株式会社リクルート",            "https://townwork.net/assets/twc/jobpost/gatepage/index.html", "リスティング広告", "", "タウンワーク求人掲載・アルバイト採用広告",  "バイト 求人 掲載 料金, アルバイト 採用 広告, タウンワーク 申し込み",  "採用メディア"),
    # ── 大学院生就活 ──
    ("株式会社アカリク",              "https://acaric.jp/special/register-now-general",           "リスティング広告", "", "大学院生・理系学生特化就活・スカウト転職",     "大学院生 就活 サイト, 理系 院卒 転職 エージェント, アカリク 評判",    "就活サービス"),
    # ── インテントセールス ──
    ("株式会社Sales Marker",          "https://sales-marker.jp/lp/demo/",                         "リスティング広告", "", "インテントデータ・BtoB営業自動化",            "インテント セールス SaaS, 法人 営業 ターゲット データ, セールスマーカー 料金", "セールスSaaS"),
]


def make_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "架電リスト"

    for col_idx, header in enumerate(HEADER, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
    ws.row_dimensions[1].height = 20

    for row_idx, row_data in enumerate(DATA, 2):
        fill = ROW_FILLS[(row_idx - 2) % 2]
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill = fill
            cell.alignment = ROW_ALIGN
            if col_idx == 2:
                cell.hyperlink = value
                cell.font = URL_FONT
            else:
                cell.font = ROW_FONT
        ws.row_dimensions[row_idx].height = 16

    for col_idx, width in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}1"

    wb.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")
    print(f"Rows: {len(DATA)}")


if __name__ == "__main__":
    make_xlsx()
