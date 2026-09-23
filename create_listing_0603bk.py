import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03bk.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 56, 16, 16, 22, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

DATA = [
    # ── 電子署名SaaS ──
    ("Docusign Japan株式会社",      "https://ecom.docusign.com/ja-JP/plans-and-pricing/esignature", "リスティング広告", "", "Docusign・電子署名・世界No.1・日本語対応", "電子署名 SaaS 比較, Docusign 料金 評判, 契約書 電子化",             "電子署名SaaS"),
    # ── クラウドストレージ ──
    ("Box Japan株式会社",           "https://japan.box.com/box-price",                             "リスティング広告", "", "Box・クラウドストレージ・セキュリティ・法人", "クラウド ストレージ 法人 比較, Box 料金 評判, ファイル 管理",       "クラウドストレージSaaS"),
    # ── クラウドストレージ ──
    ("Dropbox Japan株式会社",       "https://www.dropbox.com/ja/lp/business/smb-collaboration",    "リスティング広告", "", "Dropbox Business・コラボ・電子署名・30日",  "Dropbox 料金 評判, クラウド ファイル 共有 SaaS, Dropbox Business", "クラウドストレージSaaS"),
    # ── グループウェア ──
    ("グーグル合同会社",             "https://workspace.google.com/pricing?hl=ja",                  "リスティング広告", "", "Google Workspace・Gmail・Meet・Drive法人", "Google Workspace 料金 評判, ビジネス メール SaaS, グループウェア", "グループウェアSaaS"),
    # ── オフィスSaaS ──
    ("日本マイクロソフト株式会社",   "https://www.microsoft.com/ja-jp/microsoft-365/business/compare-all-microsoft-365-business-products-b", "リスティング広告", "", "Microsoft 365・Teams・Office・法人向け", "Microsoft 365 料金 評判, Office SaaS, Teams ビジネス プラン", "グループウェアSaaS"),
    # ── CRM/カスタマーサポート ──
    ("Freshworks Japan株式会社",    "https://www.freshworks.com/jp/pricing/",                       "リスティング広告", "", "Freshworks・CRM・カスタマーサポート・AI",   "Freshworks 料金 評判, CRM SaaS 比較, カスタマーサポート ツール",   "CRM SaaS"),
    # ── HCM/人事SaaS ──
    ("ワークデイ株式会社",           "https://www.workday.com/ja-jp/pages/lp/index.html",            "リスティング広告", "", "Workday・HCM・人事・財務・エンタープライズ", "Workday 料金 評判, HCM SaaS, 人事 管理 グローバル ERP",          "HCM SaaS"),
    # ── ITSM SaaS ──
    ("ServiceNow Japan合同会社",    "https://www.servicenow.com/lpgp/pricing-itsm.html",            "リスティング広告", "", "ServiceNow・ITSM・IT業務自動化・世界No.1", "ServiceNow 料金 評判, ITSM SaaS 比較, IT 業務 管理 ツール",      "ITSMSaaS"),
    # ── ERP SaaS ──
    ("SAPジャパン株式会社",         "https://www.sap.com/japan/products/erp/s4hana-private-edition/trial.html", "リスティング広告", "", "SAP S/4HANA・クラウドERP・無料評価版", "SAP ERP SaaS 比較, S/4HANA 料金 評判, クラウド ERP 導入",        "ERP SaaS"),
    # ── クラウドERP ──
    ("日本オラクル株式会社",         "https://cloud.oracle.com/ja_JP/erp-cloud",                    "リスティング広告", "", "Oracle Cloud ERP・財務・調達・クラウド",    "Oracle ERP 料金 評判, クラウド 財務 管理 SaaS, Oracle Cloud",    "ERP SaaS"),
    # ── カンバン管理SaaS ──
    ("Atlassian株式会社",           "https://trello.com/ja/pricing",                               "リスティング広告", "", "Trello・カンバン・タスク管理・無料プラン",   "Trello 料金 評判, カンバン SaaS 比較, タスク 管理 ツール",        "プロジェクト管理SaaS"),
    # ── ノーコードDB SaaS ──
    ("Airtable Inc.",               "https://airtable.com/pricing",                                "リスティング広告", "", "Airtable・ノーコードDB・プロジェクト管理",   "Airtable 料金 評判, ノーコード DB SaaS, チーム 管理 ツール",      "ノーコードSaaS"),
    # ── ワークマネジメントSaaS ──
    ("Smartsheet Inc.",             "https://www.smartsheet.com/pricing",                          "リスティング広告", "", "Smartsheet・プロジェクト管理・無料体験",     "Smartsheet 料金 評判, プロジェクト管理 SaaS, ワーク 管理",       "プロジェクト管理SaaS"),
    # ── プロジェクト管理SaaS ──
    ("Citrix Systems Japan合同会社", "https://www.wrike.com/ja/price-vf/",                          "リスティング広告", "", "Wrike・プロジェクト管理・AI駆動・日本語",    "Wrike 料金 評判, プロジェクト管理 SaaS 比較, チーム 作業 管理",  "プロジェクト管理SaaS"),
    # ── タスク管理SaaS ──
    ("ClickUp Inc.",                "https://clickup.com/pricing",                                 "リスティング広告", "", "ClickUp・タスク管理・無料Forever・AI連携",  "ClickUp 料金 評判, タスク 管理 SaaS 無料, プロジェクト ツール",   "プロジェクト管理SaaS"),
    # ── エンジニア向けプロジェクト管理 ──
    ("Linear Inc.",                 "https://linear.app/pricing",                                  "リスティング広告", "", "Linear・プロダクト開発・AI自動化・エンジニア", "Linear 料金 評判, 開発 プロジェクト 管理 SaaS, エンジニア ツール", "DevOps SaaS"),
    # ── DevOps SaaS ──
    ("Datadog Japan株式会社",       "https://www.datadoghq.com/free-datadog-trial/",               "リスティング広告", "", "Datadog・APM・ログ分析・14日間無料",          "Datadog 料金 評判, クラウド 監視 SaaS, APM ログ 分析 ツール",    "DevOps SaaS"),
    # ── オブザーバビリティ ──
    ("New Relic株式会社",           "https://newrelic.com/jp/lp/start-new-relic-for-free-goog-da-jp", "リスティング広告", "", "New Relic・オブザーバビリティ・無料100GB", "New Relic 料金 評判, オブザーバビリティ SaaS, システム 監視",    "DevOps SaaS"),
    # ── データウェアハウス ──
    ("Snowflake Inc.",              "https://signup.snowflake.com/?_l=ja",                         "リスティング広告", "", "Snowflake・AIデータクラウド・無料$400",        "Snowflake 料金 評判, データウェアハウス SaaS, クラウド 分析",    "データSaaS"),
    # ── IAM SaaS ──
    ("Okta Japan株式会社",          "https://www.okta.com/ja-jp/pricing/",                         "リスティング広告", "", "Okta・IAM・SSO・アクセス管理・日本語対応",    "Okta 料金 評判, IAM SaaS 比較, シングルサインオン 企業",          "セキュリティSaaS"),
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
