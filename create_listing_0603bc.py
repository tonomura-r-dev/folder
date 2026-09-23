import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03bc.xlsx"
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
    # ── スマートロック/アクセス管理SaaS ──
    ("株式会社ビットキー",              "https://bitkey.co.jp/business/",                         "リスティング広告", "", "スマートアクセス・workhub・入退室管理",      "入退室 管理 システム 比較, スマートロック 法人 導入, workhub 評判 料金",  "スマートビルSaaS"),
    # ── 経費精算・出張管理SaaS ──
    ("コンカー株式会社",                "https://www.concur.co.jp/return-lp",                     "リスティング広告", "", "SAP Concur・経費精算・出張管理・法人向け",   "コンカー 経費精算 料金 評判, 出張 管理 ツール 比較, SAP Concur 導入",  "バックオフィスSaaS"),
    # ── 電子請求・BtoBプラットフォーム ──
    ("株式会社インフォマート",          "https://lp.infomart.co.jp/seikyu/free-plan/",             "リスティング広告", "", "BtoBプラットフォーム請求書・電子帳簿保存対応", "請求書 電子化 システム 比較, インフォマート 料金 評判, 電子帳簿 保存 法",  "フィンテックSaaS"),
    # ── 採用管理SaaS ──
    ("株式会社HERP",                   "https://lp.herp.cloud/",                                  "リスティング広告", "", "採用管理システム・デジタル人材・ATS",         "採用 管理 システム ATS 比較, HERP 料金 評判, 採用 DX ツール おすすめ", "HR SaaS"),
    # ── カスタマーサービスSaaS ──
    ("Zendesk株式会社",                "https://www.zendesk.co.jp/campaigns/brand/",              "リスティング広告", "", "カスタマーサービス・チケット管理・AI搭載",    "Zendesk 料金 評判 比較, チケット 管理 SaaS, カスタマーサービス ツール",  "カスタマーサポートSaaS"),
    # ── BI SaaS ──
    ("セールスフォース・ジャパン株式会社", "https://www.tableau.com/ja-jp/products/trial",         "リスティング広告", "", "Tableau・BIダッシュボード・無料トライアル",   "Tableau 無料 試用, BI ツール データ 可視化 比較, Tableau 料金 評判",   "BI SaaS"),
    # ── セールスイネーブルメント ──
    ("株式会社ナレッジワーク",          "https://knowledgework.cloud/",                            "リスティング広告", "", "セールスイネーブルメントAI・営業力強化",       "セールスイネーブルメント ツール 比較, 営業 強化 SaaS, ナレッジワーク 料金", "セールスSaaS"),
    # ── ホームページ作成 ──
    ("株式会社ペライチ",               "https://peraichi.com/landing_pages/view/lp-seisaku/",     "リスティング広告", "", "ノーコードHP・LP作成・無料から始める",        "ホームページ 作成 無料 簡単, LP 制作 ツール 初心者, ペライチ 評判 料金",  "ノーコードSaaS"),
    # ── GHG・脱炭素SaaS ──
    ("株式会社ゼロボード",             "https://www.zeroboard.jp/zeroboard_lp",                   "リスティング広告", "", "GHG排出量算定・脱炭素・ESGデータ管理",        "CO2 排出量 算定 ツール, 脱炭素 DX SaaS, ゼロボード 料金 評判",         "サステナビリティSaaS"),
    # ── 勤怠・工数管理SaaS ──
    ("株式会社チームスピリット",        "https://www.teamspirit.co.jp/lp/expense_system/",         "リスティング広告", "", "勤怠管理・工数管理・経費精算・Salesforce上",  "チームスピリット 料金 評判, 勤怠 管理 Salesforce, TeamSpirit 比較",    "HR SaaS"),
    # ── CX・BPO支援 ──
    ("バーチャレクス・コンサルティング株式会社", "https://www.virtualex.co.jp/service/cx_outsourcing/", "リスティング広告", "", "コンタクトセンターBPO・CX改善・AI活用",    "コンタクトセンター BPO 代行, CX 改善 コンサル, バーチャレクス 評判",    "BPO"),
    # ── MAツール ──
    ("SATORI株式会社",                 "https://lp.satori.marketing/lp01",                        "リスティング広告", "", "マーケティングオートメーション・匿名MA",       "MA ツール 比較 国産, マーケ オートメーション, SATORI 料金 評判",         "MAツールSaaS"),
    # ── AI研修 ──
    ("株式会社スキルアップAI",         "https://www.skillupai.com/private-training/",             "リスティング広告", "", "AI研修・DX人材育成・企業向け",               "AI 研修 法人 おすすめ, DX 人材 育成 研修, スキルアップAI 評判 料金",   "教育・研修"),
    # ── デジタルアダプション ──
    ("テックタッチ株式会社",           "https://techtouch.jp/lp/ex-digital-adoption/",            "リスティング広告", "", "デジタルアダプションプラットフォーム・操作ガイド", "デジタルアダプション ツール 比較, システム 定着率 向上, テックタッチ 料金", "DXSaaS"),
    # ── 人材派遣（企業向け） ──
    ("株式会社リクルートスタッフィング", "https://www.r-staffing.co.jp/sol/contents/client/lp/officework_01/", "リスティング広告", "", "派遣・事務スタッフ・即戦力人材",            "人材 派遣 事務 おすすめ, 即戦力 派遣 採用, リクルートスタッフィング 評判", "人材派遣"),
    # ── 電子請求書受領 ──
    ("株式会社インフォマート",          "https://www.infomart.co.jp/contract/lp_cross.asp",        "リスティング広告", "", "BtoB契約書電子化・電子契約・ID連携",          "契約書 電子化 ツール, 電子契約 法人 比較, インフォマート 契約書",         "電子契約SaaS"),
    # ── FAXDMサービス ──
    ("株式会社ネクスウェイ",            "https://www.nexway.co.jp/lp/faxdm2/",                    "リスティング広告", "", "FAXDM・法人リスト0.5円・新規開拓",            "FAX DM 法人 リスト 安い, 一括 FAX 送信 代行, ネクスウェイ 評判",     "DM・マーケ"),
    # ── 建設SaaS ──
    ("株式会社アンドパッド",            "https://andpad.jp/",                                      "リスティング広告", "", "施工管理アプリ・建設DX・シェアNo.1",           "アンドパッド 料金 評判, 施工管理 アプリ 比較, 建設 工務店 クラウド",   "建設SaaS"),
    # ── GHG・ESG可視化 ──
    ("株式会社ROBOT PAYMENT",          "https://www.robotpayment.co.jp/service/mikata/",           "リスティング広告", "", "請求管理ロボ・自動請求・クラウド決済",          "請求管理 ロボ 料金 評判, 請求書 自動化 クラウド, 債権 管理 SaaS",      "フィンテックSaaS"),
    # ── 採用管理 ──
    ("株式会社Thinkings",              "https://sonar-ats.jp/lp/cm-2022-23/",                     "リスティング広告", "", "採用管理システム・ATS・1,100社以上導入",        "採用 管理 ツール 比較, ATS SaaS おすすめ, ソナー ATS 料金 評判",      "HR SaaS"),
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
