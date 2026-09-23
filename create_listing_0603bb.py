import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03bb.xlsx"
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
    # ── MAツール（BtoB） ──
    ("株式会社ベーシック",               "https://cloudcircus.jp/",                               "リスティング広告", "", "デジタルマーケティングSaaS・MA・CMS",       "クラウドサーカス 料金 評判, BtoB マーケ ツール 比較, MA CMS 統合",      "マーケSaaS"),
    # ── 米国株投資 ──
    ("ブルーモ証券株式会社",             "https://invest.bloomo.co.jp/lp/usstock-01",             "リスティング広告", "", "米国株投資・スマホ・ポートフォリオ",         "米国株 アプリ 投資, ブルーモ 評判 料金, ポートフォリオ 投資 スマホ",    "金融・証券"),
    # ── 不動産投資ローン ──
    ("セゾンファンデックス株式会社",     "https://www.fundex.co.jp/lp/fudousan/course2.html",     "リスティング広告", "", "不動産投資ローン・低金利・個人向け",          "不動産 投資 ローン 低金利, 不動産 担保 融資 比較, セゾン ローン 申し込み", "金融・ローン"),
    # ── 施工管理SaaS ──
    ("株式会社アンドパッド",             "https://andpad.jp/",                                    "リスティング広告", "", "施工管理アプリ・シェアNo.1・建設DX",         "施工管理 アプリ 比較, アンドパッド 料金 評判, 建設 工務店 DX ツール",  "建設SaaS"),
    # ── 採用管理SaaS ──
    ("株式会社Thinkings",               "https://sonar-ats.jp/lp/cm-2022-23/",                   "リスティング広告", "", "採用管理システム・ATS・1,100社導入",          "採用管理 システム 比較, ATS ツール おすすめ, 採用 効率化 クラウド",     "HR SaaS"),
    # ── FAXDM ──
    ("株式会社ネクスウェイ",             "https://faxdm.nexway.co.jp/lp",                         "リスティング広告", "", "FAXDM・業界シェアNo.1・新規開拓営業",        "FAX DM 送信 代行 安い, 一括 FAX 法人 リスト, ネクスウェイ 評判 料金", "DM・マーケ"),
    # ── 建設職人マッチング ──
    ("クラフトバンク株式会社",           "https://craft-bank.com/lp/plan_index",                  "リスティング広告", "", "建設工事受発注・協力会社マッチング",          "協力会社 マッチング 建設, 職人 募集 クラウド, クラフトバンク 評判",    "建設SaaS"),
    # ── FAQシステム ──
    ("株式会社Helpfeel",                "https://www.helpfeel.com/lp",                            "リスティング広告", "", "AI FAQ・問い合わせ削減・意図予測検索",         "FAQ システム 比較 AI, 問い合わせ 削減 ツール, Helpfeel 料金 評判",    "カスタマーサポートSaaS"),
    # ── BtoBリード獲得メディア ──
    ("株式会社マイナビ",                 "https://ad-lp.news.mynavi.jp/lead/generation/",          "リスティング広告", "", "TECH+ BtoBリード獲得・ホワイトペーパー",      "BtoB リード 獲得 メディア, IT 企業 広告 掲載, TECH+ 料金 評判",       "広告メディア"),
    # ── 音声解析AI ──
    ("株式会社RevComm",                 "https://miitel.com/jp/lp/",                              "リスティング広告", "", "音声解析AI・商談分析・トーク改善",            "音声解析 AI 電話 比較, MiiTel 料金 評判, 商談 分析 ツール インサイドセールス", "セールスSaaS"),
    # ── メール・LP構築SaaS ──
    ("スパイラル株式会社",               "https://www.pi-pe.co.jp/lp/mail2016b/",                 "リスティング広告", "", "SPIRAL・メール配信・LP構築・データベース",    "スパイラル メール 配信 比較, LP 構築 ツール 簡単, 顧客 DB クラウド", "マーケSaaS"),
    # ── 外国人材紹介 ──
    ("グローバルパワー株式会社",         "https://globalpower.co.jp/lp-listing/",                 "リスティング広告", "", "高度外国人材紹介・日本語N1/N2・中途特化",     "外国人 採用 紹介 会社, 高度 外国人 転職 エージェント, グローバル採用 費用", "人材紹介"),
    # ── 請求管理SaaS ──
    ("株式会社ROBOT PAYMENT",           "https://www.robotpayment.co.jp/service/mikata/",         "リスティング広告", "", "請求管理ロボ・自動請求・債権管理",            "請求書 自動化 システム, 請求管理 ロボ 料金 評判, 債権 管理 SaaS 比較", "フィンテックSaaS"),
    # ── IT資産管理SaaS ──
    ("Josys株式会社",                   "https://www.josys.com/lp",                               "リスティング広告", "", "ITデバイス・SaaS統合管理クラウド",            "SaaS 管理 ツール 比較, IT 資産 管理 クラウド, Josys 料金 評判",       "IT資産管理SaaS"),
    # ── BtoBメディア掲載 ──
    ("スマートキャンプ株式会社",         "https://majisemi.com/service/lp/",                      "リスティング広告", "", "ウェビナー集客・独自ハウスリスト・BtoB",      "ウェビナー 集客 代行, BtoB セミナー 申し込み, マジセミ 料金 評判",    "マーケ支援"),
    # ── WAFセキュリティ ──
    ("株式会社サイバーセキュリティクラウド", "https://lp.cscloud.co.jp/ksk/mail/mailmagazine/lp", "リスティング広告", "", "WAF・サイバー攻撃対策・クラウド型",          "WAF クラウド 比較 国産, サイバー攻撃 対策 申し込み, 攻撃遮断くん 料金", "セキュリティSaaS"),
    # ── 営業リストSaaS ──
    ("株式会社SalesNow",                "https://salesnow.jp/lp_c/",                              "リスティング広告", "", "法人営業リスト・1,400万件DB・ターゲット抽出", "営業 リスト 作成 ツール, 法人 データベース SaaS, SalesNow 料金 評判", "セールスSaaS"),
    # ── BIツール ──
    ("ウイングアーク1st株式会社",        "https://lp.wingarc.com/MB/001",                         "リスティング広告", "", "BIダッシュボード・データ可視化・国産シェアNo.1", "BI ツール 比較 国産, MotionBoard 料金 評判, データ 可視化 SaaS",      "BI SaaS"),
    # ── SaaS比較メディア ──
    ("スマートキャンプ株式会社",         "https://boxil.jp/campaign/partners/",                   "リスティング広告", "", "BOXIL SaaS比較掲載・リード獲得",             "SaaS 比較 サイト 掲載, BtoB リード 獲得 媒体, BOXIL 評判 費用",      "デジタルマーケ"),
    # ── CRM SaaS ──
    ("ゾーホージャパン株式会社",         "https://www.zoho.com/jp/crm/lp/crm-software-new.html", "リスティング広告", "", "Zoho CRM・MA・無料プランあり・国産対応",      "Zoho CRM 料金 評判, CRM ツール 比較, 顧客 管理 無料 使える",         "CRM SaaS"),
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
