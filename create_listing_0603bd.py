import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03bd.xlsx"
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
    # ── オンボードAI ──
    ("株式会社ビズリーチ",             "https://onboard-ai.jp/lp/",                              "リスティング広告", "", "オンボードAI・入社定着・マネジメント改革", "オンボード AI 入社 定着, 新入社員 育成 ツール, Onboard AI 料金 評判",  "HR SaaS"),
    # ── テスト自動化SaaS ──
    ("株式会社Autify",                 "https://autify.com/ja/demo",                              "リスティング広告", "", "E2Eテスト自動化・ノーコード・QA効率化",    "テスト 自動化 ノーコード, QA SaaS 比較, Autify 料金 評判 デモ",        "DevOps SaaS"),
    # ── 物流SaaS ──
    ("ハコベル株式会社",               "https://lp.hacobell.com/",                                "リスティング広告", "", "物流DX・配送マッチング・トラック手配",      "物流 DX プラットフォーム, トラック 手配 マッチング, ハコベル 料金",      "物流SaaS"),
    # ── 医師向け広告媒体 ──
    ("メドピア株式会社",               "https://medpeer.co.jp/service/company.html",              "リスティング広告", "", "医師18万人プラットフォーム・医薬品広告掲載", "医師 向け 広告 媒体, 製薬 マーケティング 支援, メドピア 掲載 費用",     "医療メディア"),
    # ── オンライン商談 ──
    ("ベルフェイス株式会社",            "https://bell-face.com/lp/",                               "リスティング広告", "", "電話×オンライン商談・金融シェアNo.1",       "オンライン 商談 ツール 比較, ベルフェイス 料金 評判, Web会議 営業",      "セールスSaaS"),
    # ── 採用求人掲載 ──
    ("株式会社リクルート",             "https://townwork.net/assets/twc/jobpost/gatepage/index.html", "リスティング広告", "", "タウンワーク求人掲載・アルバイト採用",    "タウンワーク 求人 掲載 料金, アルバイト 採用 広告 比較, 求人 掲載 申し込み", "採用メディア"),
    # ── プロシェアリング ──
    ("株式会社サーキュレーション",     "https://circu.co.jp/lp_hojin_01/",                        "リスティング広告", "", "プロシェアリング・副業プロ・経営課題解決",   "プロシェアリング 副業 専門家, 経営 課題 外部 人材, サーキュレーション 料金", "人材マッチング"),
    # ── 法人英語研修 ──
    ("株式会社リクルート",             "https://eigosapuri-biz.jp/",                              "リスティング広告", "", "スタディサプリENGLISH法人・ビジネス英語研修", "スタディサプリ 法人 英語 研修, ビジネス 英語 アプリ 法人, スタサプ 法人 料金", "語学研修"),
    # ── 大学院生就活 ──
    ("株式会社アカリク",               "https://acaric.jp/special/register-now-general",          "リスティング広告", "", "大学院生・理系学生特化就活・スカウト",      "大学院生 就活 サイト おすすめ, 理系 院生 転職 エージェント, アカリク 評判", "就活サービス"),
    # ── CRM MA SaaS ──
    ("シナジーマーケティング株式会社",  "https://www.synergy-marketing.co.jp/lp/synergy/degital_marketing/", "リスティング広告", "", "CRM・MA・国産ツール・伴走支援付き",    "CRM MA ツール 国産, Synergy 料金 評判, メール配信 CRM 一体型",          "CRM SaaS"),
    # ── インテントセールス ──
    ("株式会社Sales Marker",           "https://sales-marker.jp/lp/demo/",                        "リスティング広告", "", "インテントデータ・BtoBターゲティング営業",   "インテント セールス ツール, 企業 行動データ 営業, セールスマーカー 料金", "セールスSaaS"),
    # ── 製造業調達DX ──
    ("キャディ株式会社",               "https://caddi.jp/lp/semicon/",                            "リスティング広告", "", "加工品調達DX・製造業向け・コストダウン",     "製造 調達 DX, 加工品 一式 コスト削減, キャディ 評判 料金",             "製造業SaaS"),
    # ── ベルフェイス(採用) ──
    ("株式会社Autify",                 "https://autify.com/ja/demo",                              "リスティング広告", "", "テスト自動化・E2E・AIノーコード",           "テスト 自動化 ツール 比較, Autify 料金 評判, E2E テスト SaaS",          "DevOps SaaS"),
    # ── 医師求人 ──
    ("株式会社マイナビ",               "https://doctor.mynavi.jp/lp/agent/",                      "リスティング広告", "", "医師転職・求人代行・転職支援",               "医師 転職 エージェント 無料, ドクター 転職 サイト 比較, マイナビ 医師",   "医療人材"),
    # ── マーケプラットフォーム ──
    ("株式会社マジセミ",               "https://majisemi.com/service/lp/",                        "リスティング広告", "", "IT企業向けウェビナー集客・独自ハウスリスト", "ウェビナー 集客 代行 費用, IT BtoB セミナー 申し込み, マジセミ 評判",   "マーケ支援"),
    # ── 建設工事受発注 ──
    ("クラフトバンク株式会社",         "https://craft-bank.com/lp/plan_index",                    "リスティング広告", "", "建設工事受発注・協力会社マッチング",          "協力会社 マッチング, 建設 工事 受発注, クラフトバンク 評判 料金",        "建設SaaS"),
    # ── 物流マッチング ──
    ("ハコベル株式会社",               "https://www.hacobell.com/matching",                       "リスティング広告", "", "軽貨物〜一般貨物・運送手配・即日対応",       "トラック 手配 急ぎ, 貨物 マッチング サービス, ハコベル 運送 料金",       "物流SaaS"),
    # ── リクルートスタッフィング ──
    ("株式会社リクルートスタッフィング", "https://www.r-staffing.co.jp/sol/contents/client/lp/bpo_01/", "リスティング広告", "", "業務委託・アウトソーシング・BPO",       "業務 委託 アウトソーシング 比較, BPO 人材 紹介, リクルートスタッフィング 評判", "人材派遣"),
    # ── 採用LP制作 ──
    ("株式会社ネオキャリア",            "https://www.neo-career.co.jp/lp/rpo/",                   "リスティング広告", "", "採用代行RPO・採用プロセスアウトソーシング", "採用代行 RPO 費用, 採用 プロセス 外注, ネオキャリア 評判",             "採用支援"),
    # ── インフォマート ──
    ("株式会社インフォマート",          "https://lp.infomart.co.jp/seikyu/free-plan/",             "リスティング広告", "", "BtoBプラットフォーム請求書・電子化無料",    "請求書 電子化 無料 試し, インフォマート 料金 評判, BtoB 請求書 クラウド", "フィンテックSaaS"),
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
