import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03az.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 56, 16, 16, 22, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

# 全社 WebSearch実確認済みLP（/lp/ /trial/ /request/ 等）
DATA = [
    # ── インサイドセールス代行（LP） ──
    ("株式会社スマートキャンプ",         "https://bales.smartcamp.co.jp/",                          "リスティング広告", "", "インサイドセールス代行・BDR・商談創出",     "インサイドセールス 代行 おすすめ, BDR アウトソーシング 費用, 商談 獲得 代行", "マーケ支援"),
    # ── SFA/CRM（LP） ──
    ("株式会社ナレッジスイート",         "https://knowledgesuite.jp/service/sfa-hikaku.html",       "リスティング広告", "", "SFA/CRM・営業支援・顧客管理・無制限",       "SFA CRM ツール 比較, ナレッジスイート 料金 評判, 顧客 管理 クラウド",     "SFA SaaS"),
    # ── ネットショップ開設（LP） ──
    ("ストアーズ・ドット・ジェーピー株式会社", "https://stores.jp/lp/ec/04",                        "リスティング広告", "", "無料ネットショップ開設・EC構築",            "ネットショップ 無料 開設 簡単, EC 通販 開業 方法, STORES 評判 料金",      "ECプラットフォーム"),
    # ── 工数管理SaaS（LP） ──
    ("株式会社クラウドワークス",         "https://lp.crowdlog.jp/to-c",                            "リスティング広告", "", "工数管理・プロジェクト管理・個人向け",       "工数管理 ツール 個人 フリーランス, プロジェクト 管理 クラウド 比較, クラウドログ 評判", "プロジェクト管理SaaS"),
    # ── 経費精算SaaS（LP） ──
    ("株式会社ラクス",                  "https://www.rakurakuseisan.jp/lp/brand02.html",            "リスティング広告", "", "クラウド経費精算・導入社数No.1",            "経費精算 システム 比較 おすすめ, 楽楽精算 料金 評判, 経費 ペーパーレス DX", "バックオフィスSaaS"),
    # ── ノーコードアプリ開発（LP） ──
    ("株式会社ヤプリ",                  "https://yappli.co.jp/service/",                           "リスティング広告", "", "ノーコードアプリ開発・シェアNo.1",          "ノーコード アプリ 開発 ツール, スマホ アプリ 企業 作成 費用, Yappli 評判", "ノーコードSaaS"),
    # ── 会計クラウド（1ヶ月無料LP） ──
    ("株式会社マネーフォワード",         "https://biz.moneyforward.com/trial/",                     "リスティング広告", "", "クラウド会計・経費精算・1ヶ月無料",         "クラウド 会計 ソフト 比較, マネーフォワード 料金 評判, 経費 精算 無料 試し", "会計SaaS"),
    # ── AI契約書審査（資料請求LP） ──
    ("株式会社LegalOn Technologies",    "https://www.legalon-cloud.com/legalforce/request",        "リスティング広告", "", "AI契約書レビュー・LegalForce・法務DX",      "AI 契約書 レビュー システム, LegalForce 料金 評判, 法務 DX ツール 比較",  "LegalTech SaaS"),
    # ── 営業管理SaaS（公式LP） ──
    ("株式会社ジーニー",                "https://sfacrm.geniee.co.jp/",                            "リスティング広告", "", "SFA/CRM・営業管理・国産クラウド",           "GENIEE SFA CRM 料金 評判, 営業 管理 ツール 国産, 顧客 管理 SaaS 比較",   "SFA SaaS"),
    # ── 求人検索エンジン（採用企業LP） ──
    ("株式会社スタンバイ",              "https://jinji.stanby.co.jp/",                             "リスティング広告", "", "求人検索エンジン・採用広告・クリック課金",   "スタンバイ 求人 掲載 料金, 採用 広告 求人 検索, 中途採用 求人 エンジン",  "採用メディア"),
    # ── 店舗向け光回線（申し込みLP） ──
    ("株式会社USEN",                    "https://usen.com/service/wifi/uhikari/",                  "リスティング広告", "", "店舗向け光回線・AIR UNLIMITED・Wi-Fi",      "店舗 光回線 おすすめ, USEN AIR 評判 料金, 飲食 店 Wi-Fi 業務用",         "通信・回線"),
    # ── 年末調整BPO（LP） ──
    ("ラクラス株式会社",                "https://www.lacras.co.jp/lp/tax/",                        "リスティング広告", "", "年末調整BPO・給与アウトソーシング・大企業向け", "年末調整 代行 アウトソーシング, 給与計算 BPO 費用, 人事 労務 外注",      "人事BPO"),
    # ── ワークフロー（30日無料LP） ──
    ("株式会社コラボスタイル",          "https://lp.collabo-style.co.jp/trial.html",               "リスティング広告", "", "ワークフローシステム・30日無料・顧客満足度No.1", "ワークフロー システム 比較, 申請 承認 クラウド おすすめ, コラボフロー 評判", "業務DX SaaS"),
    # ── タレントマネジメント（LP） ──
    ("株式会社カオナビ",                "https://www.kaonavi.jp/",                                 "リスティング広告", "", "タレントマネジメント・人事評価・シェアNo.1",  "タレントマネジメント シェアNo1, カオナビ 料金 評判, 人事評価 クラウド",   "HR SaaS"),
    # ── アフィリエイトASP（広告主LP） ──
    ("株式会社インタースペース",         "https://www.accesstrade.ne.jp/merchant",                  "リスティング広告", "", "アフィリエイト広告・ASP・成果報酬型",        "アフィリエイト 広告主 掲載 ASP, 成果 報酬 広告 集客, アクセストレード 申し込み", "デジタル広告"),
    # ── 採用企業向け求人（採用担当者LP） ──
    ("マイナビ株式会社",                "https://doctor.mynavi.jp/lp/agent/",                      "リスティング広告", "", "医師転職・求人調べ代行・完全無料",           "医師 転職 求人 調べ代行, ドクター 転職 エージェント 比較, 医師 求人 紹介", "医療人材"),
    # ── デジタルチラシ（店舗集客LP） ──
    ("株式会社くふうカンパニー",         "https://biz-lp.tokubai.co.jp/",                          "リスティング広告", "", "デジタルチラシ・スーパー・ドラッグ掲載",     "デジタルチラシ 掲載 料金, 食品 スーパー 集客 アプリ, トクバイ 評判",    "流通・小売"),
    # ── HR SaaS（360度評価LP） ──
    ("株式会社シーベース",              "https://www.cbase.co.jp/lp/360_feedback/",                "リスティング広告", "", "360度評価・多面評価・クラウドシステム",       "360度評価 クラウド 比較, 多面評価 ツール おすすめ, フィードバック 人事 導入", "HR SaaS"),
    # ── 経費精算SaaS（デモLP） ──
    ("株式会社TOKIUM",                  "https://www.keihi.com/seminars/tokium_demo/",              "リスティング広告", "", "経費精算・請求書受領クラウド・デモ",          "経費精算 SaaS 比較, TOKIUM 料金 評判, 請求書 受領 電子化",               "バックオフィスSaaS"),
    # ── 電子契約（乗り換えLP） ──
    ("GMOグローバルサイン・ホールディングス株式会社", "https://www.gmosign.com/lp/transfer/",      "リスティング広告", "", "電子契約・電子印鑑・乗り換えキャンペーン",   "GMOサイン 乗り換え, 電子印鑑 サービス 比較, 電子契約 移行 キャンペーン", "電子契約SaaS"),
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
