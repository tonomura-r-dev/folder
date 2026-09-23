import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03ay.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 56, 16, 16, 22, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

# 全社 WebSearch実確認済みLP（/lp/ /campaign/ /signup/ 等）
DATA = [
    # ── 英語発音AI（アプリLP） ──
    ("ELSA Corp（ELSA Speak）",       "https://elsaspeak.com/ja/",                                  "リスティング広告", "", "AI英語発音・スピーキング練習アプリ",      "英語 発音 アプリ おすすめ, ELSA Speak 効果 口コミ, スピーキング 練習 AI",     "英語学習"),
    # ── VPS（申し込みLP） ──
    ("株式会社カゴヤ・ジャパン",      "https://www.kagoya.jp/vps/",                                 "リスティング広告", "", "クラウドVPS・レンタルサーバー・格安",     "VPS サーバー 安い おすすめ, KAGOYA クラウド 評判, レンタルサーバー 比較",   "レンタルサーバー"),
    # ── 法人英会話（LP） ──
    ("株式会社レアジョブ（法人）",    "https://www.rarejob.com/pr/lp2a/",                           "リスティング広告", "", "法人向けオンライン英会話・社員研修",       "法人 英会話 研修 おすすめ, 社員 英語 スキル アップ 費用, オンライン 英会話 法人", "語学研修"),
    # ── 電子契約SaaS（法人申し込みLP） ──
    ("弁護士ドットコム株式会社（クラウドサイン）", "https://lp.cloudsign.jp/cloudsign_bitmix_business.html", "リスティング広告", "", "電子契約・クラウドサイン・国内シェアNo.1", "電子契約 クラウド 比較, クラウドサイン 料金 評判, 契約書 電子化 方法",     "電子契約SaaS"),
    # ── 法律相談（LP） ──
    ("日本司法支援センター（法テラス）", "https://www.houterasu.or.jp/lp/shakkin2022a/",            "リスティング広告", "", "借金・法律・無料相談・国設法律窓口",       "借金 無料相談 弁護士, 法テラス 相談 無料, 過払い 法律 相談 電話",           "法律相談"),
    # ── ニュース電子版（1ヶ月無料LP） ──
    ("日本経済新聞社",                "https://www.nikkei.com/promotion/",                          "リスティング広告", "", "日経電子版・1ヶ月無料・ニュース購読",       "日経新聞 電子版 無料 体験, 日経 購読 料金 比較, ビジネス ニュース 定期読む",  "メディア・ニュース"),
    # ── フリーランスエンジニアエージェント（LP） ──
    ("株式会社クラウドワークス（テック）", "https://tech.crowdworks.jp/lp/member/engineer_sp4",     "リスティング広告", "", "フリーランスITエンジニア・案件紹介",       "フリーランス エンジニア 案件 おすすめ, IT 業務委託 案件 探す, クラウドテック 評判", "フリーランス支援"),
    # ── ジム会員管理SaaS（24hジム向けLP） ──
    ("株式会社hacomono",              "https://lp.hacomono.jp/24h",                                 "リスティング広告", "", "24時間ジム会員管理・予約決済システム",     "ジム 会員管理 システム 比較, hacomono 評判 料金, フィットネス 予約 DX",      "ウェルネスSaaS"),
    # ── マウスピース矯正（フロー申し込みLP） ──
    ("株式会社Oh my teeth",           "https://www.oh-my-teeth.com/flow",                           "リスティング広告", "", "通わないマウスピース矯正・2ヶ月～",        "マウスピース矯正 通わない 安い, 歯並び 矯正 料金 比較, Oh my teeth 評判",   "歯科矯正"),
    # ── 地域情報・採用（掲載サポートLP） ──
    ("株式会社ジモティー",            "https://jmty.co.jp/archives/requirement/11043/",             "リスティング広告", "", "ジモティー掲載・アルバイト採用・無料",     "ジモティー 採用 求人 掲載, アルバイト 無料 求人 サイト, 地域 採用 効果",    "採用メディア"),
    # ── 後払い決済BtoB（LP） ──
    ("株式会社ネットプロテクションズ","https://corp.netprotections.com/lp/common_01/",              "リスティング広告", "", "NP後払い・BtoB後払い決済・請求代行",       "後払い 決済 BtoB, 請求書払い 代行 おすすめ, NP後払い 料金 比較",           "決済サービス"),
    # ── CRM・マーケSaaS（デモ申し込みLP） ──
    ("HubSpot Japan株式会社",         "https://offers.hubspot.com/jp/demo",                         "リスティング広告", "", "CRM・マーケティング・無料デモ",            "HubSpot 無料 デモ, CRM ツール 比較 無料, マーケティング オートメーション",  "CRM SaaS"),
    # ── 医師転職（LP） ──
    ("マイナビ（マイナビDOCTOR）",    "https://doctor.mynavi.jp/lp/028.html",                       "リスティング広告", "", "医師転職・求人・スカウト・完全無料",        "医師 転職 エージェント おすすめ, ドクター 転職 無料, 医師 求人 紹介",       "医療人材"),
    # ── デジタルチラシ（店舗掲載LP） ──
    ("株式会社くふうカンパニー（トクバイ）", "https://biz-lp.tokubai.co.jp/",                      "リスティング広告", "", "デジタルチラシ・スーパー・店舗集客",       "デジタルチラシ 店舗 集客, トクバイ 掲載 料金 効果, スーパー チラシ アプリ", "店舗集客メディア"),
    # ── タレントマネジメント（資料・デモLP） ──
    ("株式会社プラスアルファ・コンサルティング（タレントパレット）", "https://www.talent-palette.com/", "リスティング広告", "", "タレントマネジメント・人材分析SaaS",    "タレントマネジメント システム 比較, 人材 管理 クラウド, タレントパレット 評判 料金", "HR SaaS"),
    # ── 360度評価SaaS（LP） ──
    ("株式会社シーベース（CBASE）",   "https://www.cbase.co.jp/lp/360_feedback/",                   "リスティング広告", "", "360度評価・多面評価クラウドシステム",      "360度評価 システム おすすめ, 多面評価 クラウド 比較, フィードバック 人事 DX", "HR SaaS"),
    # ── 経費精算SaaS（デモセミナーLP） ──
    ("株式会社TOKIUM",                "https://www.keihi.com/seminars/tokium_demo/",                 "リスティング広告", "", "経費精算・請求書受領クラウド・デモ",        "経費精算 クラウド 比較, TOKIUM 評判 料金, 請求書 受領 電子化 ペーパーレス", "バックオフィスSaaS"),
    # ── 電子契約（GMOサイン）（乗り換えLP） ──
    ("GMOグローバルサイン・ホールディングス株式会社（GMOサイン）", "https://www.gmosign.com/lp/transfer/", "リスティング広告", "", "電子契約・電子印鑑・国内シェアNo.1",   "GMOサイン 乗り換え キャンペーン, 電子印鑑 サービス 比較, 電子契約 移行 方法", "電子契約SaaS"),
    # ── エンジニア採用SaaS（企業向けLP） ──
    ("ファインディ株式会社",          "https://recruiting.findy-code.io/expert_lp0",                "リスティング広告", "", "エンジニアスカウト型採用・GitHubスキル分析", "エンジニア 採用 スカウト, Findy 評判 料金, IT 人材 採用 難易度",           "採用SaaS"),
    # ── 医師採用SaaS（法人LP） ──
    ("エムスリーキャリア株式会社（M3 Career Prime）", "https://enzine.m3career.com/lp/prime/entry/", "リスティング広告", "", "医師採用・医師紹介サービス・医療法人",    "医師 採用 紹介 サービス, 医療 クリニック 求人 掲載, M3キャリア 料金",      "医療人材採用"),
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
