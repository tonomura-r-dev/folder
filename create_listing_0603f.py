import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03f.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

DATA = [
    # ── 霊園・お墓・終活（高単価） ──────────────────────────────────────────
    ("株式会社鎌倉新書（いいお墓）",                   "https://www.e-ohaka.com/",                       "リスティング広告", "", "霊園・墓地・お墓の購入比較",             "お墓 購入 費用 相場, 霊園 探し方 近く, 樹木葬 永代供養 比較",               "霊園・お墓"),
    ("株式会社ヤシロ（霊園・墓石のヤシロ）",           "https://www.yashiro.co.jp/",                     "リスティング広告", "", "墓石・霊園・お墓のリフォーム",           "墓石 購入 費用, お墓 建てる 流れ, 霊園 近畿 おすすめ 購入",                 "霊園・お墓"),
    ("株式会社ニチリョク（民間霊園）",                 "https://www.nichryo.co.jp/",                     "リスティング広告", "", "民間霊園・納骨堂・永代供養",             "霊園 民間 購入, 納骨堂 費用 比較, 永代供養 安い おすすめ",                   "霊園・お墓"),
    ("株式会社ハウスボートクラブ（海洋散骨）",         "https://www.umi-sougi.jp/",                      "リスティング広告", "", "海洋散骨・自然葬サービス",               "海洋散骨 費用 流れ, 自然葬 散骨 おすすめ, 散骨 業者 東京",                  "終活・散骨"),
    # ── 旅行・クルーズ（高額ツアー） ──────────────────────────────────────────
    ("株式会社読売旅行",                               "https://www.yomiuri-ryoko.jp/",                  "リスティング広告", "", "国内・海外ツアー・バス旅行",             "旅行 ツアー おすすめ 安い, シニア 旅行 バス, 国内旅行 温泉 パック",          "旅行"),
    ("株式会社はとバス",                               "https://www.hatobus.com/",                       "リスティング広告", "", "東京観光バスツアー・日帰り旅行",         "東京 観光 バスツアー 日帰り, はとバス 申し込み, 東京 定期観光 ガイド",       "旅行・観光"),
    ("株式会社ユーラシア旅行社",                       "https://www.eurasia.co.jp/",                     "リスティング広告", "", "クルーズ・海外高額ツアー",               "クルーズ 旅行 おすすめ 費用, 豪華客船 ツアー 申し込み, 海外旅行 高級 パック","旅行・クルーズ"),
    ("ロイヤル・カリビアン・ジャパン株式会社",         "https://www.royalcaribbean.com/jpn/ja/",         "リスティング広告", "", "大型豪華クルーズ船・世界一周",           "クルーズ 申し込み 日本から, ロイヤルカリビアン 料金 コース, 豪華 船旅 体験","クルーズ"),
    ("MSCクルーズ・ジャパン株式会社",                 "https://www.msccruises.jp/",                     "リスティング広告", "", "地中海・世界クルーズ旅行",               "地中海 クルーズ ツアー 費用, MSCクルーズ 日本発, 欧州 クルーズ 申し込み",   "クルーズ"),
    # ── 弁護士・法律（高単価） ──────────────────────────────────────────
    ("弁護士法人みやび",                               "https://www.bengoshi-miyabi.com/",               "リスティング広告", "", "債務整理・過払い金・自己破産",           "債務整理 弁護士 無料相談, 過払い金 請求 流れ, 借金 解決 弁護士 選び方",     "法律"),
    ("弁護士法人フォーカスクライド",                   "https://focus-clide.jp/",                        "リスティング広告", "", "相続・遺産・離婚の弁護士",               "相続 弁護士 無料相談, 遺産 分割 費用, 離婚 弁護士 費用 相場",               "法律"),
    # ── 住宅ローン・金融（高単価） ──────────────────────────────────────────
    ("株式会社SBIマネープラザ",                       "https://www.sbimoneyhome.co.jp/",                "リスティング広告", "", "住宅ローン相談・借り換え比較",           "住宅ローン 比較 おすすめ, 借り換え 相談 無料, 住宅ローン 金利 低い 銀行",   "住宅ローン"),
    ("株式会社MFSモーゲージ（モゲチェック）",         "https://mogecheck.jp/",                          "リスティング広告", "", "住宅ローン一括比較・事前審査",           "住宅ローン 一括比較 無料, 事前審査 通過率, 住宅ローン おすすめ 2026",       "住宅ローン"),
    ("楽天モーゲージ株式会社",                        "https://mortgage.rakuten.co.jp/",                "リスティング広告", "", "住宅ローン・変動金利・楽天ポイント",     "楽天 住宅ローン 金利, 住宅ローン 変動 固定 比較, 楽天ポイント 住宅ローン", "住宅ローン"),
    # ── ハイクラス・専門職転職（高単価） ──────────────────────────────────────────
    ("株式会社レバレジーズ（レバテックキャリア）",     "https://career.levtech.jp/",                     "リスティング広告", "", "ITエンジニア・ハイクラス転職支援",       "エンジニア ハイクラス 転職, IT 転職エージェント 年収高い, レバテック 評判", "ITハイクラス転職"),
    ("ロバートハーフ・ジャパン株式会社",               "https://www.roberthalf.jp/",                     "リスティング広告", "", "管理職・外資系・専門職転職",             "外資系 転職 エージェント, 管理職 転職 40代, ハイクラス 転職 年収 1000万",   "ハイクラス転職"),
    ("株式会社プロフェッショナルバンク",               "https://www.professional-bank.com/",             "リスティング広告", "", "財務・経理・法務のプロ転職",             "財務 転職 エージェント, 経理 転職 30代 キャリアアップ, 専門職 転職 高年収", "ハイクラス転職"),
    ("株式会社アトラエ（Green）",                     "https://www.green-japan.com/",                   "リスティング広告", "", "IT・Web・ゲーム業界求人サービス",        "IT 転職 求人 スカウト, Web エンジニア 転職 Green, ゲーム 転職 エンジニア",  "IT転職"),
    ("キャリアトレック株式会社",                      "https://careertrek.com/",                        "リスティング広告", "", "AI×転職スカウトサービス",               "転職 スカウト AI, 年収 アップ 転職 30代, 転職サイト 比較 おすすめ",          "転職"),
    ("株式会社エリートネットワーク",                   "https://www.elite-network.co.jp/",               "リスティング広告", "", "ハイクラス・幹部候補転職支援",           "幹部候補 転職 エージェント, 年収500万 転職 管理職, ハイクラス 転職 エリート","ハイクラス転職"),
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
