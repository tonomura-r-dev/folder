import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03aa.xlsx"
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
    # ── 高級輸入車（超高単価） ──────────────────────────────────────────
    ("フォルクスワーゲン グループ ジャパン株式会社（VW）",  "https://www.volkswagen.co.jp/",                 "リスティング広告", "", "ゴルフ・ポロ・ティグアン・輸入車",                     "フォルクスワーゲン 新車 評判 おすすめ, VW ゴルフ 購入 価格, 輸入車 比較","輸入車"),
    ("プジョー・ジャポン合同会社（Peugeot）",               "https://www.peugeot.jp/",                       "リスティング広告", "", "フランス製SUV・セダン・EVシリーズ",                     "プジョー 新車 評判 おすすめ, フランス 輸入車 比較 価格, SUV 試乗 購入","輸入車"),
    ("マセラティ ジャパン株式会社（Maserati）",             "https://www.maserati.com/jp/ja/",               "リスティング広告", "", "イタリア高級スポーツカー・グレカーレ・ギブリ",           "マセラティ 新車 価格 評判, 高級 スポーツカー イタリア, Maserati 試乗","高級輸入車"),
    ("ジャガー・ランドローバー・ジャパン株式会社（JLR）",   "https://www.jaguar.co.jp/",                     "リスティング広告", "", "高級英国車・ジャガー・ランドローバー・ディフェンダー",   "ジャガー 新車 評判 価格, ランドローバー SUV 比較, Land Rover 試乗 購入","高級輸入車"),
    ("FCAジャパン合同会社（Jeep）",                         "https://www.jeep-japan.com/",                   "リスティング広告", "", "アメリカ製SUV・ラングラー・チェロキー・コンパス",       "ジープ 新車 評判 おすすめ, Jeep ラングラー 購入 価格, SUV アメリカン 試乗","輸入SUV"),
    # ── 超高単価スーツ（超高単価） ──────────────────────────────────────────
    ("エルメネジルド・ゼニア ジャパン合同会社（Zegna）",    "https://www.zegna.com/jp-ja/",                  "リスティング広告", "", "高級スーツ・カシミア・ウールファブリック",               "ゼニア スーツ 評判 おすすめ, 高級 生地 スーツ ブランド, Zegna 購入",   "高級スーツ"),
    # ── 分譲マンション（超高単価） ──────────────────────────────────────────
    ("野村不動産株式会社（PROUD）",                         "https://www.proud-web.jp/",                     "リスティング広告", "", "高級分譲マンション・プラウド・資産価値",                 "野村不動産 プラウド 評判 価格, 分譲 マンション 高級 購入, PROUD 物件","分譲マンション"),
    ("住友不動産株式会社（シティハウス）",                   "https://www.sumitomo-rd.co.jp/",                "リスティング広告", "", "高級分譲マンション・シティハウス・超高層",               "住友不動産 マンション 評判 価格, シティハウス 分譲 比較, 高層 マンション","分譲マンション"),
    # ── 健康管理・計測（高単価） ──────────────────────────────────────────
    ("株式会社タニタ（TANITA）",                            "https://www.tanita.co.jp/",                     "リスティング広告", "", "体重計・体組成計・活動量計・血圧計",                     "タニタ 体重計 おすすめ 評判, 体組成計 比較 精度, TANITA 体脂肪 計測",  "健康計測"),
    # ── 電動工具（高単価） ──────────────────────────────────────────
    ("株式会社リョービ（RYOBI）",                           "https://www.ryobitools.jp/",                    "リスティング広告", "", "電動工具・充電式ドリル・DIYツール・切断機",              "リョービ 電動工具 おすすめ 評判, DIY 充電式 ドリル 比較, RYOBI 購入","電動工具"),
    ("工機ホールディングス株式会社（HiKOKI）",              "https://www.hikoki-powertools.jp/",              "リスティング広告", "", "コードレス電動工具・マルチボルト・プロ向け工具",          "HiKOKI 電動工具 おすすめ 評判, マルチ ボルト 比較, 工機 ドリル 購入","電動工具"),
    # ── 高級洋菓子EC（高単価） ──────────────────────────────────────────
    ("株式会社モロゾフ（Morozoff）",                        "https://www.morozoff.co.jp/",                   "リスティング広告", "", "高級洋菓子・チョコレート・ギフト通販",                   "モロゾフ 洋菓子 評判 おすすめ, チョコ ギフト 通販 高品質, Morozoff 購入","高級洋菓子"),
    ("株式会社ゴンチャロフ製菓（GONCHAROFF）",              "https://www.goncharoff.co.jp/",                 "リスティング広告", "", "高級チョコレート・プラリネ・ギフト通販",                 "ゴンチャロフ チョコ 評判 おすすめ, 高級 ギフト 洋菓子, GONCHAROFF 購入","高級洋菓子"),
    ("株式会社ウエスト（西洋菓子ウエスト）",                 "https://www.ec-west.jp/",                       "リスティング広告", "", "バームクーヘン・リーフパイ・九州銘菓ギフト",             "ウエスト バームクーヘン 評判 通販, 九州 銘菓 ギフト おすすめ, リーフパイ 購入","高級洋菓子"),
    # ── 高級自転車（超高単価） ──────────────────────────────────────────
    ("ビアンキ ジャパン（Bianchi）",                        "https://www.bianchi.com/jp/",                   "リスティング広告", "", "イタリア最古の自転車ブランド・ロード・MTB",              "ビアンキ ロードバイク 評判 おすすめ, イタリア 自転車 ブランド 比較, Bianchi 購入","ロードバイク"),
    # ── 高級食品EC（高単価） ──────────────────────────────────────────
    ("株式会社ベルーナグルメ",                               "https://gourmet.belluna.jp/",                   "リスティング広告", "", "高級食材・牛肉・海産物・ギフト通販",                     "ベルーナグルメ 評判 おすすめ, 高級 牛肉 通販 お取り寄せ, ギフト 食材 購入","高級食品EC"),
    # ── カメラ・映像機器（高単価） ──────────────────────────────────────────
    ("株式会社シグマ（SIGMA）",                             "https://www.sigma-global.com/jp/",              "リスティング広告", "", "高品質カメラレンズ・単焦点・Artシリーズ",               "シグマ レンズ おすすめ 評判, カメラ 単焦点 高品質 比較, SIGMA Art 購入","カメラレンズ"),
    ("マンフロット ジャパン合同会社（Manfrotto）",          "https://www.manfrotto.com/ja-jp/",              "リスティング広告", "", "カメラ三脚・雲台・映像機材・スタビライザー",              "マンフロット 三脚 評判 おすすめ, カメラ 雲台 比較 高品質, Manfrotto 購入","映像機材"),
    ("ゴープロ ジャパン合同会社（GoPro）",                  "https://gopro.com/ja/jp/",                      "リスティング広告", "", "アクションカメラ・HERO・360°撮影・防水",               "GoPro カメラ おすすめ 評判, アクション カメラ 比較 防水, ゴープロ 購入","アクションカメラ"),
    ("RØDE マイクロフォンズ ジャパン（RØDE）",              "https://www.rode.com/ja/",                      "リスティング広告", "", "コンデンサーマイク・ワイヤレス・配信・録音機材",         "ロード マイク おすすめ 評判, 配信 マイク 高品質 比較, RODE ワイヤレス 購入","音響機材"),
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
