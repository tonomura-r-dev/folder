import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03ac.xlsx"
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
    # ── シューズ・ウェア（高単価） ──────────────────────────────────────────
    ("ニューエラ ジャパン株式会社（New Era）",              "https://www.neweracap.jp/",                     "リスティング広告", "", "キャップ・スナップバック・スポーツウェア",              "ニューエラ キャップ おすすめ 人気, New Era MLB 購入, ストリート ブランド","キャップ"),
    ("ティンバーランド ジャパン合同会社（Timberland）",     "https://www.timberland.co.jp/",                 "リスティング広告", "", "防水ブーツ・イエローブーツ・アウトドアシューズ",         "ティンバーランド ブーツ 評判 おすすめ, イエロー ブーツ 購入, Timberland 防水","ブーツ"),
    ("テバ ジャパン（Teva）",                               "https://teva.co.jp/",                           "リスティング広告", "", "スポーツサンダル・ハリケーン・アウトドアシューズ",       "テバ サンダル 評判 おすすめ, スポーツ サンダル 比較, Teva 購入 価格","サンダル"),
    ("バーケンストック ジャパン合同会社（Birkenstock）",    "https://www.birkenstock.com/jp/",               "リスティング広告", "", "健康サンダル・フッドベッド・ボストン・チューリッヒ",     "バーケンストック サンダル 評判 人気, 健康 サンダル 比較, Birkenstock 購入","健康サンダル"),
    ("ヴァイブラム ジャパン合同会社（Vibram）",             "https://www.vibram.com/jp/",                    "リスティング広告", "", "ファイブフィンガーズ・裸足感覚シューズ・アウトドア",     "ビブラム シューズ 評判 おすすめ, 裸足 感覚 ランニング, Vibram 購入",   "アウトドアシューズ"),
    # ── 旅行（高単価） ──────────────────────────────────────────
    ("ANAセールス株式会社（ANAトラベラーズ）",              "https://www.anatravel.com/",                    "リスティング広告", "", "国内外旅行パッケージ・ANA航空券セット・ホテル",          "ANAトラベラーズ 旅行 評判 おすすめ, 国内旅行 パック 安い, ANA セット ホテル","旅行"),
    # ── 高級ホテル（高単価） ──────────────────────────────────────────
    ("株式会社リーガロイヤルホテル（Rihga Royal）",         "https://www.rihga.co.jp/osaka/",                "リスティング広告", "", "高級ホテル・ブライダル・ダイニング・大阪",               "リーガロイヤルホテル 宿泊 予約 評判, 大阪 高級 ホテル 比較, ブライダル 費用","高級ホテル"),
    ("株式会社目黒雅叙園（HOTEL GAJOEN TOKYO）",            "https://www.hotelgajoen-tokyo.com/",            "リスティング広告", "", "高級ウエディング・宿泊・和の美術館ホテル",              "雅叙園 東京 宿泊 予約, ウエディング 会場 高級 費用, ホテル 雅叙園 評判","高級ホテル"),
    # ── 高級和牛EC（高単価） ──────────────────────────────────────────
    ("有限会社千成亭（近江牛）",                            "https://www.sennariteii.co.jp/",                "リスティング広告", "", "近江牛・和牛すき焼き・牛肉ギフト通販",                   "千成亭 近江牛 評判 おすすめ, 和牛 ギフト 通販, 近江牛 すき焼き お取り寄せ","高級和牛EC"),
    # ── 高級食品EC（高単価） ──────────────────────────────────────────
    ("株式会社ふくや",                                     "https://www.fukuya.com/",                       "リスティング広告", "", "博多明太子・辛子明太子・ギフト通販",                     "ふくや 明太子 評判 おすすめ, 博多 明太子 ギフト 通販, 明太子 お取り寄せ","高級食品EC"),
    # ── 日用品・ヘアケア（高単価） ──────────────────────────────────────────
    ("P&G ジャパン合同会社（パンテーン・ジレット）",        "https://jp.pg.com/",                            "リスティング広告", "", "ヘアケア・シェービング・洗剤・スキンケア",               "パンテーン シャンプー おすすめ 評判, ジレット シェーバー 比較, P&G 購入","日用品"),
    # ── 高性能タイヤ（高単価） ──────────────────────────────────────────
    ("ミシュラン ジャパン株式会社（Michelin）",              "https://www.michelin.co.jp/",                   "リスティング広告", "", "高性能タイヤ・プライマシー・スポーツタイヤ",              "ミシュラン タイヤ 評判 おすすめ, 低燃費 高性能 タイヤ 比較, Michelin 購入","タイヤ"),
    ("コンチネンタル ジャパン株式会社（Continental）",      "https://www.conti-japan.co.jp/",                "リスティング広告", "", "スポーツタイヤ・プレミアムコンタクト・スタッドレス",     "コンチネンタル タイヤ 評判 おすすめ, ドイツ 高品質 タイヤ 比較, Continental","タイヤ"),
    ("ピレリジャパン株式会社（Pirelli）",                   "https://www.pirelli.com/tyres/ja-jp/",          "リスティング広告", "", "スポーツタイヤ・Pゼロ・チンチュラート",                  "ピレリ タイヤ 評判 おすすめ, スポーツ タイヤ 高性能 比較, Pirelli 購入","タイヤ"),
    # ── 競泳水着（高単価） ──────────────────────────────────────────
    ("TYR スポーツ ジャパン（TYR）",                        "https://tyr.co.jp/",                            "リスティング広告", "", "競泳水着・スイムウェア・ゴーグル・水泳用品",              "TYR 競泳水着 評判 おすすめ, 水泳 用品 高機能 比較, TYR ゴーグル 購入","競泳用品"),
    # ── 高性能ロードバイク（超高単価） ──────────────────────────────────────────
    ("ルイガノ ジャパン株式会社（Louis Garneau）",           "https://louisgarneau.jp/",                      "リスティング広告", "", "カナダ製ロードバイク・クロスバイク・スポーツウェア",      "ルイガノ 自転車 評判 おすすめ, Louis Garneau ロード 購入, スポーツ 自転車 比較","ロードバイク"),
    ("フェルト ジャパン（Felt Bikes）",                     "https://feltbicycles.jp/",                      "リスティング広告", "", "アメリカ製ロードバイク・エアロ・クロスバイク",            "フェルト 自転車 評判 おすすめ, ロードバイク 高性能 比較, Felt 購入 価格","ロードバイク"),
    ("GT バイシクルズ ジャパン（GT Bicycles）",              "https://www.gtbicycles.com/jp/",                "リスティング広告", "", "マウンテンバイク・BMX・クロスバイク",                    "GT 自転車 評判 おすすめ, MTB マウンテンバイク 比較, GT Bicycles 購入","自転車"),
    ("コラテック ジャパン株式会社（Corratec）",              "https://www.corratec.co.jp/",                   "リスティング広告", "", "ドイツ製ロードバイク・Eバイク・カーボンフレーム",         "コラテック ロードバイク 評判, ドイツ 自転車 高品質 比較, Eバイク 購入","ロードバイク"),
    ("ラレー ジャパン（Raleigh）",                          "https://raleigh.jp/",                           "リスティング広告", "", "英国クラシックブランド自転車・クロスバイク・ロード",       "ラレー 自転車 評判 おすすめ, Raleigh クラシック ブランド 比較, 購入","自転車"),
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
