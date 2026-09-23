import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03r.xlsx"
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
    # ── 高級ホテル（超高単価） ──────────────────────────────────────────
    ("プリンスホテル株式会社（Prince Hotels）",             "https://www.princehotels.co.jp/",               "リスティング広告", "", "高級ホテル予約・宴会場・ブライダル",               "プリンスホテル 予約 東京 おすすめ, 高級 ホテル 宿泊 記念日, ブライダル 会場 費用","高級ホテル"),
    ("株式会社ニューオータニ（Hotel New Otani）",           "https://www.newotani.co.jp/tokyo/",             "リスティング広告", "", "高級ホテル・レストラン・ウエディング",              "ニューオータニ 宿泊 予約 評判, 高級 ホテル 東京 ランチ, ウエディング 費用 会場","高級ホテル"),
    ("帝国ホテル株式会社（Imperial Hotel）",                "https://www.imperialhotel.co.jp/",              "リスティング広告", "", "最高級ホテル・ダイニング・ビジネス利用",             "帝国ホテル 宿泊 予約 価格, 東京 最高級 ホテル 特別 宿泊, ランチ ビュッフェ 予約","最高級ホテル"),
    ("ホテルオークラ株式会社（Hotel Okura）",               "https://www.hotelokura.co.jp/tokyo/",           "リスティング広告", "", "高級ホテル宿泊・婚礼・宴会",                         "ホテルオークラ 宿泊 予約, 高級 ホテル 結婚式 費用, オークラ ランチ 予約 東京","高級ホテル"),
    ("ヒルトン・ワールドワイド・ジャパン合同会社（Hilton）","https://www.hilton.com/ja/hotels/japan/",      "リスティング広告", "", "高級ホテル・リゾート・ポイントプログラム",           "ヒルトン ホテル 予約 東京, Hilton 宿泊 特典 おすすめ, 高級 リゾート ホテル 会員","高級ホテル"),
    ("日本マリオット株式会社（Marriott Japan）",            "https://www.marriott.com/ja/hotels/japan.mi",   "リスティング広告", "", "高級ホテル・スパ・宿泊パッケージ",                   "マリオット ホテル 予約 日本, Marriott Bonvoy ポイント 活用, 高級 ホテル スパ 宿泊","高級ホテル"),
    # ── バックパック（高単価） ──────────────────────────────────────────
    ("オスプレー・ジャパン株式会社（Osprey）",              "https://www.ospreypacks.com/jp/ja/",            "リスティング広告", "", "登山・ハイキングバックパック・トレッキング",         "オスプレー バックパック おすすめ, 登山 リュック 選び方 容量, Osprey 評判 購入","バックパック"),
    ("グレゴリー ジャパン合同会社（Gregory）",              "https://www.gregorypacks.com/jp/",              "リスティング広告", "", "バックパック・デイパック・トレッキング",              "グレゴリー バックパック 評判, 登山 リュック ブランド 比較, Gregory 容量 選び方","バックパック"),
    ("ドイター ジャパン株式会社（Deuter）",                 "https://www.deuter.com/ja-jp/",                 "リスティング広告", "", "ドイツ製登山バックパック・自転車バッグ",              "ドイター バックパック 評判 おすすめ, 登山 ザック 高機能 比較, Deuter 購入 価格","バックパック"),
    ("カリマー ジャパン株式会社（Karrimor）",               "https://www.karrimor.jp/",                      "リスティング広告", "", "バックパック・トレッキングシューズ・アウトドア",     "カリマー バックパック おすすめ, トレッキング 靴 高機能, Karrimor 評判 購入",  "バックパック"),
    ("ミレー・マウンテン・グループ・ジャパン合同会社（Millet）","https://www.millet.jp/",                  "リスティング広告", "", "登山ウェア・ザック・アルパインクライミング",         "ミレー バックパック 評判 おすすめ, 登山 ウェア 高機能 フランス, Millet 購入","アウトドアウェア"),
    # ── プレミアムコスメ・ボディケア（高単価） ──────────────────────────────────────────
    ("ロクシタン ジャポン株式会社（L'Occitane）",           "https://jp.loccitane.com/",                     "リスティング広告", "", "プロヴァンスコスメ・ボディクリーム・ギフト",         "ロクシタン ボディクリーム おすすめ, L'Occitane ギフト セット 人気, コスメ プレゼント","プレミアムコスメ"),
    ("ポール・ジョー・ジャポン（Paul & Joe）",              "https://www.paulandjoe.com/ja-jp/",             "リスティング広告", "", "フランス製コスメ・リップ・ファンデーション",         "ポール&ジョー コスメ おすすめ, ファンデーション 評判 人気, フランス コスメ 通販","コスメ"),
    # ── ファッション（高単価） ──────────────────────────────────────────
    ("ペンドルトン ジャパン合同会社（Pendleton）",          "https://www.pendleton-jp.com/",                 "リスティング広告", "", "ウールブランケット・ネイティブ柄・ファッション",     "ペンドルトン ブランケット おすすめ, ウール 毛布 高品質, Pendleton 評判 購入","ファッション"),
    # ── 医薬品・ヘルスケア（高単価） ──────────────────────────────────────────
    ("小林製薬株式会社",                                    "https://www.kobayashi.co.jp/",                  "リスティング広告", "", "市販薬・衛生用品・健康補助食品通販",                  "小林製薬 通販 おすすめ, 市販薬 種類 効果, 健康食品 購入 小林製薬",           "医薬品"),
    ("大正製薬株式会社",                                    "https://www.taisho.co.jp/",                     "リスティング広告", "", "栄養ドリンク・市販薬・スキンケア通販",                "大正製薬 リポビタン 購入, 栄養ドリンク 効果 おすすめ, 市販薬 通販 安い",     "医薬品"),
    ("第一三共ヘルスケア株式会社",                          "https://www.daiichisankyo-hc.co.jp/",           "リスティング広告", "", "市販薬・スキンケア・ミノキシジル",                    "第一三共 スキンケア おすすめ, ミノキシジル 育毛剤 市販, 市販薬 効果 比較",   "医薬品"),
    # ── リカバリー・スポーツ（高単価） ──────────────────────────────────────────
    ("Therabody Japan合同会社（Theragun）",                 "https://www.therabody.com/ja-jp/",              "リスティング広告", "", "パーカッションマッサージガン・リカバリー",            "セラガン おすすめ 効果, マッサージガン 高性能 比較, Theragun 評判 購入",     "スポーツリカバリー"),
    ("Hyperice Japan（ハイパーアイス）",                    "https://hyperice.com/ja-jp/",                   "リスティング広告", "", "バイブレーションリカバリー・アイシング器具",          "ハイパーアイス おすすめ 評判, リカバリー 器具 筋肉 疲労, Hyperice 購入 比較","スポーツリカバリー"),
    # ── レーシング・スポーツウェア（高単価） ──────────────────────────────────────────
    ("コンプレスポーツ ジャパン（Compressport）",           "https://www.compressport.com/jp/",              "リスティング広告", "", "コンプレッションウェア・トレイルラン・マラソン",     "コンプレスポーツ タイツ 評判, トレイルラン ウェア 高機能, マラソン 着圧 比較","スポーツウェア"),
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
