import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03ab.xlsx"
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
    # ── 輸入車（高単価） ──────────────────────────────────────────
    ("シトロエン・ジャポン合同会社（Citroën）",             "https://www.citroen.jp/",                       "リスティング広告", "", "フランス製輸入車・C3・C5X・SUV",                        "シトロエン 新車 評判 おすすめ, フランス 輸入車 比較, Citroën 試乗 購入","輸入車"),
    ("アルファ ロメオ ジャパン合同会社（Alfa Romeo）",      "https://www.alfaromeo.jp/",                     "リスティング広告", "", "イタリア高級スポーツカー・ジュリア・ステルヴィオ",       "アルファロメオ 新車 評判 おすすめ, イタリア 高級車 比較, Giulia 試乗","高級輸入車"),
    ("アストン マーティン ジャパン（Aston Martin）",        "https://www.astonmartin.com/jp",                "リスティング広告", "", "英国超高級スポーツカー・DB12・ヴァンテージ",             "アストンマーチン 新車 価格 評判, 超高級 スポーツカー 英国, DB12 試乗","超高級車"),
    # ── 家電量販（高単価） ──────────────────────────────────────────
    ("ケーズホールディングス株式会社（ケーズデンキ）",       "https://www.ksdenki.com/",                      "リスティング広告", "", "家電・エアコン・冷蔵庫・洗濯機・工事込み",               "ケーズデンキ エアコン 安い 工事 費用, 家電 購入 おすすめ, ケーズ セール","家電量販"),
    ("株式会社ノジマ（Nojima）",                            "https://www.nojima.co.jp/",                     "リスティング広告", "", "家電・スマホ・PC・格安SIM・アウトレット",                "ノジマ 家電 安い おすすめ, スマホ 格安 SIM 比較, ノジマ セール 購入","家電量販"),
    ("株式会社コジマ（Kojima）",                            "https://www.kojima.net/",                       "リスティング広告", "", "家電・エアコン・TV・工事込み設置",                       "コジマ エアコン 安い 設置 工事, 家電 おすすめ 比較, コジマ セール",   "家電量販"),
    # ── コスメEC（高単価） ──────────────────────────────────────────
    ("株式会社NOIN",                                        "https://noin.io/",                              "リスティング広告", "", "コスメ・美容品・サンプル・通販EC",                       "NOIN コスメ 通販 評判, 化粧品 サンプル 購入, おすすめ コスメ 安い 比較","コスメEC"),
    # ── 金地金・貴金属（高単価） ──────────────────────────────────────────
    ("田中貴金属工業株式会社",                               "https://gold.tanaka.co.jp/",                    "リスティング広告", "", "金地金・プラチナ・貴金属投資・購入売却",                  "金 購入 価格 おすすめ, 金地金 投資 始め方, 田中貴金属 評判 プラチナ","貴金属投資"),
    # ── 市販薬（高単価） ──────────────────────────────────────────
    ("浅田飴株式会社",                                      "https://www.asadaame.co.jp/",                   "リスティング広告", "", "のど飴・風邪薬・市販薬・ドリンク",                       "浅田飴 のど 評判 効果, 市販薬 のど痛 おすすめ, 風邪 薬 購入 種類",  "市販薬"),
    # ── レンタカー（高単価） ──────────────────────────────────────────
    ("ニコニコレンタカー株式会社",                           "https://2-rent.com/",                           "リスティング広告", "", "低価格レンタカー・格安車・旅行・引越し",                  "ニコニコレンタカー 評判 料金, レンタカー 格安 比較, 車 借りる 安い",  "レンタカー"),
    # ── 民泊・宿泊（高単価） ──────────────────────────────────────────
    ("Airbnb Japan合同会社（Airbnb）",                      "https://www.airbnb.jp/",                        "リスティング広告", "", "民泊・一棟貸し・ホームステイ・ユニーク宿泊",              "エアビー 宿泊 おすすめ, 民泊 旅行 一棟貸し, Airbnb 評判 予約 比較","民泊"),
    # ── ピラティス（高単価） ──────────────────────────────────────────
    ("BASIピラティス ジャパン（BASI Pilates）",              "https://basipilates.jp/",                       "リスティング広告", "", "ピラティス・インストラクター養成・スタジオ",              "BASI ピラティス 評判 料金, ピラティス スタジオ 比較, 体幹 姿勢 改善","ピラティス"),
    # ── 語学学校（高単価） ──────────────────────────────────────────
    ("公益財団法人アテネフランセ",                           "https://www.athenee.net/",                      "リスティング広告", "", "フランス語・イタリア語・スペイン語学校",                  "アテネフランセ フランス語 評判 料金, 語学 スクール 比較, フランス語 習う","語学学校"),
    # ── 高級腕時計（超高単価） ──────────────────────────────────────────
    ("ロンジン ジャパン（Longines）",                       "https://www.longines.com/ja-jp/",               "リスティング広告", "", "スイス製高級腕時計・クラシック・エレガント",              "ロンジン 時計 評判 おすすめ, スイス 高級 腕時計 比較, Longines 購入","高級腕時計"),
    ("チューダー ジャパン（Tudor）",                        "https://www.tudorwatch.com/ja-jp/",             "リスティング広告", "", "スイス製プレミアム腕時計・ブラックベイ・デフォ",          "チューダー 時計 評判 おすすめ, Tudor ブラックベイ 購入, プレミアム 時計","プレミアム腕時計"),
    ("スウォッチ ジャパン株式会社（Swatch）",               "https://www.swatch.com/ja-jp/",                 "リスティング広告", "", "ファッション腕時計・ビッグボールド・限定モデル",          "スウォッチ 時計 評判 人気, ファッション 腕時計 比較, Swatch 新作 購入","ファッション時計"),
    # ── BTOパソコン（高単価） ──────────────────────────────────────────
    ("株式会社ユニットコム（パソコン工房）",                  "https://www.pc-koubou.jp/",                     "リスティング広告", "", "BTOパソコン・ゲーミングPC・ノートPC",                    "パソコン工房 BTO PC 評判, ゲーミングPC 安い 高性能 比較, BTOパソコン 購入","BTO PC"),
    ("ドスパラ ジャパン合同会社（Dospara）",                 "https://www.dospara.co.jp/",                    "リスティング広告", "", "BTOゲーミングPC・自作PC・PC周辺機器",                   "ドスパラ ゲーミングPC 評判 おすすめ, BTO PC 比較 高性能, 自作 PC 購入","BTO PC"),
    # ── 高品質ブーツ（高単価） ──────────────────────────────────────────
    ("ダナー ジャパン（Danner）",                            "https://www.danner.jp/",                        "リスティング広告", "", "ゴアテックスブーツ・トレッキングブーツ・米国製",          "ダナー ブーツ 評判 おすすめ, アウトドア 防水 ブーツ 比較, Danner 購入","ブーツ"),
    ("レッドウィング ジャパン株式会社（Red Wing）",          "https://www.redwingshoe.co.jp/",                "リスティング広告", "", "高品質ワークブーツ・アイリッシュセッター・アメリカ製",    "レッドウィング ブーツ 評判 おすすめ, ワークブーツ 高品質 比較, Red Wing 購入","ブーツ"),
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
