import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03t.xlsx"
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
    # ── 高級腕時計（超高単価） ──────────────────────────────────────────
    ("オメガ SA ジャパン株式会社（OMEGA）",                "https://www.omegawatches.com/ja-jp/",           "リスティング広告", "", "高級腕時計・スピードマスター・シーマスター",          "オメガ 時計 評判 おすすめ, OMEGA スピードマスター 購入, 高級 腕時計 ブランド","高級腕時計"),
    ("タグホイヤー ジャパン合同会社（TAG Heuer）",          "https://www.tagheuer.com/ja-jp/",               "リスティング広告", "", "高級腕時計・カレラ・クロノグラフ",                    "タグホイヤー 時計 評判 おすすめ, カレラ 購入 価格, TAG Heuer スポーツ 時計","高級腕時計"),
    # ── ゴルフ用品（超高単価） ──────────────────────────────────────────
    ("本間ゴルフ株式会社（HONMA）",                        "https://www.honmagolf.co.jp/",                  "リスティング広告", "", "高級ゴルフクラブ・ドライバー・アイアン",               "本間ゴルフ クラブ 評判 おすすめ, HONMA ドライバー 高級 比較, ゴルフ 高端 クラブ","ゴルフ用品"),
    ("コブラゴルフ ジャパン合同会社（Cobra Golf）",         "https://www.cobragolf.co.jp/",                  "リスティング広告", "", "ゴルフクラブ・ドライバー・ユーティリティ",             "コブラゴルフ ドライバー 評判 おすすめ, Cobra ゴルフ クラブ 比較, スピードゾーン 購入","ゴルフ用品"),
    ("ピン ゴルフ株式会社（PING）",                        "https://ping.com/ja-jp/",                       "リスティング広告", "", "ゴルフクラブ・カスタムフィッティング・アイアン",       "PING ゴルフ フィッティング おすすめ, ピン アイアン 評判 比較, ゴルフ カスタム クラブ","ゴルフ用品"),
    # ── 育児用品（高単価） ──────────────────────────────────────────
    ("ベビービョルン ジャパン株式会社（BabyBjörn）",       "https://www.babybjorn.co.jp/",                  "リスティング広告", "", "抱っこ紐・バウンサー・ベビーキャリア",                 "ベビービョルン 抱っこ紐 評判 おすすめ, BabyBjörn バウンサー 価格, 北欧 育児 用品","育児用品"),
    # ── 生命保険（高単価） ──────────────────────────────────────────
    ("日本生命保険相互会社（日本生命）",                    "https://www.nissay.co.jp/",                     "リスティング広告", "", "生命保険・医療保険・年金保険",                         "日本生命 保険 おすすめ 評判, 医療保険 比較 選び方, 生命保険 見直し 相談",    "生命保険"),
    # ── プレミアムオーディオ（高単価） ──────────────────────────────────────────
    ("ハーマン ジャパン株式会社（JBL）",                   "https://jp.jbl.com/",                           "リスティング広告", "", "ポータブルスピーカー・ヘッドフォン・サウンドバー",     "JBL スピーカー おすすめ 評判, ポータブル スピーカー 防水 比較, JBL ヘッドフォン 購入","プレミアムオーディオ"),
    # ── アウトドアシューズ（高単価） ──────────────────────────────────────────
    ("KEEN ジャパン株式会社（KEEN）",                      "https://www.keenfootwear.com/ja-jp/",           "リスティング広告", "", "アウトドアシューズ・サンダル・トレッキングブーツ",     "KEEN サンダル おすすめ 評判, アウトドア シューズ 比較, キーン トレッキング 購入","アウトドアシューズ"),
    # ── キャンプ用品（高単価） ──────────────────────────────────────────
    ("ユニフレーム株式会社",                               "https://www.uniflame.co.jp/",                   "リスティング広告", "", "キャンプ用クッカー・バーナー・焚き火台",               "ユニフレーム クッカー おすすめ 評判, キャンプ 焚き火 台 比較, バーナー アウトドア 購入","キャンプ用品"),
    # ── コーヒー器具（高単価） ──────────────────────────────────────────
    ("メリタ ジャパン合同会社（Melitta）",                 "https://www.melitta.jp/",                       "リスティング広告", "", "コーヒーメーカー・ドリッパー・コーヒー豆",             "メリタ コーヒーメーカー 評判 おすすめ, ドリッパー 高品質 比較, Melitta 購入","コーヒー器具"),
    ("カリタ株式会社（Kalita）",                           "https://www.kalita.co.jp/",                     "リスティング広告", "", "コーヒードリッパー・グラインダー・コーヒー器具",       "カリタ ドリッパー おすすめ 評判, コーヒー 器具 選び方, Kalita グラインダー 購入","コーヒー器具"),
    # ── 高級コスメ（高単価） ──────────────────────────────────────────
    ("ランコム・ジャポン合同会社（Lancôme）",              "https://www.lancome.co.jp/",                    "リスティング広告", "", "フランス高級コスメ・美容液・ファンデーション",         "ランコム 化粧品 おすすめ 評判, 高級 コスメ ブランド 比較, Lancome 美容液 購入","高級コスメ"),
    ("ジル スチュアート ビューティ（JILL STUART Beauty）","https://www.jillstuart-beauty.com/",            "リスティング広告", "", "フェミニンコスメ・リップ・フレグランス",               "ジルスチュアート コスメ 評判 人気, リップ おすすめ フェミニン, JILL STUART 購入","コスメ"),
    ("クラランス株式会社（Clarins）",                      "https://www.clarins.co.jp/",                    "リスティング広告", "", "フランス植物性スキンケア・ボディオイル・美容液",       "クラランス スキンケア 評判 おすすめ, 高級 化粧品 フランス 比較, Clarins 購入","高級スキンケア"),
    ("シュウ ウエムラ ジャパン合同会社（shu uemura）",     "https://www.shuuemura.co.jp/",                  "リスティング広告", "", "アーティスティックコスメ・クレンジング・リップ",       "シュウウエムラ クレンジング 評判, コスメ 人気 評価 おすすめ, shu uemura 購入","高級コスメ"),
    ("ローラ メルシエ ジャパン（Laura Mercier）",          "https://www.lauramercier.co.jp/",               "リスティング広告", "", "NYコスメ・フェイスパウダー・プライマー",              "ローラ メルシエ コスメ 評判, パウダー おすすめ 人気 比較, Laura Mercier 購入","コスメ"),
    # ── ファッション（高単価） ──────────────────────────────────────────
    ("セオリー ジャパン株式会社（Theory）",                "https://www.theory.co.jp/",                     "リスティング広告", "", "プレミアムカジュアル・ビジネス・モードカジュアル",     "セオリー ファッション 評判, Theory 新作 人気 コーデ, ビジネス カジュアル 購入","ファッション"),
    ("マックスマーラ ジャパン合同会社（Max Mara）",         "https://www.maxmara.com/ja-JP/",                "リスティング広告", "", "イタリア高級レディースファッション・コート",           "マックスマーラ コート 評判 人気, 高級 ブランド コート 比較, Max Mara 購入","高級ファッション"),
    ("アンテプリマ株式会社（Anteprima）",                  "https://www.anteprima.co.jp/",                  "リスティング広告", "", "ワイヤーバッグ・レディースバッグ・アクセサリー",       "アンテプリマ バッグ 評判 おすすめ, ワイヤー バッグ ブランド, Anteprima 新作 購入","バッグ"),
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
