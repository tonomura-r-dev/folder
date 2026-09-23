import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03x.xlsx"
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
    # ── 注文住宅（超高単価） ──────────────────────────────────────────
    ("パナソニックホームズ株式会社",                        "https://www.panahome.jp/",                      "リスティング広告", "", "注文住宅・スマートホーム・リフォーム",               "パナソニックホームズ 評判 費用, 注文住宅 ハウスメーカー 比較, スマートホーム 家","注文住宅"),
    ("住友林業株式会社",                                   "https://sfc.jp/",                               "リスティング広告", "", "木造注文住宅・ウッドセレクション・森林認証材",         "住友林業 評判 費用 坪単価, 木造 注文住宅 おすすめ, 住友林業 家 特徴 比較","注文住宅"),
    ("一条工務店株式会社",                                  "https://www.ichijo.co.jp/",                     "リスティング広告", "", "高断熱高気密住宅・全館床暖房・太陽光一体型",           "一条工務店 評判 坪単価, 高気密 高断熱 住宅 比較, 一条 床暖房 光熱費 評価","注文住宅"),
    ("積水化学工業株式会社（セキスイハイム）",               "https://www.sekisuiheim.com/",                  "リスティング広告", "", "ユニット工法住宅・蓄電池・ZEH住宅",                   "セキスイハイム 評判 費用, 注文住宅 ユニット工法, ZEH 蓄電 住宅 おすすめ","注文住宅"),
    ("ウィザースホーム株式会社",                            "https://www.withershome.co.jp/",                "リスティング広告", "", "4×6工法・高性能注文住宅・ツーバイシックス",            "ウィザースホーム 評判 坪単価, 4×6 住宅 高性能, 注文住宅 ハウスメーカー 選び方","注文住宅"),
    ("トヨタホーム株式会社",                                "https://www.toyotahome.co.jp/",                 "リスティング広告", "", "鉄骨ユニット工法・スマートハウス・省エネ住宅",         "トヨタホーム 評判 費用, 鉄骨 注文住宅 比較, スマートハウス 省エネ おすすめ","注文住宅"),
    ("桧家住宅株式会社",                                   "https://www.hinokiya.jp/",                      "リスティング広告", "", "Z空調付き注文住宅・全館空調・ウレタン断熱",            "桧家住宅 評判 坪単価, Z空調 全館空調 住宅, 注文住宅 気密 断熱 比較",      "注文住宅"),
    # ── 高級デザイン家具（超高単価） ──────────────────────────────────────────
    ("ヴィトラ ジャパン株式会社（Vitra）",                  "https://www.vitra.com/ja-jp/",                  "リスティング広告", "", "スイス製デザイン家具・イームズチェア・オフィス",        "ヴィトラ 家具 おすすめ 評判, イームズ チェア 購入, Vitra デザイン 家具 比較","高級家具"),
    ("アルテック ジャパン合同会社（Artek）",                "https://www.artek.fi/ja-JP/",                   "リスティング広告", "", "フィンランドデザイン家具・スツール60・テーブル",        "アルテック 家具 おすすめ 評判, スツール60 購入 価格, Artek 北欧 デザイン","高級家具"),
    # ── コワーキング（高単価） ──────────────────────────────────────────
    ("WeWork Japan合同会社（WeWork）",                      "https://www.wework.com/ja-JP/",                 "リスティング広告", "", "コワーキングスペース・シェアオフィス・会議室",          "WeWork コワーキング 料金 評判, シェアオフィス 東京 比較, 会議室 レンタル 費用","コワーキング"),
    # ── ホットヨガ（高単価） ──────────────────────────────────────────
    ("株式会社ベストプランニング（ホットヨガカルド）",       "https://yoga-kaldo.jp/",                        "リスティング広告", "", "ホットヨガスタジオ・常温ヨガ・月額会員",               "ホットヨガ カルド 料金 評判, ホットヨガ スタジオ 比較 近く, ヨガ 会員 体験","ヨガ"),
    # ── 脱毛（高単価） ──────────────────────────────────────────
    ("株式会社クオフィス（エピレ）",                        "https://epile.jp/",                             "リスティング広告", "", "医療脱毛・全身脱毛・ヒゲ脱毛",                          "エピレ 脱毛 評判 料金, 医療 脱毛 比較 安い, 全身 脱毛 おすすめ クリニック","脱毛"),
    # ── 住宅設備（高単価） ──────────────────────────────────────────
    ("トクラス株式会社",                                   "https://www.toclas.co.jp/",                     "リスティング広告", "", "システムキッチン・バスルーム・洗面化粧台",              "トクラス キッチン 評判 おすすめ, システム キッチン 比較 価格, バスルーム リフォーム","住宅設備"),
    # ── 日用品・オーラルケア（高単価） ──────────────────────────────────────────
    ("ライオン株式会社",                                   "https://www.lion.co.jp/",                       "リスティング広告", "", "歯磨き粉・洗剤・オーラルケア・スキンケア",              "ライオン 歯磨き おすすめ 評判, システマ 歯ブラシ 効果, 洗剤 高品質 比較","日用品"),
    # ── 市販薬・健康食品（高単価） ──────────────────────────────────────────
    ("武田コンシューマーヘルスケア株式会社",                "https://www.takeda-ch.co.jp/",                  "リスティング広告", "", "市販薬・アリナミン・睡眠改善・ビタミン剤",              "アリナミン 効果 評判, 武田 市販薬 おすすめ, 睡眠改善 薬 比較 購入",       "市販薬"),
    ("株式会社オリヒロ",                                   "https://www.orihiro-kenko.jp/",                 "リスティング広告", "", "コンドロイチン・グルコサミン・健康食品",               "オリヒロ サプリ おすすめ 評判, グルコサミン コンドロイチン 効果, 健康食品 購入","健康食品"),
    ("大塚製薬株式会社（ネイチャーメイド）",                "https://www.naturemade.jp/",                    "リスティング広告", "", "マルチビタミン・オメガ3・プレミアムサプリ",             "ネイチャーメイド ビタミン おすすめ, サプリ 高品質 比較, 大塚製薬 栄養 購入","サプリメント"),
    ("エーザイ株式会社",                                   "https://www.eisai.co.jp/",                      "リスティング広告", "", "認知症・睡眠改善・ビフィーナ・市販薬",                  "エーザイ サプリ おすすめ, 認知症 予防 市販薬, ビフィーナ 評判 効果 購入","市販薬・健康食品"),
    ("参天製薬株式会社",                                   "https://www.santen.co.jp/",                     "リスティング広告", "", "目薬・ドライアイケア・OTC医薬品",                       "参天 目薬 おすすめ 評判, ドライアイ 目薬 比較 効果, サンテ 購入 種類",    "目薬・眼科用品"),
    # ── 損害保険（高単価） ──────────────────────────────────────────
    ("あいおいニッセイ同和損害保険株式会社",                "https://www.aioinissaydowa.co.jp/",             "リスティング広告", "", "自動車保険・火災保険・ビジネス保険",                     "あいおいニッセイ 自動車保険 評判, 火災保険 比較 おすすめ, 損害保険 見積もり","損害保険"),
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
