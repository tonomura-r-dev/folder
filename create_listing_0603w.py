import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03w.xlsx"
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
    # ── 高級デザイン家具（超高単価） ──────────────────────────────────────────
    ("カッシーナ・イクスシー株式会社（Cassina ixc.）",     "https://www.cassina-ixc.jp/",                   "リスティング広告", "", "イタリア製高級デザイン家具・ソファ・テーブル",         "カッシーナ 家具 おすすめ 評判, イタリア 高級 家具 比較, Cassina ソファ 購入","高級家具"),
    ("フリッツハンセン ジャパン合同会社（Fritz Hansen）",  "https://www.fritzhansen.com/ja/",               "リスティング広告", "", "北欧デザイン家具・セブンチェア・エッグチェア",          "フリッツハンセン チェア 評判, 北欧 デザイン 家具 比較, セブンチェア 購入",  "高級家具"),
    ("ノール ジャパン株式会社（Knoll）",                   "https://www.knoll.co.jp/",                      "リスティング広告", "", "デザインオフィス家具・バルセロナチェア・ワークスペース","ノール オフィス 家具 おすすめ, バルセロナチェア 購入 価格, Knoll 評判",     "高級家具"),
    # ── 高級ホテル（超高単価） ──────────────────────────────────────────
    ("コンラッドホテル＆リゾーツ（Conrad Hotels）",        "https://www.hilton.com/ja/brands/conrad-hotels/","リスティング広告", "", "コンラッドホテル・スイート宿泊・ラグジュアリー",        "コンラッド ホテル 東京 予約, 高級 ホテル スイート 宿泊, Conrad 評判 特典","高級ホテル"),
    ("ハイアット ジャパン（Hyatt）",                       "https://www.hyatt.com/ja-JP/",                  "リスティング広告", "", "ハイアットリゾート・ホテル・ポイントプログラム",         "ハイアット ホテル 東京 予約, Hyatt 高級 リゾート 評判, ワールドオブハイアット","高級ホテル"),
    # ── バッグ・ファッション（高単価） ──────────────────────────────────────────
    ("ロングシャン ジャパン合同会社（Longchamp）",          "https://www.longchamp.com/ja-jp/",              "リスティング広告", "", "フランス製バッグ・ル・プリアージュ・レザーバッグ",       "ロングシャン バッグ 評判 人気, ル プリアージュ 購入 価格, Longchamp 新作","バッグ"),
    ("アニエスベー ジャパン株式会社（agnès b.）",          "https://www.agnesb.co.jp/",                     "リスティング広告", "", "フランスファッション・ストライプ・バッグ・雑貨",         "アニエスベー 新作 コーデ, フランス ブランド ファッション, agnès b. 購入","ファッション"),
    ("デシグアル ジャパン合同会社（Desigual）",             "https://www.desigual.com/ja_JP/",               "リスティング広告", "", "カラフルファッション・バッグ・スペインブランド",         "デシグアル 新作 人気, カラフル ブランド ファッション, Desigual 購入 セール","ファッション"),
    ("オンワード樫山株式会社（Onward）",                    "https://www.onward.co.jp/",                     "リスティング広告", "", "セレクトファッション・スーツ・カジュアルウェア",         "オンワード ファッション おすすめ, セレクト ショップ ブランド, Onward 購入","ファッション"),
    ("株式会社ジーナシス（JEANASIS）",                      "https://www.jeanasis.jp/",                      "リスティング広告", "", "トレンドファッション・デニム・ウィメンズカジュアル",     "ジーナシス 新作 人気 コーデ, JEANASIS セール, トレンド ファッション 通販","ファッション"),
    # ── 高級ペットフード（高単価） ──────────────────────────────────────────
    ("ヒルズペットニュートリション株式会社（Hill's Pet）",  "https://www.hillspet.co.jp/",                   "リスティング広告", "", "プレミアムペットフード・サイエンスダイエット",           "ヒルズ ペットフード おすすめ, 高品質 ドッグフード 評判, Hill's 療法食 購入","ペットフード"),
    ("ネスレ ピュリナ ペットケア ジャパン（Purina）",       "https://www.purina.jp/",                        "リスティング広告", "", "プロプランペットフード・プリナワン",                      "ピュリナ ペットフード おすすめ 評判, プロプラン 猫 犬 フード, Purina 購入","ペットフード"),
    # ── 高単価食品EC（高単価） ──────────────────────────────────────────
    ("グリーンスプーン株式会社（GreenSpoon）",              "https://greens-spoon.com/",                     "リスティング広告", "", "パーソナライズ野菜スムージー・サラダ定期宅配",           "グリーンスプーン 評判 効果, 野菜 スムージー 宅配 おすすめ, GreenSpoon 価格","食品宅配"),
    ("株式会社明治屋",                                      "https://www.meidi-ya.co.jp/",                   "リスティング広告", "", "高級輸入食品・ギフト・ワイン・チーズ通販",               "明治屋 高級 食品 通販, 輸入 食材 ギフト おすすめ, ワイン チーズ 購入 老舗","高級食品EC"),
    ("株式会社成城石井（成城石井）",                        "https://www.seijoishii.co.jp/",                 "リスティング広告", "", "プレミアム食品・ワイン・チーズ・スイーツ通販",           "成城石井 通販 おすすめ, 高級 ワイン 食材 購入, お取り寄せ スイーツ 評判","高級食品EC"),
    # ── メガネ（高単価） ──────────────────────────────────────────
    ("株式会社インターメスティック（Zoff）",                "https://www.zoff.co.jp/",                       "リスティング広告", "", "ファッションメガネ・ブルーライトカット・PC眼鏡",         "Zoff メガネ おすすめ 評判, ブルーライト カット 眼鏡 比較, ゾフ 購入 セール","メガネ"),
    ("白山眼鏡店株式会社",                                  "https://www.hakusan-megane.co.jp/",             "リスティング広告", "", "高品質セルロイド眼鏡・国産フレーム・オーダー",           "白山眼鏡 フレーム おすすめ, 高品質 国産 メガネ 比較, 白山眼鏡店 評判 購入","メガネ"),
    # ── 高品質国産家具（高単価） ──────────────────────────────────────────
    ("カリモク60株式会社（karimoku 60）",                   "https://www.karimoku60.jp/",                    "リスティング広告", "", "ロングセラー国産家具・Kチェア・ソファ",                  "カリモク60 チェア 評判 おすすめ, 国産 家具 高品質, Kチェア 購入 価格",    "高品質家具"),
    # ── 注文住宅（超高単価） ──────────────────────────────────────────
    ("大和ハウス工業株式会社（ダイワハウス）",              "https://www.daiwahouse.co.jp/",                  "リスティング広告", "", "注文住宅・分譲住宅・賃貸マンション",                     "大和ハウス 注文住宅 評判, ダイワハウス 家 費用 比較, 分譲住宅 おすすめ",  "注文住宅"),
    ("ミサワホーム株式会社",                               "https://www.misawa.co.jp/",                     "リスティング広告", "", "木質パネル住宅・注文住宅・蔵のある家",                   "ミサワホーム 評判 費用, 注文住宅 比較 おすすめ, 木質 パネル 家 省エネ",  "注文住宅"),
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
