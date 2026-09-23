import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03j.xlsx"
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
    # ── カー用品・コーティング（高単価） ──────────────────────────────────────────
    ("株式会社オートバックスセブン",                      "https://www.autobacs.com/",                      "リスティング広告", "", "タイヤ交換・車検・オイル交換・カー用品",  "タイヤ交換 費用 安い 近く, 車検 予約 格安 早い, カー用品 店舗 おすすめ",    "カー用品"),
    ("キーパー技研株式会社（KeePer LABO）",              "https://keeperlabo.jp/",                         "リスティング広告", "", "カーコーティング・ガラスコーティング・洗車","カーコーティング 費用 種類 比較, 車 ガラスコーティング 専門店, KeePer 評判 効果","カーコーティング"),
    # ── 不動産買取・中古住宅（高単価） ──────────────────────────────────────────
    ("スター・マイカ株式会社",                            "https://www.star-mica.co.jp/",                   "リスティング広告", "", "中古マンション買取・買取再販・売却",      "マンション 買取 業者 おすすめ, 中古マンション 売却 費用, 不動産 買取 相場",   "不動産買取"),
    ("株式会社リプライス",                               "https://www.reprice.co.jp/",                     "リスティング広告", "", "中古住宅再生・リノベ済み物件販売",        "中古住宅 リノベーション 購入, リフォーム済み 物件 お得, 中古一戸建て 安い",   "中古住宅"),
    # ── 投資型クラウドファンディング（高単価） ──────────────────────────────────────────
    ("株式会社FUNDINNO",                                 "https://fundinno.com/",                          "リスティング広告", "", "株式投資型クラウドファンディング",        "株式 クラウドファンディング 投資, 未上場株 投資 少額, FUNDINNO 評判 口コミ",  "投資型CF"),
    # ── ブライダルジュエリー（超高単価） ──────────────────────────────────────────
    ("株式会社銀座白石",                                 "https://www.ginzashiraishi.co.jp/",              "リスティング広告", "", "婚約指輪・結婚指輪・高品質ダイヤ",        "婚約指輪 ブランド 高品質 人気, 結婚指輪 オーダー 銀座, ダイヤ 指輪 評判",    "ブライダルジュエリー"),
    ("株式会社エクセルコダイヤモンド",                   "https://www.excelcojewelry.com/",                "リスティング広告", "", "婚約指輪・天然ダイヤ・プロポーズリング",  "婚約指輪 天然ダイヤ おすすめ, プロポーズ 指輪 ブランド, 婚約指輪 予算 相場",  "ブライダルジュエリー"),
    # ── 英語コーチング（高単価） ──────────────────────────────────────────
    ("株式会社PROGRIT",                                  "https://www.progrit.co.jp/",                     "リスティング広告", "", "短期集中英語コーチング・TOEIC対策",       "英語 短期 コーチング 費用, TOEIC 点数 上がる スクール, 英語 毎日 習慣 コーチ", "英語コーチング"),
    ("株式会社STRAIL",                                   "https://strail.jp/",                             "リスティング広告", "", "ビジネス英語・英語コーチング・外資転職対策","ビジネス英語 コーチング 費用, 英語 スキルアップ 社会人, 英語 コーチング 比較","英語コーチング"),
    # ── フィットネス（高単価） ──────────────────────────────────────────
    ("B-MONSTER株式会社",                                "https://b-monster.me/",                          "リスティング広告", "", "ボクシングフィットネス・バーチャルパーソナル","ボクシング フィットネス 東京 体験, ダイエット ボクシング 効果, B-MONSTER 料金","フィットネス"),
    # ── 太陽光発電・蓄電池（高単価） ──────────────────────────────────────────
    ("長州産業株式会社",                                 "https://www.chosuisangyo.co.jp/",                "リスティング広告", "", "太陽光発電・蓄電池・HEMSシステム",        "太陽光発電 設置 費用 補助金, 蓄電池 太陽光 セット 価格, 長州産業 太陽光 評判","太陽光発電"),
    ("株式会社ハチドリソーラー",                         "https://hachidorisolar.com/",                    "リスティング広告", "", "初期費用0円太陽光発電・PPAモデル",        "太陽光 0円 設置 仕組み, PPA 太陽光 メリット, 初期費用なし 太陽光 業者",      "太陽光発電"),
    # ── 高級家具・インテリア（高単価） ──────────────────────────────────────────
    ("大塚家具株式会社",                                 "https://www.idc-otsuka.jp/",                     "リスティング広告", "", "高級家具・ソファ・ベッド・インテリア",    "高級 家具 ブランド 通販, ソファ 高品質 おすすめ, 大塚家具 セール 購入",       "高級家具"),
    # ── 料理教室（高単価） ──────────────────────────────────────────
    ("株式会社ABC Cooking Studio",                       "https://www.abc-cooking.co.jp/",                 "リスティング広告", "", "料理・パン・スイーツ・和食教室",          "料理教室 初心者 おすすめ, パン 教室 体験 近く, スイーツ 作り方 スクール 費用","料理教室"),
    # ── 幼児教育（高単価） ──────────────────────────────────────────
    ("株式会社ベビーパーク",                             "https://www.babypark.jp/",                       "リスティング広告", "", "0〜3歳親子教室・早期知育・英語",          "0歳 親子 教室 おすすめ, 赤ちゃん 知育 スクール, ベビーパーク 料金 口コミ",   "幼児教育"),
    # ── ゴルフ用品（高単価） ──────────────────────────────────────────
    ("キャロウェイゴルフ株式会社",                       "https://www.callawaygolf.com/ja-JP/",            "リスティング広告", "", "ゴルフクラブ・ドライバー・フィッティング", "ゴルフ クラブ 最新 おすすめ, ドライバー 飛距離 アップ クラブ, キャロウェイ フィッティング","ゴルフ用品"),
    # ── 健康器具EMS（高単価） ──────────────────────────────────────────
    ("株式会社MTG（SIXPAD）",                            "https://www.sixpad.jp/",                         "リスティング広告", "", "電気刺激EMS・腹筋・ボディメイク器具",     "SIXPAD 効果 口コミ, EMS 腹筋 器具 おすすめ, 電気 腹筋 ながら トレーニング",  "健康器具"),
    # ── 学習塾・予備校（高単価） ──────────────────────────────────────────
    ("株式会社城南進学研究社",                           "https://www.johnan.jp/",                         "リスティング広告", "", "高校受験・大学受験予備校・個別指導塾",    "予備校 費用 比較 おすすめ, 高校受験 塾 合格率, 城南予備校 評判 コース 費用",  "学習塾"),
    # ── 葬儀（高単価） ──────────────────────────────────────────
    ("株式会社平安典礼",                                 "https://www.heian-tsuya.co.jp/",                 "リスティング広告", "", "葬儀・告別式・家族葬・1日葬",             "家族葬 費用 相場 地域, 葬儀 安い 業者 近く, 平安典礼 評判 葬式 プラン",      "葬儀"),
    # ── リノベーション比較（高単価） ──────────────────────────────────────────
    ("株式会社SUVACO",                                   "https://suvaco.jp/",                             "リスティング広告", "", "建築家・リノベ会社のマッチングサービス",  "リノベーション 業者 比較 おすすめ, 建築家 依頼 費用, 家 建てる 設計士 探し方","リノベ比較"),
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
