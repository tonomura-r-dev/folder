import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03i.xlsx"
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
    # ── 新電力・ガス乗り換え（高単価顧客獲得） ──────────────────────────────────────────
    ("ENEOSホールディングス株式会社（ENEOSでんき）",  "https://enekey.com/",                            "リスティング広告", "", "電気・ガスのセット切り替え",             "電気 乗り換え おすすめ, ENEOS でんき ガス 料金, 電力 切り替え 手続き",       "新電力"),
    ("KDDI株式会社（au電気）",                        "https://www.au.com/electricity/",                "リスティング広告", "", "auスマホとセットの電気・ガス",            "au でんき セット 割引, 電気 スマホ セット 安くなる, au 電力 乗り換え",       "新電力"),
    ("ソフトバンク株式会社（ソフトバンクでんき）",    "https://www.softbank.jp/utility/",               "リスティング広告", "", "ソフトバンクユーザー向け電気",            "ソフトバンク でんき 料金 安い, スマホ 電気 セット 割引, 電気 切り替え 方法", "新電力"),
    # ── 解体工事（高単価） ──────────────────────────────────────────
    ("解体の窓口（解体工事比較サービス）",             "https://kaitai-no-madoguchi.com/",               "リスティング広告", "", "解体工事・建物取り壊し費用比較",         "解体工事 費用 相場, 建物 解体 業者 比較, 家 解体 見積もり 無料",            "解体工事"),
    ("株式会社解体堂",                                "https://kaitai-do.com/",                         "リスティング広告", "", "解体工事・全国対応・廃材処理",           "解体工事 業者 おすすめ, 空き家 解体 費用, 解体 工事 流れ 費用",             "解体工事"),
    # ── 暗号資産（高単価） ──────────────────────────────────────────
    ("株式会社SBI VCトレード",                        "https://www.sbivc.co.jp/",                       "リスティング広告", "", "暗号資産・仮想通貨取引所",               "暗号資産 取引所 おすすめ, ビットコイン 買い方 初心者, 仮想通貨 口座 開設", "暗号資産"),
    ("HTX Japan株式会社（HTX）",                     "https://www.htx.co.jp/",                         "リスティング広告", "", "暗号資産・仮想通貨取引・ETH・BTC",      "仮想通貨 取引所 比較, ETH 購入 方法, 暗号資産 安全 取引所",                 "暗号資産"),
    # ── ふるさと納税（高単価・大量購入顧客） ──────────────────────────────────────────
    ("ANAホールディングス株式会社（ANAのふるさと納税）","https://furusato.ana.co.jp/",                  "リスティング広告", "", "ANAマイル貯まるふるさと納税",            "ふるさと納税 ANA マイル, 旅行 ふるさと納税 おすすめ, 返礼品 比較 ANA",      "ふるさと納税"),
    ("楽天グループ株式会社（楽天ふるさと納税）",      "https://event.rakuten.co.jp/furusato/",          "リスティング広告", "", "楽天ポイント還元ふるさと納税",            "ふるさと納税 楽天 ポイント, 楽天市場 ふるさと納税 やり方, 返礼品 人気 楽天", "ふるさと納税"),
    # ── 土地活用・アパート経営（超高単価） ──────────────────────────────────────────
    ("旭化成ホームズ株式会社（ヘーベルメゾン）",      "https://www.asahi-kasei-homes.co.jp/hebel-maison/","リスティング広告", "", "土地活用・賃貸アパート・ヘーベルメゾン", "土地活用 アパート経営 おすすめ, 相続 土地 アパート 建てる, ヘーベルメゾン 評判","土地活用"),
    ("株式会社明豊エンタープライズ",                 "https://www.meiho-ent.co.jp/",                   "リスティング広告", "", "土地活用・賃貸経営・マンション",         "土地活用 相談 無料, 更地 有効活用, マンション 建てる 土地 活用 費用",       "土地活用"),
    # ── オンライン診療・健康 ──────────────────────────────────────────
    ("LINEヘルスケア株式会社",                        "https://healthcare.line.me/",                    "リスティング広告", "", "オンライン医師相談・健康サービス",       "医師 相談 オンライン 無料, 症状 相談 アプリ, 病院 いかずに 医師 質問",       "オンライン医療"),
    # ── プレミアムペットフード（高単価） ──────────────────────────────────────────
    ("ロイヤルカナン ジャポン合同会社",               "https://www.royalcanin.com/jp/",                 "リスティング広告", "", "プレミアム獣医推奨ペットフード",         "ロイヤルカナン 犬 療法食, ペットフード 高品質 おすすめ, 猫 療法食 通販",    "ペットフード"),
    ("日本ヒルズ・コルゲート株式会社（サイエンスダイエット）","https://www.hills.co.jp/",              "リスティング広告", "", "獣医推奨プレミアムドッグ・キャットフード", "ヒルズ サイエンスダイエット 犬, 獣医師推奨 フード 通販, 犬 食事 療法",      "ペットフード"),
    # ── 住宅リフォーム・リノベ（残り高単価） ──────────────────────────────────────────
    ("株式会社スペースアール（LOHAS studio）",         "https://lohas-studio.jp/",                       "リスティング広告", "", "自然素材・無垢材リノベーション",         "自然素材 リノベーション, 無垢材 リフォーム 費用, LOHAS studio 評判",         "リノベーション"),
    ("パナソニックリフォーム株式会社",                "https://reform.panasonic.co.jp/",                "リスティング広告", "", "水廻り・キッチン・浴室リフォーム",       "キッチン リフォーム 費用, 浴室 リフォーム 工期, パナソニック リフォーム 見積", "リフォーム"),
    ("長谷工コミュニティ株式会社",                   "https://www.haseko-com.co.jp/",                  "リスティング広告", "", "マンション大規模修繕・管理・住み替え",   "マンション 大規模修繕 費用, 管理組合 修繕 相談, マンション 住み替え",        "マンション管理"),
    # ── ガス供給・切り替え（高単価） ──────────────────────────────────────────
    ("株式会社ニチガス",                              "https://www.nichigas.co.jp/",                    "リスティング広告", "", "都市ガス・LPガスの切り替え・供給",       "プロパンガス 料金 安くする, ガス 乗り換え 比較, LP ガス 業者 変える 方法",  "ガス供給"),
    ("東京ガスes株式会社",                            "https://www.tges.co.jp/",                        "リスティング広告", "", "電気・ガスのセット提供（東京ガス系）",   "東京ガス 電気 セット 料金, ガス 電気 まとめる 割引, 都市ガス 電力 セット",  "新電力・ガス"),
    ("株式会社ガスパル（プロパンガス比較）",           "https://www.gaspal.co.jp/",                      "リスティング広告", "", "プロパンガス料金比較・乗り換えサービス", "プロパンガス 高い 安くしたい, LPガス 料金 比較, プロパン ガス 乗り換え",    "ガス比較"),
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
