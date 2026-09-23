import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-02b.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

DATA = [
    # ── 不動産投資・資産運用 ──────────────────────────────────────────
    ("株式会社プロパティエージェント",      "https://www.propertyagent.co.jp/lp/",              "リスティング広告", "", "都心区分マンション投資",             "マンション投資 おすすめ 東京, 不動産投資 初心者, 区分マンション 利回り",   "不動産投資"),
    ("株式会社グローバル・リンク・マネジメント", "https://www.global-link-m.com/lp/",           "リスティング広告", "", "東京・都心マンション投資",           "東京 マンション投資 評判, 不動産投資 セミナー 無料, 一棟 区分 比較",     "不動産投資"),
    ("株式会社武蔵コーポレーション",        "https://www.musashi-corporation.com/lp/",          "リスティング広告", "", "アパート・マンション経営",           "アパート経営 始める, 土地活用 アパート 収益, マンション経営 相談",       "不動産投資"),
    ("株式会社トーシンパートナーズ",        "https://www.tohshin.co.jp/lp/",                    "リスティング広告", "", "東京マンション投資（新築）",         "新築マンション投資 東京, 不動産投資 新築 メリット, 資産形成 マンション", "不動産投資"),
    ("株式会社日本財託",                   "https://www.nihonzaidan.co.jp/lp/",                "リスティング広告", "", "中古マンション投資・管理",           "中古マンション投資 安全, 不動産投資 管理 一括, 東京 中古 利回り",       "不動産投資"),
    ("株式会社シノケングループ",            "https://www.shinoken.co.jp/lp/",                   "リスティング広告", "", "アパート経営・土地活用",             "アパート経営 成功, 土地活用 方法 比較, シノケン 評判 不動産投資",       "不動産投資"),
    ("株式会社GA technologies",            "https://www.renosy.com/lp/",                       "リスティング広告", "", "AI不動産投資（Renosy）",            "Renosy 口コミ, AI 不動産投資 初心者, マンション投資 アプリ 管理",       "不動産投資"),
    ("株式会社ウェルスナビ",               "https://www.wealthnavi.com/lp/",                   "リスティング広告", "", "ロボアドバイザー資産運用",           "ロボアドバイザー 比較, 自動 積立投資 おすすめ, ウェルスナビ 手数料",   "資産運用"),
    ("コインチェック株式会社",              "https://coincheck.com/lp/",                        "リスティング広告", "", "暗号資産取引所（コインチェック）",   "仮想通貨 始め方, ビットコイン 購入 安い, コインチェック 登録 手順",     "暗号資産"),
    ("株式会社マネーフォワード",            "https://moneyforward.com/lp/",                     "リスティング広告", "", "家計・資産管理アプリ",               "家計管理 アプリ おすすめ, 資産管理 自動 連携, マネーフォワード 口コミ", "FinTech"),
    # ── 結婚式場・ブライダル（追加） ─────────────────────────────────
    ("株式会社スタージュエリー",            "https://www.starjewelry.co.jp/lp/",                "リスティング広告", "", "結婚・婚約指輪（スタージュエリー）", "結婚指輪 おすすめ ブランド, 婚約指輪 予算 選び方, 指輪 試着 予約",     "ブライダルジュエリー"),
    ("株式会社プリモジャパン",              "https://www.primobridal.co.jp/lp/",                "リスティング広告", "", "マリッジリング・エンゲージリング",   "マリッジリング おしゃれ, 結婚指輪 ペア 人気, ブライダルリング 安い",   "ブライダルジュエリー"),
    ("株式会社ヴァンドームヤマジ",          "https://www.vendome.co.jp/lp/",                    "リスティング広告", "", "ブライダルジュエリー・婚約指輪",     "婚約指輪 高品質 ブランド, ダイヤ 指輪 オーダー, ヴァンドーム 評判",   "ブライダルジュエリー"),
    ("株式会社ラヴィ・ファクトリー",        "https://www.laviefactory.com/lp/",                 "リスティング広告", "", "ウェディングフォト・フォト婚",       "フォト婚 費用, ウェディングフォト スタジオ 屋外, 前撮り 衣装 込み",   "フォトウェディング"),
    ("株式会社ラグナヴェール",              "https://www.lagnavert.com/lp/",                    "リスティング広告", "", "ガーデンリゾート式場",               "ガーデン 結婚式場 名古屋, リゾート 式場 見学 予約, 海が見える 式場",   "ウェディング"),
    ("株式会社アルモニーアンブラッセ",      "https://www.harmonyambrace.com/lp/",               "リスティング広告", "", "独立型チャペル婚（大阪）",           "チャペル 式場 大阪 梅田, 独立型 ウェディング 関西, 式場 フェア 大阪",  "ウェディング"),
    ("株式会社クレッセント",               "https://www.crescent-w.co.jp/lp/",                  "リスティング広告", "", "総合結婚式場（東海エリア）",         "結婚式場 名古屋 東海, ウェディングフェア 無料, 式場 見学 特典 東海",   "ウェディング"),
    ("株式会社ベルサール",                 "https://www.bellesalle.co.jp/lp/",                  "リスティング広告", "", "貸し会場・ウェディングパーティー",   "パーティー会場 東京 貸切, 二次会 会場 おすすめ, 披露宴 会場 都内",    "ウェディング"),
    ("株式会社クレイジーウェディング",      "https://crazy-wedding.jp/lp/",                     "リスティング広告", "", "オリジナルウェディングプロデュース", "オリジナル 結婚式 演出, サプライズ ウェディング, 式場 なし 結婚式",    "ウェディング"),
    ("株式会社鈴乃屋",                     "https://www.suzunoya.com/lp/",                     "リスティング広告", "", "和装婚・白無垢・色打掛レンタル",     "和装婚 白無垢 費用, 色打掛 レンタル おすすめ, 神前式 着物 衣装",       "ウェディング衣装"),
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
