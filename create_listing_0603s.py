import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03s.xlsx"
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
    # ── 損害保険（高単価） ──────────────────────────────────────────
    ("損保ジャパン株式会社",                                "https://www.sompo-japan.co.jp/",                "リスティング広告", "", "自動車保険・火災保険・傷害保険",                     "損保ジャパン 自動車保険 評判 比較, 火災保険 おすすめ 安い, 保険 見積もり 損保","損害保険"),
    ("東京海上日動火災保険株式会社",                        "https://www.tokiomarine-nichido.co.jp/",        "リスティング広告", "", "自動車保険・火災保険・各種損害保険",                  "東京海上 自動車保険 評判 見積もり, 火災保険 比較 おすすめ, 損害保険 選び方","損害保険"),
    ("三井住友海上火災保険株式会社",                        "https://www.ms-ins.com/",                       "リスティング広告", "", "自動車保険・海外旅行保険・火災保険",                  "三井住友海上 自動車保険 評判, 海外旅行保険 比較 おすすめ, 損害保険 見積もり","損害保険"),
    # ── 高級アウターウェア（超高単価） ──────────────────────────────────────────
    ("カナダグース ジャパン合同会社（Canada Goose）",       "https://www.canadagoose.com/jp/ja/",            "リスティング広告", "", "高級ダウンジャケット・パーカー・ベスト",              "カナダグース ダウン 評判 おすすめ, 高級 ダウンジャケット ブランド, Canada Goose 購入","高級アウター"),
    ("ウールリッチ ジャパン株式会社（Woolrich）",           "https://www.woolrich.com/ja-jp/",               "リスティング広告", "", "アメリカ製ダウン・アウトドアウェア・ブランケット",    "ウールリッチ ダウン 評判 おすすめ, アウトドア ウェア アメリカ ブランド, Woolrich 購入","高級アウター"),
    ("ストーンアイランド ジャパン合同会社（Stone Island）", "https://www.stoneisland.com/ja-jp/",            "リスティング広告", "", "テクニカルファブリック・ジャケット・スウェット",      "ストーンアイランド ジャケット 評判, 高級 カジュアル ブランド, Stone Island 購入","高級カジュアル"),
    # ── 高級スキンケア（高単価） ──────────────────────────────────────────
    ("キールズ ジャパン株式会社（Kiehl's）",                "https://www.kiehls.jp/",                        "リスティング広告", "", "NYスキンケア・保湿クリーム・エッセンス",              "キールズ クリーム 評判 おすすめ, 男性 スキンケア 高品質, Kiehl's 化粧水 購入","高級スキンケア"),
    # ── 住宅設備（高単価） ──────────────────────────────────────────
    ("株式会社TOTO",                                        "https://www.toto.co.jp/",                       "リスティング広告", "", "ウォシュレット・システムバス・節水トイレ",            "TOTO ウォシュレット おすすめ 比較, トイレ リフォーム 費用, システムバス 価格","住宅設備"),
    ("YKK AP株式会社",                                      "https://www.ykkap.co.jp/",                      "リスティング広告", "", "高断熱窓・玄関ドア・リノベーション窓",                "YKK AP 窓 リフォーム 費用, 断熱 窓 おすすめ 交換, 玄関 ドア 交換 価格",    "住宅設備"),
    ("クリナップ株式会社",                                  "https://cleanup.jp/",                           "リスティング広告", "", "システムキッチン・洗面化粧台・お風呂",                 "クリナップ キッチン 評判 おすすめ, システムキッチン リフォーム 費用, 洗面台 比較","住宅設備"),
    # ── 給湯器（高単価） ──────────────────────────────────────────
    ("株式会社ノーリツ",                                    "https://www.noritz.co.jp/",                     "リスティング広告", "", "給湯器・エコジョーズ・ガス機器・床暖房",              "ノーリツ 給湯器 交換 費用, エコジョーズ おすすめ 比較, 給湯器 修理 業者 価格","給湯器"),
    ("リンナイ株式会社",                                    "https://www.rinnai.co.jp/",                     "リスティング広告", "", "給湯器・ガスコンロ・ビルトインコンロ・乾太くん",      "リンナイ 給湯器 評判 交換, ガスコンロ おすすめ 比較, 乾太くん 設置 費用",   "給湯器"),
    # ── アウトドアクッカー（高単価） ──────────────────────────────────────────
    ("カスケードデザインズ ジャパン合同会社（MSR）",        "https://www.cascadedesigns.com/msr/",           "リスティング広告", "", "山岳テント・ガスバーナー・浄水器・クッカー",          "MSR テント おすすめ 評判, 山岳 テント 軽量 比較, MSR クッカー バーナー 購入","アウトドア用品"),
    ("プリムス ジャパン合同会社（Primus）",                 "https://www.primus.jp/",                        "リスティング広告", "", "ガスバーナー・ランタン・アウトドアクッカー",          "プリムス バーナー おすすめ 評判, アウトドア ガスバーナー 比較, キャンプ 火器 購入","アウトドア用品"),
    # ── 楽器（高単価） ──────────────────────────────────────────
    ("株式会社山野楽器",                                    "https://www.yamano.co.jp/",                     "リスティング広告", "", "楽器・音楽教室・楽譜・音響機器",                       "山野楽器 楽器 購入 おすすめ, 音楽教室 体験 申し込み, ギター ピアノ 購入 比較","楽器"),
    ("パール楽器製造株式会社（Pearl）",                     "https://www.pearldrum.co.jp/",                  "リスティング広告", "", "ドラムセット・電子ドラム・打楽器",                     "パール ドラム おすすめ 評判, 電子ドラム セット 購入, ドラム 練習 初心者 比較","楽器"),
    # ── ランニングシューズ（高単価） ──────────────────────────────────────────
    ("HOKA Japan株式会社（HOKA）",                          "https://www.hoka.com/ja-jp/",                   "リスティング広告", "", "厚底ランニングシューズ・トレイルランニング",           "HOKA ランニングシューズ 評判 おすすめ, 厚底 シューズ マラソン 比較, ホカ 購入","ランニングシューズ"),
    ("On Japan合同会社（On Running）",                      "https://www.on-running.com/ja-jp/",             "リスティング広告", "", "スイス製ランニングシューズ・Cloudクッション",          "On ランニングシューズ 評判 おすすめ, Cloud シューズ 比較, スイス 高機能 購入","ランニングシューズ"),
    # ── ダイビング用品（高単価） ──────────────────────────────────────────
    ("TUSA株式会社",                                        "https://www.tusa.co.jp/",                       "リスティング広告", "", "スキューバダイビング機器・マスク・フィン",             "TUSA ダイビング 器材 おすすめ, スキューバ マスク 選び方, ダイビング 用品 購入","ダイビング用品"),
    ("アクアラング・ジャパン株式会社（Aqualung）",          "https://www.aqualung.co.jp/",                   "リスティング広告", "", "スキューバダイビング器材・BCD・レギュレーター",       "アクアラング 器材 評判, ダイビング BCD おすすめ 比較, 海外 ダイビング 器材","ダイビング用品"),
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
