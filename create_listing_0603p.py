import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03p.xlsx"
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
    # ── バイク（超高単価） ──────────────────────────────────────────
    ("ハーレーダビッドソン ジャパン合同会社",               "https://www.harley-davidson.com/jp/ja/",         "リスティング広告", "", "大型バイク・クルーザー・アパレル",               "ハーレーダビッドソン 新車 価格, 大型バイク 購入 費用, ハーレー おすすめ モデル","バイク"),
    ("カワサキモータース株式会社",                          "https://www.kawasaki-motors.com/ja_JP/",         "リスティング広告", "", "大型バイク・スポーツバイク・Ninja",               "カワサキ バイク おすすめ 新型, Ninja 評判 試乗, 大型バイク 購入 免許 費用",   "バイク"),
    ("BMW モトラッド・ジャパン（BMW Motorrad）",            "https://www.bmw-motorrad.co.jp/",               "リスティング広告", "", "プレミアムバイク・アドベンチャー・ツアラー",      "BMW バイク 評判 おすすめ, アドベンチャー バイク 比較, BMWモトラッド 購入 試乗","バイク"),
    # ── 高級輸入車（超高単価） ──────────────────────────────────────────
    ("ボルボ・カー・ジャパン株式会社（Volvo Cars）",        "https://www.volvocars.com/ja/",                 "リスティング広告", "", "プレミアムSUV・電動車・安全性能",                 "ボルボ SUV おすすめ 評判, 輸入車 安全 ファミリー, XC60 XC40 試乗 価格",     "輸入車"),
    ("テスラ ジャパン合同会社（Tesla）",                    "https://www.tesla.com/ja_jp/",                  "リスティング広告", "", "電気自動車・EV・自動運転・充電インフラ",           "テスラ モデル3 評判 価格, EV 電気自動車 購入 補助金, Tesla 試乗 申込み",     "EV自動車"),
    # ── タイヤ（高単価） ──────────────────────────────────────────
    ("横浜ゴム株式会社（YOKOHAMA TYRES）",                 "https://www.y-yokohama.com/",                   "リスティング広告", "", "低燃費タイヤ・スタッドレス・スポーツタイヤ",      "ヨコハマ タイヤ おすすめ 比較, スタッドレス 選び方, 低燃費 タイヤ 交換 費用","タイヤ"),
    ("東洋タイヤ株式会社（TOYO TIRES）",                   "https://www.toyotires.jp/",                     "リスティング広告", "", "スタッドレスタイヤ・スポーツタイヤ・SUVタイヤ",   "東洋タイヤ 評判 おすすめ, スタッドレス TOYO 比較, SUV タイヤ 選び方 交換",  "タイヤ"),
    # ── ペット用品EC（高単価） ──────────────────────────────────────────
    ("株式会社チャーム（charm ペット通販）",                "https://www.charm.jp/",                         "リスティング広告", "", "ペット用品・フード・水槽・アクアリウム",           "ペット 用品 通販 安い, 犬 フード おすすめ 通販, アクアリウム 水槽 セット 購入","ペット用品EC"),
    # ── 語学留学（超高単価） ──────────────────────────────────────────
    ("EF Education First Japan株式会社（EF）",             "https://www.ef.co.jp/",                         "リスティング広告", "", "語学留学・短期留学・英語コース",                   "語学 留学 費用 比較, 英語 留学 短期 おすすめ, EF 留学 評判 コース 申し込み", "語学留学"),
    # ── アウトドアウェア（高単価） ──────────────────────────────────────────
    ("株式会社ゴールドウィン（The North Face）",            "https://www.thenorthface.com/ja-jp/",           "リスティング広告", "", "アウトドアウェア・ジャケット・登山用品",           "ノースフェイス ジャケット おすすめ 人気, 登山 ウェア 高機能, The North Face 新作","アウトドアウェア"),
    ("ヘリーハンセン ジャパン株式会社（Helly Hansen）",     "https://www.hellyhansen.com/ja_jp/",            "リスティング広告", "", "マリンウェア・アウトドアウェア・レインウェア",    "ヘリーハンセン ジャケット 評判 おすすめ, マリン アウトドア ウェア 比較, Helly Hansen 購入","アウトドアウェア"),
    ("ジャック・ウルフスキン ジャパン（Jack Wolfskin）",   "https://www.jack-wolfskin.jp/",                 "リスティング広告", "", "ドイツ製アウトドアウェア・登山・ハイキング",      "ジャック・ウルフスキン ジャケット 評判, アウトドア ウェア 防水 比較, Jack Wolfskin 購入","アウトドアウェア"),
    # ── スポーツウェア（高単価） ──────────────────────────────────────────
    ("チャンピオン・ジャパン合同会社（Champion）",          "https://www.champion.co.jp/",                   "リスティング広告", "", "スウェット・スポーツウェア・アクティブウェア",    "チャンピオン スウェット おすすめ 人気, スポーツ ウェア ブランド 比較, Champion 新作","スポーツウェア"),
    ("エレッセ ジャパン株式会社（ellesse）",                "https://www.ellesse.co.jp/",                    "リスティング広告", "", "テニスウェア・スポーツファッション・シューズ",    "エレッセ テニス ウェア おすすめ, ellesse スポーツ ファッション 人気, 新作 購入","スポーツウェア"),
    # ── 高級自転車（超高単価） ──────────────────────────────────────────
    ("スペシャライズド・ジャパン合同会社（Specialized）",  "https://www.specialized.com/jp/",               "リスティング広告", "", "ロードバイク・マウンテンバイク・Eバイク",         "スペシャライズド ロードバイク おすすめ, Specialized 評判 購入, ロードバイク 高性能 比較","自転車"),
    ("トレック・ジャパン合同会社（Trek）",                  "https://www.trekbikes.com/jp/ja_JP/",           "リスティング広告", "", "ロードバイク・クロスバイク・Eバイク",             "トレック ロードバイク 評判 おすすめ, Trek 自転車 購入 価格, クロスバイク 選び方","自転車"),
    # ── ランニングシューズ（高単価） ──────────────────────────────────────────
    ("ブルックス ジャパン合同会社（Brooks）",               "https://www.brooksrunning.com/ja_JP/",          "リスティング広告", "", "プレミアムランニングシューズ・マラソン",           "ブルックス ランニングシューズ 評判 おすすめ, マラソン 靴 高性能 比較, Brooks 購入","ランニングシューズ"),
    # ── フィットネスジム（高単価） ──────────────────────────────────────────
    ("フィットネスワン株式会社",                            "https://www.fitness-one.co.jp/",                "リスティング広告", "", "フィットネスジム・24時間・プール完備",             "フィットネス ジム おすすめ 近く, 24時間 ジム 月額 安い, フィットネスワン 入会 費用","フィットネス"),
    ("JOYFIT株式会社（JOYFIT24）",                          "https://joyfit.jp/",                            "リスティング広告", "", "24時間フィットネスジム・マシントレーニング",       "JOYFIT ジム 評判 月会費, 24時間 フィットネス 入会 おすすめ, ジム 通い放題 安い","フィットネス"),
    # ── スポーツGPS時計（高単価） ──────────────────────────────────────────
    ("スント・ジャパン株式会社（Suunto）",                  "https://www.suunto.com/ja-JP/",                 "リスティング広告", "", "スポーツGPS時計・トレイルラン・登山計測",         "スント 時計 おすすめ 評判, GPS 時計 スポーツ 比較, Suunto トレイルラン 購入","スポーツ時計"),
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
