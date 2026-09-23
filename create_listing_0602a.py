import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-02a.xlsx"
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
    ("株式会社ワタベウェディング",      "https://kyoto-wakon.watabe-wedding.co.jp/lp/",         "リスティング広告", "", "京都和婚・海外ウェディング",         "結婚式 京都 和婚, 海外挙式 ハワイ 費用, リゾート婚 沖縄 格安",       "ウェディング"),
    ("株式会社オンザページ",            "https://produce.novarese.jp/lp/",                       "リスティング広告", "", "レストランウェディング（ノバレーゼ）", "レストランウェディング おすすめ, 一軒家 貸切 式場, おしゃれ 結婚式 東京", "ウェディング"),
    ("アニヴェルセル株式会社",          "https://www.anniversaire.co.jp/lp/",                   "リスティング広告", "", "独立型チャペル結婚式",               "チャペル 結婚式 おすすめ, 式場 見学 予約 表参道, アニヴェルセル 費用", "ウェディング"),
    ("株式会社Plan・Do・See",           "https://www.plan-do-see.co.jp/lp/",                    "リスティング広告", "", "レストランウェディング・貸切パーティー", "レストラン 結婚式 貸切, 少人数 挙式 おしゃれ, パーティー 結婚式 費用", "ウェディング"),
    ("株式会社マリエール",              "https://www.mariell.co.jp/lp/",                        "リスティング広告", "", "全国チェーン結婚式場",               "結婚式場 費用 比較, 神前式 チャペル 両方, ブライダルフェア 無料",     "ウェディング"),
    ("株式会社ベルクラシック",          "https://www.bellclassic.co.jp/lp/",                    "リスティング広告", "", "総合結婚式場（東海・関西）",         "結婚式場 名古屋 おすすめ, ウェディングフェア 特典, 式場 見学 申込",   "ウェディング"),
    ("株式会社アールイズウエディング",  "https://www.arluis.com/lp/",                           "リスティング広告", "", "海外・沖縄リゾートウェディング",     "グアム 挙式 費用, 沖縄 リゾート婚 格安, 海外挙式 二人だけ おすすめ", "リゾートウェディング"),
    ("株式会社スマ婚",                  "https://smakon.jp/lp/",                                "リスティング広告", "", "少人数・格安ウェディング",           "少人数 結婚式 費用, 格安 ウェディング 20人, スマ婚 口コミ 評判",     "ウェディング"),
    ("株式会社フォトワ",                "https://fotowa.com/lp/",                               "リスティング広告", "", "フォトウェディング・出張撮影",       "フォトウェディング 費用 安い, 前撮り カメラマン 派遣, ウェディングフォト 格安", "フォトウェディング"),
    ("株式会社タカミブライダル",        "https://www.takami-bridal.com/lp/",                    "リスティング広告", "", "ウェディングドレス・ブライダル衣装", "ウェディングドレス 試着 予約, ブライダル 衣装 おすすめ, 式場 ドレス セット", "ウェディング衣装"),
    ("株式会社アールベル",              "https://www.r-bell.co.jp/lp/",                         "リスティング広告", "", "総合結婚式場（東北・全国）",         "結婚式場 仙台 おすすめ, 式場 東北 費用, ブライダルフェア 無料 申込",  "ウェディング"),
    ("株式会社フィールオブジョイ",      "https://www.feelofjoy.co.jp/lp/",                      "リスティング広告", "", "ハウスウェディング",                 "ハウスウェディング 費用, 一軒家 貸切 結婚式, アットホーム 式場 人気", "ウェディング"),
    ("株式会社エルアンジェ",            "https://www.elange.co.jp/lp/",                         "リスティング広告", "", "独立型チャペル・披露宴会場（関西）", "チャペル 式場 大阪 費用, 結婚式場 関西 おすすめ, エルアンジェ フェア", "ウェディング"),
    ("株式会社ジョイントワークス",      "https://www.joint-works.jp/lp/",                       "リスティング広告", "", "ブライダルプロデュース・海外挙式",   "ハワイ 結婚式 プロデュース, 海外挙式 専門 旅行, 挙式 プランナー 依頼", "ウェディング"),
    ("株式会社ヴィラアンジェリカ",      "https://www.villa-angelica.co.jp/lp/",                 "リスティング広告", "", "ガーデンウェディング",               "ガーデンウェディング 費用, 庭 結婚式 貸切, 屋外 挙式 おしゃれ",     "ウェディング"),
    ("株式会社ジョイフル恵利",          "https://www.joyfuleli.co.jp/lp/",                      "リスティング広告", "", "ウェディング衣装レンタル",           "ウェディングドレス レンタル 安い, 和装 レンタル 結婚式, 衣装 試着 予約", "ウェディング衣装"),
    ("株式会社グレースコンチネンタル",  "https://www.grace-continental.co.jp/lp/",              "リスティング広告", "", "ブライダルドレス・フォーマルドレス", "ブライダル ドレス 購入, フォーマル ドレス 結婚式 参列, お呼ばれ ドレス 上品", "ウェディング衣装"),
    ("株式会社オリエンタルウェディング","https://www.orientalwedding.co.jp/lp/",                "リスティング広告", "", "ウェディングプランニング・比較",     "結婚式 プラン 比較, ウェディング 見積もり 無料, 式場 相談 コンシェルジュ", "ウェディング"),
    ("株式会社ケイウノ",                "https://www.k-uno.co.jp/lp/",                          "リスティング広告", "", "オーダーメイド結婚・婚約指輪",       "オーダーメイド 結婚指輪 費用, 婚約指輪 デザイン 自由, 指輪 工房 体験", "ブライダルジュエリー"),
    ("株式会社ダイヤモンドシライシ",    "https://www.diamond-shiraishi.co.jp/lp/",              "リスティング広告", "", "婚約指輪・結婚指輪（オーダー）",     "婚約指輪 おすすめ ブランド, 結婚指輪 オーダー 費用, ダイヤ 指輪 選び方", "ブライダルジュエリー"),
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
