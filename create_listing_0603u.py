import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03u.xlsx"
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
    # ── プレミアムスキンケア（高単価） ──────────────────────────────────────────
    ("イソップ ジャパン株式会社（Aesop）",                 "https://www.aesop.com/jp/",                     "リスティング広告", "", "プレミアムスキンケア・ボディケア・フレグランス",      "イソップ スキンケア 評判 おすすめ, Aesop ハンドクリーム 人気, プレミアム 化粧品 購入","プレミアムスキンケア"),
    # ── ファッション（高単価） ──────────────────────────────────────────
    ("ヴィヴィアンウエストウッド ジャパン株式会社（Vivienne Westwood）","https://www.viviennewestwood.jp/","リスティング広告", "", "ブリティッシュファッション・バッグ・アクセサリー",   "ヴィヴィアンウエストウッド バッグ 評判, ファッション ブランド 人気, Vivienne 購入","ファッション"),
    # ── 高級コスメ（高単価） ──────────────────────────────────────────
    ("NARS ジャパン合同会社（NARS Cosmetics）",            "https://www.narscosmetics.jp/",                 "リスティング広告", "", "プロ仕様コスメ・ファンデーション・リップ",              "NARS コスメ おすすめ 評判, ファンデーション 高カバー 比較, ナーズ リップ 購入","高級コスメ"),
    ("MAC コスメティックス ジャパン（MAC Cosmetics）",     "https://www.maccosmetics.jp/",                  "リスティング広告", "", "プロフェッショナルメイク・リップ・アイシャドウ",       "MAC コスメ おすすめ 評判, マック リップ 人気 色, プロ メイク ブランド 購入","高級コスメ"),
    ("ボビイ ブラウン コスメティックス（Bobbi Brown）",    "https://www.bobbibrown.co.jp/",                 "リスティング広告", "", "ナチュラルコスメ・ファンデーション・スキンケア",       "ボビイブラウン ファンデーション 評判, コスメ 自然 おすすめ, Bobbi Brown 購入","高級コスメ"),
    # ── テニス用品（高単価） ──────────────────────────────────────────
    ("ウィルソン スポーツ ジャパン合同会社（Wilson）",     "https://www.wilson.co.jp/",                     "リスティング広告", "", "テニスラケット・バドミントン・スポーツ用品",            "ウィルソン テニス ラケット おすすめ, Wilson 評判 選び方, テニス 用品 購入",  "テニス用品"),
    ("ヘッド ジャパン合同会社（HEAD）",                    "https://www.head.com/ja-JP/",                   "リスティング広告", "", "テニスラケット・スキー板・スポーツ用品",               "HEAD テニス ラケット 評判, ヘッド スキー 用品 比較, テニス ラケット 購入",   "テニス用品"),
    ("プリンス スポーツ ジャパン合同会社（Prince Sports）","https://princesports.jp/",                     "リスティング広告", "", "テニスラケット・ガット・テニスシューズ",               "プリンス ラケット おすすめ 評判, テニス 用品 ブランド 比較, Prince 購入",    "テニス用品"),
    # ── 高級スーツケース（超高単価） ──────────────────────────────────────────
    ("RIMOWA Japan合同会社（RIMOWA）",                     "https://www.rimowa.com/ja-jp/",                 "リスティング広告", "", "アルミ・ポリカーボネートスーツケース",                  "リモワ スーツケース 評判 おすすめ, RIMOWA 軽量 比較, 高級 旅行 バッグ 購入","高級スーツケース"),
    # ── 超高級腕時計（超高単価） ──────────────────────────────────────────
    ("IWCシャフハウゼン ジャパン（IWC）",                  "https://www.iwc.com/ja/",                       "リスティング広告", "", "スイス製高級腕時計・パイロットウォッチ",               "IWC 時計 評判 おすすめ, パイロット ウォッチ ブランド, スイス 高級 時計 購入","超高級腕時計"),
    ("ブライトリング ジャパン合同会社（Breitling）",        "https://www.breitling.com/jp/",                 "リスティング広告", "", "アビエーター腕時計・クロノグラフ・高精度時計",         "ブライトリング 時計 評判 おすすめ, Breitling クロノグラフ 購入, 高級 腕時計","超高級腕時計"),
    # ── 真空断熱ボトル（高単価） ──────────────────────────────────────────
    ("スタンレー・ジャパン合同会社（Stanley）",             "https://www.stanley-japan.jp/",                 "リスティング広告", "", "真空断熱マグ・クエンチャー・アウトドアボトル",         "スタンレー 水筒 評判 おすすめ, クエンチャー 保温 比較, Stanley マグ 購入",  "保温ボトル"),
    ("ハイドロフラスク ジャパン（Hydro Flask）",            "https://www.hydroflask.com/ja/",                "リスティング広告", "", "保冷保温ボトル・ワイドマウス・アウトドア",              "ハイドロフラスク 水筒 評判 おすすめ, 保冷 ボトル 比較 長持ち, Hydro Flask 購入","保温ボトル"),
    # ── アウトドア照明（高単価） ──────────────────────────────────────────
    ("ゲントス株式会社（GENTOS）",                         "https://www.gentos.jp/",                        "リスティング広告", "", "高機能ランタン・ヘッドライト・懐中電灯",               "ゲントス ランタン おすすめ 評判, アウトドア ヘッドライト 高輝度, GENTOS 購入","アウトドア照明"),
    # ── プロ用ヘアケア（高単価） ──────────────────────────────────────────
    ("株式会社アリミノ（ARIMINO）",                        "https://www.arimino.co.jp/",                    "リスティング広告", "", "プロ用ヘアワックス・シャンプー・トリートメント",       "アリミノ スタイリング 評判 おすすめ, プロ ヘアケア 美容院 用品, ARIMINO 購入","ヘアケア"),
    # ── オーガニックコスメ（高単価） ──────────────────────────────────────────
    ("株式会社シロ（SHIRO）",                              "https://shiro-shiro.jp/",                       "リスティング広告", "", "オーガニックコスメ・フレグランス・スキンケア",          "シロ コスメ 評判 おすすめ, SHIRO フレグランス 人気, オーガニック 化粧品 通販","オーガニックコスメ"),
    # ── プレミアムクレジットカード（高単価） ──────────────────────────────────────────
    ("三井住友カード株式会社",                              "https://www.smbc-card.com/",                    "リスティング広告", "", "プレミアムカード・ゴールド・NLカード",                   "三井住友カード おすすめ 評判, ゴールドカード 年会費 特典, クレジットカード 比較","クレジットカード"),
    ("株式会社ジェーシービー（JCBカード）",                "https://www.jcb.co.jp/",                        "リスティング広告", "", "JCBカード・プラチナ・ゴールド・国際ブランド",           "JCBカード おすすめ 評判, プラチナカード 審査 特典, クレジット 海外 比較",    "クレジットカード"),
    ("アメリカン・エキスプレス・インターナショナル合同会社","https://www.americanexpress.com/ja-jp/",       "リスティング広告", "", "プレミアムカード・ゴールド・プラチナ・特典",             "アメックス カード おすすめ 評判, ゴールド プラチナ 年会費 特典, American Express 比較","クレジットカード"),
    ("三井住友トラストクラブ株式会社（ダイナースクラブ）", "https://www.diners.co.jp/",                     "リスティング広告", "", "プレミアムカード・空港ラウンジ・旅行特典",              "ダイナースクラブ カード 評判 おすすめ, プレミアム カード 特典 比較, 空港 ラウンジ 無料","クレジットカード"),
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
