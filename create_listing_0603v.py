import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03v.xlsx"
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
    # ── 証券会社（高単価） ──────────────────────────────────────────
    ("野村證券株式会社（野村証券）",                        "https://www.nomura.co.jp/retail/",              "リスティング広告", "", "株式・投資信託・資産運用・NISA",                       "野村証券 評判 おすすめ, 投資信託 証券会社 比較, NISA 口座 おすすめ",         "証券会社"),
    ("株式会社大和証券（大和証券）",                        "https://www.daiwa.jp/retail/",                  "リスティング広告", "", "株式・投資信託・NISAラップ口座",                         "大和証券 評判 おすすめ, 資産運用 始め方 証券, NISAラップ 投資信託 比較",    "証券会社"),
    ("SMBC日興証券株式会社",                               "https://www.smbcnikko.co.jp/",                  "リスティング広告", "", "株式・ETF・米国株・NISA・iDeCo",                        "SMBC日興証券 評判 比較, 米国株 購入 証券, ネット 証券 手数料 おすすめ",     "証券会社"),
    ("岡三証券株式会社",                                   "https://www.okasan.co.jp/",                     "リスティング広告", "", "株式・債券・IPO・NISAセミナー",                          "岡三証券 評判 おすすめ, IPO 投資 証券会社, 株式 始め方 初心者",            "証券会社"),
    ("みずほ証券株式会社",                                 "https://www.mizuho-sc.com/",                    "リスティング広告", "", "株式・投資信託・外国債券・相続相談",                     "みずほ証券 評判 比較, 相続 資産管理 相談, 投資信託 証券 おすすめ",          "証券会社"),
    # ── 保険（高単価） ──────────────────────────────────────────
    ("株式会社かんぽ生命保険（かんぽ生命）",                "https://www.jp-life.japanpost.jp/",             "リスティング広告", "", "生命保険・医療保険・学資保険・年金",                     "かんぽ生命 保険 評判 比較, 学資保険 おすすめ, 生命保険 見直し 相談",        "生命保険"),
    ("JA共済連（JA共済）",                                 "https://www.ja-kyosai.or.jp/",                  "リスティング広告", "", "生命共済・建物火災共済・自動車共済",                     "JA共済 評判 おすすめ, 共済 保険 比較 違い, 建物 火災 保険 安い",           "共済保険"),
    # ── 高級ファッション（高単価） ──────────────────────────────────────────
    ("ヒューゴ・ボス ジャパン株式会社（HUGO BOSS）",        "https://www.hugoboss.com/ja/",                  "リスティング広告", "", "メンズビジネスウェア・スーツ・コート",                   "ヒューゴボス スーツ 評判 おすすめ, HUGO BOSS ビジネス ウェア, 高級 スーツ 購入","高級ファッション"),
    # ── 高級食品・ギフトEC（高単価） ──────────────────────────────────────────
    ("株式会社久原本家（茅乃舎）",                          "https://www.kubara.jp/",                        "リスティング広告", "", "高級だし・だしパック・調味料・ギフト",                   "茅乃舎 だし おすすめ 評判, 高級 だし パック 購入, 久原本家 ギフト 通販",    "高級食品EC"),
    ("株式会社銀座千疋屋",                                  "https://sembikiya.co.jp/",                      "リスティング広告", "", "高級フルーツ・スイーツ・ギフト通販",                     "銀座千疋屋 フルーツ ギフト 通販, 高級 フルーツ お取り寄せ, 贈り物 果物 高品質","高級食品EC"),
    # ── 高級マットレス（高単価） ──────────────────────────────────────────
    ("テンピュール・シーリー・ジャパン株式会社（Tempur）",  "https://www.tempur.co.jp/",                     "リスティング広告", "", "高反発マットレス・枕・睡眠サポート",                     "テンピュール マットレス 評判 おすすめ, 高級 マットレス 腰痛 比較, Tempur 購入","高級寝具"),
    # ── アウトドアチェア（高単価） ──────────────────────────────────────────
    ("ヘリノックス ジャパン合同会社（Helinox）",            "https://helinox.jp/",                           "リスティング広告", "", "軽量アウトドアチェア・コット・テーブル",                 "ヘリノックス チェア おすすめ 評判, アウトドア 椅子 軽量 比較, Helinox 購入","アウトドア家具"),
    # ── 高級テーブルウェア（高単価） ──────────────────────────────────────────
    ("ウェッジウッド ジャパン合同会社（Wedgwood）",         "https://www.wedgwood.co.jp/",                   "リスティング広告", "", "ボーンチャイナ・高級食器・ティーウェア・ギフト",         "ウェッジウッド 食器 おすすめ 評判, 高級 カップ ギフト, Wedgwood 購入 価格","高級食器"),
    ("株式会社ノリタケカンパニーリミテド（Noritake）",     "https://www.noritake.co.jp/",                   "リスティング広告", "", "高級磁器・食器・ギフト・工業用砥石",                     "ノリタケ 食器 評判 おすすめ, 高級 陶磁器 ギフト 購入, Noritake 食器 比較","高級食器"),
    ("ロイヤルコペンハーゲン ジャパン（Royal Copenhagen）","https://www.royalcopenhagen.jp/",              "リスティング広告", "", "デンマーク王室御用達食器・ブルーフルーテッド",           "ロイヤルコペンハーゲン 食器 評判, 北欧 食器 高級 購入, ブルーフルーテッド 価格","高級食器"),
    # ── 北欧ブランド（高単価） ──────────────────────────────────────────
    ("イッタラ ジャパン合同会社（iittala）",               "https://www.iittala.jp/",                       "リスティング広告", "", "フィンランド製ガラス食器・キャンドル・雑貨",             "イッタラ 食器 おすすめ 評判, 北欧 ガラス 食器 比較, iittala 購入 通販",    "北欧食器"),
    ("アラビア ジャパン（Arabia）",                         "https://www.arabia.jp/",                        "リスティング広告", "", "フィンランド製陶器・ムーミンコレクション",               "アラビア 食器 おすすめ 評判, ムーミン マグ 購入, 北欧 陶器 ブランド 比較", "北欧食器"),
    ("マリメッコ ジャパン合同会社（Marimekko）",            "https://www.marimekko.com/jp_ja/",              "リスティング広告", "", "北欧ファッション・バッグ・インテリアテキスタイル",       "マリメッコ バッグ 評判 おすすめ, 北欧 ブランド ファッション, Marimekko 購入","北欧ファッション"),
    # ── プログラミングスクール（高単価） ──────────────────────────────────────────
    ("DIVE INTO CODE株式会社",                             "https://diveintocode.jp/",                      "リスティング広告", "", "プログラミングスクール・AI・Webエンジニア転職",          "プログラミングスクール 評判 比較, エンジニア 転職 スクール, DIVE INTO CODE 費用","プログラミングスクール"),
    # ── 高級食器（高単価） ──────────────────────────────────────────
    ("ロイヤルドルトン ジャパン（Royal Doulton）",          "https://www.royaldoulton.co.jp/",               "リスティング広告", "", "英国高級食器・ボーンチャイナ・ギフト",                   "ロイヤルドルトン 食器 評判, 英国 ブランド 食器 高級, Royal Doulton 購入", "高級食器"),
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
