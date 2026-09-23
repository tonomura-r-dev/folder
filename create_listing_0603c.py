import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03c.xlsx"
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
    # ── アウトドア・スポーツ用品 ──────────────────────────────────────────
    ("株式会社好日山荘",                        "https://www.kojitusanso.jp/",                    "リスティング広告", "", "登山・アウトドア用品専門店",             "登山 装備 初心者 おすすめ, アウトドア 用品 店舗 東京, トレッキング シューズ 購入",  "登山・アウトドア"),
    ("株式会社石井スポーツ（ICI石井スポーツ）", "https://www.ici-sports.com/",                    "リスティング広告", "", "登山・スキー・アウトドア用品",           "石井スポーツ 登山 用品, スキー 板 購入 専門店, アウトドア ウェア おすすめ",        "登山・アウトドア"),
    ("株式会社ビクトリア（ビクトリアスポーツ）","https://www.victoria.co.jp/",                    "リスティング広告", "", "総合スポーツ用品店",                     "スポーツ用品 店舗 東京, テニス ラケット 購入, ランニングシューズ おすすめ",          "スポーツ用品"),
    ("株式会社サッカーショップKAMO",            "https://www.sskamo.co.jp/",                      "リスティング広告", "", "サッカー専門スポーツ用品",               "サッカースパイク 購入 おすすめ, フットサルシューズ 専門店, サッカーウェア ブランド",  "スポーツ用品"),
    ("株式会社ワイズロード（Y's Road）",        "https://ysroad.co.jp/",                          "リスティング広告", "", "スポーツ自転車専門店",                   "スポーツ自転車 購入 おすすめ, ロードバイク 専門店 東京, クロスバイク 相場",          "自転車"),
    ("有限会社さかいやスポーツ",                "https://www.sakaiya.com/",                       "リスティング広告", "", "登山・アウトドア用品（神保町）",         "登山用品 東京 神保町, クライミング 装備 初心者, トレッキングポール おすすめ",        "登山・アウトドア"),
    ("株式会社ワイルドワン（WILD-1）",          "https://www.wild1.co.jp/",                       "リスティング広告", "", "キャンプ・アウトドア用品",               "キャンプ 用品 揃える, アウトドア 店舗 関東, テント 購入 初心者 セット",              "キャンプ"),
    ("株式会社ニシスポーツ",                    "https://www.nishi.com/",                         "リスティング広告", "", "陸上競技・スポーツ用品",                 "陸上競技 用品 購入, スパイク 陸上 おすすめ, マット 体操 購入 学校",                  "陸上用品"),
    # ── 保険・金融 ──────────────────────────────────────────
    ("FWD生命保険株式会社",                    "https://www.fwdlife.co.jp/",                     "リスティング広告", "", "ネット完結型生命保険・収入保障",         "生命保険 比較 おすすめ, 収入保障保険 安い, 定期保険 ネット 申込 安い",              "生命保険"),
    ("プルデンシャル生命保険株式会社",          "https://www.prudential.co.jp/",                  "リスティング広告", "", "生命保険・プランニング相談",             "生命保険 見直し 相談, 保険 プランナー 相談 無料, 保険 設計 ライフプラン",            "生命保険"),
    # ── カーリース・乗り物 ──────────────────────────────────────────
    ("オリックス自動車株式会社（カーリース）",  "https://www.carlease-online.jp/",               "リスティング広告", "", "個人向けカーリース（マイカーリース）",   "カーリース 個人 安い, 新車 リース 月額, オリックスカーリース 評判",                  "カーリース"),
    ("株式会社バディカ",                       "https://buddica.com/",                           "リスティング広告", "", "中古車販売・査定",                       "中古車 販売 おすすめ, 車 購入 オンライン, 中古車 査定 高額 買取",                    "中古車"),
    ("株式会社Luup（LUUP）",                   "https://luup.sc/",                               "リスティング広告", "", "電動キックボード・シェアサイクル",       "電動キックボード シェア 乗り方, LUUP 使い方 始め方, 電動 自転車 シェア",             "シェアモビリティ"),
    ("有限会社ラフ&ロード",                    "https://www.roughandroad-net.co.jp/",            "リスティング広告", "", "バイク用品・ウェア・ツーリングギア",     "バイク 用品 通販, ライディングジャケット おすすめ, バイク グローブ 購入 おすすめ",   "バイク用品"),
    # ── 転職・仕事 ──────────────────────────────────────────
    ("株式会社Forkwell（Forkwell）",            "https://forkwell.com/",                          "リスティング広告", "", "エンジニア向け転職・求人サービス",       "エンジニア 転職 おすすめ, IT エンジニア 求人 スカウト, 開発 転職 ポートフォリオ",    "IT転職"),
    ("ウォンテッドリー株式会社（Wantedly）",   "https://www.wantedly.com/",                      "リスティング広告", "", "仕事のやりがい重視の採用・求人",         "Wantedly 会社 訪問, 共感採用 求人 スタートアップ, 転職 やりがい 仕事 探し方",        "転職・採用"),
    # ── フィットネス・マッチング・その他 ──────────────────────────────────────────
    ("FiT24 LLC（FiT24）",                     "https://www.fit24.jp/",                          "リスティング広告", "", "24時間フィットネスジム・会員入会",       "24時間 フィットネス 格安, ジム 月額 安い, フィットネスクラブ 入会 おすすめ",         "フィットネス"),
    ("株式会社Mrk&Co（with）",                 "https://with.is/",                               "リスティング広告", "", "マッチングアプリ（with）",               "マッチングアプリ 比較 おすすめ, 恋活 アプリ 無料 始め方, with 評判 使い方",          "マッチングアプリ"),
    ("株式会社DogHuggy",                       "https://doghuggy.com/",                          "リスティング広告", "", "犬専門ペットシッターサービス",           "ペットシッター 犬 おすすめ, ペットホテル 代わり, 犬 一泊 預け先 安い",              "ペット"),
    ("株式会社エブリー（DELISH KITCHEN）",     "https://delishkitchen.tv/premium",               "リスティング広告", "", "料理動画プレミアム会員（デリッシュ）",   "レシピ 動画 見放題, 料理 アプリ 有料 おすすめ, デリッシュキッチン プレミアム",       "料理動画"),
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
