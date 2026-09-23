import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03o.xlsx"
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
    # ── 予備校（高単価） ──────────────────────────────────────────
    ("代々木ゼミナール株式会社",                            "https://www.yozemi.ac.jp/",                     "リスティング広告", "", "大学受験予備校・映像授業・共通テスト対策",  "大学受験 予備校 おすすめ, 代ゼミ 費用 評判, 共通テスト 対策 スクール 選び方",  "予備校"),
    # ── ファッション（高単価） ──────────────────────────────────────────
    ("ポールスミス ジャパン株式会社（Paul Smith）",         "https://www.paulsmith.co.jp/",                  "リスティング広告", "", "メンズ・レディースファッション・バッグ・雑貨","ポールスミス 新作 コーデ おすすめ, ブリティッシュ ファッション ブランド, ポールスミス バッグ 購入","ファッション"),
    ("ラルフローレン ジャパン合同会社（Ralph Lauren）",     "https://www.ralphlauren.co.jp/",                "リスティング広告", "", "ポロシャツ・ファッション・ホームウェア",         "ラルフローレン ポロシャツ 人気 コーデ, Ralph Lauren 新作 購入, ブランド ファッション 高品質","ファッション"),
    ("マイケルコース ジャパン合同会社（Michael Kors）",    "https://www.michaelkors.co.jp/",                "リスティング広告", "", "バッグ・腕時計・ファッションアクセサリー",       "マイケルコース バッグ おすすめ 人気, MK 時計 ブランド 比較, マイケルコース セール","バッグ・ファッション"),
    ("ケイト・スペード ジャパン株式会社（Kate Spade）",    "https://www.katespade.jp/",                     "リスティング広告", "", "バッグ・ウォレット・ファッションジュエリー",     "ケイト スペード バッグ 新作 人気, Kate Spade 財布 おすすめ, ブランド バッグ 購入","バッグ"),
    ("コーチ・ジャパン合同会社（Coach）",                  "https://www.coach.com/ja-jp/",                  "リスティング広告", "", "レザーバッグ・財布・シューズ・アクセサリー",    "コーチ バッグ 新作 人気, Coach 財布 ブランド 評判, コーチ アウトレット 購入","バッグ"),
    # ── アウトドアウェア（高単価） ──────────────────────────────────────────
    ("アークテリクス ジャパン合同会社（Arc'teryx）",       "https://arcteryx.com/jp/ja/",                   "リスティング広告", "", "プレミアムアウトドアウェア・ハードシェル",       "アークテリクス ジャケット 評判 おすすめ, アウトドア 高機能 ウェア 比較, Arc'teryx 購入","アウトドアウェア"),
    ("マムートジャパン株式会社（Mammut）",                 "https://www.mammut.com/jp/ja/",                 "リスティング広告", "", "クライミングギア・登山ウェア・バックパック",     "マムート ジャケット おすすめ 評判, クライミング ウェア 高機能, 登山 バックパック 比較","アウトドアウェア"),
    ("ホグロフス ジャパン合同会社（Haglöfs）",             "https://www.haglofs.com/ja-JP/",                "リスティング広告", "", "北欧アウトドアウェア・ハイキング・レイヤリング","ホグロフス ジャケット 評判 おすすめ, 北欧 アウトドア ウェア 比較, Haglöfs 購入 価格","アウトドアウェア"),
    # ── ブーツ・シューズ（高単価） ──────────────────────────────────────────
    ("ドクターマーチン ジャパン株式会社（Dr.Martens）",    "https://www.drmartens.com/jp/ja/",              "リスティング広告", "", "レザーブーツ・8ホールブーツ・サンダル",         "ドクターマーチン ブーツ おすすめ 人気, 8ホール 選び方 サイズ, Dr.Martens 新作 購入","ブーツ"),
    ("ハンターブーツ ジャパン合同会社（Hunter）",          "https://www.hunterboots.com/jp/ja_jp/",         "リスティング広告", "", "レインブーツ・ウェリントンブーツ・レインウェア","ハンターブーツ 評判 おすすめ, レインブーツ ブランド 比較, Hunter 雨靴 購入",  "ブーツ"),
    # ── スキー・スノーボード（高単価） ──────────────────────────────────────────
    ("バートン ジャパン合同会社（Burton）",                "https://www.burton.com/jp/ja/",                 "リスティング広告", "", "スノーボード・ビンディング・ウェア",             "バートン スノーボード おすすめ 評判, ボード ビンディング セット, Burton 新作 購入","スノーボード"),
    ("サロモン ジャパン株式会社（Salomon）",               "https://www.salomon.com/ja-jp/",                "リスティング広告", "", "スキー・トレイルランニングシューズ・アウトドア", "サロモン スキー おすすめ 評判, トレイルランニング シューズ 比較, Salomon 購入","スキー・トレラン"),
    ("K2 JAPAN株式会社",                                   "https://k2snow.jp/",                            "リスティング広告", "", "スキー板・スノーボード・アウトドア用品",         "K2 スキー板 おすすめ 評判, スノーボード ブランド 比較, K2 スキー 購入 価格",  "スキー"),
    # ── アウトドア用品（高単価） ──────────────────────────────────────────
    ("コールマン ジャパン株式会社（Coleman）",             "https://www.coleman.co.jp/",                    "リスティング広告", "", "テント・寝袋・ランタン・キャンプ用品",           "コールマン テント おすすめ 評判, キャンプ 用品 ブランド 比較, Coleman 寝袋 購入","アウトドア用品"),
    # ── ゴルフ用品（高単価） ──────────────────────────────────────────
    ("住友ゴム工業株式会社（SRIXON ゴルフ）",              "https://www.srixon.co.jp/",                     "リスティング広告", "", "ゴルフボール・ドライバー・アイアン",             "スリクソン ゴルフ ボール おすすめ, ドライバー 飛距離 選び方, SRIXON 評判 購入","ゴルフ用品"),
    ("ゴルフドゥ株式会社",                                 "https://www.golfdo.com/",                       "リスティング広告", "", "中古ゴルフクラブ・ゴルフ用品EC・買取",          "中古 ゴルフクラブ 買取 おすすめ, ゴルフ用品 ネット 購入 安い, ゴルフドゥ 評判","ゴルフEC"),
    # ── プレミアムスイーツ（高単価） ──────────────────────────────────────────
    ("ゴディバ ジャパン株式会社（GODIVA）",                "https://www.godiva.co.jp/",                     "リスティング広告", "", "高級チョコレート・ギフト・季節限定スイーツ",    "ゴディバ チョコ ギフト おすすめ, 高級 チョコレート ブランド 通販, GODIVA 季節 限定","プレミアムスイーツ"),
    # ── スポーツ用品（高単価） ──────────────────────────────────────────
    ("スポーツオーソリティ合同会社",                       "https://www.sportsauthority.co.jp/",            "リスティング広告", "", "スポーツ用品総合・フィットネス・アウトドア",     "スポーツ用品 店舗 おすすめ, フィットネス 用品 購入, アウトドア 総合 スポーツ 安い","スポーツ用品"),
    # ── トレッキングシューズ（高単価） ──────────────────────────────────────────
    ("メレル ジャパン合同会社（Merrell）",                 "https://www.merrell.com/JP/ja/",                "リスティング広告", "", "トレッキングシューズ・ハイキングシューズ",        "メレル トレッキングシューズ おすすめ, ハイキング 靴 選び方 比較, Merrell 評判 購入","トレッキングシューズ"),
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
