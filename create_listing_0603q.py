import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03q.xlsx"
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
    # ── 高級家電（超高単価） ──────────────────────────────────────────
    ("ミーレ・ジャパン株式会社（Miele）",                  "https://www.miele.co.jp/",                      "リスティング広告", "", "高級洗濯機・食洗機・掃除機・オーブン",           "ミーレ 洗濯機 評判 おすすめ, Miele 食洗機 比較 価格, 高級 家電 ドイツ 購入","高級家電"),
    ("フィリップスジャパン合同会社（Philips）",            "https://www.philips.co.jp/",                    "リスティング広告", "", "電動歯ブラシ・電動シェーバー・空気清浄機",        "フィリップス 電動歯ブラシ おすすめ, 電動 シェーバー 比較, Philips ソニッケアー 購入","家電"),
    # ── ホビー（高単価） ──────────────────────────────────────────
    ("株式会社タミヤ",                                     "https://www.tamiya.com/japan/",                 "リスティング広告", "", "ラジコン・プラモデル・ミニ四駆・塗料",             "タミヤ ラジコン おすすめ 初心者, プラモデル 高品質 購入, ミニ四駆 パーツ 通販","ホビー"),
    # ── 漢方薬（高単価） ──────────────────────────────────────────
    ("株式会社ツムラ",                                     "https://www.tsumura.co.jp/",                    "リスティング広告", "", "漢方薬・エキス顆粒・ボタニカルケア",               "ツムラ 漢方 おすすめ 症状, 漢方薬 効果 種類 選び方, 冷え性 漢方 購入",      "漢方薬"),
    # ── 高級コスメ（高単価） ──────────────────────────────────────────
    ("株式会社コーセー（コスメデコルテ）",                 "https://www.cosmedecorte.com/",                 "リスティング広告", "", "プレミアム美容液・化粧水・コスメ",                  "コスメデコルテ おすすめ 評判, 高級 スキンケア ブランド 比較, COSMEDECORTÉ 購入","高級コスメ"),
    # ── ドローン（超高単価） ──────────────────────────────────────────
    ("DJI Japan株式会社（DJI）",                           "https://www.dji.com/jp/",                       "リスティング広告", "", "ドローン・空撮カメラ・ジンバル・FPV",              "DJI ドローン おすすめ 評判, 空撮 ドローン 購入 免許, DJI Mini 価格 比較",   "ドローン"),
    # ── ファッション（高単価） ──────────────────────────────────────────
    ("PVH Japan合同会社（Tommy Hilfiger）",                "https://www.tommy.co.jp/",                      "リスティング広告", "", "アメリカンカジュアルファッション・ポロシャツ",     "トミーヒルフィガー 新作 人気, アメカジ ブランド おすすめ コーデ, Tommy セール","ファッション"),
    ("PVH Japan合同会社（Calvin Klein）",                  "https://www.calvinklein.co.jp/",                "リスティング広告", "", "デニム・バッグ・下着・ファッション",                "カルバンクライン デニム 人気, Calvin Klein バッグ 評判 おすすめ, CK 下着 購入","ファッション"),
    ("リーバイ・ストラウス ジャパン株式会社（Levi's）",    "https://www.levi.co.jp/",                       "リスティング広告", "", "デニムジーンズ・ジャケット・カジュアルウェア",     "リーバイス ジーンズ 人気 モデル, デニム ブランド 選び方, Levi's 501 購入",  "ファッション"),
    # ── サーフ・アウトドア（高単価） ──────────────────────────────────────────
    ("クイックシルバー ジャパン株式会社（Quiksilver）",   "https://www.quiksilver.jp/",                    "リスティング広告", "", "サーフウェア・ボードショーツ・アウトドアウェア",   "クイックシルバー ウェア おすすめ, サーフィン ウェア ブランド, Quiksilver 購入","サーフ・アウトドア"),
    # ── 競泳水着（高単価） ──────────────────────────────────────────
    ("株式会社アリーナ（arena）",                          "https://www.arena.co.jp/",                      "リスティング広告", "", "競泳水着・水泳用品・スイミングキャップ",            "アリーナ 水着 おすすめ 競泳, arena 水着 選び方, 水泳 用品 ブランド 比較",   "水泳用品"),
    ("スピード ジャパン株式会社（SPEEDO）",                "https://www.speedo.jp/",                        "リスティング広告", "", "競泳水着・スポーツ水着・水泳小物",                  "スピード 水着 評判 おすすめ, SPEEDO 競泳 用品, 水泳 スポーツ 水着 選び方",  "水泳用品"),
    # ── ウィッグ（高単価） ──────────────────────────────────────────
    ("株式会社フォンテーヌ",                               "https://www.fontaine.co.jp/",                   "リスティング広告", "", "医療用ウィッグ・ファッションかつら・育毛",          "ウィッグ おすすめ 医療用, かつら 自然 高品質, フォンテーヌ 評判 価格",       "ウィッグ"),
    # ── リラクゼーション（高単価） ──────────────────────────────────────────
    ("株式会社ラフィネ（Raffine）",                        "https://raffine.jp/",                           "リスティング広告", "", "リラクゼーション・アロマ・ボディケア",              "リラクゼーション おすすめ 近く, ラフィネ 料金 評判, ボディケア サロン 予約","リラクゼーション"),
    # ── クライミングギア（高単価） ──────────────────────────────────────────
    ("ブラックダイヤモンド ジャパン（Black Diamond）",    "https://www.blackdiamondequipment.com/ja_JP/",  "リスティング広告", "", "クライミングギア・登山用品・ヘッドランプ",          "ブラックダイヤモンド クライミング ギア, 登山 ヘッドランプ 高性能, Black Diamond 購入","クライミング"),
    # ── バッグ（高単価） ──────────────────────────────────────────
    ("吉田カバン株式会社（PORTER）",                       "https://porter.yoshidakaban.com/",              "リスティング広告", "", "ポーターバッグ・ビジネスバッグ・財布",              "ポーター バッグ おすすめ 人気, PORTER 評判 ブランド, 吉田カバン 財布 購入","バッグ"),
    ("フルラ ジャパン株式会社（FURLA）",                   "https://www.furla.com/jp/ja/",                  "リスティング広告", "", "イタリアンレザーバッグ・財布・ハンドバッグ",        "フルラ バッグ おすすめ 評判, FURLA 財布 人気 色, イタリア バッグ ブランド","バッグ"),
    # ── ゴルフ用品（高単価） ──────────────────────────────────────────
    ("クリーブランドゴルフジャパン合同会社（Cleveland Golf）","https://www.clevelandgolf.jp/",             "リスティング広告", "", "ゴルフウェッジ・アイアン・スピン性能",              "クリーブランド ウェッジ 評判 おすすめ, ゴルフ アプローチ 上達 ウェッジ, Cleveland Golf 購入","ゴルフ用品"),
    # ── 天体・光学（高単価） ──────────────────────────────────────────
    ("ビクセン株式会社",                                   "https://www.vixen.co.jp/",                      "リスティング広告", "", "天体望遠鏡・双眼鏡・天体観測用品",                  "天体望遠鏡 おすすめ 初心者, 双眼鏡 高品質 比較, ビクセン 評判 天体 観測",   "光学機器"),
    # ── プリンター（高単価） ──────────────────────────────────────────
    ("コニカミノルタジャパン株式会社",                     "https://www.konicaminolta.jp/",                 "リスティング広告", "", "複合機・カラーレーザープリンター・業務用印刷",      "コニカミノルタ 複合機 評判 価格, カラー プリンター 業務用 比較, 印刷 機器 導入","プリンター"),
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
