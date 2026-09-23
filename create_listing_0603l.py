import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03l.xlsx"
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
    # ── ゴルフ用品（超高単価） ──────────────────────────────────────────
    ("テーラーメイドゴルフ株式会社",                      "https://www.taylormadegolf.jp/",                "リスティング広告", "", "ゴルフクラブ・ドライバー・アイアン",       "テーラーメイド ドライバー 最新 評判, ゴルフクラブ 新作 選び方, テーラーメイド 試打","ゴルフ用品"),
    ("ブリヂストンスポーツ株式会社",                      "https://www.bs-sports.co.jp/",                 "リスティング広告", "", "ゴルフボール・ゴルフクラブ・フィッティング","ゴルフ ボール 飛距離 おすすめ, ブリヂストン ゴルフ クラブ 評判, ゴルフ用品 国産","ゴルフ用品"),
    ("タイトリストジャパン合同会社",                      "https://www.titleist.co.jp/",                  "リスティング広告", "", "ゴルフボール・ウェッジ・プロ使用",         "タイトリスト ゴルフ ボール おすすめ, ウェッジ 選び方 費用, タイトリスト 試打 フィッティング","ゴルフ用品"),
    # ── 高級食材EC（高単価） ──────────────────────────────────────────
    ("株式会社かに本舗",                                 "https://www.kani-honpo.com/",                   "リスティング広告", "", "北海道産カニ・海産物・ギフト通販",         "カニ 通販 おすすめ 北海道, 蟹 お取り寄せ ギフト, ズワイガニ 格安 産直",      "高級食材EC"),
    # ── 語学スクール（高単価） ──────────────────────────────────────────
    ("株式会社K-Village Tokyo",                          "https://k-village.jp/",                         "リスティング広告", "", "韓国語スクール・K-POP・オンライン韓国語",  "韓国語 スクール おすすめ, 韓国語 初心者 学び方, K-POP 韓国語 レッスン 費用",  "語学スクール"),
    # ── 着物・和装（高単価） ──────────────────────────────────────────
    ("株式会社きものやまと",                              "https://www.kimonoyamato.co.jp/",               "リスティング広告", "", "着物・振袖・和装小物・レンタル",           "振袖 購入 レンタル 比較, 着物 成人式 価格, きものやまと 評判 振袖 費用",     "着物・和装"),
    # ── カラオケ・音楽配信（高単価） ──────────────────────────────────────────
    ("第一興商株式会社（DAM）",                           "https://www.dkkaraoke.co.jp/",                 "リスティング広告", "", "家庭用カラオケ機器・音楽配信・採点",       "家庭用 カラオケ 機器 おすすめ, DAM 採点 購入 費用, カラオケ 自宅 設置 方法", "カラオケ"),
    # ── 旅行（高単価） ──────────────────────────────────────────
    ("名鉄観光サービス株式会社",                          "https://www.mwt.co.jp/",                        "リスティング広告", "", "国内外パッケージツアー・バス旅行",         "旅行 ツアー 名鉄観光 おすすめ, 国内旅行 パック 費用, シニア 旅行 バス 申し込み","旅行"),
    ("JR東日本旅行株式会社（びゅうトラベル）",            "https://www.jreast-travel.co.jp/",             "リスティング広告", "", "新幹線パック・国内旅行・宿泊セット",       "新幹線 旅行 パック おすすめ, JR 旅行 宿泊 セット 安い, びゅう 旅行 申し込み","旅行"),
    # ── 家電（高単価） ──────────────────────────────────────────
    ("ダイキン工業株式会社",                              "https://www.daikin.co.jp/",                    "リスティング広告", "", "高機能エアコン・除湿機・空気清浄機",       "エアコン 高機能 省エネ おすすめ, 空気清浄 エアコン 比較, ダイキン うるさら 評判","家電"),
    # ── 楽器（高単価） ──────────────────────────────────────────
    ("ローランド株式会社",                               "https://www.roland.com/jp/",                    "リスティング広告", "", "電子ピアノ・デジタル楽器・シンセサイザー", "電子ピアノ 選び方 初心者, デジタル ドラム 自宅 練習, ローランド 電子ピアノ 評判","楽器"),
    # ── キッチン・保温用品（高単価） ──────────────────────────────────────────
    ("株式会社サーモス",                                 "https://www.thermos.jp/",                       "リスティング広告", "", "保温マグ・水筒・スープジャー・ランチジャー","水筒 保温 おすすめ 長持ち, スープジャー ランチ 人気, サーモス 真空断熱 比較",  "キッチン用品"),
    ("タイガー魔法瓶株式会社",                            "https://www.tiger-corporation.com/",           "リスティング広告", "", "高機能炊飯器・保温スープジャー・ポット",   "炊飯器 高機能 おすすめ 比較, タイガー 炊飯器 評判 価格, ポット 保温 長時間",  "家電"),
    ("株式会社ツヴィリングJ.A.ヘンケルスジャパン（STAUB）", "https://www.zwilling.com/jp/",              "リスティング広告", "", "高単価鋳物鍋・ストウブ・シェフナイフ",    "ストウブ 鍋 おすすめ サイズ, 鋳物 鍋 料理 美味しい, STAUB ホーロー 価格 比較","調理器具"),
    # ── 旅行用品（高単価） ──────────────────────────────────────────
    ("株式会社エース",                                   "https://www.ace.jp/",                           "リスティング広告", "", "スーツケース・旅行かばん・ビジネスバッグ", "スーツケース おすすめ 軽量, 旅行 かばん 人気 ブランド, エース スーツケース 評判","旅行用品"),
    # ── テニス用品（高単価） ──────────────────────────────────────────
    ("バボラジャパン株式会社",                            "https://www.babolat.com/ja_JP/",               "リスティング広告", "", "テニスラケット・ストリング・テニスシューズ","テニス ラケット 選び方 中級, バボラ ラケット 評判 おすすめ, テニス 用品 通販", "スポーツ用品"),
    # ── 高級腕時計（高単価） ──────────────────────────────────────────
    ("シチズン時計株式会社",                             "https://www.citizenwatch.co.jp/",               "リスティング広告", "", "エコドライブ・GPS電波時計・高精度時計",   "シチズン 時計 高級 評判, エコドライブ 腕時計 比較 おすすめ, GPS 時計 精度 購入","高級腕時計"),
    # ── カメラ（高単価） ──────────────────────────────────────────
    ("株式会社ニコン",                                   "https://www.nikon-image.com/",                  "リスティング広告", "", "ミラーレス一眼・デジタル一眼レフ・レンズ", "ニコン ミラーレス 評判 選び方, 一眼レフ 初心者 購入, カメラ レンズ おすすめ", "カメラ"),
    ("富士フイルム株式会社（Xシリーズ）",                 "https://fujifilm-x.com/ja-jp/",                "リスティング広告", "", "ミラーレスカメラ・フィルムシミュレーション","富士フイルム カメラ おすすめ, Xシリーズ 評判 フィルム感, ミラーレス 写真 綺麗 安い","カメラ"),
    # ── 腕時計（高単価） ──────────────────────────────────────────
    ("カシオ計算機株式会社（G-SHOCK）",                   "https://www.g-shock.jp/",                      "リスティング広告", "", "G-SHOCK・スポーツウォッチ・限定モデル",   "G-SHOCK おすすめ 人気 モデル, カシオ 時計 ブランド 評判, G-SHOCK 限定 購入",  "腕時計"),
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
