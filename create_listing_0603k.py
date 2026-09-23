import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03k.xlsx"
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
    # ── 投資・資産形成スクール（超高単価） ──────────────────────────────────────────
    ("グローバルファイナンシャルスクール株式会社（GFS）", "https://global-financial-school.com/",          "リスティング広告", "", "株式投資・NISA・資産形成スクール",        "株式投資 スクール 初心者, NISA 使い方 学ぶ, 資産形成 社会人 勉強",           "投資スクール"),
    # ── 船舶免許（高単価） ──────────────────────────────────────────
    ("株式会社マリンライセンスロイヤル",                  "https://www.marine-license.com/",               "リスティング広告", "", "船舶免許1・2級・水上バイク免許取得",       "船舶免許 取得 費用 最短, 1級 船舶免許 スクール, ボート 免許 試験 日程",      "船舶免許"),
    # ── グランピング（高単価） ──────────────────────────────────────────
    ("株式会社RECAMP",                                   "https://recamp.jp/",                            "リスティング広告", "", "グランピング・手ぶらキャンプ施設予約",     "グランピング 予約 関東 おすすめ, 手ぶらキャンプ 施設 比較, グランピング 値段","グランピング"),
    # ── 音楽スクール（高単価） ──────────────────────────────────────────
    ("椿音楽教室株式会社",                               "https://www.tsubaki-music.com/",                "リスティング広告", "", "ピアノ・ボイトレ・ギター個人レッスン",     "ピアノ教室 大人 初心者 マンツーマン, ボイトレ 個人 レッスン 近く, 音楽 習う","音楽教室"),
    # ── ゴルフ場・会員権（高単価） ──────────────────────────────────────────
    ("株式会社パシフィックゴルフマネジメント（PGMゴルフ）","https://www.pgm.co.jp/",                      "リスティング広告", "", "ゴルフ場予約・会員権・ゴルフレッスン",    "ゴルフ場 予約 空き 確認, ゴルフ 会員権 おすすめ, ゴルフ 初心者 コース 予約",  "ゴルフ場"),
    # ── 高級寝具（高単価） ──────────────────────────────────────────
    ("株式会社サータ・ジャパン（Serta）",                "https://www.serta.co.jp/",                      "リスティング広告", "", "プレミアムマットレス・ベッドフレーム",     "高級 マットレス おすすめ 腰痛 改善, ホテル マットレス 家庭用 購入, サータ 評判","高級寝具"),
    # ── マッサージチェア（高単価） ──────────────────────────────────────────
    ("ファミリーイナダ株式会社",                          "https://www.family-inada.com/",                "リスティング広告", "", "マッサージチェア・フルボディ・多機能",     "マッサージチェア 比較 おすすめ 高機能, 家庭用 マッサージ椅子 価格, イナダ 評判","健康器具"),
    # ── シューズ（高単価） ──────────────────────────────────────────
    ("株式会社エービーシー・マート",                      "https://www.abc-mart.net/",                    "リスティング広告", "", "スニーカー・ランニングシューズ・ブーツ",   "スニーカー おすすめ 新作 人気, ランニングシューズ 購入 おすすめ, ABCマート セール","シューズ"),
    # ── 会員制リゾート（超高単価） ──────────────────────────────────────────
    ("東急リゾーツ＆ステイ株式会社（東急ハーヴェストクラブ）","https://harvest.tokyu-resort.co.jp/",      "リスティング広告", "", "会員制リゾートホテル・別荘型・家族旅行",  "会員制 リゾート ホテル 購入, 東急ハーヴェスト 料金 評判, 別荘 リゾート 費用", "会員制リゾート"),
    # ── カーリース（高単価） ──────────────────────────────────────────
    ("株式会社カーモ（カーモくん）",                      "https://www.carmo-kun.jp/",                    "リスティング広告", "", "カーリース・定額制マイカー・新車乗換",     "カーリース 定額 新車 乗れる, カーモ 評判 審査, 車 月額 安い リース 比較",    "カーリース"),
    # ── 英会話スクール（高単価） ──────────────────────────────────────────
    ("株式会社ブライチャー",                             "https://www.breaker.jp/",                       "リスティング広告", "", "ネイティブ英会話・ビジネス英語スクール",   "英会話 ネイティブ おすすめ スクール, ビジネス英語 上達 方法, 英語 社会人 習う","英会話"),
    # ── スポーツ用品（高単価） ──────────────────────────────────────────
    ("株式会社ミズノ",                                   "https://jp.mizuno.com/",                        "リスティング広告", "", "スポーツ用品・野球・ランニングシューズ",   "ミズノ シューズ おすすめ 選び方, 野球 グラブ 購入 ミズノ, ランニング 靴 高機能","スポーツ用品"),
    ("アシックスジャパン株式会社",                        "https://www.asics.com/jp/ja-jp/",              "リスティング広告", "", "ランニングシューズ・スポーツウェア",       "ランニングシューズ 選び方 初心者, アシックス 人気 モデル おすすめ, マラソン 靴 購入","ランニング"),
    ("株式会社ヨネックス",                               "https://www.yonex.co.jp/",                      "リスティング広告", "", "テニス・バドミントンラケット・用品",        "ヨネックス テニス ラケット 選び方, バドミントン ラケット 初心者 おすすめ, スポーツ用品 通販","スポーツ用品"),
    # ── 電動自転車（高単価） ──────────────────────────────────────────
    ("ヤマハ発動機株式会社（PAS電動アシスト自転車）",       "https://www.yamaha-motor.co.jp/pas/",          "リスティング広告", "", "電動アシスト自転車・通勤・子乗せモデル",   "電動自転車 おすすめ 通勤, 電動アシスト 子乗せ 選び方, ヤマハPAS 評判 価格",  "電動自転車"),
    # ── ジュエリー（高単価） ──────────────────────────────────────────
    ("株式会社御木本（ミキモト）",                        "https://www.mikimoto.com/ja_JP/",               "リスティング広告", "", "真珠・ジュエリー・プレミアムネックレス",   "ミキモト 真珠 ネックレス 価格, 婚約 記念 真珠 ジュエリー 本物, 真珠 ネックレス 選び方","ジュエリー"),
    # ── 旅行（高単価） ──────────────────────────────────────────
    ("JTB株式会社",                                      "https://www.jtb.co.jp/",                        "リスティング広告", "", "国内外ツアー・パッケージ旅行・ハネムーン","旅行 パッケージ おすすめ 人気, JTB ツアー 海外 申し込み, ハネムーン 旅行 費用","旅行"),
    # ── 家電量販（高単価） ──────────────────────────────────────────
    ("株式会社エディオン",                               "https://www.edion.com/",                         "リスティング広告", "", "家電・エアコン設置工事・修理サービス",     "エアコン 工事 取付 安い 業者, 家電 購入 アドバイス 店舗, エディオン セール",  "家電量販"),
    # ── ロボット掃除機（高単価） ──────────────────────────────────────────
    ("アイロボットジャパン合同会社（ルンバ）",             "https://www.irobot-jp.com/",                   "リスティング広告", "", "ロボット掃除機・自動ゴミ収集・マッピング",  "ルンバ おすすめ 機種 比較, ロボット掃除機 評判 選び方, 自動掃除 ルンバ 価格",  "ロボット家電"),
    # ── 高級腕時計（高単価） ──────────────────────────────────────────
    ("セイコーグループ株式会社（SEIKOウオッチ）",          "https://www.seikowatches.com/jp-ja/",           "リスティング広告", "", "高級腕時計・プレステージ・スポーツウォッチ","セイコー 腕時計 高級 おすすめ, 時計 男性 プレゼント ブランド, プレサージュ 評判","高級腕時計"),
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
