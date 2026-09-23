import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03h.xlsx"
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
    # ── 住宅設備・水道修理（高単価） ──────────────────────────────────────────
    ("株式会社ミズテック",                             "https://mizutec.co.jp/",                         "リスティング広告", "", "給湯器・エコキュート・ガス機器交換",     "給湯器 交換 費用 業者, エコキュート 設置 費用, 給湯器 故障 交換 早い",        "住宅設備"),
    ("湯ドクター株式会社",                             "https://yu-doctor.com/",                         "リスティング広告", "", "給湯器交換・修理・無料見積",             "給湯器 交換 安い 業者, 給湯器 修理 おすすめ, お湯 出ない 給湯器 交換",       "住宅設備"),
    ("株式会社ジャパンベスト（生活救急車）",           "https://www.seikatsuq.com/",                     "リスティング広告", "", "水道・鍵・電気・ガラス修理（緊急）",     "水道 修理 業者 24時間, 鍵 開かない 業者 費用, 水漏れ 修理 緊急 近く",       "緊急修理"),
    ("株式会社グラス（ガーデンプラス）",              "https://www.garden-plus.jp/",                    "リスティング広告", "", "外構工事・エクステリア・庭工事",         "外構工事 費用 相場, エクステリア 業者 比較, 庭 リフォーム 見積もり 無料",   "外構工事"),
    ("株式会社ジュプロ（エコキュート専門）",           "https://www.jupro.co.jp/",                       "リスティング広告", "", "エコキュート・太陽光・省エネ設備",       "エコキュート 交換 費用 補助金, 太陽光 蓄電池 費用, 省エネ リフォーム 補助", "省エネ設備"),
    # ── 害虫・害獣駆除（高単価） ──────────────────────────────────────────
    ("株式会社ムシプロテック",                         "https://www.mushiprotec.com/",                   "リスティング広告", "", "害虫駆除・ゴキブリ・ハチ・ムカデ",       "害虫駆除 業者 費用, ゴキブリ 駆除 業者 おすすめ, ハチ 退治 業者 費用",       "害虫駆除"),
    ("ホームレスキュー株式会社",                       "https://home-rescue.jp/",                        "リスティング広告", "", "害獣駆除・ネズミ・ハクビシン・コウモリ", "ネズミ 駆除 業者 費用, ハクビシン 駆除 費用, コウモリ 対策 業者 おすすめ",   "害獣駆除"),
    # ── 健康宅配食・ウォーターサーバー ──────────────────────────────────────────
    ("株式会社Green Spoon",                           "https://green-spoon.jp/",                        "リスティング広告", "", "野菜たっぷり・パーソナル健康宅配食",     "健康 宅配食 ダイエット, サラダ 宅配 定期, Green Spoon 評判 口コミ",          "健康宅配食"),
    ("株式会社One Way（ワンウェイウォーター）",       "https://www.1waywater.com/",                     "リスティング広告", "", "ウォーターサーバー・天然水宅配",         "ウォーターサーバー おすすめ 安い, 天然水 宅配 比較, 水 サーバー 料金 月額",   "ウォーターサーバー"),
    ("サントリーウェルネス株式会社（うるのん）",       "https://www.urunon.jp/",                         "リスティング広告", "", "天然水ウォーターサーバー・宅配水",       "サントリー ウォーターサーバー 評判, 天然水 サーバー 安い, 水 宅配 定期",      "ウォーターサーバー"),
    # ── 共済・保険（残り高単価） ──────────────────────────────────────────
    ("日本コープ共済生活協同組合連合会（コープ共済）", "https://coopkyosai.coop/",                       "リスティング広告", "", "コープ共済・生命・医療・こども共済",     "コープ共済 加入 費用, 共済 保険 比較, こども 共済 安い おすすめ",            "共済・保険"),
    ("一般社団法人こくみん共済 coop（全労済）",        "https://www.zenrosai.coop/",                     "リスティング広告", "", "こくみん共済・火災・医療・生命共済",     "こくみん共済 評判 加入, 掛け金 安い 共済, 火災共済 おすすめ 比較",           "共済・保険"),
    ("東京都民共済生活協同組合（都民共済）",           "https://www.tomin-kyosai.or.jp/",                "リスティング広告", "", "都民共済・生命・医療・火災共済",         "都民共済 加入 評判, 共済 保険 料金 安い, 都民共済 新型 コース 選び方",        "共済・保険"),
    # ── 人間ドック・健康診断（高単価） ──────────────────────────────────────────
    ("株式会社MDV（人間ドックnet）",                  "https://www.ningen-dock.net/",                   "リスティング広告", "", "人間ドック・健康診断の予約比較",         "人間ドック 費用 相場, 健診 おすすめ 病院 比較, 人間ドック 予約 申し込み",    "人間ドック"),
    ("公益財団法人日本健診財団",                       "https://www.kenshin.org/",                       "リスティング広告", "", "人間ドック認定施設・健診専門",           "人間ドック 認定病院 探し方, 健康診断 高精度 おすすめ, 健診 受け方 費用",      "人間ドック"),
    # ── 眼鏡チェーン ──────────────────────────────────────────
    ("愛眼株式会社（アイガン）",                       "https://www.aigan.co.jp/",                       "リスティング広告", "", "眼鏡・サングラス・コンタクト",           "眼鏡 購入 おすすめ, コンタクト 買う 近く, 愛眼 店舗 メガネ 料金",           "眼鏡"),
    ("株式会社ビジョンメガネ",                         "https://www.visionmegane.co.jp/",                "リスティング広告", "", "メガネ・遠近両用・補聴器",               "メガネ 近く 店舗, 遠近両用 おすすめ 価格, 補聴器 眼鏡店 相談",              "眼鏡"),
    # ── 宅配クリーニング（残り） ──────────────────────────────────────────
    ("株式会社ダブルフロンティア（ホワイト急便）",    "https://www.white-kyubin.com/",                  "リスティング広告", "", "宅配クリーニング・集配無料",             "宅配クリーニング 安い 比較, クリーニング 宅配 おすすめ, 洋服 クリーニング 宅配","クリーニング"),
    ("ウォッシュアンドフォールド株式会社",             "https://www.wash-and-fold.jp/",                  "リスティング広告", "", "宅配洗濯・洗いあがり返却サービス",       "洗濯代行 宅配 安い, 洗濯物 まとめて 出す, 宅配 洗濯 サービス 評判",         "宅配洗濯"),
    # ── リース・カーサービス（高単価） ──────────────────────────────────────────
    ("株式会社ディーラーライン（車のサブスク比較）",  "https://car-lease-lab.com/",                     "リスティング広告", "", "カーリース・車のサブスク比較",           "カーリース 比較 おすすめ, 車 サブスク 安い 新車, カーリース 審査 通りやすい","カーリース"),
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
