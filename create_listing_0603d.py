import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03d.xlsx"
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
    # ── 化粧品・スキンケア通販（D2C）──────────────────────────────────────────
    ("株式会社ナチュラルサイエンス",                    "https://www.naturalscience.co.jp/",              "リスティング広告", "", "スキンケア・美容液・化粧品通販",         "スキンケア 通販 おすすめ 安い, 美容液 高品質 コスパ, 化粧水 保湿 D2C",       "化粧品通販"),
    ("株式会社チューンメーカーズ（TUNEMAKERS）",       "https://tunemakers.net/",                        "リスティング広告", "", "高濃度原液美容液・スキンケア通販",       "美容液 原液 高濃度, ビタミンC 美容液 通販 口コミ, エイジングケア 化粧品",    "化粧品通販"),
    ("株式会社アルソア本社（ARSOA）",                  "https://www.arsoa.co.jp/",                       "リスティング広告", "", "ハーブ自然派化粧品・スキンケア",         "自然派 化粧品 おすすめ, ハーブ スキンケア 通販, 無添加 化粧品 口コミ",       "化粧品通販"),
    # ── 家庭教師 ──────────────────────────────────────────
    ("株式会社名門会（家庭教師の名門会）",              "https://www.meimonsurvivor.jp/",                 "リスティング広告", "", "家庭教師・個別指導派遣",                 "家庭教師 おすすめ 派遣, 名門会 評判 費用, 個別指導 自宅 中学生",             "家庭教師"),
    ("家庭教師のガンバ",                               "https://gamba-katekyo.jp/",                      "リスティング広告", "", "家庭教師・学習指導・受験対策",           "家庭教師 安い 個別, ガンバ 評判 料金, 家庭教師 小学生 算数 英語",           "家庭教師"),
    ("株式会社スタンダード（家庭教師のスタンダード）",  "https://standard-tutoring.jp/",                  "リスティング広告", "", "家庭教師派遣・受験指導",                 "家庭教師 受験 対策, スタンダード 費用, 家庭教師 中学受験 おすすめ",          "家庭教師"),
    ("家庭教師のナンバー1",                            "https://www.number1katekyo.com/",                 "リスティング広告", "", "家庭教師・一対一完全個別指導",           "家庭教師 派遣 評判, ナンバー1 料金 口コミ, 家庭教師 高校受験 東京",          "家庭教師"),
    ("家庭教師のコナン",                               "https://www.conan-katekyo.jp/",                   "リスティング広告", "", "家庭教師・定期テスト対策・苦手克服",     "家庭教師 定期テスト 対策, コナン 費用 口コミ, 成績アップ 個別指導",          "家庭教師"),
    ("全国家庭教師協会",                               "https://www.zenkoku-katekyo.com/",                "リスティング広告", "", "家庭教師紹介・全国対応派遣",             "家庭教師 紹介 比較 無料, 全国 家庭教師 探し方, 家庭教師 体験 申し込み",      "家庭教師"),
    ("ベスト家庭教師",                                 "https://www.best-katekyo.jp/",                    "リスティング広告", "", "家庭教師・難関校受験特化",               "家庭教師 難関校 受験, ベスト家庭教師 評判 費用, 個別指導 自宅 プロ",         "家庭教師"),
    # ── ランドセル ──────────────────────────────────────────
    ("株式会社池田屋（池田屋ランドセル）",              "https://ikedayabag.com/",                         "リスティング広告", "", "ランドセル・新入学準備・直販",           "ランドセル おすすめ 男の子, 池田屋 ランドセル 口コミ, 小学校 入学 準備",     "ランドセル"),
    ("株式会社羅羅屋（ランドセルの羅羅屋）",           "https://www.raraya.co.jp/",                       "リスティング広告", "", "カラー豊富・防水ランドセル",             "ランドセル 女の子 かわいい, カラー ランドセル 選び方, 軽い ランドセル 人気", "ランドセル"),
    ("株式会社橋本（フィットちゃんランドセル）",       "https://fittchan.jp/",                            "リスティング広告", "", "背負いやすい・軽量ランドセル",           "フィットちゃん 評判 口コミ, ランドセル 軽い 6年保証, ランドセル 購入 時期",  "ランドセル"),
    ("株式会社鞄工房山本（山本鞄ランドセル）",         "https://www.bag-y.co.jp/",                        "リスティング広告", "", "職人手作り・工房系ランドセル",           "工房系 ランドセル おすすめ, 鞄工房山本 評判, 手作り ランドセル 職人 本革",   "ランドセル"),
    ("株式会社村瀬鞄行（村瀬鞄行ランドセル）",         "https://murase-kaban.co.jp/",                     "リスティング広告", "", "名古屋発・伝統製法ランドセル",           "ランドセル 名古屋 工房, 村瀬鞄行 評判 口コミ, 本革 ランドセル 職人 安心",   "ランドセル"),
    # ── リラクゼーション ──────────────────────────────────────────
    ("株式会社ほぐしの達人（ほぐしの達人）",           "https://www.hogushi.com/",                        "リスティング広告", "", "もみほぐし・リラクゼーションサロン",     "もみほぐし 安い 近く, リラクゼーション 駅近 予約, ほぐしの達人 料金",        "リラクゼーション"),
    ("快癒工房",                                       "https://www.kaiyu-koubou.com/",                   "リスティング広告", "", "リラクゼーション・マッサージサロン",     "整体 リラク 料金 安い, 快癒工房 予約 店舗, 肩こり 腰痛 もみほぐし",         "リラクゼーション"),
    # ── 合宿免許 ──────────────────────────────────────────
    ("株式会社ドリームジャパン（合宿免許ドリーム）",   "https://www.dream-japan.co.jp/",                  "リスティング広告", "", "合宿免許・免許取得プラン比較",           "合宿免許 安い おすすめ, 運転免許 合宿 申し込み, 免許 最短 取得 費用",        "合宿免許"),
    ("株式会社ブルーム（合宿免許受付センター）",       "https://www.goukaku.ne.jp/",                      "リスティング広告", "", "合宿免許・教習所比較・一括申込",         "合宿免許 比較 安い 一覧, 教習所 合宿 申し込み, 免許 合宿 料金 相場",        "合宿免許"),
    ("株式会社ジェック（合宿免許わかば）",             "https://wakaba.jecs.co.jp/",                      "リスティング広告", "", "合宿免許・AT限定・格安プラン",           "合宿免許 格安 女性 1人, AT 限定 合宿 費用 短期, 免許取得 合宿 口コミ",      "合宿免許"),
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
