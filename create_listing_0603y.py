import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03y.xlsx"
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
    # ── 福利厚生サービス（高単価） ──────────────────────────────────────────
    ("ベネフィット・ワン株式会社",                          "https://bs.benefit-one.co.jp/",                 "リスティング広告", "", "法人向け福利厚生・社員旅行・ポイント優待",            "ベネフィット・ワン 評判 料金, 福利厚生 サービス 法人 比較, 社員 優待 制度","福利厚生"),
    ("株式会社リロクラブ",                                  "https://www.reloclub.jp/",                      "リスティング広告", "", "法人福利厚生・旅行・飲食・育児支援",                   "リロクラブ 評判 料金, 福利厚生 法人 比較, 旅行 優待 割引 サービス",       "福利厚生"),
    # ── 高級オーダーキッチン（超高単価） ──────────────────────────────────────────
    ("株式会社キッチンハウス",                              "https://www.kitchenhouse.jp/",                  "リスティング広告", "", "フルオーダーキッチン・ステンレス・高級設計",           "キッチンハウス 評判 価格, フルオーダー キッチン 高級, ステンレス キッチン 費用","高級キッチン"),
    # ── 高級インテリア（高単価） ──────────────────────────────────────────
    ("コンランショップ ジャパン合同会社（The Conran Shop）","https://www.conranshop.jp/",                   "リスティング広告", "", "デザイン家具・インテリア雑貨・照明・アート",            "コンランショップ 家具 おすすめ, インテリア デザイン ブランド 比較, Conran 購入","インテリア"),
    ("ボーコンセプト ジャパン合同会社（BoConcept）",        "https://www.boconcept.com/ja-jp/",              "リスティング広告", "", "デンマーク製デザイン家具・ソファ・ダイニング",         "ボーコンセプト 家具 評判 おすすめ, デンマーク デザイン 家具 比較, BoConcept 購入","高級家具"),
    # ── 損害保険（高単価） ──────────────────────────────────────────
    ("AIG損害保険株式会社",                                "https://www.aig.co.jp/",                        "リスティング広告", "", "海外旅行保険・自動車保険・医療保険",                   "AIG 保険 評判 おすすめ, 海外旅行 保険 比較 安い, 損害保険 見積もり AIG","損害保険"),
    # ── 転職（高単価） ──────────────────────────────────────────
    ("株式会社エン・ジャパン（type転職）",                  "https://type.jp/",                              "リスティング広告", "", "IT・エンジニア・営業・正社員転職",                     "type転職 評判 おすすめ, IT エンジニア 転職 サイト 比較, 正社員 転職 求人","転職"),
    # ── 高級ベビーカー（超高単価） ──────────────────────────────────────────
    ("ブガブー ジャパン合同会社（Bugaboo）",               "https://www.bugaboo.com/ja-jp/",                "リスティング広告", "", "オランダ製高級ベビーカー・抱っこ紐・アクセサリー",     "ブガブー ベビーカー 評判 おすすめ, 高級 ベビーカー 比較, Bugaboo 購入 価格","高級ベビーカー"),
    ("アッパーベビー ジャパン（UPPAbaby）",                "https://uppababy.jp/",                          "リスティング広告", "", "アメリカ製高級ベビーカー・クルーザー",                  "UPPAbaby ベビーカー 評判 おすすめ, 高級 ベビーカー 選び方, アッパーベビー 購入","高級ベビーカー"),
    ("ベビーゼン ジャパン（Babyzen YOYO）",               "https://www.babyzen.com/ja/",                   "リスティング広告", "", "超軽量折りたたみベビーカー・YOYO2・機内持込",          "ベビーゼン YOYO ベビーカー 評判, 機内 持ち込み ベビーカー 比較, Babyzen 購入","高級ベビーカー"),
    # ── 高級自転車（超高単価） ──────────────────────────────────────────
    ("ピナレロ ジャパン合同会社（Pinarello）",              "https://pinarello.co.jp/",                      "リスティング広告", "", "プロ仕様ロードバイク・カーボンフレーム・イタリア製",    "ピナレロ ロードバイク 評判 おすすめ, プロ 自転車 高性能 比較, Pinarello 購入","ロードバイク"),
    ("ブロンプトン ジャパン（Brompton）",                   "https://www.brompton.com/ja",                   "リスティング広告", "", "英国製折りたたみ自転車・コンパクト・通勤",              "ブロンプトン 評判 おすすめ, 折りたたみ 自転車 高品質 比較, Brompton 購入","折りたたみ自転車"),
    ("メリダ ジャパン株式会社（Merida）",                   "https://www.merida.jp/",                        "リスティング広告", "", "ロードバイク・MTB・Eバイク・クロスバイク",              "メリダ 自転車 評判 おすすめ, ロードバイク 選び方, Merida クロスバイク 購入","自転車"),
    # ── スーツ・ビジネスウェア（高単価） ──────────────────────────────────────────
    ("株式会社ザ・スーツカンパニー（THE SUIT COMPANY）",   "https://www.the-suit-company.co.jp/",           "リスティング広告", "", "ビジネススーツ・シャツ・ネクタイ・ビジネスウェア",     "スーツカンパニー スーツ 評判 おすすめ, ビジネス スーツ 比較 安い, オーダー スーツ","スーツ"),
    ("株式会社ユニバーサルランゲージ",                      "https://www.universal-language.co.jp/",         "リスティング広告", "", "セミオーダースーツ・ビジネスウェア・サイジング",        "ユニバーサルランゲージ スーツ 評判, セミオーダー スーツ 比較, ビジネス ウェア 購入","スーツ"),
    # ── ファッション（高単価） ──────────────────────────────────────────
    ("株式会社アダストリア（GLOBAL WORK）",                "https://store.globalwork.jp/",                  "リスティング広告", "", "カジュアルファッション・大人カジュアル・通販",           "グローバルワーク ファッション 評判, GLOBAL WORK 新作 セール, アダストリア 通販","ファッション"),
    ("株式会社デイトナ・インターナショナル（FREAK'S STORE）","https://freaksstore.com/",                  "リスティング広告", "", "アメカジ・アウトドアミックスファッション・バッグ",       "フリークスストア 評判 人気, FREAK'S STORE 新作 通販, アメカジ ファッション","ファッション"),
    # ── 高品質メガネ（高単価） ──────────────────────────────────────────
    ("株式会社フォーナインズ（999.9）",                     "https://www.fournines.co.jp/",                  "リスティング広告", "", "高品質日本製眼鏡・フレーム・ノーズパッドレス",           "フォーナインズ メガネ 評判 おすすめ, 999.9 フレーム 高品質, 日本製 眼鏡 購入","高品質メガネ"),
    # ── 高級食材EC（高単価） ──────────────────────────────────────────
    ("有限会社魚久",                                       "https://www.uokyu.co.jp/",                      "リスティング広告", "", "京漬け魚・西京漬け・高級ギフト通販",                   "魚久 京漬け 評判 おすすめ, 西京漬け ギフト 通販, 高級 漬け魚 お取り寄せ","高級食材EC"),
    # ── 高級マンション（超高単価） ──────────────────────────────────────────
    ("三井不動産レジデンシャル株式会社",                    "https://www.mfr.co.jp/",                        "リスティング広告", "", "高級分譲マンション・パークホームズ・購入相談",           "三井不動産 マンション 評判 おすすめ, パークホームズ 価格 購入, 分譲 高級 マンション","高級マンション"),
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
