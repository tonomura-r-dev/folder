import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03e.xlsx"
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
    # ── 不動産売買仲介（高単価） ──────────────────────────────────────────
    ("野村不動産ソリューションズ株式会社（ノムコム）",  "https://www.nomu.com/",                         "リスティング広告", "", "マンション・戸建て売買仲介",             "マンション 売却 査定 無料, 不動産 仲介 おすすめ, 住み替え 売却 購入",        "不動産仲介"),
    ("住友不動産販売株式会社（スミフル）",              "https://www.stepon.co.jp/",                      "リスティング広告", "", "マンション・一戸建て売買仲介",           "マンション 売却 高額 査定, 住友不動産販売 評判, 不動産 売買 仲介 安心",      "不動産仲介"),
    ("三井不動産リアルティ株式会社（三井のリハウス）",  "https://www.rehouse.co.jp/",                     "リスティング広告", "", "不動産売買・賃貸仲介（全国）",           "不動産 売却 相談, 三井のリハウス 評判 査定, 一戸建て 購入 仲介 おすすめ",  "不動産仲介"),
    # ── 補聴器（超高単価） ──────────────────────────────────────────
    ("リオン株式会社（リオネット補聴器）",              "https://www.rionet.jp/",                         "リスティング広告", "", "高性能補聴器・聴力ケア製品",             "補聴器 おすすめ 種類, リオネット 評判 価格, 耳 聞こえ にくい 補聴器",        "補聴器"),
    ("デマントジャパン株式会社（オーティコン補聴器）",  "https://www.oticon.co.jp/",                      "リスティング広告", "", "デジタル補聴器・専門相談",               "補聴器 高性能 おすすめ, オーティコン 試聴 費用, 補聴器 自分に合う 選び方", "補聴器"),
    ("WS Audiology Japan株式会社（シグニア補聴器）",   "https://www.signia.net/ja-jp/",                  "リスティング広告", "", "シグニア補聴器・充電式補聴器",           "補聴器 充電式 使いやすい, シグニア 補聴器 評判, 高齢者 補聴器 軽い",        "補聴器"),
    # ── 注文住宅・工務店（高単価） ──────────────────────────────────────────
    ("株式会社サイエンスホーム",                       "https://www.sciencehome.jp/",                    "リスティング広告", "", "木の家・国産ヒノキ注文住宅",             "注文住宅 木の家 おすすめ, 自然素材 家 工務店, 国産 ヒノキ 家 費用",         "注文住宅"),
    ("株式会社アルネットホーム",                       "https://www.arnethome.co.jp/",                   "リスティング広告", "", "規格型注文住宅・ローコスト住宅",         "注文住宅 コスト 安い 規格型, アルネットホーム 評判, 1000万円台 家 建てる", "注文住宅"),
    ("株式会社ハーバーハウス",                         "https://www.harberhouse.co.jp/",                 "リスティング広告", "", "注文住宅・分譲住宅（新潟・北陸）",       "注文住宅 新潟 おすすめ, ハーバーハウス 評判 坪単価, 家 新築 北陸 工務店",  "注文住宅"),
    # ── FX・外国為替（高単価） ──────────────────────────────────────────
    ("株式会社マネーパートナーズ",                     "https://www.moneypartners.co.jp/",               "リスティング広告", "", "FX取引・外国為替・キャンペーン",         "FX 口座開設 キャンペーン, 外国為替 取引 始め方, スプレッド 狭い FX会社",   "FX"),
    ("SBIFXトレード株式会社",                         "https://www.sbifxt.co.jp/",                      "リスティング広告", "", "FX・外国為替スワップ取引",               "FX スワップ 高い 会社, 外為 積立 FX 始め方, SBI FX 評判 スプレッド",      "FX"),
    ("株式会社外為どっとコム（外貨ex）",               "https://www.gaitame.com/fx/",                    "リスティング広告", "", "FX・外国為替取引・自動売買",             "FX 自動売買 おすすめ, 外為どっとコム 評判 口コミ, FX 少額 初心者 始め方", "FX"),
    # ── 高額エステ・美容 ──────────────────────────────────────────
    ("株式会社不二ビューティ（たかの友梨）",            "https://www.takano-yuri.com/",                   "リスティング広告", "", "高級エステ・フェイシャル・全身美容",     "エステ 高級 おすすめ, たかの友梨 料金 口コミ, 痩身 フェイシャル コース", "エステサロン"),
    ("株式会社ヴィーナスウォーカー",                   "https://www.venus-walker.jp/",                   "リスティング広告", "", "痩身・エステ・ボディメイク",             "痩身エステ 効果 コース, ヴィーナスウォーカー 評判 料金, エステ 体験 無料", "エステサロン"),
    # ── 法律・弁護士（高単価） ──────────────────────────────────────────
    ("泉総合法律事務所",                               "https://izumisogo-law.jp/",                      "リスティング広告", "", "債務整理・過払い金・自己破産",           "債務整理 弁護士 相談 無料, 過払い金 請求 費用, 借金 解決 弁護士 おすすめ", "法律"),
    ("みらい総合法律事務所",                           "https://www.mirai-sogo.com/",                    "リスティング広告", "", "相続・離婚・借金の弁護士相談",           "相続 弁護士 相談 費用, 離婚 弁護士 探し方, 借金 問題 無料 法律相談",       "法律"),
    # ── 分譲マンション・高額不動産 ──────────────────────────────────────────
    ("三菱地所レジデンス株式会社",                     "https://www.mec-r.com/",                         "リスティング広告", "", "分譲マンション・ザ・パークハウス",       "分譲マンション 購入 おすすめ, 三菱地所 マンション 評判, 新築 マンション 資産価値","分譲マンション"),
    ("東急不動産株式会社（ブランズ）",                 "https://www.tokyu-land.co.jp/mansion/",          "リスティング広告", "", "分譲マンション・ブランズシリーズ",       "ブランズ マンション 評判, 東急不動産 物件 購入, 新築 マンション 東京 資産", "分譲マンション"),
    # ── 証券・投資（高単価） ──────────────────────────────────────────
    ("株式会社インヴァスト証券（くりっく365）",        "https://www.invast.jp/",                         "リスティング広告", "", "取引所FX・くりっく365・株式CFD",        "取引所FX 安全 信頼, くりっく365 口座開設, FX 会社 倒産 リスク 低い",       "FX・証券"),
    ("株式会社ネクシィーズ・トレード（NEXTIES）",      "https://www.nexties.jp/",                        "リスティング広告", "", "FX自動売買・シストレ・EA運用",           "FX 自動売買 ツール おすすめ, シストレ FX 始め方, EA 運用 おすすめ",         "FX・証券"),
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
