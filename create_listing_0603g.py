import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03g.xlsx"
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
    # ── ICL・視力矯正（超高単価） ──────────────────────────────────────────
    ("医療法人康誠会（神戸神奈川アイクリニック）",      "https://www.kobeko.jp/",                         "リスティング広告", "", "ICL・レーシック・視力矯正手術",          "ICL 手術 費用 口コミ, レーシック 比較 おすすめ, 近視 手術 安全 クリニック",  "眼科・視力矯正"),
    ("表参道近視クリニック",                           "https://omotesando-clinic.jp/",                   "リスティング広告", "", "ICL・レーシック・近視矯正手術",          "ICL 体験 口コミ 費用, 近視 手術 東京 評判, レーシック ICL 違い",             "眼科・視力矯正"),
    # ── M&A仲介（超高単価） ──────────────────────────────────────────
    ("株式会社M&Aキャピタルパートナーズ",              "https://www.ma-cp.com/",                         "リスティング広告", "", "中小企業M&A・事業承継仲介",             "会社 売却 M&A 相談, 事業承継 後継者 M&A, 中小企業 M&A 費用 流れ",          "M&A仲介"),
    ("株式会社ストライク（M&A）",                     "https://www.strike.co.jp/",                      "リスティング広告", "", "事業承継・M&A仲介マッチング",           "M&A 仲介 おすすめ, 事業売却 相談 無料, 会社 買収 売却 相場",                 "M&A仲介"),
    ("株式会社M&A総合研究所",                         "https://masouken.com/",                          "リスティング広告", "", "AI活用型M&A・事業承継サポート",         "M&A 専門 会社 比較, 事業承継 M&A 費用 相場, 会社 売りたい 相談",            "M&A仲介"),
    # ── 不妊治療（高単価） ──────────────────────────────────────────
    ("医療法人ROPC（リプロダクションクリニック）",      "https://www.repro.or.jp/",                       "リスティング広告", "", "体外受精・不妊治療・生殖医療",           "不妊治療 クリニック 費用, 体外受精 おすすめ 病院, 不妊 検査 相談 はじめて", "不妊治療"),
    ("医療法人愛誠会（桂川レディースクリニック）",      "https://www.kc-ladies.jp/",                      "リスティング広告", "", "不妊治療・婦人科・体外受精",             "不妊治療 病院 選び方, 体外受精 成功率 高い 病院, 不妊検査 費用 流れ",       "不妊治療"),
    ("医療法人社団パリティ（加藤レディスクリニック）",  "https://www.kato-reproductive.com/",             "リスティング広告", "", "体外受精・顕微授精・不妊専門",           "体外受精 費用 東京, 不妊治療 専門 クリニック, 顕微授精 成功率 費用",        "不妊治療"),
    # ── 独立開業・フランチャイズ ──────────────────────────────────────────
    ("株式会社アントレ（アントレ）",                   "https://entre.co.jp/",                           "リスティング広告", "", "独立開業・フランチャイズ情報",           "独立 開業 フランチャイズ 比較, 起業 副業 情報 探し方, FC 加盟 費用 調べ方", "独立開業"),
    ("フランチャイズ比較ネット（株式会社リクルート系）","https://fc-hikaku.net/",                         "リスティング広告", "", "フランチャイズ加盟比較・資料請求",       "フランチャイズ 比較 一覧, FC 加盟 おすすめ, フランチャイズ 資料請求 無料",  "フランチャイズ"),
    ("一般財団法人ドリームゲート（起業支援）",          "https://www.dreamgate.gr.jp/",                   "リスティング広告", "", "起業・創業支援・専門家相談",             "起業 相談 無料, 創業 資金 補助金, 事業計画書 作り方 相談",                   "起業支援"),
    # ── 害虫・特殊清掃（高単価サービス） ──────────────────────────────────────────
    ("株式会社アース環境サービス",                    "https://www.earth-kankyo.co.jp/",                "リスティング広告", "", "害虫駆除・ゴキブリ・シロアリ対策",       "ゴキブリ 駆除 業者 費用, 害虫 駆除 おすすめ, シロアリ 予防 費用 相場",      "害虫駆除"),
    ("株式会社ハウスプロテクト（シロアリ駆除）",       "https://www.house-protect.jp/",                  "リスティング広告", "", "シロアリ・白蟻・防蟻処理",               "シロアリ 駆除 費用 業者, 白蟻 対策 おすすめ, 床下 シロアリ 点検 無料",      "シロアリ駆除"),
    ("株式会社ブルークリーン（特殊清掃）",             "https://www.blue-clean.jp/",                     "リスティング広告", "", "特殊清掃・孤独死・遺品整理",             "特殊清掃 業者 費用, 孤独死 現場 清掃, 遺品整理 部屋 片付け 費用",           "特殊清掃"),
    # ── ドローンスクール（高単価） ──────────────────────────────────────────
    ("株式会社ドローンスクールジャパン",               "https://dsj.co.jp/",                             "リスティング広告", "", "ドローン免許・資格取得スクール",         "ドローン 資格 取得 費用, ドローンスクール おすすめ 比較, 国家資格 ドローン", "ドローンスクール"),
    ("株式会社ドローンエンタープライズ",               "https://drone-enterprise.com/",                  "リスティング広告", "", "ドローン操縦士資格・産業活用",           "ドローン 操縦 資格 一覧, ドローン 免許 取り方, ドローン ビジネス 活用",     "ドローンスクール"),
    # ── 資格スクール（高単価） ──────────────────────────────────────────
    ("株式会社東京リーガルマインド（LEC）",            "https://www.lec-jp.com/",                        "リスティング広告", "", "司法書士・宅建・公務員資格スクール",     "司法書士 試験 合格 スクール, 宅建 資格 通学 費用, LEC 評判 口コミ",          "資格スクール"),
    ("TAC株式会社",                                   "https://www.tac-school.co.jp/",                  "リスティング広告", "", "公認会計士・税理士・宅建資格講座",       "公認会計士 スクール 費用, 税理士 試験 予備校, TAC 評判 合格率 口コミ",       "資格スクール"),
    # ── トランクルーム・収納 ──────────────────────────────────────────
    ("株式会社コンテナスペース（コンテナスペース）",   "https://www.container-space.jp/",                "リスティング広告", "", "トランクルーム・屋外コンテナ倉庫",       "トランクルーム 月額 安い, コンテナ 倉庫 借りる 近く, 屋外 収納 レンタル",   "トランクルーム"),
    # ── 資格専門学校（高単価） ──────────────────────────────────────────
    ("学校法人大原学園（大原簿記・法律専門学校）",     "https://www.o-hara.ac.jp/",                      "リスティング広告", "", "簿記・法律・公務員資格の専門学校",       "簿記 専門学校 おすすめ, 大原 評判 費用, 公務員 学校 通学 費用 資格",         "資格専門学校"),
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
