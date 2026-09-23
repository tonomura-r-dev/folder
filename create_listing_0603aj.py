import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03aj.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 56, 16, 16, 22, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

# ── 全社、実際の広告クリック先LP（ライブショッピング/通販CVページ）をWeb確認済み ──
DATA = [
    # ── テレビショッピング（ネットライブショッピングLP） ──
    ("株式会社QVCジャパン",                    "https://qvc.jp/catalog/weblive/index.html",     "リスティング広告", "", "テレビショッピング・ネットライブ・24時間放送",     "QVC ライブ shopping 通販, テレビショッピング オンライン, ネット 通販 便利", "テレビショッピング"),
    # ── メンズ脱毛（無料カウンセリングLP） ──
    ("メンズクリア運営会社",                   "https://mens-clear.jp/lp/",                     "リスティング広告", "", "メンズ脱毛・髭脱毛・無料カウンセリング",         "メンズ脱毛 髭 おすすめ 料金, 脱毛 男性 無料 相談, メンズエステ 比較",  "メンズ脱毛"),
    # ── オンライン診療（初回診療申し込みLP） ──
    ("クリニック・フォー株式会社",              "https://www.clinicfor.life/lp/",                "リスティング広告", "", "オンライン診療・初回診療・不眠症対応",           "オンライン診療 初回 無料, 睡眠 診察 オンライン, 医者 相談 アプリ", "オンライン診療"),
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
