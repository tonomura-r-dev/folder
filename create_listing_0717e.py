import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def make_xlsx(filename, data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "架電リスト"

    headers = ["企業名", "LP URL", "電話番号", "商材", "検索KW", "業界"]

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="メイリオ", bold=True, color="FFFFFF", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for row_idx, row in enumerate(data, 2):
        fill_color = "F2F7FC" if row_idx % 2 == 0 else "FFFFFF"
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = row_fill
            cell.font = Font(name="メイリオ", size=9)
            cell.alignment = left
            cell.border = border

    col_widths = [30, 45, 16, 20, 30, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(filename)
    print(f"保存完了: {filename}  件数: {len(data)}")


# 列順: 企業名 | LP URL | 電話番号 | 商材 | 検索KW | 業界
# Meta広告ライブラリでキーワード検索し、実配信中の広告を確認できた企業のみ採用。
# 全社、公式サイトで社名電話番号確認済み・使用済み企業リストとGrep照合し重複なしを確認済み。

data_0717e = [
    ["医療法人社団上伸会", "https://www.ueno.co.jp/", "", "包茎手術・AGA治療（東京上野クリニック、全国15医院）", "AGA 治療／包茎 手術", "メンズ美容クリニック"],
    ["株式会社あい・グループ", "https://www.ai-medical.co.jp/", "06-6775-0222", "鍼灸整骨院（あい鍼灸院・接骨院、全国40店舗以上）", "整体 肩こり／鍼灸院 近く", "鍼灸整骨院FC"],
    ["株式会社ジェーアンビシャス", "https://js-seitai.com/", "03-6908-8838", "整体院（J'sメディカル整体院、全国27店舗）", "整体 骨盤矯正／腰痛 整体院", "整体院FC"],
    ["株式会社ビー・ファクトリー", "https://www.bee-music.jp/vocal/", "03-6824-5169", "ボイストレーニング・音楽教室（Beeミュージックスクール）", "ボイトレ 教室／ボーカルレッスン", "音楽教室"],
]

base = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー"
make_xlsx(f"{base}\\架電リスト_Meta_2026-07-17e.xlsx", data_0717e)
