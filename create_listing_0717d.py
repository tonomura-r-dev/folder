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
# ここからはMeta広告ライブラリ(facebook.com/ads/library)でキーワード検索し、
# 「実際に配信中」の広告クリエイティブを目視確認できた企業のみを採用（推測なし）。
# 全社、運営会社名・電話番号は公式サイトでWebSearch裏取り済み。使用済み企業リストとGrep照合し重複なしを確認済み。

data_0717d = [
    ["株式会社ブラスト", "https://kogaokyosei.com/", "0120-980-439", "小顔矯正・骨盤矯正・美容矯正サロン（美.design、全国38店舗）", "小顔矯正／頭蓋骨矯正 サロン", "美容矯正サロンFC"],
    ["株式会社さつま骨格矯正", "https://beauty.hotpepper.jp/kr/slnH000319620/", "03-5468-0910", "顎顔面矯正・小顔矯正整骨院（さつま骨格矯正）", "小顔矯正／顎顔面矯正 整骨院", "美容矯正サロン"],
    ["株式会社Racbaki", "https://racbaki.com/", "", "骨格矯正靴D2C（Racbaki）", "骨格矯正靴／脚痩せ サンダル", "美容雑貨D2C"],
    ["株式会社羅針", "https://www.kaitori-ginza.com/", "03-6252-1515", "高級ブランド腕時計買取（GINZA RASIN）", "腕時計 買取／ロレックス 売る", "時計買取"],
    ["株式会社DT", "https://estime.co.jp/", "03-4400-7963", "高級ブランド品・時計買取（ESTIME）", "時計 買取／リシャールミル 売る", "ブランド品買取"],
]

base = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー"
make_xlsx(f"{base}\\架電リスト_Meta_2026-07-17d.xlsx", data_0717d)
