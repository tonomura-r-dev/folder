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
    print(f"保存完了: {filename}  行数: {len(data)}")


# 第156弾（2026-07-16h）婚活・結婚相談所FC 7社・LP URLは予約/申込ページ

data_0716h = [
    ["株式会社トータルマリアージュサポート", "https://ssl.total-marriage.com/reservation/", "050-1744-5777", "結婚相談所（フィオーレ）", "結婚相談所 大阪, 結婚相談所 東京, 婚活 30代", "結婚相談所チェーン（直営16店舗・FC連盟運営）"],
    ["合同会社Play Works", "https://playworks.co.jp/", "", "街コン・婚活パーティー（名古屋東海街コン）", "街コン 名古屋, 婚活パーティー 東海, 街コン 愛知", "婚活パーティー運営"],
    ["有限会社アーバンワールド", "https://machicon.um-club.jp/", "", "婚活パーティー（アーバンマリッジ）", "婚活パーティー 広島, 街コン 中国地方, 婚活 広島", "婚活パーティー運営"],
    ["株式会社ブライダルドリーム", "https://duo-bridal.com/contact/", "", "結婚相談所・婚活パーティー（DUO BRIDAL デュオブライダル）", "結婚相談所 新宿, 婚活パーティー 横浜, 結婚相談所 名古屋", "結婚相談所・婚活パーティー運営"],
    ["株式会社ウィルコミュニケーションズ", "https://whitekey.co.jp/schedule/", "011-252-7889", "婚活パーティー（ホワイトキー）", "婚活パーティー 東京, お見合いパーティー 大阪, 婚活パーティー 名古屋", "婚活パーティー運営"],
    ["アイマップス株式会社", "https://konkatsu-nana.com/", "03-5566-5670", "婚活パーティー（婚活NANA）", "婚活パーティー 福岡, 婚活パーティー 大阪, 婚活パーティー 浜松", "婚活パーティー運営"],
    ["株式会社Realing", "https://realing-agent.com/contact/", "03-6805-1992", "結婚相談所（リアリングエージェント）・街コン（イベントコンタクト）", "結婚相談所 渋谷, 街コン 東京, 婚活パーティー 神奈川", "結婚相談所・街コン運営"],
]

base = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー"
make_xlsx(f"{base}\\架電リスト_Meta_2026-07-16h.xlsx", data_0716h)
