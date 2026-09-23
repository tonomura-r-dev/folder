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


# 第154弾（2026-07-16f）不動産・リノベーション・注文住宅 8社・LP URLは相談/資料請求ページ

data_0716f = [
    ["スクールバス空間設計株式会社", "https://www.school-bus.jp/contact/", "06-6205-3733", "マンション・戸建てリノベーション・オーダー家具", "リノベーション 大阪, 中古マンション リノベーション 神戸, 戸建てリノベーション 京都", "リノベーション"],
    ["株式会社アネストワン", "https://www.anestone.com/contact/", "052-777-2441", "マンション・戸建てリノベーション・新築注文住宅", "リノベーション 名古屋, 中古マンション リノベーション 愛知, 注文住宅 名古屋", "リノベーション・注文住宅"],
    ["ヤマト住建株式会社", "https://www.yamatojk.co.jp/starter_form", "", "高性能注文住宅（ZERO-CUBEシリーズ等）", "注文住宅 神戸, 高気密高断熱 注文住宅, 注文住宅 世界基準", "注文住宅"],
    ["富士住建株式会社", "https://www2.fujijuken.co.jp/contact", "048-778-3310", "完全フル装備の家（注文住宅）", "注文住宅 埼玉, 完全フル装備の家, 注文住宅 ローコスト 関東", "注文住宅"],
    ["株式会社アイ工務店", "https://www.ai-koumuten.co.jp/contact/", "06-6227-8288", "自由設計の高品質注文住宅", "注文住宅 大阪, 自由設計 注文住宅, 注文住宅 適正価格", "注文住宅"],
    ["オークラヤ住宅株式会社", "https://www.ohkuraya.co.jp/contact/sell-index", "03-3262-2684", "中古マンション売買仲介・住み替え仲介", "マンション売却 東京, 中古マンション購入 千代田区, マンション住み替え", "不動産仲介"],
    ["株式会社シアーズホーム", "https://www.justhome.jp/mh_form/", "096-370-0007", "完全注文住宅・企画型住宅（ジャストホーム）", "注文住宅 熊本, 注文住宅 福岡, ローコスト住宅 九州", "注文住宅FC"],
    ["株式会社フレッシュハウス", "https://freshhouse.co.jp/contact/consultation", "", "リフォーム・リノベーション（戸建て・マンション）", "リフォーム 横浜, 中古住宅 リノベーション 神奈川, 外壁塗装リフォーム 首都圏", "リフォーム・リノベーション"],
]

base = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー"
make_xlsx(f"{base}\\架電リスト_Meta_2026-07-16f.xlsx", data_0716f)
