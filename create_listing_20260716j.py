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


# 第158弾（2026-07-16j）美容系サロン・クリニック拡充第2弾 11社・LP URLは予約ページ

data_0716j = [
    ["株式会社はあと", "https://heartnail.jp/", "075-311-3130", "定額制ネイルサロン（はあとねいる）", "ネイルサロン フランチャイズ, 定額制ネイル, ネイルサロン 独立開業", "ネイルサロンFC"],
    ["株式会社Blanc", "https://www.blanc-lash.com/", "06-6886-7800", "まつげエクステ専門サロン（Blanc）", "まつげエクステ 大阪, マツエクサロン, まつ毛エクステ フランチャイズ", "まつげエクステサロンFC"],
    ["株式会社シンメトリー", "https://symmetry.jp/", "03-3541-1395", "小顔矯正・美容整体サロン（Symmetry）", "小顔矯正, 骨格矯正サロン, 小顔矯正 東京", "小顔矯正サロンFC"],
    ["株式会社DRN", "https://dr-nail.jp/otoiwase/", "044-201-2918", "巻き爪・魚の目フットケア専門店（ドクターネイル爪革命）", "巻き爪 治療, フットケアサロン, 魚の目 ケア", "巻き爪フットケアサロンFC"],
    ["株式会社ロレインブロウ", "https://lorraine-brow.co.jp/", "06-4400-1749", "まつ毛パーマ×眉毛サロン（ロレインブロウ）", "まつげパーマ サロン, 眉毛サロン, アイブロウ専門店", "まつげパーマ・眉毛サロンFC"],
    ["株式会社GROWBING", "https://rank-up-mens.com/", "03-5843-0521", "メンズ専門眉毛サロン（眉毛の王様）", "メンズ眉毛サロン, 眉毛 メンズ, 男性 眉毛 整形", "メンズ眉毛サロンFC"],
    ["一般社団法人ママリュクス", "https://mamaluxe.jp/", "0797-23-3177", "マタニティ・産後骨盤矯正専門サロン（mamaluxe）", "産後骨盤矯正, マタニティ整体, 産後ケア サロン", "産後骨盤矯正サロン"],
    ["株式会社SPEED spring", "https://bupura.jp/", "", "小顔・痩身サロン（BUPURA）", "小顔サロン, 痩身エステ, 小顔矯正 サブスク", "小顔痩身サロンFC"],
    ["ecxia株式会社", "https://ecxia-inc.com/whitening/", "", "定額制セルフホワイトニングサロン（エクシアホワイトニング）", "セルフホワイトニング, 歯 ホワイトニング サロン, ホワイトニング 通い放題", "ホワイトニングサロンFC"],
    ["株式会社Zina", "https://www.zina-hair.com/", "", "髪質改善トリートメント美容室（Zina）", "髪質改善 美容室, 縮毛矯正 東京, 艶髪トリートメント", "髪質改善美容室FC"],
    ["株式会社WEBSTYLE", "https://dears-salon.com/", "", "女性専門髪質改善美容室（Dears）", "髪質改善 サロン, 縮毛矯正 女性専門, 艶髪美容室", "髪質改善美容室FC"],
]

base = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー"
make_xlsx(f"{base}\\架電リスト_Meta_2026-07-16j.xlsx", data_0716j)
