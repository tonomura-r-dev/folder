import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-26.xlsx"

HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]

# LP URLはすべて /lp/ を含む実在確認済みページ
DATA = [
    (
        "SEモバイル・アンド・オンライン株式会社（スマリッジ）",
        "https://s-marriage.jp/promotion/11/",
        "リスティング広告", "",
        "オンライン結婚相談所",
        "オンライン 結婚相談所, 婚活 安い 月額, 婚活 真剣 サポート",
        "婚活・マッチングアプリ"
    ),
    (
        "SBIペット少額短期保険株式会社",
        "https://www.sbipet-ssi.co.jp/lp/",
        "リスティング広告", "0120-63-1234",
        "ペット保険（犬・猫）",
        "ペット保険 安い, 犬 保険 通院, ペット保険 比較 おすすめ",
        "ペット・ペット保険"
    ),
    (
        "株式会社デイトラ（デイトラ）",
        "https://dailytrial.net/reskilling_lp/",
        "リスティング広告", "",
        "リスキリング・ITスキルスクール（オンライン）",
        "リスキリング IT転職, プログラミング 副業 在宅, webデザイン 独学",
        "教育・スキルアップ"
    ),
    (
        "株式会社Wellness（あんしん漢方）",
        "https://www.kamposupport.com/anshin1.0/lp/",
        "リスティング広告", "",
        "AI×専門家 オーダーメイド漢方定期便",
        "漢方 オンライン 相談, 漢方薬 定期便 通販, 漢方 サブスク",
        "化粧品・健康食品通販D2C"
    ),
    (
        "株式会社ロイブ（ホットヨガスタジオ ロイブ）",
        "https://www.hotyoga-loive.com/lp/easy-hotyoga/",
        "リスティング広告", "0570-666-969",
        "ホットヨガスタジオ（0円体験）",
        "ホットヨガ 体験 0円, ヨガスタジオ 月額, ホットヨガ 女性 おすすめ",
        "フィットネス・スポーツ・ゴルフ"
    ),
    (
        "万田発酵株式会社（万田酵素）",
        "https://www.mandahakko.com/contents/ec_trial_1000_p_pluson_lp.html",
        "リスティング広告", "",
        "発酵酵素健康食品（お試しセット）",
        "万田酵素 お試し, 酵素 健康食品 定期購入, 発酵食品 サプリ",
        "化粧品・健康食品通販D2C"
    ),
    (
        "株式会社subsclife（subsclife）",
        "https://subsclife.com/lp/lp001/",
        "リスティング広告", "",
        "家具・家電サブスクリプション（月額）",
        "家具 サブスク 月額, ソファ レンタル 月額, 家具 定額 初期費用なし",
        "その他（EC・通販）"
    ),
    (
        "株式会社クロスエッジ（Dr.つるかめキッチン）",
        "https://tsurukame-kitchen.com/lp/002/",
        "リスティング広告", "",
        "制限食・健康宅配弁当",
        "カロリー制限 食事 宅配, 制限食 弁当 通販, 医師監修 宅配弁当",
        "食品EC・宅配弁当・ミールキット"
    ),
    (
        "株式会社SOYOKAZE（食のそよ風）",
        "https://shokunosoyokaze.com/lp/lp01/",
        "リスティング広告", "0120-253-831",
        "高齢者向け冷凍宅配弁当",
        "高齢者 宅配弁当, 介護食 冷凍 宅配, シニア 食事 宅配 毎日",
        "食品EC・宅配弁当・ミールキット"
    ),
    (
        "Photoback株式会社（Photoback）",
        "https://www.photoback.jp/lp/photoback_beginner",
        "リスティング広告", "",
        "フォトブック・写真アルバム作成",
        "フォトブック 高品質 おしゃれ, 写真アルバム 作成 通販, フォトブック 無料クーポン",
        "その他（EC・通販）"
    ),
    (
        "株式会社キカガク（キカガク）",
        "https://lp-school.kikagaku.ai/",
        "リスティング広告", "",
        "AIスクール・データサイエンス長期コース",
        "AIスクール 社会人, データサイエンス 学習, AI 資格 取得",
        "教育・スキルアップ"
    ),
    (
        "医療法人社団クリニクフォア（クリニックフォア）",
        "https://www.clinicfor.life/lp/online-insurance/",
        "リスティング広告", "",
        "オンライン診療（皮膚科・内科・アレルギー科）",
        "オンライン診療 予約, 皮膚科 オンライン 保険適用, 内科 オンライン 処方",
        "美容クリニック・脱毛・エステ"
    ),
]

def make_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "架電リスト"

    header_font   = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
    header_fill   = PatternFill("solid", fgColor="1F4E79")
    header_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)

    fill_odd  = PatternFill("solid", fgColor="F2F7FC")
    fill_even = PatternFill("solid", fgColor="FFFFFF")
    data_font = Font(name="メイリオ", size=9)
    data_align = Alignment(vertical="center", wrap_text=False)
    url_font  = Font(name="メイリオ", size=9, color="0563C1", underline="single")

    thin = Side(style="thin", color="D0D7DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, name in enumerate(HEADER, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    for row_idx, row in enumerate(DATA, start=2):
        fill = fill_odd if row_idx % 2 == 0 else fill_even
        for col, val in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.fill = fill
            cell.border = border
            cell.alignment = data_align
            if col == 2 and val:
                cell.hyperlink = val
                cell.font = url_font
            else:
                cell.font = data_font

    for col, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[1].height = 22
    for r in range(2, len(DATA) + 2):
        ws.row_dimensions[r].height = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}1"

    wb.save(OUTPUT_PATH)
    print(f"保存完了: {OUTPUT_PATH}")
    print(f"件数: {len(DATA)}社")

if __name__ == "__main__":
    make_xlsx()
