import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-26c.xlsx"

HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]

# LP URLはすべて /lp/ を含む実在確認済みページ
DATA = [
    (
        "株式会社GAKKEN CC（CCレッスン）",
        "https://www.cclesson.com/lp/lp-a8",
        "リスティング広告", "",
        "オンライン中国語教室（ネイティブ講師）",
        "中国語 オンライン レッスン, 中国語 会話 個人 スクール, オンライン中国語 無料体験",
        "英会話・語学・資格スクール"
    ),
    (
        "株式会社JOY&LIFE Creation（JOYミュージックスクール）",
        "https://joy-music.jp/vocal_lp/",
        "リスティング広告", "",
        "ボイトレ・音楽スクール（東京・大阪6校+オンライン）",
        "ボイトレ 教室 渋谷, 歌 上手くなる スクール, ボイストレーニング 社会人",
        "教育・スキルアップ"
    ),
    (
        "イエローハットグループ（cyma -サイマ-）",
        "https://cyclemarket.jp/shop/y_s/lp",
        "リスティング広告", "",
        "自転車・電動アシスト自転車 通販（組立整備済み配送）",
        "電動自転車 通販 安い, 自転車 ネット 整備済み, クロスバイク 通販 送料無料",
        "その他（EC・通販）"
    ),
    (
        "株式会社ソーシャルテック（チャップアップ CHAPUP）",
        "https://chapup.jp/shopping/lp.php",
        "リスティング広告", "",
        "育毛剤・育毛シャンプー（Web売上No.1）",
        "育毛剤 おすすめ 男性, 薬用 育毛 ローション 通販, 抜け毛 薄毛 対策 サプリ",
        "化粧品・健康食品通販D2C"
    ),
    (
        "株式会社資格スクエア（資格スクエア）",
        "https://www.shikaku-square.com/lp/sale_pdf_16",
        "リスティング広告", "",
        "司法試験・予備試験・行政書士・宅建 オンライン通信講座",
        "司法試験予備試験 通信講座, 行政書士 オンライン スクール, 宅建 最短合格 通信",
        "教育・スキルアップ"
    ),
    (
        "株式会社アビバ（パソコン教室アビバ AVIVA）",
        "https://www.aviva.co.jp/lp/chatgpt/",
        "リスティング広告", "",
        "AI・PC・ITスキルスクール（通学・オンライン、全国展開）",
        "パソコン教室 AI ChatGPT, ITパスポート 取得 通学, Excel スキルアップ 資格",
        "教育・スキルアップ"
    ),
    (
        "株式会社マウスコンピューター（マウスコンピューター）",
        "https://mouse-jp.co.jp/lp/high_spec_notepc",
        "リスティング広告", "",
        "BTOパソコン・ノートPC通販（3年保証・国内生産）",
        "BTOパソコン 通販 高性能 安い, ゲーミングPC 予算 おすすめ, ノートPC カスタマイズ 注文",
        "IT・デジタルサービス"
    ),
    (
        "GMOインターネット株式会社（GMOとくとくBB）",
        "https://gmobb.jp/lp/gmohikari/",
        "リスティング広告", "",
        "光回線プロバイダー（GMOとくとくBB光）",
        "光回線 乗り換え キャッシュバック, GMO光 申し込み 特典, ドコモ光 プロバイダ 安い",
        "格安スマホ・通信"
    ),
    (
        "株式会社ラクリ（ラクリ）",
        "https://www.lacuri.jp/s/lp/lp_202005/",
        "リスティング広告", "",
        "宅配クリーニング・衣類保管サービス",
        "宅配クリーニング 安い 口コミ, 衣類 クリーニング 保管, コート クリーニング 宅配 全国",
        "家事代行・クリーニング"
    ),
    (
        "株式会社キタムラ（カメラのキタムラ）",
        "https://lp.kitamura.jp/kit_takamatsu-minami/",
        "リスティング広告", "",
        "カメラ・レンズ販売・買取・フォトブック（全国800店）",
        "カメラ 買取 高価 近く, 中古カメラ レンズ 販売, フォトブック 作成 安い",
        "その他（EC・通販）"
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
