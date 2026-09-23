import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-26b.xlsx"

HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]

# LP URLはすべて /lp/ を含む実在確認済みページ
DATA = [
    (
        "株式会社ルラ（ルラ美容クリニック）",
        "https://www.lula-beauty.jp/lp/",
        "リスティング広告", "",
        "医療脱毛・美容整形（多拠点クリニック）",
        "医療脱毛 全国 クリニック, 美容整形 相談 無料, 脱毛 全身 安い",
        "美容クリニック・脱毛・エステ"
    ),
    (
        "株式会社K Village Tokyo（K Village 韓国語教室）",
        "https://kvillage.jp/lp01/",
        "リスティング広告", "",
        "韓国語教室（全国25校＋オンライン）",
        "韓国語教室 全国, 韓国語 スクール 通い, 韓国語 上達 最短",
        "英会話・語学・資格スクール"
    ),
    (
        "株式会社ブランドゥール（Branduru）",
        "https://branduru.jp/lp.php",
        "リスティング広告", "",
        "ブランド品宅配買取（洋服・バッグ・小物）",
        "ブランド品 宅配買取 無料, 洋服 バッグ 買取, ブランド 高価買取 宅配",
        "買取・リユース・フリマ"
    ),
    (
        "合同会社ドリーマーズギルド（コードアドベンチャー）",
        "https://codeadventure.jp/lp/fujinomiya/",
        "リスティング広告", "",
        "子どもプログラミング教室（全国100教室以上）",
        "子ども プログラミング教室 近く, プログラミング 習い事 小学生, コーディング スクール 子供",
        "教育・スキルアップ"
    ),
    (
        "株式会社小さな結婚式（小さな結婚式）",
        "https://www.petitwedding.com/LP/wedding_photo/",
        "リスティング広告", "",
        "少人数婚・フォトウェディング（全国展開）",
        "少人数婚 費用 格安, フォトウェディング 全国 安い, 結婚式 小規模 おすすめ",
        "ウエディング・葬儀"
    ),
    (
        "株式会社ファーストステージ（ファーストステージ）",
        "https://www.firststage.co.jp/photostudio/lp/bridal/",
        "リスティング広告", "",
        "ブライダル写真スタジオ（関西多拠点）",
        "前撮り スタジオ 関西 安い, ウェディングフォト 大阪, 和装 前撮り 費用",
        "ウエディング・葬儀"
    ),
    (
        "株式会社ハンコヤドットコム（Hankoya.com）",
        "https://www.hankoya.com/lp/gift/shop/",
        "リスティング広告", "",
        "印鑑・はんこ通販（国内最大級）",
        "印鑑 通販 即日発送, はんこ 作成 安い, 実印 法人印 注文",
        "その他（EC・通販）"
    ),
    (
        "たんす屋株式会社（たんす屋）",
        "https://tansuya.jp/lp-202512-asakusa/",
        "リスティング広告", "",
        "リユース着物・和装品の販売・買取（全国43店）",
        "リサイクル着物 購入 安い, 中古着物 通販, 着物 買取 宅配 無料",
        "買取・リユース・フリマ"
    ),
    (
        "株式会社留学ジャーナル（留学ジャーナル）",
        "https://www.ryugaku.co.jp/lp/internship/",
        "リスティング広告", "",
        "海外留学・ワーキングホリデー手配（KDDIグループ）",
        "海外留学 エージェント 無料相談, ワーキングホリデー サポート 費用, 留学 手続き 代行",
        "英会話・語学・資格スクール"
    ),
    (
        "日本PCサービス株式会社（PCホスピタル）",
        "https://www.4900.co.jp/lp/mass-retailer/",
        "リスティング広告", "0120-49-0041",
        "パソコン修理・Mac修理・データ復旧（全国400拠点）",
        "パソコン 修理 出張 即日, データ復旧 安い 早い, Mac 修理 店 近く",
        "IT・デジタルサービス"
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
