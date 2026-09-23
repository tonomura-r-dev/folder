import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-27a.xlsx"

HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]

# LP URLはすべて /lp/ を含む実在確認済みページ
DATA = [
    (
        "株式会社ネクスト（保険マンモス）",
        "https://www.hoken-mammoth.jp/lp/survey_jmty/",
        "リスティング広告", "",
        "保険相談・見直しサービス（FPによる無料相談）",
        "保険 相談 無料 FP おすすめ, 生命保険 見直し 比較, 保険 見直し 主婦",
        "保険・金融・証券・カード"
    ),
    (
        "ロードスターキャピタル株式会社（OwnersBook）",
        "https://www.ownersbook.jp/lp/sociallending//",
        "リスティング広告", "",
        "不動産特化型ソーシャルレンディング（1万円〜）",
        "不動産投資 ソーシャルレンディング 少額, 不動産クラウドファンディング 利回り, 1万円 不動産 投資 始め方",
        "不動産・賃貸・リノベーション・投資"
    ),
    (
        "キラメックス株式会社（TechAcademyジュニア）",
        "https://junior.techacademy.jp/lp/curriculum_trial",
        "リスティング広告", "",
        "小中学生向けプログラミング・デジタル学習（オンライン）",
        "子供 プログラミング オンライン 無料体験, 小学生 コーディング 習い事, ジュニア IT 教室 自宅",
        "教育・スキルアップ"
    ),
    (
        "株式会社バイクランド",
        "https://www.bike-kaitori.com/lp/sp/",
        "リスティング広告", "",
        "バイク宅配買取・高価買取専門サービス",
        "バイク 買取 高い 業者, 原付 バイク 売る 宅配, バイク 査定 無料 自宅",
        "車・カーシェア・バイク"
    ),
    (
        "出光興産株式会社（出光ポチモ）",
        "https://pochi-mo.com/lp/cp_202306_1",
        "リスティング広告", "",
        "新車カーリース・定額マイカー（頭金なし・全込み）",
        "カーリース 新車 頭金なし, 車 月額 定額 全込み, マイカーリース おすすめ 審査",
        "車・カーシェア・バイク"
    ),
    (
        "株式会社リフォームガイド",
        "https://www.reform-guide.jp/lp/gaihekitosou/",
        "リスティング広告", "",
        "外壁塗装・リフォーム業者の一括比較・見積もりサービス",
        "外壁塗装 業者 比較 一括, リフォーム 相見積もり 無料, 屋根 塗装 費用 相場 地域",
        "不動産・賃貸・リノベーション・投資"
    ),
    (
        "株式会社ナイチンゲール",
        "https://www.nightingale-web.com/lp/",
        "リスティング広告", "",
        "看護師・准看護師専門転職エージェント",
        "看護師 転職 サイト おすすめ, 准看護師 求人 夜勤なし, 看護師 高給 正社員 紹介",
        "転職・就職・人材派遣"
    ),
    (
        "つばきグループ（医療法人社団つばき会）",
        "https://tsubaki-grp.com/lp/depi/free_select/",
        "リスティング広告", "",
        "全身脱毛・美容医療クリニック（都度払い対応）",
        "全身脱毛 安い 都度払い, 医療脱毛 クリニック 比較 東京, 脱毛 永久 コース 料金",
        "美容クリニック・脱毛・エステ"
    ),
    (
        "株式会社ハピすむ",
        "https://hapisumu.jp/lp/reform_wall_1sp",
        "リスティング広告", "",
        "外壁塗装・水回り・リフォームの業者比較・一括見積もり",
        "外壁塗装 一括 見積もり 無料, 水回り リフォーム 費用 比較, リフォーム 業者 口コミ 地域",
        "不動産・賃貸・リノベーション・投資"
    ),
    (
        "株式会社ベストライフ",
        "https://www.e-kaitori.jp/shop/lp/gold/",
        "リスティング広告", "",
        "ブランド品・金・プラチナ・時計の宅配買取",
        "ブランド 買取 宅配 高価, 金 買取 相場 高い 業者, 時計 ブランド 売る 宅配 査定",
        "買取・リユース・フリマ"
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
