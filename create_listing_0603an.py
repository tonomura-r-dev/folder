import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03an.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 56, 16, 16, 22, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

# ── 全社、WebSearchで現在出稿中の広告LP（/lp/ /campaign/ 等）を実確認済み ──
DATA = [
    # ── 婚活写真（撮影予約LP） ──
    ("株式会社スタジオインディ",       "https://konkatsu.studioindi.jp/",                  "リスティング広告", "", "婚活・お見合い写真撮影・全額返金保証", "婚活写真 スタジオ おすすめ, お見合い 写真 撮影 東京, 婚活 写真 安い",   "フォトスタジオ"),
    # ── レンタルサーバー（公式LP） ──
    ("エックスサーバー株式会社",       "https://www.xserver.ne.jp/lp/service02/",          "リスティング広告", "", "高性能レンタルサーバー・WordPress対応", "レンタルサーバー おすすめ 比較, サーバー 安い 高速, エックスサーバー 評判", "レンタルサーバー"),
    # ── 保育ICT（キャンペーンLP） ──
    ("株式会社コドモン",               "https://www.codmon.com/campaign/",                 "リスティング広告", "", "保育・幼稚園ICTシステム・無料トライアル", "保育 ICT システム おすすめ, 保育園 DX ツール 比較, コドモン 評判",    "EdTech・保育ICT"),
    # ── ASP（広告主向け登録LP） ──
    ("バリューコマース株式会社",       "https://mer.valuecommerce.ne.jp/online-signup/vc_ec01_form.php", "リスティング広告", "", "アフィリエイトASP・広告主登録",       "アフィリエイト ASP 広告主 登録, アフィリ 集客 EC 費用, バリューコマース 申し込み", "デジタル広告"),
    # ── 銀行カードローン（申し込みLP） ──
    ("オリックス銀行株式会社",         "https://www.orixbank.co.jp/personal/cardloan/",    "リスティング広告", "", "銀行カードローン・低金利・最短翌営業日",  "カードローン おすすめ 銀行, 低金利 ローン 比較, オリックス銀行 審査",  "金融・カードローン"),
    # ── ペット葬儀（24時間受付LP） ──
    ("ペットの旅立ち（株式会社旅立ち）", "https://pet-tabi.jp/",                           "リスティング広告", "", "ペット訪問火葬・24時間受付・葬儀一式", "ペット 葬儀 費用, ペット 火葬 訪問 24時間, ペット 旅立ち 申し込み", "ペット葬儀"),
    # ── 太陽光発電（初期費用0円LP） ──
    ("東京電力ホールディングス株式会社", "https://www.tepco-ht.co.jp/enekari/lp/zero/",   "リスティング広告", "", "太陽光発電・蓄電池・初期費用0円設置", "太陽光発電 初期費用 0円, ソーラーパネル 設置 無料, 蓄電池 セット 補助金", "再生可能エネルギー"),
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
