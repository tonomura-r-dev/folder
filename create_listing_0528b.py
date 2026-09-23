import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-28b.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]

DATA = [
    [
        "株式会社クラブツーリズム",
        "https://www.club-t.com/special/common/lp/tabikomachi/gourmet/",
        "リスティング広告",
        "0120-834-834",
        "旅行ツアー・バスツアー（シニア・国内旅行専門）",
        "旅行 バスツアー 国内旅行 シニア旅行 旅行ツアー 一人旅",
        "旅行・宿泊",
    ],
    [
        "ドリームベッド株式会社",
        "https://www.dreambed.jp/lp/",
        "リスティング広告",
        "0120-652-231",
        "ベッド・マットレス直販（快眠・睡眠改善）",
        "ベッド マットレス 快眠 睡眠 ベッドフレーム 寝具通販",
        "家具・インテリアEC",
    ],
    [
        "楽天エナジー株式会社（楽天でんき）",
        "https://energy.rakuten.co.jp/campaign/lp/mama/",
        "リスティング広告",
        "0120-800-847",
        "電力・ガスサービス（楽天ポイントがたまる電気）",
        "電気 楽天でんき 電力 格安電気 電気代節約 楽天ポイント",
        "電力・ガス・エネルギー",
    ],
    [
        "医療法人社団十二会（ゴリラクリニック）",
        "https://gorilla.clinic/lp/",
        "リスティング広告",
        "03-5291-5270",
        "メンズ脱毛・AGA治療・医療痩身（男性専門クリニック）",
        "メンズ脱毛 AGA 薄毛治療 医療脱毛 男性脱毛 ゴリラクリニック",
        "美容クリニック・脱毛・エステ",
    ],
    [
        "株式会社クレディセゾン（ネット保険）",
        "https://nethoken.saisoncard.co.jp/lp/faq/",
        "リスティング広告",
        "",
        "ネット保険（自動車・がん保険オンライン加入）",
        "ネット保険 自動車保険 がん保険 生命保険 セゾン保険",
        "保険・金融・証券・カード",
    ],
    [
        "株式会社ビズリーチ",
        "https://www.bizreach.jp/lp/official/pc/",
        "リスティング広告",
        "",
        "ハイクラス転職サービス（年収600万円以上向け）",
        "ハイクラス転職 転職サイト 転職 年収アップ 管理職 転職エージェント",
        "転職・就職・人材派遣",
    ],
    [
        "弁護士法人ベリーベスト法律事務所",
        "https://www.vbest.jp/roudoumondai/lp/zangyou01/",
        "リスティング広告",
        "0120-117-059",
        "労働問題・残業代請求の弁護士無料相談",
        "残業代請求 労働問題 弁護士相談 残業代 無料相談 労働弁護士",
        "法律・税務・相談",
    ],
    [
        "株式会社明光ネットワークジャパン（明光義塾）",
        "https://www.meikogijuku.jp/lp/",
        "リスティング広告",
        "0120-334-117",
        "個別指導学習塾（小・中・高・大学受験対応・全国2,000教室）",
        "個別指導 学習塾 塾 受験対策 個別指導塾 明光義塾",
        "教育・スキルアップ",
    ],
    [
        "株式会社アンファー（スカルプD）",
        "https://www.angfa-store.jp/lp/bmps",
        "リスティング広告",
        "0120-866-866",
        "スカルプDシャンプー・育毛剤（医師監修メンズヘアケア）",
        "スカルプD 薄毛 育毛 抜け毛 シャンプー 育毛シャンプー",
        "化粧品・健康食品通販D2C",
    ],
    [
        "SMBCコンシューマーファイナンス株式会社（SMBCモビット）",
        "https://www.mobit.ne.jp/lp/al04/index.html",
        "リスティング広告",
        "0120-03-5000",
        "カードローン（SMBC系・在籍確認なし・最短即日融資）",
        "カードローン 消費者金融 キャッシング 借入 即日融資 モビット",
        "保険・金融・証券・カード",
    ],
]


def make_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "架電リスト"

    header_font = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    fill_odd = PatternFill("solid", fgColor="F2F7FC")
    fill_even = PatternFill("solid", fgColor="FFFFFF")
    data_font = Font(name="メイリオ", size=9)
    data_align = Alignment(vertical="center", wrap_text=False)
    url_font = Font(name="メイリオ", size=9, color="0563C1", underline="single")

    thin = Side(style="thin", color="D0D7DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ヘッダー行
    for col_idx, col_name in enumerate(HEADER, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    ws.row_dimensions[1].height = 22

    # データ行
    for row_idx, row_data in enumerate(DATA, start=2):
        fill = fill_odd if (row_idx % 2 == 0) else fill_even
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill = fill
            cell.border = border
            if col_idx == 2:  # LP URL列
                cell.font = url_font
                cell.hyperlink = value
                cell.alignment = data_align
            else:
                cell.font = data_font
                cell.alignment = data_align
        ws.row_dimensions[row_idx].height = 18

    # 列幅
    for col_idx, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # 先頭行固定・オートフィルター
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}1"

    wb.save(OUTPUT_PATH)
    print(f"保存完了: {OUTPUT_PATH}")
    print(f"件数: {len(DATA)}社")


if __name__ == "__main__":
    make_xlsx()
