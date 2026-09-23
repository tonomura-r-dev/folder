import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-29g.xlsx"

HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

ROWS = [
    ["東急リバブル株式会社",         "https://www.livable.co.jp/kounyu/s/lp/lgente/",              "リスティング広告", "",               "新築マンション購入・リノベーション相談", "マンション 購入, 新築マンション 比較, リノベーション 相談",  "不動産・売買"],
    ["AIG損害保険株式会社",           "https://www.aig.co.jp/sonpo/lp/ind/2komahoken-auto",         "リスティング広告", "",               "自動車保険・損害保険",                   "自動車保険 比較, 車 保険 安い, 損保 おすすめ",               "損害保険"],
    ["株式会社土屋ホーム",             "https://www.tsuchiyahome.jp/lp/campaign02/",                 "リスティング広告", "",               "注文住宅",                               "注文住宅 北海道, 家 建てる 費用, 工務店 おすすめ",           "住宅・建設"],
    ["フジ住宅株式会社",               "https://fuji-ie.com/lp/new_bukken/",                         "リスティング広告", "",               "分譲住宅・新築一戸建て",                 "新築一戸建て 大阪, 建売住宅 購入, 住宅 分譲",                "不動産・住宅"],
    ["イオンライフ株式会社",           "https://www.aeonlife.jp/lp/funeral-support",                 "リスティング広告", "",               "葬儀・家族葬サービス",                   "家族葬 費用, 葬儀 安い, 葬式 費用 相場",                     "葬儀"],
    ["ハウスメイトパートナーズ株式会社", "https://www.housemate.co.jp/owner/lp/index.html",           "リスティング広告", "",               "賃貸管理・オーナーサポート",             "賃貸 管理会社, 不動産 管理, アパート 管理 委託",             "不動産管理"],
    ["株式会社TSON",                   "https://www.tson.co.jp/futoku-sonae/lp/sonae_37/",           "リスティング広告", "",               "不動産小口化商品",                       "不動産投資 小口, 相続対策 不動産, 不動産小口化",             "不動産投資"],
    ["株式会社COSPAウエルネス",        "https://www.cospa-wellness.co.jp/lp/cospa-fitness-taiken/", "リスティング広告", "",               "フィットネスクラブ体験入会",             "ジム 体験, フィットネス 入会, スポーツクラブ 近く",          "フィットネス"],
    ["西川株式会社",                   "https://www.nishikawa1566.com/contents/nishikawa-down/lp/select/", "リスティング広告", "",      "羽毛布団・寝具",                         "羽毛布団 おすすめ, 布団 比較, 羽毛布団 選び方",              "寝具・インテリア"],
    ["ヒノキヤグループ株式会社",       "https://www.hinokiya.jp/lp/01/",                             "リスティング広告", "",               "注文住宅（Z空調の家）",                  "注文住宅 東京, 高気密高断熱 住宅, Z空調",                    "住宅・建設"],
    ["医療法人恵生会",                 "https://www.keiseikai.or.jp/lp/index.html",                  "リスティング広告", "",               "婦人科内視鏡手術",                       "腹腔鏡手術 婦人科, 子宮筋腫 手術, 内視鏡手術 病院",         "医療・婦人科"],
    ["株式会社トライエ",               "https://venus-walker.com/lp/?page=ad",                       "リスティング広告", "",               "美容モニターサービス",                   "美容モニター 謝礼, 化粧品 モニター 募集, 美容 体験",         "マーケティング・美容"],
    ["学校法人国際学園",               "https://www.clark.ed.jp/lp/clark-smart-pbl/",                "リスティング広告", "",               "通信制高校（クラーク記念国際）",         "通信制高校 おすすめ, 高校 転校 通信, 不登校 高校",           "教育"],
    ["株式会社ニッセンライフ",         "https://www.nissen-life.co.jp/lp/cancer20/",                 "リスティング広告", "",               "生命保険・医療保険比較",                 "医療保険 比較, がん保険 おすすめ, 保険 見直し",              "保険代理"],
    ["FTC株式会社",                    "https://www.recyclemart.jp/lp/franchise-recruit/",           "リスティング広告", "",               "フランチャイズ加盟店募集（リサイクルマート）", "リサイクルショップ FC加盟, 独立 開業, フランチャイズ 募集", "リユース・FC"],
    ["学校法人産業能率大学",           "https://www.sanno.ac.jp/tukyo/lp/index.html",                "リスティング広告", "",               "通信制大学（通信教育課程）",             "通信制大学 社会人, 大学 通信 編入, 働きながら 大学",         "教育（通信制大学）"],
    ["株式会社メガネスーパー",         "https://www.meganesuper.co.jp/lp/",                          "リスティング広告", "",               "眼鏡・コンタクトレンズ・補聴器",         "メガネ 度付き, コンタクトレンズ 購入, 眼鏡 おすすめ",        "眼鏡・補聴器"],
    ["株式会社ソラシドエア",           "https://www.solaseedair.jp/smileclub/lp/",                   "リスティング広告", "0570-037-283",   "航空券・マイレージサービス",             "格安航空券 九州, 宮崎 飛行機 予約, 那覇 航空券 安い",        "航空"],
    ["株式会社ティア",                 "https://www.tear.co.jp/lp/tearnokai/",                       "リスティング広告", "",               "葬儀・家族葬（会員サービス）",           "葬儀 費用 名古屋, 家族葬 会館, 葬式 費用 相場",              "葬儀"],
    ["株式会社セイバン",               "https://www.seiban.co.jp/lp/karukutekarui/",                 "リスティング広告", "",               "ランドセル（天使のはね）",               "ランドセル 軽い, ランドセル 人気, 天使のはね 2027",           "教育用品"],
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "架電リスト"

# ヘッダー行
for col, (h, w) in enumerate(zip(HEADER, COL_WIDTHS), 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    ws.column_dimensions[get_column_letter(col)].width = w

ws.row_dimensions[1].height = 20

# データ行
for row_idx, row_data in enumerate(ROWS, 2):
    fill = ROW_FILLS[(row_idx - 2) % 2]
    for col_idx, value in enumerate(row_data, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.fill = fill
        cell.alignment = ROW_ALIGN
        if col_idx == 2 and value:
            cell.font = URL_FONT
            cell.hyperlink = value
        else:
            cell.font = ROW_FONT
    ws.row_dimensions[row_idx].height = 16

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}1"

wb.save(OUTPUT_PATH)
print(f"保存完了: {OUTPUT_PATH}")
print(f"行数: {len(ROWS)}社")
