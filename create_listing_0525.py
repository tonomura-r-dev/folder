import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-25.xlsx"

HEADER = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 45, 16, 16, 20, 30, 18]

DATA = [
    ("エリアリンク株式会社（ハローストレージ）", "https://www.hello-storage.com/campaign/", "リスティング広告", "", "トランクルーム・レンタル収納", "トランクルーム おすすめ, レンタル収納 月額, トランクルーム 東京", "収納・トランクルーム"),
    ("株式会社キュラーズ（Quraz）", "https://www.quraz.com/", "リスティング広告", "", "屋内型トランクルーム", "トランクルーム バイク保管, 屋内型トランクルーム, 収納スペース 月額", "収納・トランクルーム"),
    ("株式会社寺田倉庫（minikura）", "https://minikura.com/feature-lp/", "リスティング広告", "", "宅配型トランクルーム", "宅配型トランクルーム, 荷物 預ける 月額, 宅配保管サービス", "収納・トランクルーム"),
    ("ニッポンレンタカーサービス株式会社", "https://www.nipponrentacar.co.jp/", "リスティング広告", "", "レンタカー予約", "レンタカー 予約, 格安レンタカー, レンタカー キャンペーン", "レンタカー"),
    ("スカイレンタカー株式会社", "https://skyrentacar.com/", "リスティング広告", "", "格安レンタカー（沖縄・北海道）", "格安レンタカー 沖縄, レンタカー 那覇, 沖縄 レンタカー 予約", "レンタカー"),
    ("ユーカーパック株式会社", "https://ucarpac.com/", "リスティング広告", "", "車買取オークション", "車 高く売る, 車買取 一括査定, 愛車 売却", "車・カーシェア・バイク"),
    ("株式会社コバック（車検のコバック）", "https://www.kobac.co.jp/", "リスティング広告", "", "車検・車メンテナンス", "車検 安い, 車検 予約, 車検 近く", "車・カーシェア・バイク"),
    ("ニコニコ車検株式会社", "https://2525syaken.com/", "リスティング広告", "", "格安車検", "車検 格安, 軽自動車 車検, 車検 キャンペーン", "車・カーシェア・バイク"),
    ("廃車王（株式会社オールドカーサポートジャパン）", "https://www.haishaou.com/", "リスティング広告", "", "廃車・故障車買取", "廃車 買取, 故障車 売る, 廃車 無料引き取り", "車・カーシェア・バイク"),
    ("株式会社クルカ", "https://newcar.shop/", "リスティング広告", "", "新車カーリース", "カーリース 月額, 新車 サブスク, 車 リース 個人", "車・カーシェア・バイク"),
    ("Mt.Fuji GX Holdings株式会社（エブリィフレシャス）", "https://every.frecious.jp/", "リスティング広告", "", "浄水型ウォーターサーバー", "浄水型ウォーターサーバー, ウォーターサーバー 水道水, ウォーターサーバー 月額", "ウォーターサーバー"),
    ("株式会社アルピナウォーター", "https://www.alpina-water.co.jp/", "リスティング広告", "", "天然水ウォーターサーバー", "天然水 ウォーターサーバー, ウォーターサーバー 宅配, アルプス天然水", "ウォーターサーバー"),
    ("株式会社マーキュロップ", "https://www.mercurop.co.jp/", "リスティング広告", "", "天然水ウォーターサーバー", "ウォーターサーバー 比較, 宅配水 おすすめ, ウォーターサーバー 申し込み", "ウォーターサーバー"),
    ("株式会社ふじざくら命水", "https://www.fujizakurameisui.jp/", "リスティング広告", "", "富士山天然水ウォーターサーバー", "富士山 天然水 宅配, ウォーターサーバー 富士山, 天然水 送料無料", "ウォーターサーバー"),
    ("株式会社トーエル", "https://www.toell.co.jp/", "リスティング広告", "", "天然水ウォーターサーバー", "ウォーターサーバー 天然水, 宅配水 比較, ウォーターサーバー 無料体験", "ウォーターサーバー"),
    ("ホワイトニングバー株式会社", "https://whiteningbar.jp/", "リスティング広告", "", "セルフホワイトニング専門店", "セルフホワイトニング 月額, 歯 白くする サロン, ホワイトニング 通い放題", "美容クリニック・脱毛・エステ"),
    ("スリムビューティハウス株式会社（SBH Medical）", "https://www.slim.co.jp/", "リスティング広告", "", "痩身エステ・脱毛サロン", "痩身エステ おすすめ, 体型管理 エステ, 痩身 無料体験", "美容クリニック・脱毛・エステ"),
    ("株式会社ジェイエステティック（ジェイエステ）", "https://www.j-esthe.com/", "リスティング広告", "", "エステ・脱毛・フェイシャル", "エステ 全国, 脱毛サロン おすすめ, エステ 無料体験", "美容クリニック・脱毛・エステ"),
    ("医療法人社団大美会（大美会クリニック）", "https://osaka-bc.com/", "リスティング広告", "", "医療脱毛・美容外科", "医療脱毛 大阪, 全身脱毛 クリニック, 医療脱毛 安い", "美容クリニック・脱毛・エステ"),
    ("医療法人社団英和会（エミシアクリニック）", "https://emishia-clinic.jp/", "リスティング広告", "", "医療脱毛", "医療脱毛 全身, 脱毛クリニック 比較, 医療脱毛 初回", "美容クリニック・脱毛・エステ"),
    ("医療法人社団レナトゥス（レナトゥスクリニック）", "https://renatusclinic.jp/", "リスティング広告", "", "医療脱毛", "医療脱毛 安い, 全身脱毛 安い クリニック, 医療脱毛 比較", "美容クリニック・脱毛・エステ"),
    ("医療法人社団（ブランクリニック）", "https://bccl.jp/", "リスティング広告", "", "医療脱毛・美容皮膚科", "医療脱毛 東京, 脱毛 クリニック, 全身脱毛 料金", "美容クリニック・脱毛・エステ"),
    ("医療法人社団グロウ会（AGAヘアクリニック）", "https://aga-hair.jp/", "リスティング広告", "", "AGA治療・薄毛治療", "AGA 治療, 薄毛 クリニック, 抜け毛 治療", "美容クリニック・脱毛・エステ"),
    ("株式会社レバレジーズメディカルケア（レバクリ）", "https://levacli.jp/", "リスティング広告", "", "AGAオンライン診療", "AGA オンライン診療, AGAクリニック 安い, 薄毛 オンライン治療", "美容クリニック・脱毛・エステ"),
    ("医療法人社団（イースト駅前クリニック）", "https://east-ekimae.jp/", "リスティング広告", "", "AGA治療・薄毛対策", "AGA 全国, 薄毛 治療 安い, AGA クリニック 予約", "美容クリニック・脱毛・エステ"),
    ("医療法人社団よじかい（よつば会クリニック）", "https://yotsubakai-group.com/", "リスティング広告", "", "医療脱毛・美容皮膚科", "医療脱毛 クリニック, 全身脱毛 おすすめ, 医療脱毛 予約", "美容クリニック・脱毛・エステ"),
    ("家庭教師の銀河（株式会社GINGAシステム）", "https://well-stone.info/", "リスティング広告", "", "家庭教師派遣・不登校支援", "家庭教師 おすすめ, 家庭教師 不登校, 家庭教師 無料体験", "教育・スキルアップ"),
    ("株式会社アルファ（家庭教師のアルファ）", "https://www.kateikyoushi-alpha.jp/", "リスティング広告", "", "家庭教師派遣サービス", "家庭教師 派遣, 家庭教師 無料体験, 家庭教師 中学生", "教育・スキルアップ"),
    ("株式会社クラウティ（クラウティ英会話）", "https://clouty.jp/", "リスティング広告", "", "子ども向けオンライン英会話", "子供 英会話 オンライン, こども英会話 月額, 子供 英語 習い事", "英会話・語学・資格スクール"),
    ("株式会社エイゴックス（エイゴックス）", "https://www.eigox.com/", "リスティング広告", "", "格安オンライン英会話", "オンライン英会話 格安, 英会話 月額, 英語 練習 オンライン", "英会話・語学・資格スクール"),
    ("株式会社ベストティーチャー（Best Teacher）", "https://bestteacher.us/", "リスティング広告", "", "英文添削オンライン英会話", "ビジネス英語 オンライン, 英語 添削, 英文 ライティング", "英会話・語学・資格スクール"),
    ("Hupro株式会社（ヒュープロ）", "https://hupro-job.com/", "リスティング広告", "", "士業・管理部門特化転職", "税理士 転職, 会計士 求人, 経理 転職サイト", "転職・就職・人材派遣"),
    ("M3キャリア株式会社（M3キャリアエージェント）", "https://career.m3.com/", "リスティング広告", "", "医師・医療職転職", "医師 転職, 医師 求人, 医師 転職エージェント", "転職・就職・人材派遣"),
    ("株式会社ファルマスタッフ", "https://www.pharma-staff.com/", "リスティング広告", "", "薬剤師特化転職", "薬剤師 転職, 薬剤師 求人, 薬剤師 パート", "転職・就職・人材派遣"),
    ("株式会社ネクストビート（保育のお仕事）", "https://hoiku-shigoto.com/", "リスティング広告", "", "保育士特化転職", "保育士 転職, 保育士 求人, 保育士 給料", "転職・就職・人材派遣"),
    ("株式会社マイビジョン（MyVision）", "https://my-vision.co.jp/", "リスティング広告", "", "コンサル業界転職", "コンサル 転職, コンサルタント 求人, 戦略コンサル 未経験", "転職・就職・人材派遣"),
    ("フォースタートアップス株式会社", "https://forstartups.com/", "リスティング広告", "", "スタートアップ・ベンチャー転職", "スタートアップ 転職, ベンチャー 転職エージェント, IPO 求人", "転職・就職・人材派遣"),
    ("株式会社ムービン・ストラテジック・キャリア", "https://www.movin.co.jp/", "リスティング広告", "", "コンサル・戦略系転職", "コンサル 転職 未経験, 外資 コンサル 求人, 経営コンサルタント", "転職・就職・人材派遣"),
    ("弁護士法人デイライト法律事務所", "https://daylight-law.jp/", "リスティング広告", "", "離婚・相続・交通事故法律相談", "弁護士 相談 無料, 離婚 弁護士, 慰謝料 請求 弁護士", "法律・税務・相談"),
    ("株式会社アシロ（ベンナビ弁護士）", "https://best.bengo4.com/", "リスティング広告", "", "弁護士費用比較・法律相談", "弁護士 費用 比較, 法律相談 無料, 弁護士 評判", "法律・税務・相談"),
    ("株式会社ソーラーパートナーズ", "https://www.solar-partners.jp/", "リスティング広告", "", "太陽光発電一括見積もり", "太陽光発電 見積もり, ソーラーパネル 設置 費用, 太陽光発電 比較", "不動産・賃貸・リノベーション・投資"),
    ("株式会社タイナビ（タイナビ）", "https://www.tainavi.com/", "リスティング広告", "", "太陽光発電・蓄電池一括見積もり", "太陽光発電 一括見積もり, 蓄電池 比較, 太陽光発電 業者 比較", "不動産・賃貸・リノベーション・投資"),
    ("株式会社エコ発", "https://www.eco-hatsu.com/", "リスティング広告", "", "太陽光発電価格比較", "太陽光発電 安い, ソーラーパネル 価格 比較, 太陽光発電 導入 費用", "不動産・賃貸・リノベーション・投資"),
    ("株式会社グリエネ", "https://greenene.jp/", "リスティング広告", "", "太陽光発電・蓄電池比較見積もり", "太陽光発電 業者 おすすめ, 蓄電池 見積もり, 太陽光パネル 比較", "不動産・賃貸・リノベーション・投資"),
    ("外壁塗装パートナーズ（株式会社スタンダード）", "https://gaiheki-partners.com/", "リスティング広告", "", "外壁塗装一括見積もり", "外壁塗装 見積もり, 外壁塗装 業者 比較, 外壁塗装 費用", "不動産・賃貸・リノベーション・投資"),
    ("ダンロップスポーツ株式会社（ダンロップゴルフスクール）", "https://sports.dunlop.co.jp/golfschool/", "リスティング広告", "", "ゴルフスクール・レッスン", "ゴルフスクール 初心者, ゴルフレッスン 東京, ゴルフ 体験レッスン", "フィットネス・スポーツ・ゴルフ"),
    ("東建コーポレーション株式会社", "https://www.token.co.jp/estate/", "リスティング広告", "", "アパート経営・土地活用", "土地活用 アパート, 賃貸経営 始め方, アパート経営 セミナー", "不動産・賃貸・リノベーション・投資"),
    ("買取王子（株式会社流通革命）", "https://kaitori-ouji.com/", "リスティング広告", "", "出張買取・宅配買取", "出張買取 おすすめ, 家電 買取, 不用品 買取", "買取・リユース・フリマ"),
    ("アサンテ株式会社", "https://www.asante.co.jp/", "リスティング広告", "", "シロアリ防除・害虫駆除", "シロアリ 駆除, 害虫駆除 業者, 白蟻 防除 費用", "家事代行・クリーニング"),
    ("株式会社あさひ（サイクルベースあさひ）", "https://ec.cb-asahi.co.jp/", "リスティング広告", "", "自転車・電動自転車販売", "電動自転車 おすすめ, 自転車 通販, クロスバイク 購入", "車・カーシェア・バイク"),
]

def make_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "架電リスト"

    # スタイル定義
    header_font = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    fill_odd  = PatternFill("solid", fgColor="F2F7FC")
    fill_even = PatternFill("solid", fgColor="FFFFFF")
    data_font = Font(name="メイリオ", size=9)
    data_align = Alignment(vertical="center", wrap_text=False)
    url_font_base = Font(name="メイリオ", size=9, color="0563C1", underline="single")

    thin_side = Side(style="thin", color="D0D7DE")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # ヘッダー行
    for col_idx, col_name in enumerate(HEADER, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # データ行
    for row_idx, row_data in enumerate(DATA, start=2):
        fill = fill_odd if (row_idx % 2 == 0) else fill_even
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill = fill
            cell.border = thin_border
            cell.alignment = data_align
            if col_idx == 2 and value:  # LP URL列
                cell.hyperlink = value
                cell.font = url_font_base
            else:
                cell.font = data_font

    # 列幅設定
    for col_idx, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # 行の高さ
    ws.row_dimensions[1].height = 22
    for row_idx in range(2, len(DATA) + 2):
        ws.row_dimensions[row_idx].height = 18

    # 先頭行固定・オートフィルター
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}1"

    wb.save(OUTPUT_PATH)
    print(f"保存完了: {OUTPUT_PATH}")
    print(f"件数: {len(DATA)}社")

if __name__ == "__main__":
    make_xlsx()
