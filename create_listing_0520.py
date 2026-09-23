import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def make_xlsx(filename, data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "架電リスト"
    headers = ["企業名", "LP URL", "広告種別", "電話番号", "商材", "検索KW", "業界"]
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
            cell.alignment = left
            cell.border = border
            # URL列（2列目）はハイパーリンク設定
            if col_idx == 2 and val and str(val).startswith("http"):
                cell.hyperlink = val
                cell.font = Font(name="メイリオ", size=9, color="0563C1", underline="single")
            else:
                cell.font = Font(name="メイリオ", size=9)
    col_widths = [30, 45, 16, 16, 20, 30, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(filename)
    print(f"保存完了: {filename}  ({len(data)}社)")

rows = [
    # 企業名, LP URL, 広告種別, 電話番号, 商材, 検索KW, 業界
    ["株式会社ブライダルネット", "https://www.bridalnet.co.jp/", "Meta", "03-5324-5666", "婚活サービス", "婚活サイト おすすめ", "婚活"],
    ["株式会社Aidemy", "https://aidemy.net/", "Meta", "03-6868-0998", "AIプログラミング講座", "AI プログラミング スクール", "教育"],
    ["株式会社COACHTECH", "https://coachtech.site/", "Meta", "050-1780-2465", "Webエンジニア育成スクール", "Webエンジニア なり方", "教育"],
    ["株式会社Coconala", "https://coconala.com/", "Meta", "03-6417-3390", "スキルマーケット", "スキル販売 副業", "フリーランス"],
    ["株式会社ビザスク", "https://visasq.co.jp/", "Meta", "03-3344-0660", "スポットコンサル", "スポットコンサル 副業", "フリーランス"],
    ["株式会社コデアル", "https://codeal.jp/", "Meta", "", "フリーランスエンジニア案件", "エンジニア 案件 直請け", "フリーランス"],
    ["株式会社エムール", "https://www.emoor.jp/", "Meta", "042-686-0171", "寝具・枕通販", "低反発枕 おすすめ", "インテリア"],
    ["株式会社モダンデコ", "https://www.modern-deco.jp/", "Meta", "0570-039-777", "家具・インテリアEC", "ソファ 通販 安い", "インテリア"],
    ["株式会社タンスのゲン", "https://www.tansu-gen.com/", "Meta", "0944-86-3871", "家具EC", "ベッド 安い 通販", "インテリア"],
    ["株式会社アクメファニチャー", "https://acme.co.jp/", "Meta", "03-5720-1071", "ヴィンテージ家具", "アメリカン 家具 通販", "インテリア"],
    ["アコーディア・ゴルフ株式会社", "https://www.accordiagolf.com/", "Meta", "03-6688-1500", "ゴルフ場予約・レッスン", "ゴルフ場 予約 格安", "ゴルフ"],
    ["株式会社ナチュラム", "https://www.naturum.co.jp/", "Meta", "06-6910-0031", "アウトドア用品EC", "キャンプ用品 通販", "アウトドア"],
    ["株式会社エルブレス", "https://www.supersports.com/", "Meta", "024-927-7719", "登山・アウトドア用品", "登山用品 おすすめ 店舗", "アウトドア"],
    ["株式会社なっぷ", "https://www.nap-camp.com/", "Meta", "03-6231-0760", "キャンプ場予約", "キャンプ場 予約 おすすめ", "アウトドア"],
    ["株式会社ゴールドウイン（THE NORTH FACE）", "https://www.goldwin.co.jp/tnf/", "Meta", "0766-61-4802", "アウトドアウェア", "THE NORTH FACE 公式", "アウトドア"],
    ["株式会社スヴェンソン", "https://mens-svenson.net/", "Meta", "03-3586-0011", "ウィッグ・増毛サービス", "薄毛 解決 男性 ウィッグ", "ウィッグ・育毛"],
    ["株式会社くらしの友", "https://www.k-tomo.co.jp/", "Meta", "03-3735-3101", "葬儀・互助会", "互助会 葬儀 おすすめ", "葬儀・終活"],
    ["株式会社Curon（クロン）", "https://curon.co/", "Meta", "042-313-7590", "オンライン診療", "オンライン診療 使い方", "医療テック"],
    ["株式会社ミナカラ", "https://minakara.com/", "Meta", "03-5544-8282", "オンライン薬局", "薬 処方箋 オンライン", "医療テック"],
    ["株式会社SOKUYAKU", "https://sokuyaku.jp/", "Meta", "0120-060-203", "オンライン診療・薬配送", "オンライン診療 薬 すぐ届く", "医療テック"],
    ["株式会社mederi", "https://mederi.jp/", "Meta", "", "婦人科オンライン診療", "ピル 処方 オンライン", "医療テック"],
    ["株式会社OwnersBook", "https://www.ownersbook.jp/", "Meta", "03-6630-6690", "不動産クラウドファンディング", "不動産CF 利回り", "不動産投資"],
    ["株式会社クラウドバンク", "https://crowdbank.jp/", "Meta", "03-6447-0011", "ソーシャルレンディング", "利回り 高い 投資 比較", "金融"],
    ["株式会社FUNDROP", "https://fundrop.jp/", "Meta", "03-6441-3086", "不動産クラウドファンディング", "不動産投資 少額 始め方", "不動産投資"],
    ["株式会社SOMPOケア", "https://sompocare.com/", "Meta", "03-5715-5532", "有料老人ホーム", "老人ホーム 費用 相場", "介護"],
    ["株式会社ソラスト", "https://www.solasto.co.jp/", "Meta", "03-3450-2610", "介護・医療サービス", "介護 仕事 資格", "介護"],
    ["株式会社山田養蜂場", "https://www.3838.com/", "Meta", "0120-383-830", "ローヤルゼリー・健康食品通販", "ローヤルゼリー 効果", "健康食品"],
    ["株式会社エバーライフ", "https://www.ever-life.jp/", "Meta", "0120-783-793", "コラーゲン・健康食品通販", "コラーゲン サプリ 効果 おすすめ", "健康食品"],
    ["株式会社ノエビアホールディングス", "https://sp.noevirstyle.jp/", "Meta", "03-5568-0300", "化粧品通販", "敏感肌 化粧品 おすすめ", "コスメ"],
    ["株式会社アクシージア", "https://www.axxzia.com/", "Meta", "03-6304-5840", "エイジングケア化粧品", "エイジングケア 化粧水 おすすめ", "コスメ"],
    ["株式会社ナップス", "https://shops.naps-jp.com/", "Meta", "045-441-1171", "バイク用品販売", "バイク用品 おすすめ 店舗", "バイク"],
    ["株式会社ウェビック", "https://www.webike.net/", "Meta", "03-5431-3219", "バイク用品・パーツEC", "バイクパーツ 通販 安い", "バイク"],
    ["株式会社アニヴェルセル", "https://www.anniversaire.co.jp/", "Meta", "03-5786-1088", "結婚式場・ブライダル", "結婚式場 おしゃれ 東京", "ウエディング"],
    ["株式会社テイクアンドギブ・ニーズ", "https://www.tgn.co.jp/", "Meta", "03-6833-1122", "結婚式プロデュース", "結婚式 費用 プロデュース会社", "ウエディング"],
    ["株式会社ベストブライダル", "https://www.bestbridal.co.jp/", "Meta", "03-5464-6606", "ブライダルプロデュース", "結婚式場 相談 無料", "ウエディング"],
    ["税理士法人チェスター", "https://chester-tax.com/", "Meta", "03-6869-5040", "相続税申告", "相続税 税理士 おすすめ", "法律・税務"],
    ["司法書士法人みどり法務事務所", "https://www.midori-law.com/", "Meta", "03-5212-1821", "債務整理・過払い金請求", "借金 相談 無料 司法書士", "法律・税務"],
    ["弁護士法人サリュ", "https://salyu.co.jp/", "Meta", "0120-181-398", "交通事故弁護士", "交通事故 弁護士 費用 無料", "法律"],
    ["株式会社スタッフサービス", "https://www.staffservice.co.jp/", "Meta", "03-5224-1581", "人材派遣・求人", "派遣 仕事 探し 登録", "人材"],
    ["パーソルテンプスタッフ株式会社", "https://www.tempstaff.co.jp/", "Meta", "03-5350-1212", "人材派遣", "派遣会社 大手 おすすめ", "人材"],
    ["株式会社アウトソーシングテクノロジー", "https://www.oc-t.jp/", "Meta", "03-3286-4777", "エンジニア派遣・アウトソーシング", "ITエンジニア 派遣 求人", "人材"],
    ["株式会社UTエイム", "https://www.ut-aim.com/", "Meta", "03-5447-1715", "製造派遣・アウトソーシング", "工場 派遣 高収入 求人", "人材"],
    ["株式会社フルキャスト", "https://fc.fullcast.co.jp/", "Meta", "03-4530-4848", "単発バイト・人材派遣", "単発 バイト すぐ働ける", "人材"],
    ["株式会社エイチームライフデザイン（引越し侍）", "https://hikkoshi-samurai.jp/", "Meta", "052-533-2098", "引越し一括比較", "引越し 費用 一括比較", "引越し"],
    ["株式会社エス・エム・エス", "https://kango-oshigoto.jp/", "Meta", "03-6721-2400", "看護師転職エージェント", "看護師 転職 おすすめ", "転職"],
    ["OWNDAYS株式会社", "https://www.owndays.com/jp/ja", "Meta", "0120-900-298", "眼鏡チェーン", "メガネ 安い おしゃれ 店舗", "眼鏡"],
    ["株式会社シードアイ", "https://www.seedeye.co.jp/", "Meta", "03-5842-7371", "コンタクトレンズ通販", "コンタクトレンズ 通販 安い", "コンタクト"],
    ["株式会社プリンセスウィンク", "https://princesswink.co.jp/", "Meta", "", "まつ毛エクステFC", "マツエク サロン 資格", "美容"],
    ["株式会社ネイルクイック", "https://www.nailquick.co.jp/", "Meta", "03-5447-5583", "ネイルサロンFC", "ネイルサロン 予約 安い", "美容"],
    ["株式会社アキュラホーム", "https://aqura.co.jp/", "Meta", "03-6302-5001", "注文住宅", "注文住宅 価格 坪単価", "住宅"],
    ["株式会社アイダ設計", "https://www.aidagroup.co.jp/", "Meta", "050-3100-2611", "ローコスト注文住宅", "安い 家 建てる ローコスト", "住宅"],
    ["株式会社エスリード", "https://www.eslead.co.jp/", "Meta", "06-6345-1368", "投資用マンション", "マンション投資 大阪 おすすめ", "不動産"],
    ["株式会社資格スクウェア", "https://shikakusquare.com/", "Meta", "", "難関資格オンライン講座", "司法試験 予備試験 通信 講座", "教育"],
    ["株式会社コナミスポーツクラブ", "https://www.konamisportsclub.jp/", "Meta", "03-5769-0573", "フィットネスクラブ", "スポーツジム 月会費 おすすめ", "フィットネス"],
    ["株式会社ジェクサー・フィットネス＆スパ", "https://www.jexer.jp/", "Meta", "0570-00-3535", "都心型フィットネスジム", "都心 ジム おすすめ 24時間", "フィットネス"],
    ["株式会社リゾートトラスト", "https://www.resorttrust.co.jp/", "Meta", "052-933-6000", "リゾートホテル会員権", "リゾートホテル 会員権 おすすめ", "旅行"],
    ["株式会社東急バケーションズ", "https://www.tokyu-vacations.com/", "Meta", "0120-618-109", "リゾート会員サービス", "東急 リゾート 会員 価格", "旅行"],
    ["株式会社ABホテル", "https://www.ab-hotel.jp/", "Meta", "0566-79-3013", "ビジネスホテルチェーン", "ビジネスホテル 格安 全国", "ホテル"],
    ["株式会社リースナブル", "https://leasnable.com/", "Meta", "0120-064-581", "個人向けカーリース", "車 リース 個人 安い", "自動車"],
    ["株式会社COREFIT", "https://corefit.jp/", "Meta", "03-5544-9825", "EMS美顔器", "美顔器 おすすめ 効果", "美容家電"],
    ["弁護士法人きさらぎ", "https://kisaragi-law.jp/", "Meta", "03-6694-1980", "離婚・男女問題弁護士", "離婚 弁護士 費用 相談", "法律"],
    ["株式会社スマートバンク", "https://smartbank.co.jp/", "Meta", "03-4221-0007", "家計管理アプリ・プリペイドカード", "家計管理 カード おすすめ", "フィンテック"],
    ["株式会社charm", "https://www.charm.jp/", "Meta", "0276-88-4828", "ペット用品EC", "ペット用品 通販 安い", "ペット"],
    ["株式会社シュアラスター", "https://www.surluster.com/", "Meta", "03-5733-4189", "洗車・カーケア用品", "洗車 コーティング おすすめ", "自動車"],
    ["株式会社オートウェイ", "https://www.autoway.jp/", "Meta", "093-436-4800", "タイヤ通販", "タイヤ 通販 安い 交換", "自動車"],
    ["株式会社オーダースーツSADA", "https://www.ordersuit.info/", "Meta", "03-5809-2536", "オーダースーツ", "スーツ オーダーメイド 安い", "ファッション"],
    ["株式会社FXプライムbyGMO", "https://www.fxprime.com/", "Meta", "03-5489-2511", "FX取引サービス", "FX 始め方 スプレッド 比較", "金融"],
    ["株式会社SBI FXトレード", "https://www.sbifxt.co.jp/", "Meta", "03-3589-3403", "FX取引サービス", "FX 少額 始める", "金融"],
    ["楽天銀行株式会社", "https://www.rakuten-bank.co.jp/", "Meta", "0120-776-910", "ネット銀行", "ネット銀行 金利 おすすめ", "金融"],
    ["株式会社ミラブル", "https://www.mirabelle.co.jp/", "Meta", "0120-3434-26", "ウルトラファインバブルシャワーヘッド", "シャワーヘッド 美肌 おすすめ", "美容家電"],
    ["株式会社ブレインスリープ", "https://brainsleep.com/", "Meta", "0120-088-885", "枕・睡眠グッズ通販", "枕 おすすめ 首こり 快眠", "寝具"],
    ["EMMA SLEEP JAPAN株式会社", "https://www.emmasleep.jp/", "Meta", "03-4579-5916", "マットレス通販", "マットレス おすすめ 口コミ", "寝具"],
    ["株式会社ジオス（GEOS英会話）", "https://www.geosenglish.com/", "Meta", "", "英会話スクール", "英会話 大人 スクール 料金", "語学"],
    ["株式会社ファイナンシャルアカデミー", "https://www.f-academy.jp/", "Meta", "03-6206-3960", "資産運用・お金の学校", "資産運用 セミナー 初心者", "教育"],
    ["株式会社ブルックスホールディングス", "https://www.brooks.co.jp/", "Meta", "045-902-0990", "コーヒー通販定期便", "コーヒー 定期便 おすすめ", "食品EC"],
    ["株式会社珈琲きゃろっと", "https://www.coffee-karrot.com/", "Meta", "050-3188-8388", "スペシャルティコーヒー通販", "コーヒー豆 おすすめ 通販", "食品EC"],
    ["株式会社アプラス", "https://www.aplus.co.jp/", "Meta", "0570-008-789", "クレジットカード・ショッピングローン", "ショッピングローン 審査 おすすめ", "金融"],
    ["株式会社アーバネットコーポレーション", "https://www.urbanet.co.jp/", "Meta", "03-6550-9160", "投資用マンション", "ワンルームマンション投資 利回り", "不動産投資"],
    ["株式会社カバーズ", "https://www.covers.jp/", "Meta", "", "リフォームFC加盟", "リフォーム FC 独立", "リフォーム"],
    ["株式会社リフォームの窓口", "https://www.reform-madoguchi.com/", "Meta", "", "リフォーム一括見積もり", "リフォーム 業者 比較 見積もり", "リフォーム"],
    ["株式会社ゼット（ZETT）", "https://www.zett.co.jp/", "Meta", "06-6779-1171", "野球・スポーツ用品", "野球 グローブ おすすめ", "スポーツ"],
    ["株式会社エスエスケイ（SSK）", "https://www.ssk-baseball.co.jp/", "Meta", "06-6768-1111", "野球用品", "野球 バット 選び方", "スポーツ"],
    ["株式会社4℃ホールディングス", "https://www.fdh.co.jp/", "Meta", "03-5719-3295", "ジュエリー・アクセサリー", "ジュエリー プレゼント 女性", "ジュエリー"],
    ["株式会社TASAKI", "https://www.tasaki.co.jp/", "Meta", "078-302-3321", "高級ジュエリー・真珠", "真珠 ネックレス おすすめ ブランド", "ジュエリー"],
    ["株式会社シモンズ", "https://www.simmons.co.jp/", "Meta", "0120-316-066", "高級マットレス・ベッド", "シモンズ マットレス 価格", "寝具"],
    ["株式会社カリモク家具", "https://www.karimoku.com/", "Meta", "0562-83-1111", "国産高級家具", "カリモク 家具 ソファ", "インテリア"],
    ["株式会社シーリーベッドジャパン", "https://www.sealy.co.jp/", "Meta", "03-5413-6600", "高品質マットレス", "シーリー マットレス 評判", "寝具"],
    ["株式会社プリントネット", "https://www.print-net.info/", "Meta", "050-3734-6495", "印刷EC・チラシ印刷", "チラシ 印刷 安い 激安", "印刷"],
    ["株式会社セキュリティハウス", "https://www.security-house.co.jp/", "Meta", "0120-8484-24", "防犯カメラ・警備システム", "防犯カメラ 設置 業者", "セキュリティ"],
    ["株式会社イロドリ", "https://irodori.co.jp/", "Meta", "03-6821-2233", "フォトブック・写真プリント", "フォトブック おすすめ 安い", "フォトサービス"],
    ["株式会社デロンギ・ジャパン", "https://www.delonghi.com/ja-jp", "Meta", "03-5256-6321", "全自動コーヒーメーカー", "全自動 コーヒーメーカー おすすめ", "家電"],
    ["株式会社ティーライフ", "https://www.t-life.co.jp/", "Meta", "0547-46-3459", "健康茶・サプリ通販", "国産 ほうじ茶 健康 通販", "健康食品"],
    ["株式会社石澤研究所", "https://www.ishizawa-lab.co.jp/", "Meta", "03-3796-2821", "スキンケア・日焼け止め", "日焼け止め 透明白肌 おすすめ", "コスメ"],
    ["株式会社アルビオン", "https://www.albion.co.jp/", "Meta", "03-5524-1711", "エイジングケア化粧品", "化粧品 エクサージュ 口コミ", "コスメ"],
    ["株式会社イミュ（CANMAKE）", "https://www.canmake.com/", "Meta", "0120-441-184", "プチプラコスメ", "キャンメイク 人気 コスメ", "コスメ"],
    ["株式会社マザーハウス", "https://www.mother-house.jp/", "Meta", "03-6240-1415", "バッグ・革小物EC", "革 バッグ 通販 おすすめ", "ファッション"],
    ["株式会社ヒロセ通商（LION FX）", "https://hirose-fx.co.jp/", "Meta", "06-6534-0708", "FX取引サービス", "LION FX 評判 スプレッド", "金融"],
    ["株式会社トレイダーズ証券（みんなのFX）", "https://min-fx.jp/", "Meta", "0120-637-104", "FX取引サービス", "みんなのFX 評判 初心者", "金融"],
    ["株式会社ヘアスタジオIWASAKI", "https://www.iwasaki.co.jp/", "Meta", "045-909-1538", "低価格ヘアカット", "美容室 格安 カット 1000円", "美容室"],
    ["株式会社コーセーコスメポート", "https://www.kose-cosmeport.co.jp/", "Meta", "03-6771-7611", "プチプラスキンケア・コスメ", "化粧品 コスパ おすすめ", "コスメ"],
]

assert len(rows) == 100, f"100社ではなく{len(rows)}社です"

out = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_Meta_2026-05-20_new.xlsx"
make_xlsx(out, rows)
