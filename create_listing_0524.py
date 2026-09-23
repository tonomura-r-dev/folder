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
            if col_idx == 2 and val and str(val).startswith("http"):
                cell.hyperlink = val
                cell.font = Font(name="メイリオ", size=9, color="0563C1", underline="single")
            else:
                cell.font = Font(name="メイリオ", size=9)
    col_widths = [32, 45, 16, 16, 20, 30, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(filename)
    print(f"保存完了: {filename}  ({len(data)}社)")

rows = [
    # 企業名, LP URL, 広告種別, 電話番号, 商材, 検索KW, 業界
    # 第22弾
    ["株式会社ナノユニバース（nano・universe）", "https://nano-universe.jp/", "Meta", "0120-988-163", "トレンドカジュアルウェア", "レディース メンズ ファッション おしゃれ 通販", "ファッション"],
    ["株式会社ジュン（ROPE' PICNIC）", "https://www.ropepicnic.com/", "Meta", "03-5745-0661", "ナチュラルフェミニンウェア", "レディース ナチュラル おしゃれ 通販 きれいめ", "ファッション"],
    ["株式会社ゴルフパートナー（GOLF PARTNER）", "https://www.golfpartner.co.jp/", "Meta", "0120-009-660", "ゴルフクラブ買取・販売チェーン", "ゴルフクラブ 買取 中古 おすすめ 安い", "スポーツ・ゴルフ"],
    ["株式会社ニューバランスジャパン（New Balance）", "https://www.newbalance.jp/", "Meta", "0120-85-7598", "スニーカー・スポーツウェア", "スニーカー おすすめ ランニング New Balance 靴", "スポーツ・ファッション"],
    ["株式会社プーマジャパン（PUMA）", "https://jp.puma.com/", "Meta", "0120-051-776", "スポーツウェア・スニーカー通販", "スニーカー スポーツ おすすめ PUMA ウェア", "スポーツ・ファッション"],
    ["株式会社クロックスジャパン（Crocs）", "https://www.crocs.co.jp/", "Meta", "0120-200-625", "サンダル・クロッグシューズ", "クロックス サンダル おすすめ 夏 軽い 快適", "ファッション・靴"],
    ["株式会社明光ネットワークジャパン（明光義塾）", "https://www.meikogijuku.jp/", "Meta", "0120-934-681", "個別指導学習塾", "個別指導 塾 おすすめ 料金 小学生 中学生", "教育"],
    ["株式会社ナガセ（東進ハイスクール）", "https://www.toshin.com/", "Meta", "0120-104-555", "大学受験予備校・映像授業", "大学受験 予備校 おすすめ 東進 映像授業 合格", "教育"],
    ["株式会社Z会（Z-KAI）", "https://www.zkai.co.jp/", "Meta", "0120-977-844", "通信教育・学習サービス", "通信教育 おすすめ Z会 中学受験 高校受験", "教育"],
    ["株式会社アガルートアカデミー（アガルート）", "https://agaroot.jp/", "Meta", "0120-758-702", "資格試験対策通信講座", "資格 通信講座 おすすめ 司法書士 行政書士 宅建", "教育"],
    ["株式会社ミュゼプラチナム（ミュゼプラチナム）", "https://musee-pla.com/", "Meta", "0120-329-900", "女性向け脱毛サロン", "脱毛 サロン おすすめ 安い 全身 料金", "美容サロン・脱毛"],
    ["株式会社ジョンマスターオーガニックジャパン（John Masters Organics）", "https://www.johnmasters.jp/", "Meta", "0120-207-322", "オーガニックヘアケア通販", "オーガニック ヘアケア シャンプー おすすめ 自然派", "コスメ・ヘアケア"],
    ["株式会社セザンヌ化粧品（CEZANNE）", "https://www.cezanne.co.jp/", "Meta", "0120-44-1184", "プチプラコスメ・ファンデーション", "プチプラ コスメ ファンデーション 安い おすすめ", "コスメ"],
    ["株式会社ミルボン（MILBON）", "https://www.milbon.co.jp/", "Meta", "0120-373-249", "サロン品質ヘアケア通販", "ヘアケア サロン品質 おすすめ シャンプー ダメージ", "コスメ・ヘアケア"],
    ["株式会社ドクターシーラボ（Dr.Ci:Labo）", "https://www.drcilabo.jp/", "Meta", "0120-371-217", "美容液・スキンケア通販", "美容液 スキンケア おすすめ 敏感肌 シミ シワ", "コスメ"],
    ["株式会社オルビス（ORBIS）", "https://www.orbis.co.jp/", "Meta", "0120-010-010", "スキンケア・コスメ通販", "スキンケア 通販 おすすめ オルビス 乾燥肌 毛穴", "コスメ"],
    ["株式会社プレミアムウォーター（PREMIUM WATER）", "https://www.premiumwater.co.jp/", "Meta", "0120-990-834", "天然水ウォーターサーバー", "ウォーターサーバー おすすめ 天然水 安い 定期", "宅配水"],
    ["株式会社フレシャス（FRECIOUS）", "https://frecious.jp/", "Meta", "0120-937-290", "天然水・ウォーターサーバー", "ウォーターサーバー おすすめ 富士山 天然水 料金", "宅配水"],
    ["株式会社松井証券（松井証券）", "https://www.matsui.co.jp/", "Meta", "0120-021-906", "ネット証券・株式投資", "株 ネット証券 おすすめ 初心者 手数料 NISA", "証券・投資"],
    ["株式会社DMM.com（DMM FX）", "https://fx.dmm.com/", "Meta", "0120-961-488", "FX・外国為替取引", "FX おすすめ 初心者 スプレッド 口座 比較", "FX・投資"],
    ["株式会社OANDA Japan（OANDA）", "https://www.oanda.jp/", "Meta", "0120-867-177", "FX・CFD取引", "FX おすすめ CFD スプレッド 口座開設 比較", "FX・投資"],
    ["株式会社エン・ジャパン（エン転職）", "https://employment.en-japan.com/", "Meta", "0120-114-706", "転職求人サイト", "転職 求人 おすすめ 正社員 未経験 サイト", "転職"],
    ["株式会社ビズリーチ（BIZREACH）", "https://www.bizreach.jp/", "Meta", "0120-961-842", "ハイクラス転職サービス", "転職 ハイクラス 年収アップ スカウト おすすめ", "転職"],
    ["株式会社パーソルキャリア（doda）", "https://doda.jp/", "Meta", "0120-366-494", "転職・求人エージェントサービス", "転職 求人 エージェント おすすめ 年収 サポート", "転職"],
    ["エクスペディア・ジャパン合同会社（Expedia）", "https://www.expedia.co.jp/", "Meta", "0120-863-390", "オンライン旅行予約", "旅行 予約 ホテル おすすめ 格安 海外", "旅行"],
    ["株式会社一休（一休.com）", "https://www.ikyu.com/", "Meta", "0120-919-989", "高級ホテル・旅館予約", "旅館 ホテル 高級 おすすめ 予約 温泉", "旅行・ホテル"],
    ["株式会社エイチ・アイ・エス（H.I.S.）", "https://www.his-j.com/", "Meta", "0120-956-440", "格安旅行・海外ツアー", "旅行 格安 ツアー 海外 おすすめ HIS", "旅行"],
    ["株式会社オイシックス・ラ・大地（Oisix）", "https://www.oisix.com/", "Meta", "0120-020-892", "有機野菜・ミールキット定期便", "ミールキット 有機野菜 おすすめ 宅配 食材 定期", "食品EC"],
    ["株式会社らでぃっしゅぼーや（らでぃっしゅぼーや）", "https://radishbo-ya.co.jp/", "Meta", "0120-372-370", "有機野菜・食品宅配サービス", "有機野菜 宅配 おすすめ 無農薬 食材 安全", "食品EC"],
    ["株式会社千趣会（Belle Maison）", "https://www.bellemaison.jp/", "Meta", "0120-311-000", "ライフスタイル・ファッション通販", "通販 おすすめ 生活雑貨 レディース ファッション", "通販EC"],
    ["株式会社森下仁丹（仁丹）", "https://www.jintan.co.jp/", "Meta", "0120-81-1245", "腸活サプリ・健康食品通販", "腸活 サプリ おすすめ 乳酸菌 健康 通販", "健康食品"],
    ["株式会社井藤漢方製薬（井藤漢方）", "https://www.itohkampo.co.jp/", "Meta", "0120-85-0107", "ダイエット・健康サプリ通販", "ダイエット サプリ おすすめ 健康 コラーゲン 通販", "健康食品"],
    ["株式会社ティップネス（TIPNESS）", "https://www.tipness.co.jp/", "Meta", "0120-87-4131", "総合フィットネスクラブ", "フィットネス ジム おすすめ 会費 入会 プール", "フィットネス"],
    ["株式会社ゴールドジム・ジャパン（Gold's Gym）", "https://www.goldsgym.jp/", "Meta", "03-3416-5161", "本格フィットネスジム", "筋トレ ジム おすすめ 24時間 マシン トレーニング", "フィットネス"],
    ["株式会社カーブスジャパン（CURVES）", "https://www.curves.co.jp/", "Meta", "0120-533-748", "女性向けサーキットトレーニング", "フィットネス 女性 おすすめ 短時間 運動 40代", "フィットネス"],
    ["株式会社リネット（Lenet）", "https://lenet.jp/", "Meta", "0120-033-229", "宅配クリーニングサービス", "クリーニング 宅配 おすすめ 安い スーツ コート", "クリーニング"],
    ["株式会社よりそう（よりそうお葬式）", "https://www.yorisou.jp/", "Meta", "0120-556-789", "格安葬儀・家族葬サービス", "葬儀 家族葬 おすすめ 安い 料金 比較", "葬儀"],
    ["株式会社ベアーズ（Bears）", "https://www.happy-bears.com/", "Meta", "0120-432-888", "家事代行・ハウスクリーニング", "家事代行 おすすめ 料金 掃除 共働き 東京", "家事代行"],
    ["株式会社カジー（Casy）", "https://casy.co.jp/", "Meta", "0120-000-218", "家事代行サービス", "家事代行 おすすめ 安い 料金 掃除 スポット", "家事代行"],
    ["株式会社ビューティーキャデンツ（モグワン）", "https://mogwan.jp/", "Meta", "0120-838-852", "総合栄養食ドッグフード通販", "ドッグフード おすすめ 総合栄養食 安全 原材料", "ペット食品"],
    ["株式会社ライフネット生命保険（ライフネット生命）", "https://www.lifenet-seimei.co.jp/", "Meta", "0120-205-594", "ネット完結生命保険", "生命保険 おすすめ ネット 安い 掛け捨て 比較", "保険"],
    ["株式会社あみあみ（AmiAmi）", "https://www.amiami.jp/", "Meta", "03-5353-8687", "アニメ・フィギュア・ホビー通販", "フィギュア 通販 おすすめ アニメ 安い あみあみ", "ホビーEC"],
    ["株式会社グラニフ（graniph）", "https://www.graniph.com/", "Meta", "0120-048-339", "キャラクターTシャツ・アパレル通販", "Tシャツ おしゃれ キャラクター おすすめ 通販 デザイン", "ファッション"],
    ["株式会社ルームクリップ（RoomClip）", "https://roomclip.me/", "Meta", "03-6555-2801", "インテリア・住まいSNS・EC", "インテリア おしゃれ おすすめ 部屋 収納 実例", "インテリア"],
    ["株式会社ハウスドゥ（HOUSE DO）", "https://www.housedo.co.jp/", "Meta", "0120-200-899", "不動産売買・リノベーションFC", "不動産 売却 おすすめ 査定 リノベーション 買取", "不動産"],
    ["株式会社ナプラ（napla）", "https://www.napla.co.jp/", "Meta", "06-6150-8888", "サロン品質ヘアケア・スタイリング剤", "ヘアケア おすすめ サロン トリートメント スタイリング", "コスメ・ヘアケア"],
    ["株式会社コタ（YOLU）", "https://yolu-care.com/", "Meta", "0120-200-777", "ナイトケアヘアシャンプー・D2C", "シャンプー おすすめ ナイトケア 寝ている間 ヘアケア", "コスメ・ヘアケア"],
    ["株式会社ロフト（LOFT）", "https://www.loft.co.jp/", "Meta", "0120-56-7840", "生活雑貨・文具・コスメセレクト", "雑貨 おしゃれ コスメ おすすめ 文具 ギフト", "生活雑貨"],
    ["株式会社キャン★ドゥ（Can Do）", "https://www.cando-web.co.jp/", "Meta", "0120-720-031", "100円均一ライフスタイルショップ", "100均 おすすめ 雑貨 安い 収納 キャンドゥ", "生活雑貨"],
    ["株式会社ジャストシステム（スマイルゼミ）", "https://smile-zemi.jp/", "Meta", "0120-123-185", "子ども向けタブレット通信教育", "タブレット 通信教育 おすすめ 小学生 スマイルゼミ", "教育"],
    ["株式会社ハッシュパピージャパン（Hush Puppies）", "https://www.hushpuppies.co.jp/", "Meta", "0120-77-8025", "ウォーキング・コンフォートシューズ", "ウォーキングシューズ おすすめ 歩きやすい 靴 幅広", "ファッション・靴"],
]

assert len(rows) == 51, f"51社ではなく{len(rows)}社です"

out = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_Meta_2026-05-24.xlsx"
make_xlsx(out, rows)
