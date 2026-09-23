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
    # 第23弾（リスティング広告）
    ["弁護士法人アディーレ法律事務所", "https://www.adire.jp/", "リスティング広告", "0120-783-184", "債務整理・過払い金請求", "過払い金 弁護士 無料相談 債務整理 おすすめ", "法律"],
    ["弁護士法人ベリーベスト法律事務所", "https://www.vbest.jp/", "リスティング広告", "0120-773-048", "離婚・相続・労働問題相談", "弁護士 無料相談 離婚 相続 おすすめ 費用", "法律"],
    ["株式会社エイブル（ABLE）", "https://www.able.co.jp/", "リスティング広告", "0120-558-008", "賃貸物件仲介・部屋探し", "賃貸 部屋探し 仲介手数料 安い おすすめ 物件", "不動産"],
    ["株式会社アパマンショップ（APAMAN）", "https://www.apaman.jp/", "リスティング広告", "0120-123-817", "賃貸・不動産仲介サービス", "賃貸 アパート マンション 物件 検索 おすすめ", "不動産"],
    ["株式会社ピタットハウス（スターツピタットハウス）", "https://www.pitatthome.co.jp/", "リスティング広告", "0120-500-600", "賃貸・売買不動産仲介", "賃貸 売買 不動産 マンション 探し方 おすすめ", "不動産"],
    ["メットライフ生命保険株式会社", "https://www.metlife.co.jp/", "リスティング広告", "0120-655-955", "外資系生命保険・医療保険", "生命保険 おすすめ 外資 医療 比較 終身保険", "保険"],
    ["株式会社セコム損害保険", "https://www.secom-sonpo.co.jp/", "リスティング広告", "0120-201-815", "火災保険・地震保険", "火災保険 おすすめ 安い 比較 地震保険 見積もり", "保険"],
    ["株式会社FiT Easy（フィットイージー）", "https://fiteasy.jp/", "リスティング広告", "0120-963-669", "低価格24時間フィットネスジム", "ジム 安い 24時間 おすすめ 月額 近く フィットネス", "フィットネス"],
    ["株式会社ワークポート（WORKPORT）", "https://www.workport.co.jp/", "リスティング広告", "0120-987-545", "IT・Web系転職エージェント", "転職 IT エンジニア おすすめ エージェント 未経験", "転職"],
    ["株式会社MS-Japan（MS-Japan）", "https://www.jmsc.co.jp/", "リスティング広告", "0120-813-810", "管理部門・士業特化転職", "転職 経理 財務 法務 管理部門 おすすめ 士業", "転職"],
    ["株式会社TAC（TAC）", "https://www.tac-school.co.jp/", "リスティング広告", "0120-509-117", "資格試験・公務員対策スクール", "資格 公務員 簿記 税理士 おすすめ 合格率 スクール", "教育"],
    ["株式会社LEC東京リーガルマインド（LEC）", "https://www.lec-jp.com/", "リスティング広告", "0120-509-108", "資格・司法試験対策スクール", "資格 司法書士 行政書士 おすすめ 合格 通信講座", "教育"],
    ["学校法人大原学園（大原）", "https://www.o-hara.ac.jp/", "リスティング広告", "0120-333-700", "簿記・公務員・資格専門学校", "簿記 公務員 資格 専門学校 おすすめ 合格実績", "教育"],
    ["株式会社銀座カラー（銀座カラー）", "https://ginzacolor.jp/", "リスティング広告", "0120-218-400", "女性向け脱毛サロン", "脱毛 サロン おすすめ 安い 全身 料金 銀座", "美容サロン・脱毛"],
    ["株式会社エスクリ（ESCRIT）", "https://escrit.jp/", "リスティング広告", "0120-899-419", "ウェディング・結婚式場運営", "結婚式 費用 安い おすすめ 式場 ウェディング 少人数", "ウェディング"],
    ["株式会社IBJ（日本結婚相談所連盟）", "https://www.ibj.co.jp/", "リスティング広告", "0120-800-163", "結婚相談所・婚活マッチング", "結婚相談所 おすすめ 費用 口コミ 婚活 成婚率", "婚活"],
    ["株式会社パートナーエージェント", "https://www.p-a.jp/", "リスティング広告", "0120-919-987", "結婚相談所・婚活支援", "結婚相談所 おすすめ 料金 婚活 成婚 30代 40代", "婚活"],
    ["株式会社ツヴァイ（ZWEI）", "https://www.zwei.com/", "リスティング広告", "0120-214-981", "結婚相談所・婚活サービス", "結婚相談所 おすすめ 比較 料金 婚活 年齢層", "婚活"],
    ["株式会社ヨシケイ（ヨシケイ）", "https://yoshikei-dvlp.co.jp/", "リスティング広告", "0120-771-841", "夕食食材宅配・ミールキット", "食材宅配 ミールキット おすすめ 夕食 安い 時短 料理", "食品EC"],
    ["生活協同組合連合会パルシステム（パルシステム）", "https://www.pal-system.co.jp/", "リスティング広告", "0120-572-633", "有機・安心食材宅配コープ", "食材宅配 有機 安心 コープ おすすめ 野菜 無農薬", "食品EC"],
    ["株式会社ワタミ（ワタミの宅食）", "https://www.watami-takushoku.co.jp/", "リスティング広告", "0120-924-444", "高齢者・シニア向け宅配弁当", "宅配弁当 高齢者 おすすめ シニア 一人暮らし 安い", "食品EC"],
    ["株式会社アリさんマークの引越社", "https://www.the-arisanmark.com/", "リスティング広告", "0120-535-595", "引越しサービス全般", "引越し 見積もり おすすめ 安い 比較 料金 単身", "引越し"],
    ["株式会社サカイ引越センター", "https://www.hikkoshi-sakai.co.jp/", "リスティング広告", "0120-00-8378", "引越しサービス全般", "引越し おすすめ 見積もり 安い 口コミ サービス", "引越し"],
    ["株式会社アーク引越センター", "https://www.the-ark.jp/", "リスティング広告", "0120-33-8383", "引越しサービス全般", "引越し 料金 安い 見積もり おすすめ 口コミ", "引越し"],
    ["株式会社インターネットイニシアティブ（IIJmio）", "https://www.iijmio.jp/", "リスティング広告", "0120-919-233", "格安SIM・スマホ", "格安SIM おすすめ 安い 料金 比較 乗り換え 月額", "通信"],
    ["ソニーネットワークコミュニケーションズ株式会社（NUROモバイル）", "https://mobile.nuro.jp/", "リスティング広告", "0570-960-850", "格安SIM・データSIM", "格安SIM おすすめ 安い 月額 データ無制限 比較", "通信"],
    ["株式会社BIGLOBE（BIGLOBEモバイル）", "https://mobile.biglobe.ne.jp/", "リスティング広告", "0120-68-0962", "格安SIM・スマホ通信", "格安SIM おすすめ BIGLOBE 安い 乗り換え ポイント", "通信"],
    ["株式会社クラシアン", "https://www.crasian.co.jp/", "リスティング広告", "0120-08-1010", "水道・トイレ緊急修理サービス", "水道修理 緊急 24時間 水漏れ おすすめ 料金 トイレ", "住宅・修理"],
    ["医療法人社団AGAスキンクリニック", "https://www.aga-skinclinic.jp/", "リスティング広告", "0120-976-878", "AGA治療・薄毛治療クリニック", "AGA 薄毛 治療 クリニック おすすめ 料金 効果", "美容クリニック・AGA"],
    ["医療法人社団クリニックフォア（クリニックフォア）", "https://clinicfor.life/", "リスティング広告", "0120-958-539", "オンライン診療・ピル処方", "オンライン診療 ピル 処方 おすすめ 月額 安い", "医療サービス"],
    ["楽天カード株式会社（楽天カード）", "https://card.rakuten.co.jp/", "リスティング広告", "0120-86-6010", "クレジットカード・楽天ポイント", "クレジットカード おすすめ ポイント 楽天 年会費無料", "カード"],
    ["株式会社クレディセゾン（セゾンカード）", "https://www.saisoncard.co.jp/", "リスティング広告", "0120-084-009", "クレジットカード・セゾンポイント", "クレジットカード おすすめ セゾン ポイント 優待 永久不滅", "カード"],
    ["株式会社レイク（レイク）", "https://www.lake.jp/", "リスティング広告", "0120-166-108", "カードローン・フリーローン", "カードローン おすすめ 即日 審査 金利 消費者金融", "金融・ローン"],
    ["株式会社ダスキン（ダスキン）", "https://www.duskin.co.jp/", "リスティング広告", "0120-100-100", "ハウスクリーニング・レンタルモップ", "ハウスクリーニング おすすめ 料金 定期 清掃 プロ", "家事代行"],
    ["株式会社エネカリ（エネカリ）", "https://enecarry.jp/", "リスティング広告", "0120-979-680", "太陽光・蓄電池サービス", "太陽光 蓄電池 おすすめ 初期費用0 補助金 節電", "エネルギー"],
    ["株式会社夢真ビーネックスグループ（俺の転職）", "https://www.yumeshin.co.jp/", "リスティング広告", "0120-327-930", "建設・施工管理特化転職", "施工管理 転職 おすすめ 建設 資格 求人 年収アップ", "転職"],
    ["株式会社クリーク・アンド・リバー社（C&R社）", "https://www.c-r.com/", "リスティング広告", "0120-414-190", "クリエイター・医師特化転職", "転職 クリエイター 専門職 高収入 おすすめ 医師 弁護士", "転職"],
    ["株式会社ジェイック（JAIC）", "https://www.jaic-college.jp/", "リスティング広告", "0120-117-227", "未経験・フリーター向け就職支援", "就職 未経験 フリーター おすすめ 内定 研修 就活", "就職支援"],
    ["株式会社DYM（DYM就職）", "https://www.dshu.jp/", "リスティング広告", "0120-974-174", "若手・既卒向け就職支援", "就職 既卒 フリーター おすすめ 未経験 就活 内定率", "就職支援"],
    ["株式会社コジマ（Kojima）", "https://www.kojima.net/", "リスティング広告", "0120-210-401", "家電・PCネット通販", "家電 ネット通販 おすすめ 安い PC テレビ 比較", "家電EC"],
    ["株式会社ブックオフコーポレーション（BOOKOFF ONLINE）", "https://www.bookoffonline.co.jp/", "リスティング広告", "0120-976-066", "中古本・ゲーム・CD買取・通販", "買取 本 ゲーム 安い 中古 おすすめ ブックオフ", "買取・リユース"],
    ["株式会社駿河屋（Suruga-ya）", "https://www.suruga-ya.jp/", "リスティング広告", "054-266-4310", "ゲーム・アニメ・フィギュア買取EC", "ゲーム 買取 フィギュア アニメ 中古 おすすめ 高額", "買取・リユース"],
    ["株式会社IDOM（ガリバー）", "https://221616.com/", "リスティング広告", "0120-011-788", "中古車買取・販売", "中古車 買取 おすすめ 査定 高く売る 無料 車", "中古車"],
    ["株式会社カーセブン（CarShop7）", "https://www.carseven.co.jp/", "リスティング広告", "0120-081-082", "中古車買取・販売チェーン", "中古車 買取 高額 おすすめ 査定 無料 比較", "中古車"],
    ["株式会社MOTA（MOTA car）", "https://car.mota.jp/", "リスティング広告", "0120-198-855", "AI中古車買取サービス", "車 買取 AI 高額 おすすめ 査定 一括比較", "中古車"],
    ["株式会社Coo&RIKU（Coo&RIKU）", "https://www.cooandriku.co.jp/", "リスティング広告", "0120-28-3564", "ペットショップ・トリミングサービス", "子犬 子猫 ペットショップ おすすめ トリミング 近く", "ペット"],
    ["株式会社エクシング（JOYSOUND）", "https://www.joysound.com/", "リスティング広告", "0120-973-973", "カラオケ配信・家庭用カラオケ", "カラオケ 家庭用 おすすめ 配信 月額 曲数 安い", "エンタメ"],
    ["株式会社イオンクレジットサービス（イオンカード）", "https://www.aeon.co.jp/card/", "リスティング広告", "0120-72-8765", "クレジットカード・WAONポイント", "クレジットカード おすすめ イオン WAON ポイント 年会費無料", "カード"],
    ["株式会社さとふる（さとふる）", "https://www.satofull.jp/", "リスティング広告", "0120-530-440", "ふるさと納税ポータルサイト", "ふるさと納税 おすすめ 返礼品 比較 申し込み 節税", "ふるさと納税"],
    ["株式会社トラストバンク（ふるさとチョイス）", "https://www.furusato-tax.jp/", "リスティング広告", "0120-300-031", "ふるさと納税サイト", "ふるさと納税 おすすめ 人気 返礼品 節税 限度額", "ふるさと納税"],
]

assert len(rows) == 50, f"50社ではなく{len(rows)}社です"

out = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-05-24.xlsx"
make_xlsx(out, rows)
