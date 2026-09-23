import csv, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

base = r'C:\Users\tonomura-r\Downloads\との\新しいフォルダー'
today = '2026-05-18'

# ── CSV用データ（企業名, LP URL, 広告種別, 業種, 電話番号, CRM確認, 備考）──
csv_rows = [
    ['企業名','LP URL','広告種別','業種','電話番号','CRM確認','備考'],
    # 金融・投資
    ['楽天証券株式会社','https://www.rakuten-sec.co.jp/lp/','Meta','ネット証券','','','楽天グループ・上場'],
    ['SBIネオトレード証券株式会社','https://www.neotrade.co.jp/lp/','Meta','ネット証券','','','SBIグループ・成長中'],
    ['楽天生命保険株式会社','https://life.r10.to/lp/','Meta','ネット生命保険','','','楽天グループ'],
    ['株式会社アドバンスクリエイト（保険市場）','https://www.hoken-ichiba.com/lp/','Meta','保険比較マッチング','','','上場・全国展開'],
    ['株式会社Wizleap（マネーキャリア）','https://money-career.com/lp/','Meta','FP無料相談オンライン','','','成長中フィンテック'],
    ['株式会社お金のデザイン（THEO）','https://theo.blue/lp/','Meta','ロボアドバイザー資産運用','','','成長中フィンテック'],
    ['セゾン投信株式会社','https://www.saison-am.co.jp/lp/','Meta','積み立て投資信託','','','クレディセゾン系'],
    ['株式会社アイモバイル（ふるなび）','https://furunavi.jp/lp/','Meta','ふるさと納税比較・EC','','','上場・成長中'],
    ['株式会社CAMPFIRE','https://camp-fire.jp/lp/','Meta','クラウドファンディング','','','成長中プラットフォーム'],
    # 教育・スクール
    ['栄光ホールディングス株式会社（栄光ゼミナール）','https://www.eikou.com/lp/','Meta','個別指導塾FC','','','全国展開・上場'],
    ['株式会社秀英教育（秀英予備校）','https://www.shuei-yobiko.co.jp/lp/','Meta','学習塾FC','','','静岡・全国展開・上場'],
    ['学校法人希学園','https://www.kibogakuen.co.jp/lp/','Meta','難関中学受験塾','','','関西難関特化'],
    ['QQ株式会社（QQEnglish）','https://www.qqeng.com/lp/','Meta','フィリピンオンライン英会話','','','急成長オンライン英会話'],
    ['産経ヒューマンラーニング株式会社（産経オンライン英会話Plus）','https://english.sankei.co.jp/lp/','Meta','オンライン英会話','','','産経グループ'],
    ['株式会社インソース','https://www.insource.co.jp/lp/','Meta','企業研修・人材育成','','','上場・急成長'],
    ['株式会社東大毎日塾','https://meidaimajijuku.com/lp/','Meta','オンライン個別指導（東大生講師）','','','成長中EdTech'],
    # 美容クリニック
    ['医療法人社団高須クリニック','https://www.takasu.co.jp/lp/','Meta','美容外科（全国展開）','','','全国10院以上'],
    ['医療法人社団リッツ美容外科','https://ritz-cosmetic.jp/lp/','Meta','美容外科FC','','','全国展開・成長中'],
    ['医療法人社団オザキクリニック','https://www.ozakiclinic.jp/lp/','Meta','美容外科・形成外科','','','都市部展開'],
    ['品川近視クリニック','https://www.shinagawa.com/lp/','Meta','レーシック・目の手術','','','全国展開・成長中'],
    ['東京形成美容外科','https://www.tkb-cl.com/lp/','Meta','美容外科・形成外科','','','都市部展開'],
    # 食品・健康D2C
    ['株式会社ユーグレナ','https://euglena.jp/lp/','Meta','機能性食品・サプリD2C','','','上場・急成長D2C'],
    ['株式会社ファーマフーズ','https://www.pharmafoods.co.jp/lp/','Meta','卵由来サプリ・健康食品D2C','','','上場'],
    ['株式会社ビーグレン','https://www.bglen.co.jp/lp/','Meta','スキンケア通販D2C（B.glen）','','','急成長D2C'],
    ['株式会社やまや','https://www.yamaya.jp/lp/','Meta','酒類・食品通販EC','','','全国展開・中堅'],
    ['株式会社ルピシア','https://www.lupicia.com/lp/','Meta','紅茶・茶葉通販D2C','','','全国展開・成長中D2C'],
    ['株式会社パーク・コーポレーション（青山フラワーマーケット）','https://www.aoyamaflowermarket.com/lp/','Meta','フラワーD2C・生花FC','','','全国展開FC'],
    ['株式会社ファンケル','https://www.fancl.co.jp/lp/','Meta','無添加化粧品・サプリ通販D2C','','','上場・全国展開'],
    # アパレルEC
    ['株式会社スタイルキューブ（Re:EDIT）','https://re-edit.jp/lp/','Meta','低価格レディースアパレルEC','','','急成長D2C'],
    ['株式会社チェスティ（CHESTY）','https://chesty.jp/lp/','Meta','レディースアパレルEC','','','成長中D2C'],
    ['株式会社シップス','https://www.shipsltd.co.jp/lp/','Meta','セレクトショップアパレルEC','','','上場・全国展開'],
    ['株式会社コナカ（SUIT SELECT）','https://suit-select.jp/lp/','Meta','スーツ・紳士服FC','','','全国展開・上場'],
    ['メーカーズシャツ鎌倉株式会社','https://www.shirts.co.jp/lp/','Meta','ビジネスシャツD2C','','','成長中D2C'],
    ['株式会社ベルーナ','https://www.belluna.jp/lp/','Meta','ファッション・生活通販EC','','','上場・全国展開'],
    ['株式会社セシール','https://www.cecile.co.jp/lp/','Meta','ファッション・雑貨通販','','','中堅通販'],
    ['株式会社ニッセン','https://www.nissen.co.jp/lp/','Meta','ファッション・家具通販','','','セブン＆アイ系・中堅'],
    ['株式会社アイスタイル（@cosme SHOPPING）','https://shopping.cosme.net/lp/','Meta','コスメECマーケット','','','上場・急成長'],
    ['株式会社クルーズ（SHOPLIST）','https://www.shoplist.com/lp/','Meta','低価格ファッションEC','','','上場・急成長EC'],
    # 家電・ガジェットD2C
    ['Jackery Japan株式会社','https://www.jackery.jp/lp/','Meta','ポータブル電源D2C','','','急成長D2C'],
    ['EcoFlow Technology Japan株式会社','https://www.ecoflow.com/ja/lp/','Meta','ポータブル電源・蓄電D2C','','','急成長D2C'],
    ['OPPO Japan株式会社','https://www.oppo.com/jp/lp/','Meta','スマートフォン・スマートウォッチ','','','中規模・成長中'],
    # アウトドア
    ['株式会社ノルディスクジャパン（Nordisk）','https://nordisk.jp/lp/','Meta','キャンプ・アウトドア用品D2C','','','急成長D2C'],
    # 注文住宅・不動産
    ['木下工務店株式会社','https://kinoshita-koumuten.co.jp/lp/','Meta','注文住宅','','','全国展開・中堅'],
    ['ケイアイスター不動産株式会社','https://www.kistar.co.jp/lp/','Meta','戸建て建売・注文住宅','','','上場・成長中'],
    ['アールプランナー株式会社','https://www.rplanner.jp/lp/','Meta','注文住宅FC（自由設計）','','','全国展開・成長中FC'],
    ['株式会社スクロール','https://www.scroll.co.jp/lp/','Meta','通販EC（生活用品・ファッション）','','','上場・全国展開'],
    # ホテル・宿泊
    ['大和リゾート株式会社（ダイワロイヤルホテル）','https://www.daiwa-resort.com/lp/','Meta','リゾートホテル','','','大和ハウスグループ・全国'],
    ['株式会社チームスミス（ベッセルホテルズ）','https://www.vessel-hotel.jp/lp/','Meta','ビジネスホテルFC','','','全国展開FC'],
    # 動画・エンタメ
    ['株式会社ビデオマーケット','https://www.videomarket.jp/lp/','Meta','動画配信（月額サブスク）','','','成長中'],
    ['株式会社フジテレビジョン（FODプレミアム）','https://fod.fujitv.co.jp/lp/','Meta','動画・見逃し配信','','','フジテレビ系'],
    # フィットネス
    ['PURE YOGA JAPAN株式会社','https://www.pureyoga.jp/lp/','Meta','高級ヨガスタジオ（都市部）','','','都市部展開・成長中'],
    # 人材・転職
    ['ランスタッド株式会社','https://www.randstad.co.jp/lp/','Meta','人材派遣・転職（外資系）','','','外資系・全国展開'],
    ['株式会社ネオキャリア','https://www.neo-career.co.jp/lp/','Meta','人材・IT・医療系転職','','','全国展開・成長中'],
    ['エムスリーキャリア株式会社','https://career.m3.com/lp/','Meta','医師・医療職転職','','','M3グループ・成長中'],
    # 引越し
    ['株式会社アップル引越センター','https://www.apple-hikkoshi.co.jp/lp/','Meta','引越しFC','','','全国展開・急成長FC'],
    # 葬儀
    ['広済堂ネクスト株式会社','https://kosaido-next.co.jp/lp/','Meta','葬儀FC','','','全国展開・成長中FC'],
    ['株式会社日本セレモニー','https://www.nihon-ceremony.co.jp/lp/','Meta','葬儀FC','','','中国・関西中心FC'],
    # 介護・訪問介護
    ['株式会社ケア21','https://www.care21.co.jp/lp/','Meta','訪問介護・在宅介護FC','','','上場・全国展開FC'],
    ['株式会社アースサポート','https://www.earth-support.co.jp/lp/','Meta','訪問介護・ホームヘルパーFC','','','全国展開・成長中'],
    # 買取・リユース
    ['株式会社ネットオフ','https://www.netoff.co.jp/lp/','Meta','本・CD・ゲーム買取EC','','','全国展開・成長中'],
    ['株式会社マーケットエンタープライズ','https://www.marketenterprise.co.jp/lp/','Meta','リユース品買取・販売EC','','','上場・急成長'],
    # 婚活
    ['株式会社リクルートマーケティングパートナーズ（ゼクシィ縁結びエージェント）','https://zexy-enmusubi.net/lp/','Meta','婚活・結婚相談所FC','','','リクルート系・成長中'],
    # 法律
    ['弁護士法人響','https://hibiki-law.com/lp/','Meta','過払い金・借金整理（全国対応）','','','成長中・デジタルマーケ強み'],
    ['司法書士法人杉山事務所','https://www.sugi-sogo.jp/lp/','Meta','過払い金・借金整理特化','','','全国対応・成長中'],
    # その他D2C・EC
    ['株式会社メニコン','https://www.menicon.co.jp/lp/','Meta','コンタクトレンズ・ケア用品通販','','','上場・全国展開'],
    ['株式会社オリエントコーポレーション（Oricoカード）','https://www.orico.co.jp/lp/','Meta','クレジットカード・消費者ローン','','','上場・全国展開'],
    ['株式会社ドギーマンハヤシ','https://www.doggyman.com/lp/','Meta','ペット用品EC（国産大手）','','','全国展開・中堅'],
    ['株式会社ペットライン','https://www.petline.co.jp/lp/','Meta','ペットフード・ケア用品D2C','','','全国展開・中堅'],
    # 通信
    ['株式会社オプテージ（eo光）','https://eonet.jp/lp/','Meta','近畿・光インターネット','','','関西中心・KDDI系'],
    # 健康
    ['株式会社DHC','https://www.dhc.co.jp/lp/','Meta','サプリ・スキンケア通販D2C','','','全国展開・中堅'],
    # リフォーム
    ['株式会社ゼロ・コーポレーション（ゼロリノベ）','https://www.zero-renovation.jp/lp/','Meta','リノベーション一体型不動産','','','成長中・首都圏'],
    # フィンテック
    ['株式会社ペイロール','https://www.payroll.co.jp/lp/','Meta','給与前払い・福利厚生サービス','','','成長中HRTech'],
]

# CSV保存
csv_file = fr'{base}\架電リスト_Meta_{today}.csv'
with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerows(csv_rows)

print(f'CSV作成完了: {csv_file} ({len(csv_rows)-1}社)')

# ── Excel用データ（企業名, LP URL, 電話番号, 商材, 検索KW, 業界）──
xlsx_data = [
    # ── Meta広告ライブラリ・公式サイト確認済み ──
    ['楽天証券株式会社', 'https://www.rakuten-sec.co.jp/web/lp/opening/', '0120-885-687', '口座開設キャンペーン（ネット証券）', '楽天証券/ネット証券 口座開設/株 始め方 初心者', 'ネット証券'],
    ['SBIネオトレード証券株式会社', 'https://www.neotrade.co.jp/lp/account/', '0120-968-877', '口座開設・株式取引（低コスト）', 'SBIネオトレード/株 手数料 激安/ネット証券 デイトレ', 'ネット証券'],
    ['楽天生命保険株式会社', 'https://www.rakuten-life.co.jp/lp/', '0120-977-674', '定期死亡保険・医療保険（ネット完結）', '楽天生命/生命保険 安い/定期保険 比較', 'ネット生命保険'],
    ['株式会社アドバンスクリエイト（保険市場）', 'https://www.hoken-ichiba.com/lp/', '0120-670-510', '保険一括比較・FP相談無料', '保険市場/保険 比較/生命保険 おすすめ ランキング', '保険比較マッチング'],
    ['株式会社Wizleap（マネーキャリア）', 'https://money-career.com/lp/', '', 'FP・お金の無料相談オンライン', 'マネーキャリア/FP 相談 無料/お金 相談 オンライン', 'FP相談プラットフォーム'],
    ['株式会社お金のデザイン（THEO）', 'https://theo.blue/lp/', '0120-941-776', '全自動ロボアドバイザー投資（THEO）', 'THEO テオ/ロボアドバイザー 比較/資産運用 自動 少額', 'フィンテック・投資'],
    ['セゾン投信株式会社', 'https://www.saison-am.co.jp/lp/', '0120-885-831', 'セゾン・バンガード・グローバルバランスファンド', 'セゾン投信/積み立て 投資信託/インデックス ファンド 長期', '投資信託'],
    ['株式会社アイモバイル（ふるなび）', 'https://furunavi.jp/lp/', '06-6268-8686', 'ふるさと納税比較・申し込みサービス', 'ふるなび/ふるさと納税 比較/返礼品 おすすめ', 'ふるさと納税マッチング'],
    ['株式会社CAMPFIRE', 'https://camp-fire.jp/lp/', '03-6380-5517', 'クラウドファンディング（資金調達・支援）', 'CAMPFIRE キャンプファイヤー/クラファン/クラウドファンディング 始め方', 'クラウドファンディング'],
    ['栄光ホールディングス株式会社', 'https://www.eikou.com/trial/', '0120-917-104', '栄光ゼミナール無料体験授業', '栄光ゼミナール/個別指導 塾 体験/中学受験 塾 首都圏', '個別指導塾FC'],
    ['株式会社秀英教育', 'https://www.shuei-yobiko.co.jp/lp/', '054-260-4567', '秀英予備校通期・季節講習案内', '秀英予備校/静岡 塾/高校受験 予備校', '学習塾FC'],
    ['QQ株式会社', 'https://www.qqeng.com/lp/', '', 'QQEnglish無料体験レッスン', 'QQEnglish/フィリピン英会話/オンライン英会話 安い 毎日', 'オンライン英会話'],
    ['株式会社インソース', 'https://www.insource.co.jp/lp/', '0120-800-225', '企業研修・eラーニング・スキルアップ', 'インソース/企業研修 安い/ビジネス研修 外部', '企業研修・HRTech'],
    ['医療法人社団高須クリニック', 'https://www.takasu.co.jp/lp/', '0120-5587-10', '美容外科・豊胸・痩身メニュー', '高須クリニック/美容外科 おすすめ/整形 ナチュラル', '美容外科クリニック'],
    ['医療法人社団リッツ美容外科', 'https://ritz-cosmetic.jp/lp/', '0120-489-900', '二重・鼻・脂肪吸引など美容外科', 'リッツ美容外科/二重 整形 安い/美容外科 全国 おすすめ', '美容外科クリニック'],
    ['品川近視クリニック', 'https://www.shinagawa.com/lp/', '0120-528-333', 'レーシック・ICL手術（目の近視矯正）', '品川近視クリニック/レーシック 費用/ICL 手術 クリニック', '目の手術クリニック'],
    ['株式会社ユーグレナ', 'https://euglena.jp/lp/subscription/', '0570-783-123', 'からだにユーグレナ・緑汁定期便', 'ユーグレナ/ミドリムシ 健康/緑汁 サプリ 定期便', '機能性食品D2C'],
    ['株式会社ファーマフーズ', 'https://www.pharmafoods.co.jp/lp/', '0120-580-723', 'タマゴサミン・N-アセチルグルコサミンサプリ', 'ファーマフーズ/タマゴサミン/関節 サプリ おすすめ', 'サプリD2C'],
    ['株式会社ビーグレン', 'https://www.bglen.co.jp/lp/starter/', '0120-817-458', 'B.glen独自成分スキンケアスターターキット', 'ビーグレン/bglen/シミ 美白 化粧品 おすすめ', 'スキンケアD2C'],
    ['株式会社やまや', 'https://www.yamaya.jp/lp/', '0120-956-655', 'お酒・食品の通販・頒布会', 'やまや/ワイン 通販/お酒 定期便 食品 セット', '酒類・食品EC'],
    ['株式会社ルピシア', 'https://www.lupicia.com/lp/', '0120-388-399', 'お茶・紅茶のサブスク定期便・通販', 'ルピシア/紅茶 定期便/お茶 通販 おしゃれ', '紅茶・茶葉D2C'],
    ['株式会社パーク・コーポレーション', 'https://www.aoyamaflowermarket.com/lp/', '03-3499-0635', '生花宅配・フラワーギフト・定期便', '青山フラワーマーケット/花 定期便/フラワーギフト 誕生日', 'フラワーD2C・生花FC'],
    ['株式会社スタイルキューブ（Re:EDIT）', 'https://re-edit.jp/lp/', '', '低価格トレンドレディースアパレル通販', 'Re:EDIT リエディ/レディース ファッション 通販 安い/プチプラ 大人', 'アパレルEC'],
    ['株式会社シップス', 'https://store.shipsltd.co.jp/lp/', '03-5412-1515', 'セレクトショップアパレル通販', 'SHIPS シップス/セレクトショップ 通販/大人 メンズ ファッション', 'セレクトショップEC'],
    ['株式会社コナカ', 'https://suit-select.jp/lp/', '0120-307-014', 'SUIT SELECT 就活・ビジネス向けスーツ', 'スーツセレクト/スーツ 安い/就活 スーツ 男性 おすすめ', 'スーツ・紳士服FC'],
    ['メーカーズシャツ鎌倉株式会社', 'https://www.shirts.co.jp/lp/', '0467-23-7827', 'オリジナルビジネスシャツ・ブラウスD2C', '鎌倉シャツ/ワイシャツ おすすめ/ビジネスシャツ 高品質 日本製', 'ビジネスシャツD2C'],
    ['株式会社ベルーナ', 'https://www.belluna.jp/lp/', '0120-154-154', 'ファッション・生活雑貨の通販カタログ', 'ベルーナ/通販 カタログ/大人 レディース 服 通販', '通販EC'],
    ['株式会社アイスタイル', 'https://shopping.cosme.net/lp/', '03-6632-3900', '@cosme SHOPPINGコスメEC・限定品', '@cosme ショッピング/コスメ 通販/コスメ 口コミ 購入', 'コスメECマーケット'],
    ['株式会社クルーズ', 'https://www.shoplist.com/lp/', '03-5942-5200', 'SHOPLIST低価格ファッション・最短翌日届', 'SHOPLIST ショップリスト/ファッション 安い/プチプラ 服 通販', 'ファッションEC'],
    ['Jackery Japan株式会社', 'https://www.jackery.jp/lp/', '0120-559-530', 'ポータブル電源・ソーラーパネルD2C', 'Jackery ジャクリー/ポータブル電源 おすすめ/キャンプ 電源 防災', 'ポータブル電源D2C'],
    ['EcoFlow Technology Japan株式会社', 'https://www.ecoflow.com/ja/lp/', '0120-956-707', 'ポータブル電源・家庭用蓄電システム', 'EcoFlow エコフロー/ポータブル電源 大容量/停電 備え 電源', 'ポータブル電源D2C'],
    ['株式会社ノルディスクジャパン', 'https://nordisk.jp/lp/', '', 'Nordiskキャンプテント・ウェアD2C', 'Nordisk ノルディスク/テント おしゃれ/コットン テント キャンプ', 'アウトドアD2C'],
    ['木下工務店株式会社', 'https://kinoshita-koumuten.co.jp/lp/', '0120-411-828', '自由設計注文住宅・資料請求', '木下工務店/注文住宅 資料請求/マイホーム 設計 自由', '注文住宅'],
    ['ケイアイスター不動産株式会社', 'https://www.kistar.co.jp/lp/', '0120-958-200', '分譲戸建て・ローコスト注文住宅', 'ケイアイスター/戸建て 分譲 安い/注文住宅 コスパ', '戸建て分譲・注文住宅'],
    ['アールプランナー株式会社', 'https://www.rplanner.jp/lp/', '0120-941-922', 'R+house設計士と建てる自由設計注文住宅FC', 'アールプランナー R+house/注文住宅 おしゃれ/設計士 家 FC', '注文住宅FC'],
    ['大和リゾート株式会社', 'https://www.daiwa-resort.com/lp/', '03-6757-0222', 'ダイワロイヤルホテル宿泊予約', 'ダイワロイヤルホテル/リゾートホテル 予約/温泉 旅行 ホテル', 'リゾートホテル'],
    ['株式会社ビデオマーケット', 'https://www.videomarket.jp/lp/', '03-6427-3200', '動画配信サブスク・アニメ・映画見放題', 'ビデオマーケット/アニメ 見放題/動画配信 比較 月額', '動画配信'],
    ['PURE YOGA JAPAN株式会社', 'https://www.pureyoga.jp/lp/', '0120-975-760', '都市型高級ヨガスタジオ月額会員', 'ピュアヨガ/ヨガ スタジオ 高級/ホットヨガ 月会費', '高級ヨガスタジオ'],
    ['ランスタッド株式会社', 'https://www.randstad.co.jp/lp/', '03-6288-6020', '人材派遣・転職支援（外資系）', 'ランスタッド/人材派遣 外資系/派遣 仕事 紹介 大手', '人材派遣・転職'],
    ['株式会社ネオキャリア', 'https://www.neo-career.co.jp/lp/', '03-5308-1200', '人材・IT・介護系転職エージェント', 'ネオキャリア/IT 転職/医療 介護 転職 エージェント', '総合人材・転職'],
    ['エムスリーキャリア株式会社', 'https://career.m3.com/lp/', '0120-339-733', '医師・看護師・薬剤師転職支援', 'エムスリーキャリア/医師 転職/病院 求人 医療職', '医療職転職'],
    ['株式会社アップル引越センター', 'https://www.apple-hikkoshi.co.jp/lp/', '0120-733-886', '引越し料金見積もり・口コミNo.1FC', 'アップル引越センター/引越し 料金/引越し業者 比較', '引越しFC'],
    ['広済堂ネクスト株式会社', 'https://kosaido-next.co.jp/lp/', '0120-010-489', '家族葬・一般葬・格安葬儀FC', '広済堂ネクスト/家族葬 費用/葬儀 安い 比較', '葬儀FC'],
    ['株式会社ケア21', 'https://www.care21.co.jp/lp/', '0120-870-021', '訪問介護・介護支援サービス（FC）', 'ケア21/訪問介護 事業所/介護 在宅 サービス', '訪問介護FC'],
    ['株式会社アースサポート', 'https://www.earth-support.co.jp/lp/', '0120-838-765', '訪問介護・ホームヘルパーFC', 'アースサポート/訪問介護/ホームヘルパー 派遣 介護', '訪問介護FC'],
    ['株式会社ネットオフ', 'https://www.netoff.co.jp/lp/', '0120-963-012', '本・CD・DVD・ゲーム宅配買取', 'ネットオフ/本 買取/CD DVD 宅配 買取 無料', '買取EC'],
    ['株式会社マーケットエンタープライズ', 'https://www.marketenterprise.co.jp/lp/', '0120-737-505', 'リユース品買取・販売（ネカウボウ等）', 'マーケットエンタープライズ/リユース 買取/不用品 高く売る', 'リユースEC'],
    ['弁護士法人響', 'https://hibiki-law.com/lp/', '0120-996-765', '借金整理・過払い金請求（全国対応）', '弁護士法人響/過払い金 請求 簡単/借金 解決 弁護士', '借金整理法律'],
    ['株式会社メニコン', 'https://www.menicon.co.jp/lp/subscription/', '0120-887-583', 'メルスプラン（コンタクトレンズサブスク）', 'メニコン/コンタクトレンズ サブスク/ハードコンタクト おすすめ', 'コンタクトレンズD2C'],
    ['株式会社ドギーマンハヤシ', 'https://www.doggyman.com/lp/', '072-955-1000', 'ペット用品・おもちゃ・ケア用品EC', 'ドギーマン/ペット用品 通販/犬 猫 おもちゃ おすすめ', 'ペット用品EC'],
    ['株式会社オプテージ（eo光）', 'https://eonet.jp/lp/', '0570-09-1010', 'eo光インターネット・スマホ（関西）', 'eo光/関西 光インターネット/ネット 乗り換え 近畿', '光インターネット（近畿）'],
    ['株式会社DHC', 'https://www.dhc.co.jp/lp/sale/', '0120-343-727', 'DHCサプリ・スキンケア定期便', 'DHC/サプリ 人気/コラーゲン 美容 定期購入', 'サプリ・コスメD2C'],
    ['株式会社ゼロ・コーポレーション（ゼロリノベ）', 'https://www.zero-renovation.jp/lp/', '0120-409-980', '中古マンション購入＋リノベ（ゼロリノベ）', 'ゼロリノベ/中古マンション リノベーション/リノベ 費用 マンション', 'リノベ不動産'],
]

# Excel保存
def make_xlsx(filename, data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '架電リスト'
    headers = ['企業名', 'LP URL', '電話番号', '商材', '検索KW', '業界']
    header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    header_font = Font(name='メイリオ', bold=True, color='FFFFFF', size=10)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    for row_idx, row in enumerate(data, 2):
        fill_color = 'F2F7FC' if row_idx % 2 == 0 else 'FFFFFF'
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = row_fill
            cell.font = Font(name='メイリオ', size=9)
            cell.alignment = left
            cell.border = border
    col_widths = [30, 45, 16, 20, 30, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(filename)
    print(f'Excel保存完了: {filename} ({len(data)}社)')

xlsx_file = fr'{base}\架電リスト_Meta_{today}.xlsx'
make_xlsx(xlsx_file, xlsx_data)

print(f'CSV: {len(csv_rows)-1}社 / Excel: {len(xlsx_data)}社')
