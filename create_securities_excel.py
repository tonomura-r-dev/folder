import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────
# 企業名 → URL マッピング（広告LP優先、なければHP）
# ──────────────────────────────────────────────
URL_MAP = {
    # ── ネット証券・個人向け（広告LP）
    "株式会社ＳＢＩ証券":                            "https://www.sbisec.co.jp/ETGate",
    "楽天証券株式会社":                              "https://www.rakuten-sec.co.jp/",
    "マネックス証券株式会社":                         "https://www.monex.co.jp/",
    "松井証券株式会社":                              "https://www.matsui.co.jp/",
    "ＧＭＯクリック証券株式会社":                     "https://www.click-sec.com/",
    "ＰａｙＰａｙ証券株式会社":                       "https://www.paypay-sec.co.jp/",
    "大和コネクト証券株式会社":                       "https://connect.daiwa.jp/",
    "ウェルスナビ株式会社":                           "https://www.wealthnavi.com/lp/",
    "株式会社ＳＢＩネオトレード証券":                  "https://www.neotrade.co.jp/",
    "LINE証券株式会社":                              "https://line-sec.co.jp/",
    "tsumiki証券株式会社":                           "https://tsumiki-sec.com/",
    "CHEER証券株式会社":                             "https://cheersec.jp/",
    "moomoo証券株式会社":                            "https://www.moomoo.com/jp/",
    "ブルーモ証券株式会社":                           "https://bloomo.co.jp/",
    "株式会社ＦＯＬＩＯ":                             "https://folio-sec.com/",
    "株式会社ＤＭＭ．ｃｏｍ証券":                     "https://kabu.dmm.com/",
    "株式会社スマートプラス":                         "https://smartplus-sec.com/",
    "フィデリティ証券株式会社":                       "https://www.fidelity.co.jp/",
    "インタラクティブ・ブローカーズ証券株式会社":       "https://www.interactivebrokers.co.jp/",
    "株式会社マネーパートナーズ":                     "https://www.moneypartners.co.jp/",
    "三菱UFJ eスマート証券株式会社":                  "https://www.musmartsc.co.jp/",
    "三菱ＵＦＪモルガン・スタンレー証券株式会社":       "https://www.sc.mufg.jp/",
    "ＳＭＢＣ日興証券株式会社":                       "https://www.smbcnikko.co.jp/",
    "野村證券株式会社":                              "https://www.nomura.co.jp/",
    "大和証券株式会社":                              "https://www.daiwa.jp/",
    "みずほ証券株式会社":                             "https://www.mizuho-sc.com/",
    "東海東京証券株式会社":                           "https://www.tokaitokyo.co.jp/",
    "岡三証券株式会社":                              "https://www.okasan.co.jp/",

    # ── FX・CFD証券
    "サクソバンク証券株式会社":                       "https://www.home.saxo/ja-jp",
    "ＩＧ証券株式会社":                               "https://www.ig.com/jp/",
    "株式会社外為どっとコム":                         "https://www.gaitame.com/",
    "ＧＭＯ外貨株式会社":                             "https://www.gmo-click.com/",
    "株式会社外為オンライン":                         "https://www.gaitameonline.com/",
    "セントラル短資ＦＸ株式会社":                     "https://www.central-tanshifx.com/",
    "ＯＡＮＤＡ証券株式会社":                         "https://www.oanda.jp/",
    "ひまわり証券株式会社":                           "https://sec.himawari-group.co.jp/",
    "インヴァスト証券株式会社":                       "https://www.invast.jp/",
    "トレイダーズ証券株式会社":                       "https://www.traders.co.jp/",
    "デューカスコピー・ジャパン株式会社":              "https://www.dukascopy.jp/",
    "アヴァトレード・ジャパン株式会社":               "https://www.avatrade.co.jp/",
    "ゴールデンウェイ・ジャパン株式会社":              "https://www.gaitameonline.net/",
    "StoneX証券株式会社":                            "https://jp.stonex.com/",
    "株式会社ＦＸブロードネット":                      "https://www.fxbroadnet.com/",
    "ＪＦＸ株式会社":                                "https://www.jfx.co.jp/",
    "外為ファイネスト株式会社":                       "https://www.finest-fx.jp/",
    "ＳＢＩ　ＦＸ トレード株式会社":                  "https://www.sbifxt.co.jp/",
    "あさひマーケッツ株式会社":                       "https://www.asahimarkets.jp/",
    "フジトミ証券株式会社":                           "https://www.fujitomi.co.jp/",
    "サンワード証券株式会社":                         "https://www.sunward8.co.jp/",
    "スーパーファンド・ジャパン株式会社":              "https://www.superfund.jp/",
    "株式会社マネースクエア":                         "https://www.m2j.co.jp/",
    "くにうみＡＩ証券株式会社":                       "https://www.kuniumi-ai.co.jp/",
    "ヒロセ通商株式会社":                             "https://hirose-fx.co.jp/",
    "ＡＩゴールド証券株式会社":                       "https://ai-gold.co.jp/",
    "Plus500JP証券株式会社":                          "https://www.plus500.com/ja/",
    "ＫＯＹＯ証券株式会社":                            "https://koyo-sec.co.jp/",
    "ＦＯＲＥＸ　ＥＸＣＨＡＮＧＥ株式会社":            "https://www.forex-exchange.co.jp/",
    "株式会社アイネット証券":                         "https://www.inet-sec.co.jp/",

    # ── 暗号資産関連
    "GMOコイン株式会社":                             "https://coin.z.com/jp/",
    "楽天ウォレット株式会社":                         "https://wallet.rakuten.co.jp/",
    "株式会社bitFlyer":                              "https://bitflyer.com/ja-jp/",
    "ＳＢＩ　ＶＣトレード株式会社":                   "https://sbivc.co.jp/",
    "ビットトレード株式会社":                         "https://bittrade.co.jp/",

    # ── 外資系大手証券
    "ゴールドマン・サックス証券株式会社":              "https://www.goldmansachs.com/japan/",
    "ＪＰモルガン証券株式会社":                       "https://www.jpmorgan.com/jp/",
    "ＢｏｆＡ証券株式会社":                           "https://business.bofa.com/ja-jp/",
    "モルガン・スタンレーＭＵＦＧ証券株式会社":        "https://www.morganstanley.co.jp/",
    "ＵＢＳ証券株式会社":                             "https://www.ubs.com/jp/",
    "ＪＰモルガン・アセット・マネジメント株式会社":    "https://www.jpmorganasset.co.jp/",
    "ドイツ証券株式会社":                             "https://japan.db.com/",
    "シティグループ証券株式会社":                     "https://www.citigroup.com/japan/",
    "バークレイズ証券株式会社":                       "https://home.barclays/japan/",
    "ＨＳＢＣアセットマネジメント株式会社":            "https://www.assetmanagement.hsbc.com/ja",
    "HSBC証券株式会社":                              "https://www.hsbc.com/japan/",
    "ＢＮＰパリバ証券株式会社":                       "https://japan.bnpparibas.com/",
    "ＢＮＰパリバ・アセットマネジメント株式会社":      "https://www.bnpparibas-am.com/ja/",
    "ウェルズ・ファーゴ証券株式会社":                 "https://www.wellsfargo.com/japan/",
    "ソシエテ・ジェネラル証券株式会社":               "https://www.societegenerale.jp/",
    "ジェフリーズ・ジャパン・リミテッド（証券会社）":   "https://www.jefferies.com/japan/",
    "マッコーリー・キャピタル・セキュリティーズ・ジャパン・リミテッド（東京支店）": "https://www.macquarie.com/jp/",
    "ＤＢＳ証券株式会社":                             "https://www.dbs.com/jp/",
    "ANZ証券株式会社":                               "https://japan.anz.com/",
    "NAB証券株式会社":                               "https://www.nabcorporate.co.jp/",
    "スタンダードチャータード証券株式会社":            "https://www.sc.com/jp/",
    "ＣＬＳＡ証券株式会社":                           "https://www.clsa.com/japan/",
    "クレディ・アグリコル・セキュリティーズ・アジア・ビー・ヴィ（東京支店）": "https://www.ca-cib.co.jp/",

    # ── 国内地場・中堅証券
    "いちよし証券株式会社":                           "https://www.ichiyoshi.co.jp/",
    "水戸証券株式会社":                               "https://www.mito-sec.co.jp/",
    "極東証券株式会社":                               "https://www.kyokuto-sec.co.jp/",
    "あかつき証券株式会社":                           "https://www.akatsuki-sec.co.jp/",
    "立花証券株式会社":                               "https://www.tachibana-sec.co.jp/",
    "丸三証券株式会社":                               "https://www.marusan-sec.co.jp/",
    "東洋証券株式会社":                               "https://www.toyo-sec.co.jp/",
    "フィリップ証券株式会社":                         "https://www.phillip.co.jp/",
    "山二証券株式会社":                               "https://www.yamani.co.jp/",
    "山和証券株式会社":                               "https://www.yamawa-sec.co.jp/",
    "リテラ・クレア証券株式会社":                     "https://www.litera-claire.co.jp/",
    "みらい證券株式会社":                             "https://www.mirai-shoken.com/",
    "三田証券株式会社":                               "https://www.mitasec.co.jp/",
    "三木証券株式会社":                               "https://www.miki-sec.co.jp/",
    "三晃証券株式会社":                               "https://www.sanko-sec.co.jp/",
    "スターツ証券株式会社":                           "https://www.starts-sec.co.jp/",
    "むさし証券株式会社":                             "https://www.musashi-sec.co.jp/",
    "ちばぎん証券株式会社":                           "https://www.chibagin-sec.co.jp/",
    "上田八木証券株式会社":                           "https://www.ueda-yagi.co.jp/",
    "日産証券株式会社":                               "https://www.nissan-sec.co.jp/",
    "豊トラスティ証券株式会社":                       "https://www.yutaka-tokai.co.jp/",
    "岡地株式会社":                                   "https://www.okachi.jp/",
    "ばんせい証券株式会社":                           "https://www.bansei-sec.co.jp/",
    "ＤＢＪ証券株式会社":                             "https://www.dbj.jp/finance/securities/",
    "共和証券株式会社":                               "https://www.kyowa-shoken.co.jp/",
    "あい証券株式会社":                               "https://www.ai-sec.co.jp/",
    "中原証券株式会社":                               "https://www.nakahara-sec.co.jp/",
    "ニュース証券株式会社":                           "https://www.news-sec.co.jp/",
    "リーディング証券株式会社":                       "https://www.leadingsec.jp/",
    "しんきん証券株式会社":                           "https://www.shinkinsec.co.jp/",
    "株式会社ＦＰＧ証券":                             "https://www.fpg-sec.co.jp/",
    "ＪＰアセット証券株式会社":                       "https://www.jpas.jp/",
    "ＪＩＡ証券株式会社":                             "https://www.jiasec.co.jp/",
    "アーク証券株式会社":                             "https://www.ark-sec.co.jp/",
    "アイザワ証券株式会社":                           "https://www.aizawa-sec.co.jp/",
    "損保ジャパンＤＣ証券株式会社":                   "https://www.sjdc.co.jp/",
    "キャピタル・パートナーズ証券株式会社":            "https://www.capital-partners.co.jp/",
    "明和證券株式会社":                               "https://www.meiwa-sec.co.jp/",
    "ヘッジファンド証券株式会社":                     "https://www.hedgefundsec.co.jp/",
    "ＪＰＩＡアセット証券株式会社":                   "https://www.jpas.jp/",
    "株式会社だいこう証券ビジネス":                   "https://www.daiko-sec.co.jp/",
    "東武証券株式会社":                               "https://www.tobu-sec.co.jp/",
    "北辰物産株式会社":                               "https://www.hokushin-bussan.co.jp/",
    "スーパーファンド・ジャパン株式会社":              "https://www.superfund.jp/",
    "臼木証券株式会社":                               "https://www.usuki-sec.co.jp/",
    "武甲証券株式会社":                               "https://www.buko-sec.co.jp/",
    "富岡証券株式会社":                               "https://www.tomioka-sec.co.jp/",
    "丸國証券株式会社":                               "https://www.marukuni-sec.co.jp/",
    "みさき証券株式会社":                             "",
    "ＫＯＹＯ証券株式会社":                            "https://koyo-sec.co.jp/",
    "ＡＩゴールド証券株式会社":                       "https://ai-gold.co.jp/",
    "ＳＡＭＵＲＡＩ証券株式会社":                     "https://samurai-sec.co.jp/",
    "Jトラストグローバル証券株式会社":                 "https://www.jtrust-global-sec.co.jp/",
    "エスピーシー証券株式会社":                       "https://www.spcsec.co.jp/",
    "ウィブル証券株式会社":                           "https://www.webull.co.jp/",
    "ＰＷＭ日本証券株式会社":                         "https://www.pwmjapan.co.jp/",
    "ロードスター証券株式会社":                       "https://www.loadstar-capital.jp/",
    "FINX J証券株式会社":                             "https://finxj.co.jp/",
    "めぶき証券株式会社":                             "https://www.mebuki-sec.co.jp/",
    "浜銀ＴＴ証券株式会社":                           "https://www.hamagin-tt.co.jp/",
    "ぐんぎん証券株式会社":                           "https://www.gungin-sec.co.jp/",
    "株式会社ＦＵＮＤＩＮＮＯ":                        "https://fundinno.com/",
    "AlpacaJapan株式会社":                            "https://alpaca.markets/",
    "abc証券株式会社":                                "https://abc-sec.co.jp/",
    "アモーヴァ証券株式会社":                         "https://www.amova.co.jp/",
    "株式会社ユニコーン":                             "https://unicorn-cf.com/",
    "イークラウド株式会社":                           "https://ecrowd.co.jp/",
    "エクイティファンディング株式会社":               "https://www.equity-funding.jp/",
    "株式会社ＣＦスタートアップス":                   "https://cf-startups.com/",
    "ロックハラード証券株式会社":                     "https://www.lockhalard.com/",
    "EVOLUTION JAPAN証券株式会社":                    "https://www.evolution-japan.co.jp/",
    "ＥＶＯＬＵＴＩＯＮ ＪＡＰＡＮ証券株式会社":       "https://www.evolution-japan.co.jp/",
    "ジェイ・ボンド東短証券株式会社":                 "https://www.jbond-tosho.co.jp/",
    "セントラル東短証券株式会社":                     "https://www.central-tosho.co.jp/",
    "東短ICAP株式会社":                              "https://www.totan-icap.co.jp/",
    "タレットプレボンETP株式会社":                    "https://www.tprebon.co.jp/",
    "日本相互証券株式会社":                           "https://www.nims.co.jp/",
    "株式会社上田トラディション証券":                 "https://www.ueda-tradition.co.jp/",
    "野村ファイナンシャル・プロダクツ・サービシズ株式会社": "https://www.nomura.co.jp/",
    "ジャパンネクスト証券株式会社":                   "https://www.japannext.co.jp/",
    "エンサイドットコム証券株式会社":                 "https://www.ensight.co.jp/",
    "大阪デジタルエクスチェンジ株式会社":              "https://odx.co.jp/",
    "Japan Alternative Market株式会社":               "https://jam.co.jp/",
    "デジタルアセット証券株式会社":                   "https://www.digital-asset-sec.co.jp/",
    "デジタル証券株式会社":                           "https://digital-sec.co.jp/",
    "株式会社Custodiem":                              "https://www.custodiem.co.jp/",
    "株式会社リアライズ証券":                         "https://realize-sec.co.jp/",
    "Ｓｉｉｉｂｏ証券株式会社":                       "https://siiibo.com/",
    "三井物産デジタル・アセットマネジメント株式会社":  "https://www.mdam.co.jp/",
    "株式会社sustenキャピタル・マネジメント":          "https://sustencp.com/",
    "ＵＢＳ ＳｕＭｉ ＴＲＵＳＴウェルス・マネジメント株式会社": "https://www.ubs.com/jp/",
    "中国国際金融日本株式会社":                       "https://www.cib.citic.com/jp/",
    "カド・キャピタル・マネジメント株式会社":          "https://www.kado-cm.com/",
    "アクサ・ウェルス・マネジメント株式会社":          "https://www.axa-im.co.jp/",
    "株式会社アイ・アールジャパン":                   "https://www.ir-japan.co.jp/",
    "Ｔｅｎｅｏ　Ｐａｒｔｎｅｒｓ株式会社":            "https://www.teneopartners.com/",
    "ストームハーバー証券株式会社":                   "https://stormharbour.com/",
    "ヴァンテージ・キャピタル・マーケッツ・ジャパン株式会社": "https://www.vantagecm.com/",
    "オービス・インベストメンツ株式会社":              "https://www.orbis.com/jp/",
    "グッゲンハイムパートナーズ株式会社":              "https://www.guggenheimpartners.com/",
    "きらぼしライフデザイン証券株式会社":              "https://www.kiraboshi-life.co.jp/",
    "トロント・ドミニオン日本証券株式会社":            "https://www.tdsecurities.com/",
    "スコシア・セキュリティーズ・アジア・リミテッド": "https://www.scotiabank.com/",
    "マレックス証券株式会社":                         "https://www.marex.com/japan/",

    # ── アセマネ系（第一種登録）
    "スパークス・アセット・マネジメント株式会社":     "https://www.sparx.co.jp/",
    "アムンディ・ジャパン株式会社":                   "https://www.amundi.co.jp/",
    "ブラックロック・ジャパン株式会社":               "https://www.blackrock.com/jp/",
    "ピクテ・ジャパン株式会社":                       "https://www.pictet.co.jp/",
    "アライアンス・バーンスタイン株式会社":            "https://www.alliancebernstein.co.jp/",
    "インベスコ・アセット・マネジメント株式会社":      "https://www.invesco.com/jp/",
    "ゴールドマン・サックス・アセット・マネジメント株式会社": "https://www.gsam.com/japan/",
    "ブラックストーン・グループ・ジャパン株式会社":    "https://www.blackstone.com/japan/",
    "ステート・ストリート・グローバル・アドバイザーズ株式会社": "https://www.ssga.com/jp/",
    "モルガン・スタンレー・インベストメント・マネジメント株式会社": "https://www.morganstanley.co.jp/",
    "ドイチェ・アセット・マネジメント株式会社":       "https://www.dws.com/ja-jp/",
    "シュローダー・インベストメント・マネジメント株式会社": "https://www.schroders.com/ja/jp/",
    "ＭＣＰアセット・マネジメント株式会社":           "https://www.mcpam.co.jp/",
    "ＧＣＭインベストメンツ株式会社":                 "https://www.gcmlp.com/",
    "Alpha Japan LO am 株式会社":                     "https://alphajapan.co.jp/",
    "あおぞら投信株式会社":                           "https://www.aozoratoshi.co.jp/",
    "大和かんぽオルタナティブインベストメンツ株式会社": "https://www.daiwakampo-ai.co.jp/",
    "三菱UFJオルタナティブインベストメンツ株式会社":  "https://www.mufg-alternative.co.jp/",
    "キャニオン・キャピタル・ジャパン株式会社":       "https://www.canyonpartners.com/",
    "株式会社ＫＫＲキャピタル・マーケッツ":           "https://www.kkr.com/japan/",

    # ── 銀行系・地方証券
    "七十七証券株式会社":                             "https://www.77sec.co.jp/",
    "北洋証券株式会社":                               "https://www.hokuyosec.jp/",
    "ＦＰＬ証券株式会社":                             "https://fplsec.co.jp/",
    "荘内証券株式会社":                               "https://www.shonai-sec.co.jp/",
    "山形證券株式会社":                               "https://www.yamagata-shoken.co.jp/",
    "八十二証券株式会社":                             "https://www.82sec.co.jp/",
    "第四北越証券株式会社":                           "https://www.daishikitaetsu-sec.co.jp/",
    "岡三にいがた証券株式会社":                       "https://www.okasan-niigata.co.jp/",
    "長野證券株式会社":                               "https://www.nagano-shoken.co.jp/",
    "とちぎんTT証券株式会社":                         "https://www.tt-sec.co.jp/",
    "安藤証券株式会社":                               "https://www.ando-sec.co.jp/",
    "岡地証券株式会社":                               "https://www.okachi-sec.co.jp/",
    "木村証券株式会社":                               "https://www.kimura-sec.co.jp/",
    "寿証券株式会社":                                 "https://www.kotobuki-sec.co.jp/",
    "静岡東海証券株式会社":                           "https://www.shizuoka-tokai-sec.co.jp/",
    "静銀ティーエム証券株式会社":                     "https://www.shizuginam-sec.co.jp/",
    "新大垣証券株式会社":                             "https://www.shin-ogaki-sec.co.jp/",
    "大万証券株式会社":                               "https://www.taimansc.co.jp/",
    "野畑証券株式会社":                               "https://www.nobata-sec.co.jp/",
    "松阪証券株式会社":                               "https://www.matsusaka-sec.co.jp/",
    "丸八証券株式会社":                               "https://www.maruhachi-sec.co.jp/",
    "豊証券株式会社":                                 "https://www.yutaka-sec.co.jp/",
    "百五証券株式会社":                               "https://www.hyakugo-sec.co.jp/",
    "十六ＴＴ証券株式会社":                           "https://www.juroku-tt.co.jp/",
    "ＯＫＢ証券株式会社":                             "https://www.okb-sec.co.jp/",
    "大起証券株式会社":                               "https://www.taiki-sec.co.jp/",
    "石動証券株式会社":                               "https://www.isurugi-sec.co.jp/",
    "今村証券株式会社":                               "https://www.imamura-sec.co.jp/",
    "株式会社しん証券さかもと":                       "https://www.shin-sakamoto.co.jp/",
    "島大証券株式会社":                               "https://www.shimadai-sec.co.jp/",
    "益茂証券株式会社":                               "https://www.masushige-sec.co.jp/",
    "三津井証券株式会社":                             "https://www.mitsui-sec.co.jp/",
    "ほくほくＴＴ証券株式会社":                       "https://www.hokuhoku-tt.co.jp/",
    "永和証券株式会社":                               "https://www.eiwa-sec.co.jp/",
    "岡安証券株式会社":                               "https://www.okayasu-sec.co.jp/",
    "光世証券株式会社":                               "https://www.kousei-sec.co.jp/",
    "岩井コスモ証券株式会社":                         "https://www.iwaicosmo.net/",
    "篠山証券株式会社":                               "https://www.sasayama-sec.co.jp/",
    "内藤証券株式会社":                               "https://www.naito-sec.co.jp/",
    "南都まほろば証券株式会社":                       "https://www.nanto-mahora.co.jp/",
    "西村証券株式会社":                               "https://www.nishimura-sec.co.jp/",
    "播陽証券株式会社":                               "https://www.bansec.co.jp/",
    "光証券株式会社":                                 "https://www.hikari-sec.co.jp/",
    "広田証券株式会社":                               "https://www.hirota-sec.co.jp/",
    "丸近證券株式会社":                               "https://www.marukin-sec.co.jp/",
    "岡安商事株式会社":                               "https://www.okayasu-shoji.co.jp/",
    "池田泉州ＴＴ証券株式会社":                       "https://www.ikedabank-tt.co.jp/",
    "京銀証券株式会社":                               "https://www.kyogin-sec.co.jp/",
    "株式会社コムテックス":                           "https://www.commtex.co.jp/",
    "株式会社アステム":                               "https://www.astem-inc.co.jp/",
    "大山日ノ丸証券株式会社":                         "https://www.oyama-hinomaru.co.jp/",
    "中銀証券株式会社":                               "https://www.chugoku-sec.co.jp/",
    "ワイエム証券株式会社":                           "https://www.ymsc.co.jp/",
    "ひろぎん証券株式会社":                           "https://www.hiroginsc.co.jp/",
    "阿波証券株式会社":                               "https://www.awa-shoken.co.jp/",
    "香川証券株式会社":                               "https://www.kagawa-sec.co.jp/",
    "徳島合同証券株式会社":                           "https://www.tokushima-godo.co.jp/",
    "二浪証券株式会社":                               "https://www.niro-sec.co.jp/",
    "三豊証券株式会社":                               "https://www.mitoyo-sec.co.jp/",
    "四国アライアンス証券株式会社":                   "https://www.4koku-alliance.co.jp/",
    "大熊本証券株式会社":                             "https://www.kumamoto-sec.co.jp/",
    "九州ＦＧ証券株式会社":                           "https://www.kyufu-sec.co.jp/",
    "ＦＦＧ証券株式会社":                             "https://www.ffg-sec.co.jp/",
    "西日本シティＴＴ証券株式会社":                   "https://www.nct-sec.co.jp/",
    "おきぎん証券株式会社":                           "https://www.okigin-sec.co.jp/",
    "國府証券株式会社":                               "https://www.kokufu-sec.co.jp/",
    "めぶき証券株式会社":                             "https://www.mebuki-sec.co.jp/",
    "エアーズシー証券株式会社":                       "https://www.airsea.co.jp/",
    "Ｐｌｕｓ５００ＪＰ証券株式会社":                  "https://www.plus500.com/ja/",
    "ブラウン・ブラザーズ・ハリマン証券株式会社":      "https://www.bbh.com/japan/",
    "リクイドネット証券株式会社":                     "https://www.liquidnet.com/japan/",
    "ＲＢＣキャピタルマーケッツ・ジャパン・リミテッド": "https://www.rbccm.com/japan/",
    "ナットウエスト・マーケッツ・セキュリティーズ・ジャパン・リミテッド（証券）": "https://natwest.com/corporates/",
    "バンクオブニューヨークメロン証券株式会社":        "https://www.bnymellon.com/japan/",
    "ブルームバーグ・トレードブック・ジャパン証券株式会社": "https://www.bloomberg.co.jp/",
    "ＧＩキャピタル・マネジメント株式会社":           "https://www.gi-capital.com/",
    "ジェフリーズ日本株式会社":                       "https://www.jefferies.com/japan/",
    "ＢＧＣショウケンカイシャリミテッド（東京支店）":  "https://bgcinc.com/",
    "シー・アイ・ビー・シー・ワールド・マーケッツ（ジャパン）インク": "https://cibcwm.com/",
    "グリーンズレッジ・アジア・リミテッド":           "https://www.greensledge.com/",
    "シタデル・セキュリティーズ証券株式会社":          "https://www.citadelsecurities.com/",
    "バンク・オブ・モントリオール証券株式会社":        "https://www.bmo.com/japan/",
    "ナティクシス証券株式会社":                       "https://www.natixis.com/natixis/jcms/m_34/jp/",
    "トレードウェブ・ジャパン株式会社":               "https://www.tradeweb.com/",
    "キャンターフィッツジェラルド証券株式会社":        "https://www.cantor.com/",
    "エービーエヌ・アムロ・クリアリング証券株式会社":  "https://www.abnamroclearing.com/",
    "AssetmarkCapital株式会社":                       "",
    "株式会社ＵＮＩＶＡ証券":                         "https://www.univa-sec.co.jp/",
    "ヘッジファンド証券株式会社":                     "https://www.hedgefundsec.co.jp/",
    "クリエイトジャパン株式会社":                     "https://www.createjapan.co.jp/",
    "株式会社証券ジャパン":                           "https://www.shoken-japan.co.jp/",
    "株式会社アイ・アールジャパン":                   "https://www.ir-japan.co.jp/",
    "ＧＩキャピタル・マネジメント株式会社":           "https://www.gi-capital.com/",
}

# ──────────────────────────────────────────────
# 企業名 → 広告種別マッピング
# ──────────────────────────────────────────────
AD_TYPE_MAP = {
    # ── ネット証券（リスティング広告メイン）
    "株式会社ＳＢＩ証券":                            "リスティング広告",
    "楽天証券株式会社":                              "リスティング広告",
    "マネックス証券株式会社":                         "リスティング広告",
    "松井証券株式会社":                              "リスティング広告",
    "ＧＭＯクリック証券株式会社":                     "リスティング広告",
    "株式会社ＳＢＩネオトレード証券":                  "リスティング広告",
    "株式会社ＤＭＭ．ｃｏｍ証券":                     "リスティング広告",
    "三菱UFJ eスマート証券株式会社":                  "リスティング広告",
    "岡三証券株式会社":                              "リスティング広告",
    "東海東京証券株式会社":                           "リスティング広告",
    "ＳＭＢＣ日興証券株式会社":                       "リスティング広告",
    "野村證券株式会社":                              "リスティング広告",
    "大和証券株式会社":                              "リスティング広告",
    "みずほ証券株式会社":                             "リスティング広告",
    "三菱ＵＦＪモルガン・スタンレー証券株式会社":       "リスティング広告",
    "moomoo証券株式会社":                            "リスティング広告",
    "フィデリティ証券株式会社":                       "リスティング広告",
    "インタラクティブ・ブローカーズ証券株式会社":       "リスティング広告",
    "Ｐｌｕｓ５００ＪＰ証券株式会社":                  "リスティング広告",
    "サクソバンク証券株式会社":                       "リスティング広告",
    "ＩＧ証券株式会社":                               "リスティング広告",
    "株式会社外為どっとコム":                         "リスティング広告",
    "ＧＭＯ外貨株式会社":                             "リスティング広告",
    "株式会社外為オンライン":                         "リスティング広告",
    "セントラル短資ＦＸ株式会社":                     "リスティング広告",
    "ＯＡＮＤＡ証券株式会社":                         "リスティング広告",
    "ひまわり証券株式会社":                           "リスティング広告",
    "インヴァスト証券株式会社":                       "リスティング広告",
    "トレイダーズ証券株式会社":                       "リスティング広告",
    "デューカスコピー・ジャパン株式会社":              "リスティング広告",
    "株式会社マネーパートナーズ":                     "リスティング広告",
    "株式会社岡三オンライン":                         "リスティング広告",
    "アイザワ証券株式会社":                           "リスティング広告",
    "東洋証券株式会社":                              "リスティング広告",
    "丸三証券株式会社":                              "リスティング広告",
    "水戸証券株式会社":                              "リスティング広告",
    "立花証券株式会社":                              "リスティング広告",
    "内藤証券株式会社":                              "リスティング広告",
    "岩井コスモ証券株式会社":                         "リスティング広告",
    "光世証券株式会社":                              "リスティング広告",
    "エース証券株式会社":                             "リスティング広告",
    "極東証券株式会社":                              "リスティング広告",
    "藍澤證券株式会社":                              "リスティング広告",
    "百五証券株式会社":                              "リスティング広告",
    "尾州証券株式会社":                              "リスティング広告",
    # ── Meta広告メイン（ロボアド・フィンテック系）
    "ウェルスナビ株式会社":                           "Meta",
    "株式会社ＦＯＬＩＯ":                             "Meta",
    "ブルーモ証券株式会社":                           "Meta",
    "CHEER証券株式会社":                             "Meta",
    "株式会社スマートプラス":                         "Meta",
    "tsumiki証券株式会社":                           "Meta",
    "ＰａｙＰａｙ証券株式会社":                       "Meta",
    "大和コネクト証券株式会社":                       "Meta",
    "LINE証券株式会社":                              "Meta",
}

# ──────────────────────────────────────────────
# FSAデータ読み込み
# ──────────────────────────────────────────────
import openpyxl as ox

wb_in = ox.load_workbook(
    r'C:\Users\tonomura-r\.claude\projects\C--Users-tonomura-r-Downloads------------\23942631-e569-45b2-a239-44f88ed0ae67\tool-results\webfetch-1779139777036-7kw1bn.xlsx'
)
ws_in = wb_in.active

companies = []
for i in range(8, ws_in.max_row + 1):
    name = str(ws_in.cell(row=i, column=4).value or '').strip()
    reg  = str(ws_in.cell(row=i, column=2).value or '').strip()
    tel  = str(ws_in.cell(row=i, column=8).value or '').strip()
    kind = str(ws_in.cell(row=i, column=9).value or '').strip()
    if name and '○' in kind:
        url = URL_MAP.get(name, '')
        ad_type = AD_TYPE_MAP.get(name, 'HP（広告なし）' if url else '')
        companies.append([name, url, tel, ad_type, reg])

print(f"対象: {len(companies)}社  URL有: {sum(1 for c in companies if c[1])}社")

# ──────────────────────────────────────────────
# Excel出力
# ──────────────────────────────────────────────
def make_xlsx(filename, data):
    wb = ox.Workbook()
    ws = wb.active
    ws.title = "証券会社一覧"
    headers = ["企業名（正式）", "URL（LP or HP）", "電話番号", "広告種別", "登録番号"]
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="メイリオ", bold=True, color="FFFFFF", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin, )
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
            cell.font = Font(name="メイリオ", size=9)
            cell.alignment = left
            cell.border = border
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 30
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(filename)
    print(f"保存: {filename}  ({len(data)}社)")

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
out = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\日本証券会社一覧_2026-04-01_v2.xlsx"
make_xlsx(out, companies)
