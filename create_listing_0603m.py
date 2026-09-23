import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03m.xlsx"
HEADER = ["企業名", "LP URL", "広告種別", "法人番号", "商材", "検索KW", "業界"]
COL_WIDTHS = [32, 50, 16, 16, 20, 32, 22]
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="メイリオ", size=10, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_FILLS = [PatternFill("solid", fgColor="F2F7FC"), PatternFill("solid", fgColor="FFFFFF")]
ROW_FONT = Font(name="メイリオ", size=9)
ROW_ALIGN = Alignment(vertical="center", wrap_text=False)
URL_FONT = Font(name="メイリオ", size=9, color="0563C1", underline="single")

DATA = [
    # ── 専門学校・スクール（高単価） ──────────────────────────────────────────
    ("学校法人大和学園（京都調理師専門学校）",            "https://www.taiwa.ac.jp/",                      "リスティング広告", "", "調理師免許・料理・製菓専門学校",          "調理師 専門学校 費用 おすすめ, 料理 資格 取得 最短, 調理師免許 学校 選び方",  "調理師専門学校"),
    ("学校法人文化学園（文化服装学院）",                  "https://bfc.bunka.ac.jp/",                      "リスティング広告", "", "ファッション・服飾デザイン専門学校",       "ファッション 専門学校 東京 おすすめ, 服飾 デザイン 学校 費用, ファッション 資格 取得","ファッション専門学校"),
    ("学校法人大阪モード学園",                           "https://mode.osaka-mode.ac.jp/",                "リスティング広告", "", "ファッション・ヘア・メイクアップ専門学校", "美容師 学校 大阪, メイクアップ スクール 資格 費用, ヘアメイク 専門学校 入学",  "美容専門学校"),
    # ── アウトドアウェア（高単価） ──────────────────────────────────────────
    ("合同会社パタゴニア日本支社",                        "https://www.patagonia.com/jp/",                "リスティング広告", "", "アウトドアウェア・フリース・登山用品",     "パタゴニア フリース 人気 モデル, アウトドア ウェア 高機能 おすすめ, パタゴニア セール","アウトドアウェア"),
    ("コロンビアスポーツウェアジャパン株式会社",          "https://www.columbiasportswear.jp/",            "リスティング広告", "", "アウトドアウェア・ハイキング・登山用品",   "コロンビア ジャケット 評判 おすすめ, アウトドア 防水 ウェア 比較, 登山 着替え 選び方","アウトドアウェア"),
    # ── セレクトショップ・ファッション（高単価） ──────────────────────────────────────────
    ("株式会社シップス（SHIPS）",                         "https://www.shipsltd.co.jp/",                  "リスティング広告", "", "セレクトショップ・ファッション・雑貨",     "SHIPS セール 人気 商品, セレクトショップ ファッション 高品質, シップス 新作 コーデ","ファッション"),
    # ── メンズコスメ（高単価） ──────────────────────────────────────────
    ("株式会社DISM",                                     "https://www.dism.jp/",                          "リスティング広告", "", "メンズスキンケア・洗顔・保湿・美肌",       "メンズ スキンケア おすすめ 初心者, 男性 洗顔 保湿 毛穴, DISM 評判 効果 口コミ","メンズコスメ"),
    # ── スポーツブランド（高単価） ──────────────────────────────────────────
    ("アディダスジャパン株式会社",                        "https://www.adidas.com/jp",                    "リスティング広告", "", "スポーツウェア・ランニングシューズ・スニーカー","アディダス 新作 スニーカー おすすめ, スポーツ ウェア 高機能 比較, アディダス セール","スポーツブランド"),
    ("株式会社コンバースジャパン",                        "https://www.converse.co.jp/",                  "リスティング広告", "", "スニーカー・オールスター・キャンバスシューズ","コンバース オールスター 新作 人気, スニーカー キャンバス おすすめ, コンバース 限定 購入","スニーカー"),
    # ── カメラ（高単価） ──────────────────────────────────────────
    ("OMデジタルソリューションズ株式会社（OM SYSTEM）",   "https://om-digitalsolutions.com/",             "リスティング広告", "", "ミラーレスカメラ・防水一眼・レンズ",       "OM SYSTEM カメラ 評判 比較, ミラーレス 防塵防滴 おすすめ, 一眼 カメラ 初心者 選び方","カメラ"),
    # ── スポーツ栄養（高単価） ──────────────────────────────────────────
    ("株式会社DNS",                                      "https://www.dnszone.jp/",                       "リスティング広告", "", "プロテイン・スポーツサプリ・筋トレ",        "プロテイン おすすめ 筋トレ, DNS プロテイン 評判 効果, スポーツサプリ 選び方",  "スポーツ栄養"),
    # ── プレミアム紅茶・ワイン（高単価） ──────────────────────────────────────────
    ("マリアージュフレール（Mariage Frères）",            "https://www.mariagefreres.com/JP/",             "リスティング広告", "", "プレミアム紅茶・茶葉・ティーギフト",       "紅茶 高級 ブランド 通販, ギフト 紅茶 おすすめ 高品質, マリアージュフレール 評判","プレミアム紅茶"),
    ("エノテカ株式会社",                                  "https://www.enoteca.co.jp/",                   "リスティング広告", "", "プレミアムワイン・ボトルギフト通販",       "ワイン 高級 通販 ギフト, プレミアム ワイン おすすめ, エノテカ 評判 ワイン 購入","プレミアムワイン"),
    # ── フィッシング用品（高単価） ──────────────────────────────────────────
    ("株式会社シマノ（フィッシング）",                    "https://fish.shimano.com/ja-JP/",               "リスティング広告", "", "釣り竿・リール・ルアー・フィッシング用品", "シマノ リール おすすめ 選び方, 釣り竿 ロッド 購入, フィッシング 用品 高機能",  "フィッシング"),
    ("株式会社グローブライド（ダイワフィッシング）",       "https://www.daiwa.com/jp/",                    "リスティング広告", "", "釣り竿・電動リール・釣り用品",             "ダイワ リール 評判 比較, ロッド 選び方 釣り, 電動 リール 海釣り おすすめ",    "フィッシング"),
    # ── 育児用品（高単価） ──────────────────────────────────────────
    ("コンビ株式会社",                                   "https://www.combi.co.jp/",                      "リスティング広告", "", "ベビーカー・チャイルドシート・ハイチェア",  "ベビーカー おすすめ 選び方 軽量, チャイルドシート 安全 比較, コンビ 評判 人気",  "育児用品"),
    # ── 腕時計・アイウェア（高単価） ──────────────────────────────────────────
    ("ダニエルウェリントン（Daniel Wellington）",         "https://www.danielwellington.com/jp/",          "リスティング広告", "", "ファッション腕時計・メッシュベルト",        "ダニエルウェリントン 評判 人気 モデル, ファッション 時計 おすすめ ブランド, DW セール","腕時計"),
    ("株式会社オークリージャパン（Oakley）",              "https://www.oakley.com/ja-JP/",                "リスティング広告", "", "スポーツサングラス・プリズムレンズ",        "オークリー サングラス おすすめ 人気, スポーツ サングラス 高機能 比較, オークリー ゴルフ","サングラス"),
    # ── 水泳用品（高単価） ──────────────────────────────────────────
    ("山本光学株式会社（SWANS）",                         "https://www.swans.co.jp/",                      "リスティング広告", "", "水泳ゴーグル・水中メガネ・競泳用品",        "水泳 ゴーグル おすすめ 選び方, 競泳 用品 高機能, スワンズ 評判 ゴーグル 度入り","水泳用品"),
    # ── ファッションウォッチ（高単価） ──────────────────────────────────────────
    ("フォッシル・ジャパン合同会社（FOSSIL）",            "https://www.fossil.com/ja-jp/",                "リスティング広告", "", "ファッション腕時計・スマートウォッチ",      "FOSSIL 時計 おすすめ 人気 モデル, ファッション スマートウォッチ 比較, フォッシル 購入","腕時計"),
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
