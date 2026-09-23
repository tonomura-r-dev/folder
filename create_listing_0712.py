import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def make_xlsx(filename, data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "架電リスト"

    headers = ["企業名", "LP URL", "電話番号", "商材", "検索KW", "業界"]

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
            cell.font = Font(name="メイリオ", size=9)
            cell.alignment = left
            cell.border = border

    col_widths = [30, 45, 16, 20, 30, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(filename)
    print(f"保存完了: {filename}")


# 列順: 企業名 | LP URL | 電話番号 | 商材 | 検索KW | 業界
# 2026-07-12: Google検索リスティング広告主から抽出・全社LP実確認済み・Grep重複確認済み

data_0712 = [
    ["株式会社REGATE（買取福ちゃん）", "https://lp.fuku-chan.jp", "0120-947-295", "着物・骨董品・貴金属の出張買取", "着物 出張買取 / 着物 高額買取", "着物・骨董買取"],
    ["株式会社マックスガイ（ザ・ゴールド）", "https://www.the-gold.jp", "0120-355-525", "ブランド品・時計・貴金属の出張買取（全国76店舗）", "時計 高額買取 出張 / ブランド品 買取", "ブランド買取"],
    ["株式会社remental（Kimochi）", "https://kimochi-care.net", "", "オンライン心理カウンセリング（サブスク）", "オンラインカウンセリング / 心理カウンセリング オンライン", "メンタルヘルスオンライン"],
    ["株式会社ライズ（GYMS）", "https://www.p-gyms.jp", "0120-634-777", "女性専用パーソナルジム（全国25店舗）", "パーソナルジム 女性専用 / ダイエット ジム 女性", "女性専用パーソナルジムFC"],
    ["株式会社フィットクルー（UNDEUX SUPERBODY）", "https://www.diet-undeux.jp", "", "女性専用パーソナルジム（全国44店舗）", "女性専用ジム / ボディメイク パーソナル", "女性専用パーソナルジムFC"],
    ["株式会社プリッツジム", "https://plizgym.co.jp", "", "女性専用セミパーソナルジム（全国37店舗）", "セミパーソナルジム 女性 / 女性専用ジム 安い", "女性専用パーソナルジムFC"],
    ["日本無人航空機免許センター株式会社（JULC）", "https://julc.co.jp", "03-6661-1920", "ドローン国家資格講習（全国28拠点）", "ドローンスクール 国家資格 / ドローン免許 取得", "ドローンスクール"],
    ["株式会社givers（こころ整体院グループ）", "https://seitai.co.jp", "03-6380-2381", "整体・整骨院（全国125店舗）", "整体 骨盤矯正 / 整体院 肩こり 腰痛", "整体・整骨院FC"],
    ["株式会社KMC（小林整骨院グループ）", "https://www.seikotsuin-kobayashi.com", "072-998-9630", "整骨院（関西・東海・東京展開）", "産後骨盤矯正 大阪 / 整骨院 交通事故", "整骨院FC"],
    ["Carstay株式会社", "https://carstay.jp", "", "キャンピングカーレンタル・車中泊スポット予約", "キャンピングカー レンタル 全国 / 車中泊 予約", "キャンピングカーシェア"],
    ["有限会社ワンラブ（ペットショップワンラブ）", "https://www.pet-onelove.com", "052-222-1131", "子犬・子猫販売、ペットホテル、トリミング（全国167店舗）", "ペットホテル 全国 チェーン / 子犬 販売", "ペットショップFC"],
    ["株式会社メディビューティー（ラココ LACOCO）", "https://la-coco.com", "", "全身脱毛サロン（全国90店舗）", "脱毛サロン 女性 全国展開 / 全身脱毛 サロン", "脱毛サロンFC"],
    ["株式会社リンリン（Rin Rin）", "https://happyrinrin.com", "052-243-5577", "脱毛×肌管理サロン（全国52店舗）", "脱毛サロン 女性 全国展開 / 肌管理 サロン", "脱毛サロンFC"],
    ["株式会社シャン・クレール（シャンクレール）", "https://www.2400.co.jp", "050-5533-1000", "婚活パーティー・街コンの企画運営", "婚活パーティー 大手 全国 / 街コン 出会い", "婚活・恋活イベント"],
    ["株式会社リビアス（ビューティーアイラッシュ）", "https://www.ribias.net", "06-6301-1138", "まつげエクステサロン（全国62店舗）", "まつげエクステ サロン 全国展開 / マツエク 予約", "まつげエクステサロンFC"],
    ["株式会社CS（CS Inc.）", "https://cs-sa.jp", "", "ヘアサロン・アイサロン・ネイルサロン（全国200店舗以上）", "まつげエクステ サロン 全国展開 / アイサロン 予約", "トータルビューティーFC"],
    ["株式会社ベンリーコーポレーション（ベンリー）", "https://www.benry.com", "052-505-8702", "生活支援サービス・便利屋（全国280店舗）", "便利屋 全国展開 FC / 家事代行 便利屋", "便利屋・生活支援FC"],
    ["株式会社ボルテックス", "https://www.vortex-net.com", "0120-285-191", "収益不動産（区分所有オフィス）の資産形成コンサルティング", "資産運用セミナー 無料 全国 / 不動産小口化商品", "不動産投資"],
    ["iCureテクノロジー株式会社", "https://i-cure.co.jp", "", "鍼灸接骨院（全国30店舗以上）", "姿勢矯正 猫背 全国展開 / 整骨院 保険適用", "整骨院FC"],
    ["株式会社すずらん健康倶楽部（すずらん鍼灸接骨院）", "https://suzuran-758.com", "", "鍼灸接骨院（全国100店舗以上）", "猫背矯正 整骨院 / 鍼灸接骨院 肩こり", "整骨院FC"],
    ["株式会社ザウルス（全力ストレッチ）", "https://zn-stretch.com", "", "パーソナルストレッチサロン（全国70店舗）", "姿勢矯正 猫背 女性 / ストレッチ専門店", "ストレッチサロンFC"],
    ["株式会社ファストノット（ベルミス）", "https://belmise.com", "0120-000-235", "着圧レギンス・着圧ソックスD2C", "着圧ソックス 通販 D2C / 着圧レギンス", "着圧インナーD2C"],
    ["株式会社Lino Che'ri（Kiratt）", "https://kiratt.jp", "080-9784-8882", "歯のホワイトニング専門サロン（全国30店舗）", "ホワイトニング サロン 全国展開 / セルフホワイトニング", "ホワイトニングサロンFC"],
    ["株式会社ピベルダ（Whitening BAR）", "https://whiteningbar.jp", "", "セルフホワイトニング専門店（月額制通い放題・全国展開）", "ホワイトニング サロン 全国展開 / セルフホワイトニング 通い放題", "ホワイトニングサロンFC"],
]

base = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー"
make_xlsx(f"{base}\\架電リスト_Meta_2026-07-12.xlsx", data_0712)
