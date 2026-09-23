import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = r"C:\Users\tonomura-r\Downloads\との\新しいフォルダー\架電リスト_リスティング_2026-06-03b.xlsx"
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
    # ── ITスクール・プログラミング ──────────────────────────────────────────
    ("株式会社ギークジョブ（GEEK JOB）",         "https://camp.geekjob.jp/",                       "リスティング広告", "", "プログラミング転職スクール",             "プログラミング 転職 スクール 無料, IT転職 未経験 エンジニア, GEEK JOB 評判",     "プログラミングスクール"),
    ("Winスクール",                              "https://www.winschool.jp/",                      "リスティング広告", "", "パソコン・Webデザインスクール",           "パソコン教室 資格 おすすめ, Webデザイン 学ぶ 通学, 社会人 スキルアップ IT",     "ITスクール"),
    ("インターノウス株式会社（プロエンジニア）",   "https://proengineer.internous.co.jp/lp_kiso2/", "リスティング広告", "", "未経験エンジニア就職スクール",           "未経験 エンジニア 就職, プログラミング 独学 限界 スクール, 就職保証 IT無料",    "プログラミングスクール"),
    ("株式会社LIG（デジLIG）",                   "https://liginc.co.jp/studioueno/web_designer",  "リスティング広告", "", "Webデザイン・制作スクール",              "Webデザイン スクール おすすめ, デザイナー 転職 未経験, デジLIG 評判 口コミ",    "Webデザインスクール"),
    ("株式会社データミックス（datamix）",          "https://datamix.co.jp/school/data-scientist/", "リスティング広告", "", "データサイエンティスト育成講座",         "データサイエンス 学ぶ スクール, AI 機械学習 資格, データ分析 転職 エンジニア",  "データサイエンス"),
    ("ラグザス・クリエイト株式会社（忍者CODE）",  "https://ninjacode.work/",                        "リスティング広告", "", "オンラインプログラミングスクール",       "プログラミング 独学 オンライン 安い, 忍者CODE 評判, 副業 プログラミング 始め方", "プログラミングスクール"),
    ("株式会社TOMAP（ZeroPlus）",                "https://zero-plus.io/",                          "リスティング広告", "", "フリーランス特化プログラミングスクール", "フリーランス エンジニア スクール, Web制作 独立 スキル, ZeroPlus 評判 口コミ",   "プログラミングスクール"),
    # ── 英会話・語学スクール ──────────────────────────────────────────
    ("株式会社リンゲージ",                       "https://www.linguage-school.jp/lp/",             "リスティング広告", "", "英会話スクール・体験レッスン",           "英会話 安い おすすめ, リンゲージ 評判 口コミ, 英語 マンツーマン 格安",            "英会話スクール"),
    ("株式会社イングリッシュカンパニー",          "https://englishcompany.jp/",                     "リスティング広告", "", "ビジネス英語コーチング",                 "英語コーチング おすすめ, ビジネス英語 上達 最短, 英会話 コーチング 比較",        "英語コーチング"),
    ("株式会社スピークバディ（コーチバディ）",    "https://speakbuddy-personalcoaching.com/",       "リスティング広告", "", "AI×コーチング英会話サービス",           "英語コーチング AI, スピーキング 上達 オンライン, 英会話 短期 集中 おすすめ",    "英語コーチング"),
    ("株式会社アムスジャパン（AEON英会話）",      "https://www.aeonet.co.jp/lp/online_lesson/",    "リスティング広告", "", "オンライン英会話レッスン（イーオン）",   "英会話 体験 無料 オンライン, イーオン 評判 口コミ, 英語 上達 レッスン 安い",    "英会話スクール"),
    ("ガバナーズ語学学院",                       "https://www.governors.ac.jp/",                   "リスティング広告", "", "英語学校・ビジネス英語コース",          "英語学校 横浜 おすすめ, ビジネス英語 学校 通学, 英会話 社会人 充実",            "語学スクール"),
    # ── 料理教室・料理スクール ──────────────────────────────────────────
    ("一般財団法人ベターホーム協会",             "https://www.betterhome.jp/",                     "リスティング広告", "", "ベターホームのお料理教室",               "料理教室 東京 おすすめ, 料理 習い事 社会人, クッキングスクール 入会 費用",      "料理教室"),
    ("学校法人辻料理師専門学校（エコール辻）",    "https://www.tsuji.ac.jp/college/tokyo/",        "リスティング広告", "", "調理師・パティシエ専門学校",             "調理師 専門学校 東京, パティシエ 資格 取得 学校, 料理 プロ 学ぶ 費用",         "料理専門学校"),
    ("ルコルドンブルー日本校（Le Cordon Bleu）", "https://www.cordonbleu.edu/tokyo/home/ja",       "リスティング広告", "", "料理・製菓プロフェッショナル養成",       "フランス料理 スクール 東京, ルコルドンブルー 入学 費用, 製菓 プロ 養成 東京",  "料理スクール"),
    # ── 家事代行・買取・宅食・リフォーム ──────────────────────────────────────────
    ("ミニメイド・サービス株式会社",             "https://www.minimaid.co.jp/",                    "リスティング広告", "", "定期家事代行サービス",                   "家事代行 定期 おすすめ, 掃除代行 料金 比較, 家事 外注 初めて 安い",            "家事代行"),
    ("株式会社いーふらん（おたからや）",          "https://www.otakaraya.jp/lp/iroiro_lp01/",      "リスティング広告", "", "ブランド品・貴金属・切手買取",           "買取 専門店 おすすめ, ブランド 売る 高額査定, 貴金属 金 買取 近く",            "買取"),
    ("株式会社ライフデリ",                       "https://lifedeli.jp/",                           "リスティング広告", "", "高齢者向け宅配弁当・配食サービス",       "高齢者 宅配弁当 おすすめ, 介護食 宅配 安い, 配食サービス 高齢者 近く",         "宅食配達"),
    ("株式会社ライフアイ（増改築.com）",          "https://www.zoukaichiku.com/",                   "リスティング広告", "", "戸建てフルリフォーム・リノベーション",   "フルリフォーム 費用 一戸建て, 増改築 業者 おすすめ, リノベーション 見積もり",  "リフォーム"),
    ("株式会社ヌリカエ",                         "https://www.nurikae.com/",                       "リスティング広告", "", "外壁塗装・屋根塗装の比較サービス",       "外壁塗装 費用 相場, 外壁 業者 比較 一括見積もり, 屋根塗装 見積もり 無料",     "外壁塗装"),
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
