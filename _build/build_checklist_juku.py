# -*- coding: utf-8 -*-
"""教育（塾）業界資料 公表判断チェックリスト（xlsx）

  python _build/build_checklist_juku.py

全36枚を「事実／推定／意見」で仕分け、赤黄緑で公表可否を判定する。
判定基準：
  緑 = 出典のある事実、またはDYMの標準メニュー・設計（そのまま公表可）
  黄 = 推定・仮説・業界推計値。データが入れば緑になる（「未取得」「推計値」の明記が必要）
  赤 = 出典未記載の数値が載っている（出典を確定 or 削除するまで公表不可）
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ROOT = Path(__file__).resolve().parent.parent
DECK = ROOT / "教育（塾）業界_LINEOA施策提案.pptx"
OUT = ROOT / "教育（塾）業界資料_公表判断チェックリスト.xlsx"

FONT = "Arial"
NAVY = "1F285A"
G_FILL = PatternFill("solid", fgColor="E2EFDA")   # 緑
Y_FILL = PatternFill("solid", fgColor="FFF2CC")   # 黄
R_FILL = PatternFill("solid", fgColor="FCE4E4")   # 赤
HDR_FILL = PatternFill("solid", fgColor=NAVY)
IN_FILL = PatternFill("solid", fgColor="FFFF00")  # 記入セル
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# (枚, タイトル, 主張の種類, 判定, 根拠・出典, 残タスク)
ROWS = [
    (1, "表紙", "意見", "緑", "提案タイトル。数値なし", "宛先社名を差し込む"),
    (2, "資料アジェンダ", "意見", "緑", "全8章の目次", "―"),
    (3, "市場データ｜ニーズ全体図", "事実", "黄", "厚労省 人口動態統計（確定数）／経産省 特定サービス産業動態統計 2024年確報／文科省 学校基本調査 令和6年度／矢野経済研究所2024。定義の違い（経産省は受講生500人以上・市場の4〜5割）も併記済み", "★提出前にe-Statで統計表を開いて数値を照合（Gemini経由取得のため直リンク未確認）"),
    (4, "現状の構造｜3つの壁", "推定", "黄", "CPC/CPA・入会率は業界推計値（出典元の特定不可）。LP CVR 2.0%はモデル値", "社内の塾・スクール系実績でCPC/CVR/CPAを確定（上司確認①）"),
    (5, "エンドユーザー行動心理", "事実", "緑", "LINEヤフー前後検索（実測・性別/年代）／塾ナラ2025 n=200／塾シル2026 n=249／明光ネットワークジャパン2020／DeltaX2025", "★明光義塾は競合。その調査を引用してよいか確認（上司確認③）"),
    (6, "本提案のスコープ", "意見", "緑", "DYM提案のスコープ宣言。数値なし", "―"),
    (7, "広告の実態｜CPCの相場が存在しない", "推定", "黄", "公的統計・業界統一データの不在はGemini再調査で確認（事実）。推計値は出典元の特定不可", "社内実績でCPC/CPAを確定（上司確認①）"),
    (8, "広告実績CVR（CV地点別）", "推定", "黄", "★差込枠。数値は未記入（捏造なし）", "DYM社内の塾・スクール系実績をCV地点別に取得（上司確認①）"),
    (9, "広告審査・表現面の懸念", "事実", "黄", "景品表示法（優良誤認・有利誤認）の一般的な考え方に基づく整理", "★配信文面は法務確認を通す。塾業界固有の公正競争規約の有無は未確認"),
    (10, "シーズナリティ", "事実", "黄", "Googleトレンド 日本 2026年年初来・週次（実測・_data/trends/教育塾/）。塾ピーク2/8・夏期講習7/12", "取得期間が年初来のみ。5年推移が要るなら再取得（「◯年で△%上昇」は現データでは言えない）"),
    (11, "前後検索｜26日タイムライン", "事実", "緑", "LINEヤフー前後検索 起点KW「塾 費用」2022-05-09〜2023-05-14・関連度3.0以上（実測・_data/journey/教育塾/）", "―（本資料の背骨。同音異義語「熟」16件は除外済み）"),
    (12, "他社分析｜友だち数とビザビ凍結", "事実", "黄", "page.line.me・実機 2026-09-15取得（実測・証跡は_data/lineoa_competitors/）。いずれも認証済アカウント", "★競合の実名掲載OKか確認（上司確認②）。★KUMON 2,525万人の内訳は未確認のため推定と明記済み"),
    (13, "全体設計｜CJ8フェーズ×CV3段", "意見", "緑", "CV3段はDYM設計。数値なし", "貴社のCV定義とすり合わせ"),
    (14, "施策全体像｜二毛作", "意見", "緑", "DYM提案の対策領域マップ。退会1人＝50万円はモデル値と明記", "貴社の月謝・平均在籍年数で再計算"),
    (15, "施策展開図（初期・月次）", "意見", "緑", "DYM標準の展開図を塾施策に置換", "―"),
    (16, "構築｜友だち追加動線", "意見", "緑", "DYM提案の動線設計。訴求軸（料金でなく「いつから」）はS11の実測に基づく", "―"),
    (17, "構築｜あいさつメッセージ", "意見", "緑", "配信文面は初稿", "担当者名・対応時間・配信頻度を貴社運用に差し替え"),
    (18, "構築｜30秒ぴったり塾診断", "意見", "緑", "診断設計。継続率への効果はDeltaX2024（事実）で裏付け", "★事例数値（58点→82点）は貴社実績に差し替え、条件を併記（S09）"),
    (19, "構築｜リッチメニュー", "意見", "緑", "DYM提案の設計。KUMONの実機リッチメニューは実測（証跡あり）", "―"),
    (20, "構築｜生徒情報・CRM連携", "意見", "緑", "仕組みの説明。数値なし", "★貴社が使用中の塾管理システムを確認（要件定義）"),
    (21, "配信設計｜シナリオ2本", "意見", "緑", "配信設計。14日に寄せる根拠はS11の実測", "―"),
    (22, "実文面①｜未入会", "意見", "緑", "配信文面は初稿", "★事例の数値を貴社実績に差し替え、学年・期間・通塾回数を併記"),
    (23, "実文面②｜在籍生の保護者", "事実", "緑", "配信文面は初稿。根拠はPOPER2022 n=300／DeltaX2024 n=65", "―"),
    (24, "歩留まり①｜体験予約リマインド", "推定", "黄", "予約→体験 80%→92%は弊社支援実績レンジ（保守置き）", "社内数値で実績レンジを確定（上司確認①）"),
    (25, "年間の企画カレンダー", "事実", "黄", "Googleトレンド実測＋前後検索の26日。退会ピーク（3月・8〜10月）は業界推計値", "貴社の入会・退会の月別実績があれば実データに置換"),
    (26, "配信設計｜通知メッセージ", "意見", "緑", "利用シーンの設計。別途費用である旨を明記済み", "利用可否・単価はLINEヤフーの最新規定を要件定義時に確認"),
    (27, "歩留まり②｜退会防止（本資料の山）", "事実", "緑", "POPER「Comiru」2022年 n=300／DeltaX「塾選」2024年 n=65。退会率10〜20%は業界推計値と明記", "―"),
    (28, "現場の工数削減", "意見", "緑", "DYM提案の改善モデル。時給換算はあえてしない旨を明記", "―"),
    (29, "成果の見方｜効果測定の設計", "意見", "緑", "KPI設計。数値は入れない（業界汎用のためSIMは作らない）", "貴社の実績をいただければSIMを作成（lineoa-simスキル）"),
    (30, "費用プラン", "意見", "緑", "DYM標準プラン（6ヶ月〜・税抜）。別途費用も明記", "―"),
    (31, "導入スケジュール", "意見", "緑", "DYM標準。逆算日はS10の実測から算出", "―"),
    (32, "体制｜誰が何をやるか", "意見", "緑", "DYM標準のサポート体制", "体制人数・担当者名は契約時に確定"),
    (33, "飛び道具", "意見", "黄", "優先度◎○△はDYM見解。根拠データはS03〜S12", "△のAIボットは優先度低として提示。採否は貴社判断"),
    (34, "第2の提案軸｜紹介・口コミの動線化", "事実", "黄", "POPER2022 n=300／ネオマーケティング2019（事実）。入会率の差70〜80% vs 40〜50%は業界推計値", "入会率の差を社内実績で裏付けできるか確認（上司確認①）"),
    (35, "LINEOA実績と出典一覧", "事実", "黄", "LINE国内MAU・開封率はLINEヤフー公式値（緑相当）。教育・スクール事例は★未取得", "lycbiz.com/jp/case-study から該当事例を取得。該当なしなら正直にそう書く"),
    (36, "裏表紙", "事実", "緑", "DYM会社情報（原本のまま）", "―"),
]
assert len(ROWS) == 36, len(ROWS)

wb = Workbook()

# ---------------- サマリ ----------------
ws = wb.active
ws.title = "サマリ"
ws.sheet_view.showGridLines = False

ws["B2"] = "教育（塾）業界_LINEOA施策提案（36枚）｜公表判断チェックリスト"
ws["B2"].font = Font(name=FONT, size=14, bold=True, color=NAVY)
ws["B3"] = "作成日：2026-09-15　対象ファイル：教育（塾）業界_LINEOA施策提案.pptx"
ws["B3"].font = Font(name=FONT, size=9, color="7F7F7F")

ws["B5"] = "背骨（この資料が言っていること・1文）"
ws["B5"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
ws["B6"] = ("保護者の検討は、塾を探す26日前に始まっている。塾が姿を現せるのは最後の数日だけ。"
            "そして最初の問いは「いくら」ではなく「いつから」。"
            "友だちは広告で買えるが、運用設計がなければ資産にならない（ビザビ371万人の凍結）。")
ws["B6"].font = Font(name=FONT, size=10)
ws["B6"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("B6:F6")
ws.row_dimensions[6].height = 32

ws["B8"] = "判定サマリ"
ws["B8"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
for i, (lab, cond, fill) in enumerate([
    ("緑（そのまま公表可）", "緑", G_FILL),
    ("黄（未取得・仮説明記のうえ公表可）", "黄", Y_FILL),
    ("赤（出典確定まで公表不可）", "赤", R_FILL),
]):
    r = 9 + i
    ws.cell(r, 2, lab).font = Font(name=FONT, size=10)
    c = ws.cell(r, 4, f'=COUNTIF(判定一覧!$D$3:$D$38,"{cond}")')
    c.font = Font(name=FONT, size=10, bold=True)
    c.alignment = Alignment(horizontal="center")
    c.fill = fill
    c.border = BOX
    ws.cell(r, 5, "枚").font = Font(name=FONT, size=10)
ws.cell(12, 2, "合計").font = Font(name=FONT, size=10, bold=True)
c = ws.cell(12, 4, "=SUM(D9:D11)")
c.font = Font(name=FONT, size=10, bold=True)
c.alignment = Alignment(horizontal="center")
c.border = BOX
ws.cell(12, 5, "枚").font = Font(name=FONT, size=10)

ws["B14"] = "納品条件"
ws["B14"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
ws["B15"] = '=IF(D11=0,"OK：赤0件。対外提出可（黄は未取得データの明記つきで提出可）","NG：赤"&D11&"件。出典を確定または該当数値を削除するまで対外提出不可")'
ws["B15"].font = Font(name=FONT, size=10, bold=True, color="C00000")
ws.merge_cells("B15:F15")

ws["B17"] = "★ 上司確認が必要な3点"
ws["B17"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
for i, t in enumerate([
    "① DYM社内の塾・スクール系広告実績（CV地点別のCPC・CVR・CPA）→ S4・S7・S8・S24・S34が黄→緑。塾は公的統計が存在しない領域なので、社内実績が唯一の一次根拠になる（最優先）",
    "② 競合8社の実名掲載OK（S12）→ NGなら「A社／B社」表記に変更。KUMON・ビザビは実測値つきなので特に確認が要る",
    "③ 民間調査の引用可否（塾ナラ／DeltaX「塾選」／POPER／明光ネットワークジャパン／オリコンME／インタースペース）→ 特に明光義塾は競合であり、その調査を引用している点の扱い",
]):
    r = 18 + i
    ws.cell(r, 2, t).font = Font(name=FONT, size=9.5)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)

ws["B22"] = "△ データ待ち（人が集める素材・5種に限定）"
ws["B22"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
for i, t in enumerate([
    "1. ✅取得済み Googleトレンド（塾／個別指導／予備校／家庭教師・夏期講習／冬期講習／春期講習・中学受験／大学受験／高校受験）→ S10・S25に反映済み。※年初来のみなので5年推移が要るなら再取得",
    "2. 上司確認3点セット（上記★）→ S4・S7・S8・S12・S24・S34が黄→緑",
    "3. ✅取得済み LINEヤフー前後検索（塾／塾 費用／教育費／奨学金／教育ローン）→ S5・S11に反映済み",
    "4. ✅取得済み 競合の友だち数実測（page.line.me・実機・2026-09-15）→ S12に反映済み。森塾のみ未確認",
    "5. トライ・早稲田アカデミーの診断の実物スクショ（捨て垢で1回通す）→ S18・S12の裏取り。残る最後の素材",
    "6. LINEヤフー公式の教育・スクール導入事例（lycbiz.com/jp/case-study）→ S35が黄→緑",
]):
    r = 23 + i
    ws.cell(r, 2, t).font = Font(name=FONT, size=9.5)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)

ws["B30"] = "凡例（このファイルの使い方）"
ws["B30"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
for i, t in enumerate([
    "・「判定一覧」タブのF列（残タスク）が、この資料を完成させるためのTODOです。",
    "・黄色で塗られたセルが記入欄です。データが入ったらD列の判定を「緑」に書き換えてください。",
    "・D列を書き換えると、このサマリの集計と納品条件の判定が自動で更新されます。",
]):
    ws.cell(31 + i, 2, t).font = Font(name=FONT, size=9.5)

for col, w in zip("ABCDEF", [2.5, 62, 14, 10, 8, 40]):
    ws.column_dimensions[col].width = w

# ---------------- 判定一覧 ----------------
ws2 = wb.create_sheet("判定一覧")
ws2.sheet_view.showGridLines = False
hdr = ["枚", "スライドタイトル", "主張の種類", "判定", "根拠・出典", "残タスク（これが埋まれば緑）"]
for j, h in enumerate(hdr, start=1):
    c = ws2.cell(2, j, h)
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BOX
ws2.freeze_panes = "A3"

FILLS = {"緑": G_FILL, "黄": Y_FILL, "赤": R_FILL}
for i, row in enumerate(ROWS):
    r = 3 + i
    for j, v in enumerate(row, start=1):
        c = ws2.cell(r, j, v)
        c.font = Font(name=FONT, size=9.5, bold=(j == 4))
        c.border = BOX
        c.alignment = Alignment(
            wrap_text=(j in (2, 5, 6)),
            horizontal="center" if j in (1, 3, 4) else "left",
            vertical="center")
    ws2.cell(r, 4).fill = FILLS[row[3]]
    if row[3] != "緑":
        ws2.cell(r, 6).fill = IN_FILL
    ws2.row_dimensions[r].height = 30

for col, w in zip("ABCDEF", [5, 42, 11, 7, 52, 46]):
    ws2.column_dimensions[col].width = w

wb.save(OUT)
print("saved:", OUT)
print("赤:", sum(1 for r in ROWS if r[3] == "赤"),
      "／黄:", sum(1 for r in ROWS if r[3] == "黄"),
      "／緑:", sum(1 for r in ROWS if r[3] == "緑"))
