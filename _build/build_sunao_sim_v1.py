# -*- coding: utf-8 -*-
"""株式会社すなおコンサルティング LINEOA施策SIM ver1.1。新FMT ver3.00・1シート（XML直接編集でチェックボックスを保持）。

与件（2026-10-15〜19 殿村さん）：保険営業の方向けオンラインサロン「ハッピーライフ」。Meta広告 → UTAGEのLP → 無料セミナー（年4回ほど）。
  広告は1回約500万円のショット（ADの予算なのでSIMには入れない）。LINE公式アカウントの有無は未確認（0人スタートで作る）。
殿村さんの判断（2026-10-19）：
  ・友だち 月400人（殿村さんの表の数字）は、イベントの月（1・4か月目＝3か月おき）だけ。平常期は広告が無いので新しい友だちは入らない
  ・CVは殿村さんのCVR目安（クリックの0.3〜1.0%）で作る。表の「有効友だちの15%」は使わない
  ・ブロック率は殿村さんの表の数字（28.0%→31.0%）
友だちの入口：ポップアップツール（離脱防止）＋サンクスLINE誘導（資料の「友だちの入口」と同じ。LPのLINEボタン・CPFは入れない）。
  イベント月のLP UU 仮に46,000・セミナー予約 仮に1,400件 → FMTの式でポップアップ約191人＋サンクスLINE約210人＝約400人。
CV：①セミナー予約（イベント月だけ）／②サロン入会（サロンはいつでも申し込めるので平常期も）。
  CVの名前に「LINE経由」は書かない（2026-10-19 殿村さん）。
  合計CV・CVR・CPAは①＋②（タナカホームver1.1と同じ考え方）。
費用：資料の料金「プラン② イベントに合わせた運用」（初期10万円〜／平常期 月5〜10万円→10万で計算／イベント期 月30万円程度）
  ＋ポップアップツール（初期1.5万・月3万）＋サンクスLINE誘導（初期15万・月3万〜）。

使い方:
    python _build/build_sunao_sim_v1.py [出力.xlsx]
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import build_workbook, set_b, set_f, set_n, set_s  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_SIM_FMT_ver3.00.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "株式会社すなおコンサルティング御中_LINEOA施策提案【SIM】ver1.1.xlsx"

LP_UU, RESERVE = 46000, 1400                       # イベント月のLP UU・セミナー予約（仮に）
EVENT = [True, False, False, True, False, False]   # 3か月おきのイベント月
BLOCK = [0.28, 0.29, 0.295, 0.30, 0.305, 0.31]     # 殿村さんの表
POSTS = [4 if e else 2 for e in EVENT]             # 企画配信：イベント期 月4回／平常期 月2回
STEPS = 5                                          # ステップ配信（その月の新しい友だち）
CV_SEMINAR = [5, 0, 0, 7, 0, 0]                    # ①セミナー予約
CV_JOIN = [1, 1, 1, 1, 1, 1]                       # ②サロン入会
CONSULT_INIT = 100000
CONSULT_M = [300000 if e else 100000 for e in EVENT]
EXIT_INIT, EXIT_M = 15000, 30000                   # ポップアップツール（離脱防止）
THANKS_INIT, THANKS_M = 150000, 30000              # サンクスLINE誘導
SIM_NAME, CMP_NAME = "SIM_ポップアップ＋サンクスLINE", "SIM_比較"
COLS = "EFGHIJ"


def fill_sim(root):
    set_s(root, "A4", "株式会社すなおコンサルティング御中_LINEOA施策提案【SIM】（ポップアップツール＋サンクスLINE誘導）")
    set_n(root, "B2", 46314)                 # 作成日 2026-10-19
    set_s(root, "C8", "セミナー予約")
    set_s(root, "C9", "サロン入会")
    for ref in ("C10", "C11", "G11"):
        set_s(root, ref, "-")
    for ref in ("D8", "D9", "D10", "D11", "H11", "K8"):
        set_n(root, ref, 0)
    set_s(root, "G8", "LP UU（イベント月）")
    set_s(root, "G9", "うち広告")
    set_s(root, "G10", "セミナー予約（イベント月）")
    set_n(root, "H8", LP_UU)
    set_n(root, "H9", LP_UU)
    set_n(root, "H10", RESERVE)
    set_b(root, "P9", False)
    for ref in ("P32", "P33", "P34", "P35"):
        set_b(root, ref, True)
    # 現状（LINE公式アカウントの有無は未確認 → 0人スタート）
    set_n(root, "D27", 0)
    set_n(root, "D28", 0)
    set_f(root, "D29", "+D28*(1-D27)")
    set_f(root, "D14", "+H8")
    set_f(root, "D15", "+H9")
    set_f(root, "D16", "+H10")
    for i, c in enumerate(COLS):
        prev = "D" if i == 0 else COLS[i - 1]
        ev = EVENT[i]
        set_s(root, f"{c}12", "イベント期" if ev else "平常期")
        set_s(root, f"{c}13", f"{i + 1}か月目")
        # 広告が出るのはイベント月だけ（平常期はLPの受付も止まっている）
        if ev:
            set_f(root, f"{c}14", "+$H$8")
            set_f(root, f"{c}15", "+$H$9")
            set_f(root, f"{c}16", "+$H$10")
        else:
            for r in (14, 15, 16):
                set_n(root, f"{c}{r}", 0)
        set_b(root, f"{c}17", False)          # CPF
        set_b(root, f"{c}18", True)           # ポップアップツール（離脱防止）
        set_b(root, f"{c}19", True)           # サンクスLINE誘導
        set_b(root, f"{c}20", False)
        set_b(root, f"{c}21", False)
        set_n(root, f"{c}22", 0)
        set_f(root, f"{c}23", f"IF({c}18=FALSE,0,{c}15*$P$18*0.08*0.08+({c}14-{c}15)*$P$19*0.12*0.1)")
        set_f(root, f"{c}24", f"IF({c}19=FALSE,0,{c}16*0.15)")
        set_f(root, f"{c}26", f"+SUM({c}22:{c}25)")          # LPのLINEボタン・プレスリリースは入れない
        set_n(root, f"{c}27", BLOCK[i])
        set_n(root, f"{c}30", POSTS[i])
        set_n(root, f"{c}32", STEPS)
        # ステップ配信の対象＝その月の新しい友だち（開始条件＝友だち追加）
        q24 = "*$Q$24" if i >= 3 else ""
        set_f(root, f"{c}33", f"+{c}26*{c}32*$P$26*$P$27*$P$36")
        set_f(root, f"{c}34", f"+{c}29*{c}30*$P$24{q24}+{c}26*{c}32")
        set_n(root, f"{c}36", CV_SEMINAR[i])
        set_n(root, f"{c}37", CV_JOIN[i])
        set_n(root, f"{c}38", 0)
        set_n(root, f"{c}39", 0)
        set_f(root, f"{c}46", f'IFERROR(SUM({c}65:{c}67)/({c}28-{prev}28),"-")')
        set_f(root, f"{c}56", f'IFERROR(SUM({c}65:{c}67)/({c}28-{prev}28),"-")')
        set_n(root, f"{c}65", 0)
        set_n(root, f"{c}66", EXIT_M)
        set_n(root, f"{c}67", THANKS_M)
        set_n(root, f"{c}69", CONSULT_M[i])
        for r in (70, 71, 72):
            set_n(root, f"{c}{r}", 0)
        set_f(root, f"{c}73", f"SUM({c}65:{c}72)+{c}54")
    for ref in ("D65", "D70", "D71", "D72"):
        set_n(root, ref, 0)
    set_n(root, "D66", EXIT_INIT)
    set_n(root, "D67", THANKS_INIT)
    set_n(root, "D69", CONSULT_INIT)
    for ref in ("B18", "B23", "B66"):                           # 資料と同じ呼び方
        set_s(root, ref, "ポップアップツール")
    set_s(root, "B67", "サンクスLINE誘導")
    set_s(root, "B69", "LINE運用コンサル（プラン②）")
    # 根拠メモ（N列）
    notes = {
        "N8": "CV①セミナー予約。CPAは①＋②の合計で計算",
        "N9": "CV②サロン入会。サロンはサイトからいつでも申し込めるので、平常期も入会がある",
        "N10": f"イベント月（3か月おき・1/4か月目）に広告からLP UU 仮に{LP_UU:,}・セミナー予約 仮に{RESERVE:,}件"
               "（友だち 月400人の目安から逆算）。広告費はADの予算なので入れていない",
        "N14": "LP UU：広告が出るイベント月だけ。平常期は0（セミナーの受付も止まっている）",
        "N22": "友だち追加広告（CPF）は入れていない",
        "N23": "ポップアップツール＝LP UU×表示65%×クリック8%×追加8%（FMTの式）＝イベント月 約191人",
        "N24": "サンクスLINE誘導＝セミナー予約×15%（FMTの式）＝イベント月 約210人。予約した方に、日程ごとに前日・当日のお知らせ",
        "N26": "友だちの入口はポップアップツール＋サンクスLINE誘導だけ（資料の「友だちの入口」と同じ）。イベント月 約400人",
        "N27": "ブロック率：殿村さんの表（28.0%→31.0%）",
        "N30": "企画配信：イベント期 月4回／平常期 月2回（資料の料金 プラン②と同じ）",
        "N32": "ステップ配信5通（その月の新しい友だち。セミナーの見どころ→参加特典→予約の案内）",
        "N36": "セミナー予約＝イベント月だけ。クリック×0.7〜0.8%",
        "N37": "サロン入会＝毎月1件（イベント直後の申込みフォームの案内＋平常期の配信から）。CVRは①＋②で0.6〜0.9%",
        "N65": "従量課金は54行（新料金テーブル）で計算。65行の旧式は二重計上になるため0",
        "N66": "ポップアップツール：初期1.5万・月3万",
        "N67": "サンクスLINE誘導：初期15万・月3万〜。予約がイベント月1,400件（仮に）なので、1人あたり約430円（3か月分÷210人）",
        "N68": "アカウント費：FMTの式（開設3万円）",
        "N69": "LINE運用コンサル（プラン②）：初期10万円〜／平常期 月5〜10万円（10万で計算）／イベント期 月30万円程度",
    }
    for ref, text in notes.items():
        set_s(root, ref, text)


def fill_cmp(root, sim_name):
    set_f(root, "D7", f"+'{sim_name}'!G27")
    set_f(root, "E7", f"+'{sim_name}'!J27")
    set_s(root, "A15", "セミナー予約（広告・イベント1回あたり）")
    for col in "CDE":
        set_f(root, f"{col}15", f"+'{sim_name}'!$H$10")
        set_f(root, f"{col}16", f"SUM({col}17:{col}20)")
    set_s(root, "A16", "CV数（+LINEOA施策）")
    set_s(root, "A17", "CV①_セミナー予約")
    set_s(root, "A18", "CV②_サロン入会")
    set_s(root, "A19", "-")
    set_s(root, "A20", "-")
    set_f(root, "C25", f"+'{sim_name}'!E68+C29*C30")
    set_f(root, "D25", f"+'{sim_name}'!G68+D29*D30")
    set_f(root, "E25", f"+'{sim_name}'!J68+E29*E30")


names = build_workbook(SRC, OUT, [(SIM_NAME, CMP_NAME, fill_sim, fill_cmp)], uid_seed="sunao-sim-v1")
print("saved:", OUT.name)
print("sheets:", names)
