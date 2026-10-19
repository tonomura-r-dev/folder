# -*- coding: utf-8 -*-
"""清水グループ様 LINEOA施策SIM ver1.0。新FMT ver3.00（XML直接編集でチェックボックスを保持）。

与件（2026-10-19 殿村さん）：飲食店スタッフ・イベント設営／警備／施設スタッフ・正社員総合職など、グループ全体の求人。
  求人サイト 月8,947UU（与件）。LINE公式アカウントは0人スタート。
友だちの入口：サイトのLINEボタン（FMTの係数 0.5%）＋離脱防止ポップアップ（定義テキストの係数 表示65%×12.5%×12.5%）
  ＋応募完了画面のサンクスLINE誘導（定義テキストの係数 遷移80%×追加45%）。
  サイトからの応募数は与件に無いので、仮にUUの1.5%（月134件）で置く。
CV：①応募（LINEのリンクからの応募だけ。サンクスLINEで友だちになった方の応募は数えない）。CVRはクリックの0.3〜1.0%前後。
費用：概算は「竹（A：コンサル基本）」で計算。月額20万円は業界資料・AIの案どちらも同じで、初期だけ違う。
  料金はAIの案の金額（松D 初期50万・月35万／竹A 初期30万・月20万／梅C 初期60万・月0）で確定（2026-10-19 殿村さん「Bで進めていい」）。
  ＋ポップアップツール（初期1.5万・月3万）＋サンクスLINE誘導（初期15万・月3万〜）。

使い方:
    python _build/build_shimizu_sim_v1.py [出力.xlsx]
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import build_workbook, set_b, set_f, set_n, set_s  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_SIM_FMT_ver3.00.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "清水グループ御中_LINEOA施策提案【SIM】ver1.0.xlsx"

UU = 8947                                             # 求人サイトの月間UU（与件）
APPLY = round(UU * 0.015)                             # サイトからの応募 仮にUUの1.5%＝月134件
BLOCK_START, BLOCK_UP = 0.34, 1.025                   # ブロック率 1か月目34%（2026-10-19 殿村さん）。2か月目以降は前の月×1.025（FMTの式）
POSTS = 3                                             # 企画配信 月3回（業界資料の竹＝コンサル基本 月20万円は月3投稿）
STEPS = 3                                             # ステップ配信3通
CV_APPLY = [1, 1, 2, 2, 3, 3]                         # ①応募（クリックの0.3〜1.0%前後）
CONSULT_M = 200000                                    # 竹（A：コンサル基本）の月額
EXIT_INIT, EXIT_M = 15000, 30000                      # ポップアップツール（離脱防止）
THANKS_INIT, THANKS_M = 150000, 30000                 # サンクスLINE誘導
COLS = "EFGHIJ"
PLANS = [("SIM_コンサル基本", "SIM_比較", 300000, "初期30万円")]   # 2026-10-19 殿村さん「Bで進めていい」＝AIの案の金額


def make_fill_sim(consult_init, consult_note):
    def fill_sim(root):
        set_s(root, "A4", "清水グループ御中_LINEOA施策提案【SIM】（離脱防止＋サイトのLINEボタン＋サンクスLINE誘導）")
        set_n(root, "B2", 46314)                      # 作成日 2026-10-19
        set_s(root, "C8", "応募")
        for ref in ("C9", "C10", "C11"):
            set_s(root, ref, "-")
        for ref in ("D8", "D9", "D10", "D11", "K8"):
            set_n(root, ref, 0)
        set_s(root, "G8", "サイトUU")
        set_s(root, "G9", "うち広告")
        set_s(root, "G10", "応募（サイト・仮に）")
        set_n(root, "H8", UU)
        set_n(root, "H9", 0)
        set_n(root, "H10", APPLY)
        set_b(root, "P9", False)
        set_n(root, "D27", 0)
        set_n(root, "D28", 0)
        set_f(root, "D29", "+D28*(1-D27)")
        for i, c in enumerate(COLS):
            prev = "D" if i == 0 else COLS[i - 1]
            set_s(root, f"{c}13", f"{i + 1}か月目")
            set_b(root, f"{c}17", False)         # CPF
            set_b(root, f"{c}18", True)          # 離脱防止ポップアップ
            set_b(root, f"{c}19", True)          # サンクスLINE誘導
            set_b(root, f"{c}20", False)
            set_b(root, f"{c}21", False)
            set_n(root, f"{c}22", 0)
            set_f(root, f"{c}23", f"IF({c}18=FALSE,0,{c}14*$P$18*0.125*0.125)")
            set_f(root, f"{c}24", f"IF({c}19=FALSE,0,{c}16*0.8*0.45)")
            if i == 0:
                set_n(root, f"{c}27", BLOCK_START)
            else:
                set_f(root, f"{c}27", f"+{prev}27*{BLOCK_UP}")
            set_n(root, f"{c}30", POSTS)
            set_n(root, f"{c}32", STEPS)
            set_n(root, f"{c}36", CV_APPLY[i])
            for r in (37, 38, 39):
                set_n(root, f"{c}{r}", 0)
            set_f(root, f"{c}46", f'IFERROR(SUM({c}65:{c}67)/({c}28-{prev}28),"-")')
            set_f(root, f"{c}56", f'IFERROR(SUM({c}65:{c}67)/({c}28-{prev}28),"-")')
            set_n(root, f"{c}65", 0)
            set_n(root, f"{c}66", EXIT_M)
            set_n(root, f"{c}67", THANKS_M)
            set_n(root, f"{c}69", CONSULT_M)
            for r in (70, 71, 72):
                set_n(root, f"{c}{r}", 0)
            set_f(root, f"{c}73", f"SUM({c}65:{c}72)+{c}54")
        for ref in ("D65", "D70", "D71", "D72"):
            set_n(root, ref, 0)
        set_n(root, "D66", EXIT_INIT)
        set_n(root, "D67", THANKS_INIT)
        set_n(root, "D69", consult_init)
        for ref in ("B18", "B23", "B66"):
            set_s(root, ref, "離脱防止ポップアップ")
        set_s(root, "B67", "サンクスLINE誘導")
        set_s(root, "B69", "LINE運用コンサル（コンサル基本）")
        notes = {
            "N8": "CV①応募：LINEのリンクからの応募だけ。サンクスLINEで友だちになった方（応募済み）の応募は数えない",
            "N10": f"サイトからの応募は与件に無いので、仮にUUの1.5%（月{APPLY}件）",
            "N14": f"求人サイトの月間UU {UU:,}（与件）。6か月同じ",
            "N23": "離脱防止ポップアップ＝UU×表示65%×クリック12.5%×追加12.5%（定義テキストの係数）",
            "N24": "サンクスLINE誘導＝応募×遷移80%×追加45%（定義テキストの係数）",
            "N26": "サイトのLINEボタン（UUの0.5%・FMTの係数）＋離脱防止＋サンクスLINE誘導",
            "N27": "ブロック率：1か月目34%（殿村さん）。2か月目以降は前の月×1.025（FMTの式）",
            "N30": "企画配信 月3回（業界資料のコンサル基本 月20万円は月3投稿。新着求人・職種ごとの案内）",
            "N32": "ステップ配信3通（友だち追加の翌日以降。職種ごとに中身を分ける）",
            "N36": "応募：クリックの0.3〜1.0%前後",
            "N66": "離脱防止ポップアップ：初期1.5万円・月3万円",
            "N67": "サンクスLINE誘導：初期15万円・月3万円〜",
            "N68": "アカウント費：FMTの式（開設3万円）",
            "N69": f"LINE運用コンサル（コンサル基本）月20万円。初期は{consult_note}",
        }
        for ref, text in notes.items():
            set_s(root, ref, text)
    return fill_sim


def fill_cmp(root, sim_name):
    set_f(root, "D7", f"+'{sim_name}'!G27")
    set_f(root, "E7", f"+'{sim_name}'!J27")
    set_s(root, "A15", "応募（サイト・仮に）")
    for col in "CDE":
        set_f(root, f"{col}15", f"+'{sim_name}'!$H$10")
        set_f(root, f"{col}16", f"SUM({col}17:{col}20)")
    set_s(root, "A16", "CV数（+LINEOA施策）")
    set_s(root, "A17", "CV①_応募")
    for ref in ("A18", "A19", "A20"):
        set_s(root, ref, "-")
    set_f(root, "C25", f"+'{sim_name}'!E68+C29*C30")
    set_f(root, "D25", f"+'{sim_name}'!G68+D29*D30")
    set_f(root, "E25", f"+'{sim_name}'!J68+E29*E30")


plans = [(sim, cmp_, make_fill_sim(init, note), fill_cmp) for sim, cmp_, init, note in PLANS]
names = build_workbook(SRC, OUT, plans, uid_seed="shimizu-sim-v1")
print("saved:", OUT.name)
print("sheets:", names)
