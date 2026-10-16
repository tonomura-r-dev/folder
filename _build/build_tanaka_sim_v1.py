# -*- coding: utf-8 -*-
"""タナカホーム株式会社 LINEOA施策SIM ver1.0。新FMT ver3.00・1シート（XML直接編集でチェックボックスを保持）。

与件（2026-10-15〜16 殿村さん・商談議事録）：
  サイトUU 13,228／LINE友だち417人／来場予約 月35〜40件（目標50〜60件）／自社サイトの資料請求 月一桁
  今のLINE：来場特典のキャンペーンで友だちを集め、見学会・来場予約のご案内が中心
資料の軸：資料請求の時点からLINEでつながり、来場・契約まで追う → SIMのCVは来場予約（CPAもこれで割る）。
  資料請求（LINE経由）は参考として2行目に出すが、合計CV・CPAには入れない（CLAUDE.md 2026-09-29）。
友だちの入口：離脱防止（情報で追加してもらう導線に作り直す）＋サイトのLINEボタン＋サンクスLINE（資料請求・来場予約の完了ページ）。
  友だち追加広告（CPF）は入れない（殿村さんが入れると言ったときだけ）。
仮置き：現状ブロック率28%（殿村さんの表）から月1.02倍／資料請求 月8件／来場予約 月35件。

使い方:
    python _build/build_tanaka_sim_v1.py [出力.xlsx]
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import build_workbook, set_b, set_f, set_n, set_s  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_SIM_FMT_ver3.00.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "タナカホーム株式会社御中_LINEOA施策提案【SIM】ver1.0.xlsx"

UU, FRIENDS0, BLOCK_NOW, BLOCK_K = 13228, 417, 0.28, 1.02
VISIT, DOC = 35, 8                       # 来場予約・資料請求（月）＝サンクスLINEの母数（FMTの24行が SUM(H10:H11)×0.15）
CVR_VISIT = [0.004, 0.005, 0.006, 0.007, 0.008, 0.008]   # 来場予約＝クリック×CVR（0.3〜1.0%）
CVR_DOC = [0.006, 0.007, 0.008, 0.009, 0.010, 0.010]     # 資料請求＝クリック×CVR（参考。合計CVには入れない）
CONSULT_INIT, CONSULT_M = 100000, 200000     # LINE運用コンサル：初期10万円〜／月額20万円〜（費用目安・殿村さん確定）
EXIT_INIT, EXIT_M = 15000, 30000             # 離脱防止
THANKS_INIT, THANKS_M = 150000, 30000        # サンクスLINE誘導：初期15万円・月額3万円〜
POSTS, STEPS = 4, 5
SIM_NAME, CMP_NAME = "SIM_離脱防止＋サンクスLINE", "SIM_比較"
COLS = "EFGHIJ"


def fill_sim(root):
    set_s(root, "A4", "タナカホーム株式会社御中_LINEOA施策提案【SIM】（離脱防止＋サイトのボタン＋サンクスLINE）")
    set_n(root, "B2", 46311)                 # 作成日 2026-10-16
    set_s(root, "C8", "来場予約（LINE経由）")
    set_n(root, "D8", 0)
    set_s(root, "C9", "資料請求（LINE経由・参考）")
    set_n(root, "D9", 0)
    for ref in ("C10", "C11"):
        set_s(root, ref, "-")
    for ref in ("D10", "D11", "H9", "K8"):
        set_n(root, ref, 0)
    set_n(root, "H8", UU)
    set_n(root, "H10", VISIT)
    set_n(root, "H11", DOC)
    set_s(root, "G10", "来場予約（月）")
    set_s(root, "G11", "資料請求（月）")
    set_b(root, "P9", False)
    for ref in ("P32", "P33", "P34", "P35"):
        set_b(root, ref, True)
    # 現状
    set_n(root, "D27", BLOCK_NOW)
    set_n(root, "D28", FRIENDS0)
    set_f(root, "D29", "+D28*(1-D27)")
    for i, c in enumerate(COLS):
        prev = "D" if i == 0 else COLS[i - 1]
        set_s(root, f"{c}13", f"{i + 1}か月目")
        set_b(root, f"{c}17", False)          # CPF
        set_b(root, f"{c}18", True)           # 離脱防止
        set_b(root, f"{c}19", True)           # サンクスLINE誘導
        set_b(root, f"{c}20", False)
        set_b(root, f"{c}21", False)
        set_n(root, f"{c}22", 0)
        if i == 0:
            set_n(root, f"{c}27", BLOCK_NOW)
        else:
            set_f(root, f"{c}27", f"ROUND({prev}27*{BLOCK_K},3)")
        set_n(root, f"{c}30", POSTS)
        set_n(root, f"{c}32", STEPS)
        # ステップ配信の対象＝その月の新しい友だち（開始条件＝友だち追加）
        q24 = "*$Q$24" if i >= 3 else ""
        set_f(root, f"{c}33", f"+{c}26*{c}32*$P$26*$P$27*$P$36")
        set_f(root, f"{c}34", f"+{c}29*{c}30*$P$24{q24}+{c}26*{c}32")
        set_f(root, f"{c}36", f"ROUND({c}35*{CVR_VISIT[i]},0)")
        set_f(root, f"{c}37", f"ROUND({c}35*{CVR_DOC[i]},0)")
        set_n(root, f"{c}38", 0)
        set_n(root, f"{c}39", 0)
        # 合計CV・CVR・CPAは来場予約だけ（資料請求は参考）
        set_f(root, f"{c}40", f"+{c}36")
        set_f(root, f"{c}41", f"+IFERROR({c}36/{c}35,0)")
        set_f(root, f"{c}48", f"+IFERROR({c}36/{c}35,0)")
        set_f(root, f"{c}58", f"+IFERROR({c}36/{c}35,0)")
        set_f(root, f"{c}49", f"+IFERROR({c}51/{c}36,0)")
        set_f(root, f"{c}59", f"+IFERROR({c}55/{c}36,0)")
        set_f(root, f"{c}46", f"=SUM({c}65:{c}67)/({c}28-{prev}28)")
        set_f(root, f"{c}56", f"=SUM({c}65:{c}67)/({c}28-{prev}28)")
        set_n(root, f"{c}65", 0)
        set_n(root, f"{c}66", EXIT_M)
        set_n(root, f"{c}67", THANKS_M)
        set_n(root, f"{c}69", CONSULT_M)
        for r in (70, 71, 72):
            set_n(root, f"{c}{r}", 0)
        set_f(root, f"{c}73", f"SUM({c}65:{c}72)+{c}54")
    set_f(root, "D40", "+D36")
    for ref in ("D65", "D68", "D70", "D71", "D72"):
        set_n(root, ref, 0)
    set_n(root, "D66", EXIT_INIT)
    set_n(root, "D67", THANKS_INIT)
    set_n(root, "D69", CONSULT_INIT)
    set_s(root, "B67", "サンクスLINE誘導")
    set_s(root, "B69", "LINE運用コンサル")
    # 根拠メモ（N列）
    notes = {
        "N8": "サイトUU 13,228／友だち417人（2026-10-15〜16 殿村さん）。CVは来場予約（資料の軸＝資料請求の時点から来場・契約まで追う）",
        "N9": "資料請求（LINE経由）は参考。合計CV・CVR・CPAには入れない",
        "N10": f"来場予約 月{VISIT}件（議事録の現状35〜40件の下限）・資料請求 月{DOC}件（自社サイトは月一桁・仮に8件）＝サンクスLINEの母数",
        "N22": "友だち追加広告（CPF）は入れていない。友だちを月500人規模にするなら、LINE広告で月300人強（1人200円で月6〜7万円）が別に要る",
        "N23": "離脱防止＝UU×表示55%×クリック12%×追加12%（FMTの式）＝月約105人。今の来場特典のポップアップに加えて、来場前の方向けに情報で追加してもらう案内に作り直す想定",
        "N24": "サンクスLINE誘導＝（来場予約＋資料請求）×15%（FMTの式）＝月約6人。資料請求・来場予約の完了ページからLINEへ",
        "N26": "サイトのLINEボタン＝UU×0.5%（FMTのP15）＝月約66人。今はサイトにLINEの導線が無いので新しく置く",
        "N27": "ブロック率：28%（殿村さんの表）から前月の1.02倍ずつ上昇",
        "N30": "企画配信 月4本（有効な友だち×80%。4か月目からは反応しない方を外してさらに8割）",
        "N32": "ステップ配信5通（会社の特長→費用・住宅ローン→断熱性能・間取り→ルームツアー→見学会・来場予約）。対象＝その月の新しい友だち",
        "N36": "来場予約＝クリック×0.4〜0.8%（運用が育つ想定で月ごとに上げる）",
        "N37": "資料請求＝クリック×0.6〜1.0%（参考）",
        "N65": "従量課金は54行（新料金テーブル）で計算。65行の旧式は二重計上になるため0",
        "N66": "離脱防止：初期1.5万・月3万",
        "N67": "サンクスLINE誘導：初期15万・月3万〜。完了ページの件数が月40件ほどなので、1人あたりは約5,000円（目安の月100件を下回る）",
        "N68": "アカウント費：FMTの式。初期は0（アカウントは既にある）",
        "N69": "LINE運用コンサル：初期10万円〜・月額20万円〜（費用目安）",
    }
    for ref, text in notes.items():
        set_s(root, ref, text)


def fill_cmp(root, sim_name):
    set_f(root, "D7", f"+'{sim_name}'!G27")
    set_f(root, "E7", f"+'{sim_name}'!J27")
    set_s(root, "A15", "来場予約（もともと・月）")
    set_s(root, "A16", "来場予約（+LINEOA施策）")
    set_f(root, "C16", "+C17")
    set_f(root, "D16", "+D17")
    set_f(root, "E16", "+E17")
    set_s(root, "A17", "CV①_来場予約（LINE経由）")
    set_s(root, "A18", "参考_資料請求（LINE経由）")
    set_s(root, "A19", "-")
    set_s(root, "A20", "-")
    set_f(root, "D15", f"+'{sim_name}'!G16")
    set_f(root, "E15", f"+'{sim_name}'!J16")
    set_f(root, "C25", f"+'{sim_name}'!E68+C29*C30")
    set_f(root, "D25", f"+'{sim_name}'!G68+D29*D30")
    set_f(root, "E25", f"+'{sim_name}'!J68+E29*E30")


names = build_workbook(SRC, OUT, [(SIM_NAME, CMP_NAME, fill_sim, fill_cmp)], uid_seed="tanaka-sim-v1")
print("saved:", OUT.name)
print("sheets:", names)
