# -*- coding: utf-8 -*-
"""株式会社ワクトリ LINEOA施策SIM ver1.0。新FMT ver3.00・プラン2つ（XML直接編集でチェックボックスを保持）。

与件（2026-10-25〜26 殿村さん。`_drafts/ワクトリ_与件整理（2026-10-25）.md` の「決まったこと」7）：
  サイトUU 85,887／LINE友だち 79,904／12月開始（仮）の半年。SIMはざっくりでよい
  CV＝応募（サイトの「ENTRY」と求人ページの「応募」は同じ登録フォーム）。LINE経由はLINE計測用のフォームで数える
  友だちの入口：サイトのLINEボタン（FMTのUU×0.5%）＋ポップアップ（FMTの式）＋サンクスLINE（応募×遷移80%×追加45%）
  サンクスLINEで追加した方は応募済みなので、ステップ配信からの応募には数えない（33行の対象から外す）
  プラン①（月3投稿）：月20万円（ポップアップ込み＝コンサル17万＋ポップアップ3万）＋サンクスLINE 月3万円
  プラン②（月5投稿）：月30万円（ポップアップ込み＝コンサル27万＋ポップアップ3万）。サンクスLINEはオプションなので入れない
  成果報酬型：応募1件 4,000円（仮）。各シートの「■成果報酬運営」と比較シートの成果報酬プランの行で見る
仮置き：サイトからの応募 月400件／今の友だちのブロック率40%（その後は前の月×1.02）／広告UUは不明なので0

    python _build/build_wakutori_sim_v1.py [出力.xlsx]
"""
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import build_workbook, set_b, set_f, set_n, set_s  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_SIM_FMT_ver3.00.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "株式会社ワクトリ御中_LINEOA施策提案【SIM】ver1.0.xlsx"

UU, FRIENDS0, BLOCK_NOW, BLOCK_K = 85887, 79904, 0.40, 1.02
APPLY = 400                                    # サイトからの応募（月・仮）＝サンクスLINEの母数
UNIT = 4000                                    # 成果報酬型の応募1件の単価（仮）
CVR = [0.003, 0.0035, 0.004, 0.005, 0.0055, 0.006]   # 応募＝クリック×CVR（0.3〜1.0%の中で、運用が育つ想定で上げる）
INIT, POP_INIT, POP_M = 100000, 15000, 30000          # 運用設計・構築／ポップアップ
THANKS_INIT, THANKS_M = 150000, 30000                 # サンクスLINE誘導
STEPS, NOW_POSTS = 5, 4
COLS = "EFGHIJ"
PLANS = [  # (シート名, 比較シート名, 企画配信の本数, コンサル月額, サンクスLINEを入れるか, 見出し)
    ("SIM①_月3投稿", "SIM①_比較", 3, 170000, True, "プラン①（月3投稿）：ポップアップ＋サンクスLINE"),
    ("SIM②_月5投稿", "SIM②_比較", 5, 270000, False, "プラン②（月5投稿）：ポップアップ"),
]


def make_fill_sim(posts, consult, thanks, label):
    def fill(root):
        set_s(root, "A4", f"株式会社ワクトリ御中_LINEOA施策提案【SIM】{label}")
        set_n(root, "B2", (datetime.date(2026, 10, 26) - datetime.date(1899, 12, 30)).days)
        # 要件定義
        set_s(root, "C8", "応募")
        set_n(root, "D8", UNIT)
        for ref in ("C9", "C10", "C11", "G11"):
            set_s(root, ref, "-")
        for ref in ("D9", "D10", "D11", "H9", "H11", "K8"):
            set_n(root, ref, 0)
        set_n(root, "H8", UU)
        set_s(root, "G10", "応募（月）")
        set_n(root, "H10", APPLY)
        set_b(root, "P9", False)
        for ref in ("P32", "P33", "P34", "P35"):
            set_b(root, ref, True)
        # 現状
        set_n(root, "D27", BLOCK_NOW)
        set_n(root, "D28", FRIENDS0)
        set_f(root, "D29", "+D28*(1-D27)")
        set_n(root, "D30", NOW_POSTS)
        for i, c in enumerate(COLS):
            prev = "D" if i == 0 else COLS[i - 1]
            set_s(root, f"{c}13", f"{i + 1}か月目")
            set_b(root, f"{c}17", False)               # CPF
            set_b(root, f"{c}18", True)                # ポップアップ
            set_b(root, f"{c}19", thanks)              # サンクスLINE誘導
            set_b(root, f"{c}20", False)
            set_b(root, f"{c}21", False)
            set_n(root, f"{c}22", 0)
            set_f(root, f"{c}24", f"IF({c}19=FALSE,0,$H$10*0.8*0.45)")
            if i == 0:
                set_n(root, f"{c}27", BLOCK_NOW)
            else:
                set_f(root, f"{c}27", f"ROUND({prev}27*{BLOCK_K},3)")
            set_n(root, f"{c}30", posts)
            set_n(root, f"{c}32", STEPS)
            q24 = "*$Q$24" if i >= 3 else ""
            # ステップ配信の対象＝その月の新しい友だち（サンクスLINEで追加した応募済みの方は除く）
            set_f(root, f"{c}33", f"+({c}26-{c}24)*{c}32*$P$26*$P$27*$P$36")
            set_f(root, f"{c}34", f"+{c}29*{c}30*$P$24{q24}+{c}26*{c}32")
            set_f(root, f"{c}36", f"ROUND({c}35*{CVR[i]},0)")
            for r in (37, 38, 39):
                set_n(root, f"{c}{r}", 0)
            # 成果報酬プランの固定費にコンサル（69行）を入れない（CLAUDE.md 2026-10-20）
            set_f(root, f"{c}45", f"SUM({c}65:{c}68)+SUM({c}71:{c}72)")
            # 費用
            set_n(root, f"{c}65", 0)                   # 旧料金の式は54行（新料金テーブル）と二重になるので0
            set_n(root, f"{c}66", POP_M)
            set_n(root, f"{c}67", THANKS_M if thanks else 0)
            set_n(root, f"{c}69", consult)
            for r in (70, 71, 72):
                set_n(root, f"{c}{r}", 0)
            set_f(root, f"{c}73", f"SUM({c}65:{c}72)+{c}54")
        for ref in ("D65", "D68", "D70", "D71", "D72"):
            set_n(root, ref, 0)
        set_n(root, "D66", POP_INIT)
        set_n(root, "D67", THANKS_INIT if thanks else 0)
        set_n(root, "D69", INIT)
        set_s(root, "B66", "ポップアップ")
        set_s(root, "B67", "サンクスLINE誘導")
        set_s(root, "B69", "LINE運用コンサル")
        # 根拠メモ（N列）
        notes = {
            "N8": f"CV＝応募（サイトの「ENTRY」と求人ページの「応募」は同じ登録フォーム）。LINE経由は計測用のフォームで数える。成果報酬型の単価 {UNIT:,}円（仮）",
            "N10": f"サイトからの応募は与件に無いので、仮に月{APPLY}件（広告費 月約250万円÷1件あたり約6,000円の目安）＝サンクスLINEの母数",
            "N15": "広告UUは不明なので0（UU全体をサイトのLINEボタン0.5%・ポップアップの式で計算）",
            "N18": "ポップアップ＝UU×表示55%×クリック12%×追加12%（FMTの式）",
            "N22": "友だち追加広告（CPF）は入れていない",
            "N24": "サンクスLINE誘導＝応募×遷移80%×追加45%。マイページ登録（応募）の完了画面からLINEへ" if thanks
                   else "サンクスLINE誘導はオプション（このプランのSIMには入れていない）",
            "N26": "サイトのLINEボタン（UU×0.5%・FMTのP15）＋ポップアップ＋サンクスLINE",
            "N27": f"ブロック率：今は不明なので仮に{BLOCK_NOW:.0%}。その後は前の月×{BLOCK_K}",
            "N28": f"今の友だち {FRIENDS0:,}人（プロフィールの表示・2026年10月）",
            "N30": f"企画配信 月{posts}本（開始時期・エリア別に出し分け。対象は有効な友だちの80%、4か月目から反応しない方を外してさらに8割）。今は週1回（月4本）",
            "N32": f"ステップ配信{STEPS}通（友だち追加から15日以内）。対象＝その月の新しい友だち（サンクスLINEで追加した応募済みの方は除く）",
            "N36": "応募＝想定クリック×0.3〜0.6%（運用が育つ想定で月ごとに上げる）。サンクスLINEで追加した方は応募済みなので数えない",
            "N45": "成果報酬型の固定費＝ツール・アカウント費・通数費（コンサル費は入れない）。成果報酬＝応募×単価",
            "N54": "通数費は54行（新料金テーブル）で計算",
            "N66": "ポップアップ：初期1.5万円・月3万円（プランの月額に込み）",
            "N67": "サンクスLINE誘導：初期15万円・月3万円〜" if thanks else "サンクスLINE誘導：オプション（初期15万円・月3万円〜）",
            "N68": "アカウント費：FMTの式。初期は0（アカウントは既にある）",
            "N69": f"LINE運用コンサル：初期10万円〜（運用設計・構築）・月額{consult // 10000}万円（ポップアップ3万円と合わせてプランの月額{(consult + POP_M) // 10000}万円）",
        }
        for ref, text in notes.items():
            set_s(root, ref, text)
    return fill


def fill_cmp(root, sim_name):
    set_f(root, "D7", f"+'{sim_name}'!G27")
    set_f(root, "E7", f"+'{sim_name}'!J27")
    set_s(root, "A15", "応募（サイト全体・月）")
    set_s(root, "A16", "応募（+LINEOA施策）")
    for c in "CDE":
        set_f(root, f"{c}16", f"+{c}17")
    set_s(root, "A17", "CV①_応募（LINE経由）")
    for ref in ("A18", "A19", "A20"):
        set_s(root, ref, "-")
    set_f(root, "D15", f"+'{sim_name}'!G16")
    set_f(root, "E15", f"+'{sim_name}'!J16")
    set_f(root, "C25", f"+'{sim_name}'!E68+C29*C30")
    set_f(root, "D25", f"+'{sim_name}'!G68+D29*D30")
    set_f(root, "E25", f"+'{sim_name}'!J68+E29*E30")


plans = [(n, c, make_fill_sim(p, k, t, lab), fill_cmp) for n, c, p, k, t, lab in PLANS]
names = build_workbook(SRC, OUT, plans, uid_seed="wakutori-sim-v1")
print("saved:", OUT.name)
