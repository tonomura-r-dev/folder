# -*- coding: utf-8 -*-
"""株式会社あらうち（100円賃貸・仲介手数料100円）LINEOA施策SIM ver1.2。新FMT ver3.00・SNS予算30万円／100万円の2パターン。

ver1.1（別セッション `_build/build_arauchi_sim.py`・openpyxlで保存）から変えたこと：
  ・XMLを直接書き換えて作る → FMTのチェックボックス（featurePropertyBag）を残す（ver1.1では消えていた）
  ・65行（旧式の従量課金）を0にする → 54行（新料金テーブル）と二重計上になるのを防ぐ（100万は配信が月2.4万通を超える）
  ・CVR（41・48・58行）は問い合わせ②（配信経由）だけで計算 → 0.4〜0.8%（①は友だち追加直後のURL送付でクリック経由ではない）
  ・ラベル：B69「LINE運用コンサル」、比較シートのCV内訳①②
  前提（与件の仮置き）は ver1.1 と同じ。

与件（2026-10-09 殿村さん）：目標CPA 1,000円（問い合わせ）／提案予算 月30万・100万／現状広告なし／
  ターゲット＝都内在住の女性オフィスワーカー／UU 1,737／LINE友だち1,241人（LINEで問い合わせ可）。
設計（9/25に提案した案B）：
  問い合わせ①＝新規友だち×URL送付率25%（100円賃貸の問い合わせは物件URLを送るだけで成立）
  問い合わせ②＝配信（企画4本＋ステップ10本）のクリック×0.4〜0.8%
  友だち獲得＝SNS予算を全額 LINE友だち追加広告（CPF 300円/人）。離脱防止はOFF。
  費用＝LINE運用コンサル 初期20万・月10万＋LINE公式アカウント費（FMTの自動式）。成果報酬は置かない。

使い方:
    python _build/build_arauchi_sim_xml.py [出力.xlsx]
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import build_workbook, set_b, set_f, set_n, set_s  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_SIM_FMT_ver3.00.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "株式会社あらうち御中_LINEOA施策提案【SIM】ver1.2.xlsx"

CPF_UNIT = 300                      # 友だち追加広告の単価（円/人）仮置き：LINEヤフー公式事例（UZUZ）の約300円
URL_RATE = 0.25                     # 新規友だちのうち、LINEで物件URLを送ってくる率（問い合わせ①）仮置き
CVR2 = [0.004, 0.005, 0.006, 0.007, 0.008, 0.008]   # 問い合わせ②＝配信クリック×CVR（0.3〜1.0%の範囲）
BLOCK0, BLOCK_K = 0.412, 1.015      # ブロック率：用済み型＋CPF流入で高め。初月41.2%→月1.5%ずつ上昇
UU, FRIENDS0, BLOCK_NOW = 1737, 1241, 0.35   # サイトUU・現状の友だち数（2026-10-09 殿村さん提供）／現状ブロック率（仮置き）
CONSULT_INIT, CONSULT_M = 200000, 100000     # LINE運用コンサル：初期／月額
POSTS, STEPS = 4, 10
PLANS = (("SIM1_SNS予算30万円", "SIM1_比較", 300000, "30万円"),
         ("SIM2_SNS予算100万円", "SIM2_比較", 1000000, "100万円"))
COLS = "EFGHIJ"


def fill_sim(root, budget, label):
    set_s(root, "A4", f"株式会社あらうち御中_LINEOA施策提案【SIM】（SNS予算 {label}／月）")
    # 要件定義
    set_s(root, "C8", "問い合わせ①（LINEで物件URL送付）")
    set_n(root, "D8", 0)
    set_s(root, "C9", "問い合わせ②（配信経由）")
    set_n(root, "D9", 0)
    for ref in ("C10", "C11"):
        set_s(root, ref, "-")
    for ref in ("D10", "D11", "H9", "H10", "H11", "K8"):
        set_n(root, ref, 0)
    set_n(root, "H8", UU)
    set_b(root, "P9", False)
    set_n(root, "P30", CPF_UNIT)
    # 現状
    set_n(root, "D27", BLOCK_NOW)
    set_n(root, "D28", FRIENDS0)
    set_f(root, "D29", "+D28*(1-D27)")   # FMTは0固定（比較シートの「実装前」に出る）
    for i, c in enumerate(COLS):
        prev = "D" if i == 0 else COLS[i - 1]
        set_s(root, f"{c}13", f"{i + 1}か月目")
        # 施策トグル（チェックボックス）：CPFだけON
        set_b(root, f"{c}17", True)
        for r in (18, 19, 20, 21):
            set_b(root, f"{c}{r}", False)
        # CPF増分＝CPF広告費（71行）÷単価（FMTの式は65行＝従量課金を見ていて通らない）
        set_f(root, f"{c}22", f"IF({c}17=FALSE,0,{c}71/$P$30)")
        # ブロック率
        if i == 0:
            set_n(root, f"{c}27", BLOCK0)
        else:
            set_f(root, f"{c}27", f"ROUND({prev}27*{BLOCK_K},3)")
        set_n(root, f"{c}30", POSTS)
        set_n(root, f"{c}32", STEPS)
        # 問い合わせ①②
        set_f(root, f"{c}36", f"ROUND({c}26*{URL_RATE},0)")
        set_f(root, f"{c}37", f"ROUND({c}35*{CVR2[i]},0)")
        set_n(root, f"{c}38", 0)
        set_n(root, f"{c}39", 0)
        # CVRは②（配信経由）だけ。①はクリック経由ではないので分母に合わない
        set_f(root, f"{c}41", f"+IFERROR({c}37/{c}35,0)")
        set_f(root, f"{c}48", f"+IFERROR({c}37/{c}35,0)")
        set_f(root, f"{c}58", f"+IFERROR({c}37/{c}35,0)")
        # 友だち獲得単価＝（LINEの月額＋CPF広告費）÷増えた友だち
        set_f(root, f"{c}46", f"=(SUM({c}65:{c}67)+{c}71)/({c}28-{prev}28)")
        set_f(root, f"{c}56", f"=(SUM({c}65:{c}67)+{c}71)/({c}28-{prev}28)")
        # 費用
        set_n(root, f"{c}65", 0)             # 旧式の従量課金。54行（新料金テーブル）で計算するので0
        set_n(root, f"{c}66", 0)             # 離脱防止 OFF
        set_n(root, f"{c}67", 0)             # サンクス誘導なし
        set_n(root, f"{c}69", CONSULT_M)
        set_n(root, f"{c}70", 0)
        set_n(root, f"{c}71", budget)
        set_n(root, f"{c}72", 0)
    for ref in ("D65", "D66", "D67", "D70", "D71", "D72"):
        set_n(root, ref, 0)
    set_n(root, "D69", CONSULT_INIT)
    set_s(root, "B69", "LINE運用コンサル")
    set_s(root, "B71", "CPF広告費（SNS予算）")
    # 根拠メモ（N列＝PDF範囲外）
    set_s(root, "N8", "UU 1,737/月・友だち1,241人（2026-10-09 殿村さん提供）。ターゲット＝都内在住の女性オフィスワーカー")
    set_s(root, "N22", f"CPF：SNS予算÷{CPF_UNIT}円/人（LINE友だち追加広告・仮置き。LINEヤフー公式事例UZUZの約300円）")
    set_s(root, "N23", "離脱防止はOFF：UU1,737では月14人・約2,200円/人でCPF（300円）より割高")
    set_s(root, "N27", "ブロック率：用済み型（引っ越しが済めば離脱）＋CPF流入で高め。41.2%→月1.5%ずつ上昇（現状35%は仮置き）")
    set_s(root, "N36", f"問い合わせ①＝新規友だち×{int(URL_RATE * 100)}%（LINEで物件URLを送る＝問い合わせ。仮置き）")
    set_s(root, "N37", "問い合わせ②＝配信クリック×0.4〜0.8%（運用が育つ想定で月ごとに上げる）")
    set_s(root, "N41", "CVRは②（配信経由）だけで計算。①は友だち追加の直後に物件URLを送る問い合わせで、配信のクリック経由ではない")
    set_s(root, "N65", "従量課金は54行（新料金テーブル）で計算。65行の旧式（2.75円）は二重計上になるため0")
    set_s(root, "N69", "LINE運用コンサル：初期20万（あいさつ・リッチメニュー・URL受付の自動応答・CPF設定）／月10万（CPF運用＋企画4本＋ステップ10本）")
    set_s(root, "N71", "SNS予算は全額CPF（LINE友だち追加広告）に充当。広告運用手数料は含まない（AD側のSIMに合わせる）")


def fill_cmp(root, sim_name):
    # ブロック率はSIMの値をそのまま出す（FMTの比較シートは友だち数から逆算していて合わない）
    set_f(root, "D7", f"+'{sim_name}'!G27")
    set_f(root, "E7", f"+'{sim_name}'!J27")
    set_s(root, "A17", "CV①_問い合わせ①（LINEで物件URL送付）")
    set_s(root, "A18", "CV②_問い合わせ②（配信経由）")


plans = []
for sim_name, cmp_name, budget, label in PLANS:
    plans.append((sim_name, cmp_name,
                  (lambda b, l: lambda root: fill_sim(root, b, l))(budget, label),
                  fill_cmp))
names = build_workbook(SRC, OUT, plans, uid_seed="arauchi-sim")
print("saved:", OUT.name)
print("sheets:", names)
