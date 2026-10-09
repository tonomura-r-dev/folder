# -*- coding: utf-8 -*-
"""株式会社あらうち（100円賃貸・仲介手数料100円）LINEOA施策SIM ver1.4。新FMT ver3.00・1シート。

ver1.4（2026-10-09 殿村さん）：サンクスLINEを外した（問い合わせ済みの方が友だちになるだけでCVは増えず、費用だけCPAに乗る。
  問い合わせ月30件（仮）×15%＝月4〜5人、月3万で1人約6,700円＝離脱防止の3〜5倍で見合わない）。CPAは資料に載せない。
  入口は離脱防止＋サイトのLINEボタン（定番の組み合わせ）。CVRは「LINEのクリック経由の問い合わせ÷クリック」で確定（②だけ）。
  コンサル初期20万・月10万で確定。サンクスLINEを入れる場合の費用は初期10万・月3万〜（今は入れない）。

以下は ver1.3 の説明：

ver1.2 からの変更（2026-10-09 殿村さん指摘「前提がズレてる」）：
  ・友だち追加広告（CPF）をやめた。30万／100万はADの提案予算で、LINEのSIMに入れるものではなかった（CHETと同じ取り違え）
  ・友だちの入口は「離脱防止」「サンクスLINE」「サイトのLINEボタン」の3つ（FMTの式をそのまま使う）
  ・予算のパターンが無くなったので1シート
  ・ブロック率：CPF流入が無くなったので、現状35%（仮置き）から月1%ずつ上昇（用済み型）
  ・問い合わせ①＝（新規友だち − サンクスLINE経由）×25%。サンクスLINE経由の方はフォームで問い合わせ済みなので数えない
検算（ワークフローで独立に再計算）を受けて直したこと：
  ・ステップ配信の対象を「その月の新しい友だち」に（FMTの「ターゲット数×20%」だと月27人の案件で約6倍に出ていた）
  ・初期のアカウント費3万（FMTの残り）を0に／知名度の係数をOFF（反応率1.46倍→1.26倍）／CPFの22行を0固定（#DIV/0!防止）
  ・73行に新料金テーブル（54行）を足す／比較シートのアカウント費・参照先・ラベルをSIMに合わせる／根拠メモを式に合わせる

与件（2026-10-09 殿村さん）：UU 1,737／LINE友だち1,241人／ターゲット＝都内在住の女性オフィスワーカー。
仮置き：サイトのフォームの問い合わせ 月30件（UUの約1.7%）／現状ブロック率35%／URL送付率25%。

使い方:
    python _build/build_arauchi_sim_v14.py [出力.xlsx]
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import build_workbook, set_b, set_f, set_n, set_s  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_SIM_FMT_ver3.00.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "株式会社あらうち御中_LINEOA施策提案【SIM】ver1.4.xlsx"

UU, FRIENDS0, BLOCK_NOW = 1737, 1241, 0.35   # サイトUU・現状の友だち数（2026-10-09 殿村さん提供）／現状ブロック率（仮置き）
SITE_CV = 30                        # サイトのフォームの問い合わせ（月）仮置き＝サンクスLINEの母数
BLOCK_K = 1.01                      # ブロック率の月ごとの上昇（用済み型：引っ越しが済めば離脱）
URL_RATE = 0.25                     # 新規友だち（サンクスLINE経由を除く）のうち、LINEで物件URLを送ってくる率（問い合わせ①）仮置き
CVR2 = [0.004, 0.005, 0.006, 0.008, 0.008, 0.008]   # 問い合わせ②＝配信クリック×CVR（0.3〜1.0%の範囲）。4か月目はFMTの係数でクリックが減るので、件数が下がらないよう0.8%に
CONSULT_INIT, CONSULT_M = 200000, 100000     # LINE運用コンサル：初期／月額
EXIT_INIT, EXIT_M = 15000, 30000             # 離脱防止：初期／月額（資料S19）
THANKS_INIT, THANKS_M = 0, 0                 # サンクスLINEは入れない（入れる場合は初期10万・月3万〜）
POSTS, STEPS = 4, 10
SIM_NAME, CMP_NAME = "SIM_離脱防止＋サイトのボタン", "SIM_比較"
COLS = "EFGHIJ"


def fill_sim(root):
    set_s(root, "A4", "株式会社あらうち御中_LINEOA施策提案【SIM】（離脱防止＋サイトのLINEボタン）")
    # 要件定義
    set_s(root, "C8", "問い合わせ①（LINEで物件URL送付）")
    set_n(root, "D8", 0)
    set_s(root, "C9", "問い合わせ②（配信経由）")
    set_n(root, "D9", 0)
    for ref in ("C10", "C11"):
        set_s(root, ref, "-")
    for ref in ("D10", "D11", "H9", "H11", "K8"):
        set_n(root, ref, 0)
    set_n(root, "H8", UU)
    set_n(root, "H10", SITE_CV)          # サンクスLINEの母数（FMTの24行が SUM(H10:H11)×0.15 で使う）
    set_b(root, "P9", False)
    set_b(root, "P34", False)            # 知名度強さ：OFF（友だち1,241人・広告なし。特典＝100円、クリエイティブ・動線＝作り直し後の想定でON）
    set_n(root, "B2", 46304)             # 作成日 2026-10-09（Excelの日付シリアル）
    # 現状
    set_n(root, "D27", BLOCK_NOW)
    set_n(root, "D28", FRIENDS0)
    set_f(root, "D29", "+D28*(1-D27)")   # FMTは0固定（比較シートの「実装前」に出る）
    for i, c in enumerate(COLS):
        prev = "D" if i == 0 else COLS[i - 1]
        set_s(root, f"{c}13", f"{i + 1}か月目")
        # 施策トグル（チェックボックス）：離脱防止だけON（サイトのボタンはFMTの26行で常に入る）。CPF・サンクスLINE・通知メッセージはOFF
        set_b(root, f"{c}17", False)
        set_b(root, f"{c}18", True)
        set_b(root, f"{c}19", False)
        set_b(root, f"{c}20", False)
        set_b(root, f"{c}21", False)
        set_n(root, f"{c}22", 0)             # CPFは使わない（P30が空なので式にすると #DIV/0! の元になる）
        # 23行（離脱防止）・24行（サンクスLINE）・26行（合計）はFMTの式のまま
        # ブロック率
        if i == 0:
            set_n(root, f"{c}27", BLOCK_NOW)
        else:
            set_f(root, f"{c}27", f"ROUND({prev}27*{BLOCK_K},3)")
        set_n(root, f"{c}30", POSTS)
        set_n(root, f"{c}32", STEPS)
        # ステップ配信の対象＝その月の新しい友だち（FMTは「ターゲット数×20%」で、月27人の案件では約6倍に出る）
        q24 = "*$Q$24" if i >= 3 else ""
        set_f(root, f"{c}33", f"+{c}26*{c}32*$P$26*$P$27*$P$36")
        set_f(root, f"{c}34", f"+{c}29*{c}30*$P$24{q24}+{c}26*{c}32")
        # 問い合わせ①：サンクスLINE経由（24行）はフォームで問い合わせ済みなので除く
        set_f(root, f"{c}36", f"ROUND(({c}26-{c}24)*{URL_RATE},0)")
        set_f(root, f"{c}37", f"ROUND({c}35*{CVR2[i]},0)")
        set_n(root, f"{c}38", 0)
        set_n(root, f"{c}39", 0)
        # CVRは②（配信経由）だけ。①はクリック経由ではないので分母に合わない
        set_f(root, f"{c}41", f"+IFERROR({c}37/{c}35,0)")
        set_f(root, f"{c}48", f"+IFERROR({c}37/{c}35,0)")
        set_f(root, f"{c}58", f"+IFERROR({c}37/{c}35,0)")
        # 友だち獲得単価＝（離脱防止＋サンクスLINEの月額）÷増えた友だち
        set_f(root, f"{c}46", f"=(SUM({c}65:{c}67)+{c}71)/({c}28-{prev}28)")
        set_f(root, f"{c}56", f"=(SUM({c}65:{c}67)+{c}71)/({c}28-{prev}28)")
        # 費用
        set_n(root, f"{c}65", 0)             # 旧式の従量課金。54行（新料金テーブル）で計算するので0
        set_n(root, f"{c}66", EXIT_M)
        set_n(root, f"{c}67", THANKS_M)
        set_n(root, f"{c}69", CONSULT_M)
        set_n(root, f"{c}70", 0)
        set_n(root, f"{c}71", 0)
        set_n(root, f"{c}72", 0)
        set_f(root, f"{c}73", f"SUM({c}65:{c}72)+{c}54")
    for ref in ("D65", "D68", "D70", "D71", "D72"):
        set_n(root, ref, 0)
    set_n(root, "D66", EXIT_INIT)
    set_n(root, "D67", THANKS_INIT)
    set_n(root, "D69", CONSULT_INIT)
    set_s(root, "B67", "サンクスLINE")
    set_s(root, "B69", "LINE運用コンサル")
    # 根拠メモ（N列＝PDF範囲外）
    set_s(root, "N8", "UU 1,737/月・友だち1,241人（2026-10-09 殿村さん提供）。ターゲット＝都内在住の女性オフィスワーカー")
    set_s(root, "N10", f"サイトのフォームの問い合わせ 月{SITE_CV}件（仮置き。UUの約1.7%）。今の状態の参考値（サンクスLINEは入れないのでSIMの計算には使わない）")
    set_s(root, "N22", "友だち追加広告（CPF）は使わない（30万／100万はADの提案予算でLINEのSIMには入れない）")
    set_s(root, "N23", "離脱防止＝UU×表示55%×クリック12%×追加12%（FMTの式）。3か月目以降は前月の1.01倍（FMTの式）")
    set_s(root, "N24", "サンクスLINEは入れない：問い合わせ済みの方が友だちになるのでCVは増えず、費用だけCPAに乗る。問い合わせ月30件（仮）×15%＝月4〜5人、月3万で1人約6,700円（初期10万込みで約1万円）＝離脱防止（約2,200円）の3〜5倍")
    set_s(root, "N26", "サイトのLINEボタン＝UU×0.5%（FMTのP15）＝月約9人。リンク先を友だち追加URLに差し替える前提")
    set_s(root, "N27", "ブロック率：現状35%（仮置き）から前月の1.01倍ずつ上昇（6か月目で37%。用済み型。友だち追加広告の流入が無いので上げ幅は小さい）")
    set_s(root, "N31", "企画配信：有効な友だち×80%（FMT）。4か月目からは反応しない方・問い合わせ済みの方を外して、さらに8割に絞る想定（FMTのQ24）")
    set_s(root, "N32", "ステップ配信10通（サイト経由・サンクスLINE経由で各10通。企画案の表2）")
    set_s(root, "N33", "ステップ配信の対象＝その月の新しい友だち（開始条件＝友だち追加）。FMTの「ターゲット数×20%」は使わない")
    set_s(root, "N46", "友だち獲得単価＝離脱防止の月額÷増えた友だち（無料のサイトのボタン経由も含む）")
    set_s(root, "N36", f"問い合わせ①＝（新規友だち−サンクスLINE経由）×{int(URL_RATE * 100)}%（LINEで物件URLを送る＝問い合わせ。仮置き）。サンクスLINE経由はフォームで問い合わせ済みなので数えない")
    set_s(root, "N37", "問い合わせ②＝配信クリック×0.4〜0.8%（運用が育つ想定で月ごとに上げる）")
    set_s(root, "N41", "CVR＝LINEのクリック経由の問い合わせ（②）÷クリック（2026-10-09 殿村さん確定）。①は友だち追加の直後に物件URLを送る問い合わせで、クリックを通らないので入れない")
    set_s(root, "N65", "従量課金は54行（新料金テーブル）で計算。65行の旧式（2.75円）は二重計上になるため0")
    set_s(root, "N66", f"離脱防止：初期{EXIT_INIT // 10000 if EXIT_INIT % 10000 == 0 else EXIT_INIT / 10000}万・月{EXIT_M // 10000}万（資料S19）")
    set_s(root, "N67", "サンクスLINE：入れない（0）。入れる目安＝問い合わせ後の来店・内見・成約を追いたい案件で、フォームの問い合わせが月100件以上（友だち月15人・1人約2,000円で離脱防止並み）。費用は初期10万・月3万〜")
    set_s(root, "N68", "アカウント費：配信5,000通未満は月5,000円（FMTの式）。初期は0（アカウントは既にある）")
    set_s(root, "N69", "LINE運用コンサル：初期20万（アカウントの作り直し：あいさつ・リッチメニュー・自動応答・ステップ配信）／月10万（企画4本＋ステップ10本の運用）。2026-10-09 殿村さん確定")


def fill_cmp(root, sim_name):
    # ブロック率はSIMの値をそのまま出す（FMTの比較シートは友だち数から逆算していて合わない）
    set_f(root, "D7", f"+'{sim_name}'!G27")
    set_f(root, "E7", f"+'{sim_name}'!J27")
    set_s(root, "A17", "CV①_問い合わせ①（LINEで物件URL送付）")
    set_s(root, "A18", "CV②_問い合わせ②（配信経由）")
    set_s(root, "A19", "-")
    set_s(root, "A20", "-")
    set_f(root, "D15", f"+'{sim_name}'!G16")
    set_f(root, "E15", f"+'{sim_name}'!J16")
    set_f(root, "C25", f"+'{sim_name}'!E68+C29*C30")
    set_f(root, "D25", f"+'{sim_name}'!G68+D29*D30")
    set_f(root, "E25", f"+'{sim_name}'!J68+E29*E30")


names = build_workbook(SRC, OUT, [(SIM_NAME, CMP_NAME, fill_sim, fill_cmp)], uid_seed="arauchi-sim-v14")
print("saved:", OUT.name)
print("sheets:", names)
