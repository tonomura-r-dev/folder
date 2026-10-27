# -*- coding: utf-8 -*-
"""タナカホーム SIM：殿村さんのPC保存版（中身は ver1.2 と同じ）を、A_サンクス有／B_サンクス無 の2プランに分ける（2026-10-27 殿村さん）。

    python _build/build_tanaka_sim_v12_ab.py <PC保存版SIM.xlsx> <出力.xlsx>

  佐村さん「サンクスLINE誘導はOptionで入れといて、最悪そこだけの導入でもOK」→ 資料ではプランから外してオプションに。
  A_サンクス有：PC保存版のまま（シート名だけ変える）。
  A：24行（サンクスLINEの友だち）を FMTの15% → 遷移80%×追加40%（32%）に。
  B_サンクス無：19行（サンクスLINE誘導）を全部FALSE、67行（費用）を0に。CVの件数（36・37行）はAのまま
    （サンクスLINEの効果＝来場予約した方のキャンセルを減らす＝「来場」の数なので、SIMには出さない。殿村さん確認）。
  XMLを直接書き換える（チェックボックスを消さない）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import build_workbook, replace_sheet_ref, set_b, set_f, set_n, set_s  # noqa: E402

OLD = "'SIM_離脱防止＋サンクスLINE'!"
COLS = "EFGHIJ"


def fill_a(root):
    set_s(root, "A4", "タナカホーム株式会社御中_LINEOA施策提案【SIM】ver1.2　A_サンクス有（離脱防止＋サイトのボタン＋サンクスLINE誘導）")
    for c in COLS:                                   # サンクスLINEの友だち＝（来場予約＋資料請求）×遷移80%×追加40%（横浜シミズ・ワクトリと同じ係数。2026-10-27 殿村さんOK）
        set_f(root, f"{c}24", f"IF({c}19=FALSE,0,SUM($H$10:$H$11)*0.8*0.4)")
    set_s(root, "N24", "サンクスLINE誘導＝（来場予約 月35件＋資料請求 月8件。全部フォームの想定）×遷移80%×追加40%＝月約14人。"
                       "資料請求・来場予約の件数はBと同じ（来場予約した方のキャンセルを減らす効果は「来場」の数なのでSIMに出さない）")


def fill_b(root):
    set_s(root, "A4", "タナカホーム株式会社御中_LINEOA施策提案【SIM】ver1.2　B_サンクス無（離脱防止＋サイトのボタン）")
    for c in COLS:
        set_b(root, f"{c}19", False)
    for c in "D" + COLS:
        set_n(root, f"{c}67", 0)
    set_s(root, "N24", "B_サンクス無：サンクスLINE誘導は入れない（オプション。Aのシートを参照）")
    set_s(root, "N67", "サンクスLINE誘導：入れない（オプション：初期15万・月3万〜）")
    set_s(root, "N36", "CVの件数はAと同じ（サンクスLINEの効果＝来場予約した方のキャンセルを減らす＝「来場」の数なので、SIMには出さない）")


def cmp_fill(root, sim_name):
    replace_sheet_ref(root, OLD, f"'{sim_name}'!")


src, out = sys.argv[1], sys.argv[2]
build_workbook(src, out, [("A_サンクス有", "A_比較", fill_a, cmp_fill), ("B_サンクス無", "B_比較", fill_b, cmp_fill)],
               uid_seed="tanaka-sim-v12-ab")
print("saved:", out)
