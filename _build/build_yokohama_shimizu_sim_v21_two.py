# -*- coding: utf-8 -*-
"""横浜シミズ SIM ver2.1：殿村さんのPC保存版 ver2.0（成果単価 6,000円入り）を、コンサル版と成果報酬版の2シートに分ける。

    python _build/build_yokohama_shimizu_sim_v21_two.py <PC保存版 ver2.0.xlsx> <出力.xlsx>

  2026-10-27 殿村さん「コンサルverと成果報酬verで組みたい」（佐村さんFB「成果報酬の単発利益大きめに乗せたパターンも提示しようか」）。
  ・SIM①_コンサル：月額 215,000円（プラン②）。成果単価は0。成果報酬運営（43〜52行）は非表示
  ・SIM②_成果報酬：月額 65,000円（ツール・アカウント費）＋応募1件 6,000円（新規応募・再応募とも）。コンサル運営（53〜62行）は非表示
  ・比較シートも、もう一方のプランの行（コスト・CPC・CPA）を非表示にする
  友だち・CV・費用の数字は ver2.0 のまま。行は消さずに非表示にするだけ（FMTの形はそのまま）。XML直接編集でチェックボックスを残す。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import N, build_workbook, get_cell, q, set_n, set_s  # noqa: E402

UNIT = 20000
TITLE = "株式会社横浜シミズ御中_LINEOA施策提案【SIM】"


def hide_rows(root, rows):
    for r in root.iter(q("row")):
        if int(r.get("r")) in rows:
            r.set("hidden", "1")


def make_fill_sim(unit, label, hide):
    def fill(root):
        set_s(root, "A4", f"{TITLE}{label}（サイトのLINEボタン＋離脱防止＋サンクスLINE誘導）")
        set_n(root, "D8", unit)
        set_n(root, "D10", unit)
        hide_rows(root, hide)
    return fill


def make_fill_cmp(hide):
    def fill(root, sim_name):
        for f in root.iter(q("f")):
            if f.text and "SIM!" in f.text:
                f.text = f.text.replace("SIM!", f"'{sim_name}'!")
        hide_rows(root, hide)
    return fill


def main(src, out):
    plans = [
        ("SIM①_コンサル", "SIM①_比較", make_fill_sim(0, "コンサル", set(range(43, 53))), make_fill_cmp({4, 12, 22})),
        ("SIM②_成果報酬", "SIM②_比較", make_fill_sim(UNIT, f"成果報酬（応募1件 {UNIT:,}円）", set(range(53, 63))), make_fill_cmp({5, 13, 23})),
    ]
    names = build_workbook(src, out, plans, uid_seed="yokohama-shimizu-sim-v21")
    print("saved:", out, names)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
