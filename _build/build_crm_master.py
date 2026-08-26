# -*- coding: utf-8 -*-
"""CRMスプシに移す用のマスタを、ストックとショットで別々に作る。

  python _build/build_crm_master.py <月次管理ブック.xlsx> [出力先ベース名]

佐村さんの指示（2026-08-25）
  ・CRMのスプシに移行する
  ・ストックとショットは別々に作る（アタックの切り口が違うため）

出力
  <ベース名>.xlsx        : ストック / ショット の2シート
  <ベース名>_ストック.tsv : CRMの新規タブに貼る用
  <ベース名>_ショット.tsv : 同上

1行=1エンドクライアント。両方の契約形態を持つ会社は両方のシートに出る
（それぞれ該当ぶんの金額だけを集計）。

元ブックは読むだけで一切変更しない。
"""
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from build_attack_list import extract  # 抽出処理は共通

FONT = "メイリオ"
NAVY = "1F3864"
YEN = "#,##0;-#,##0;\"-\""
PCT = "0.0%"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

KEIS = ["ストック", "ショット"]
HEADERS = ["エンドクライアント名", "社名（請求先）", "売上", "粗利", "粗利率",
           "案件数", "媒体", "商材カテゴリ", "担当", "最新対象月", "もう一方の契約形態"]
WIDTHS = [34, 30, 14, 13, 9, 8, 20, 26, 28, 12, 18]


def f(sz=10, bold=False, color="000000"):
    return Font(name=FONT, size=sz, bold=bold, color=color)


def month_key(s):
    """「8月分」などから並べ替え用の数値を作る。判別できないものは0。"""
    digits = ""
    for ch in s:
        if ch.isdigit():
            digits += ch
        elif digits:
            break
    return int(digits) if digits else 0


def aggregate(recs, kei):
    """契約形態 kei の案件だけを、エンドクライアント名ごとに1行にまとめる。"""
    both = defaultdict(set)  # そのエンドが持つ契約形態の一覧
    for r in recs:
        if r["契約形態"] in KEIS:
            both[r["エンドクライアント名"]].add(r["契約形態"])

    g = defaultdict(lambda: {"売上": 0.0, "粗利": 0.0, "件数": 0, "社名": [],
                             "媒体": [], "商材": [], "担当": [], "月": []})
    for r in recs:
        if r["契約形態"] != kei or not r["エンドクライアント名"]:
            continue
        d = g[r["エンドクライアント名"]]
        d["件数"] += 1
        d["売上"] += r["売上"]
        d["粗利"] += r["粗利"]
        for src, dst in (("社名", "社名"), ("媒体", "媒体"), ("商材カテゴリ", "商材")):
            if r[src] and r[src] not in d[dst]:
                d[dst].append(r[src])
        for t in r["担当"].split(" / "):
            if t and t not in d["担当"]:
                d["担当"].append(t)
        if r["対象月"] and r["対象月"] not in d["月"]:
            d["月"].append(r["対象月"])

    other = "ショット" if kei == "ストック" else "ストック"
    rows = []
    for k in sorted(g, key=lambda x: -g[x]["売上"]):
        d = g[k]
        rows.append([
            k, " / ".join(d["社名"]), d["売上"], d["粗利"],
            (d["粗利"] / d["売上"] if d["売上"] else None),
            d["件数"], " / ".join(d["媒体"]), " / ".join(d["商材"]),
            " / ".join(d["担当"]),
            max(d["月"], key=month_key) if d["月"] else "",
            f"{other}あり" if other in both[k] else "",
        ])
    return rows


def write_sheet(wb, kei, rows):
    ws = wb.create_sheet(kei)
    ws["A1"] = f"{kei}｜CRM移行用マスタ（1行=1エンドクライアント）"
    ws["A1"].font = f(14, bold=True, color=NAVY)
    dual = sum(1 for r in rows if r[10])
    ws["A2"] = (f"{len(rows)}社／金額はこの契約形態ぶんだけの合計（税抜）。"
                f"うち{dual}社はもう一方の契約形態も持っています。")
    ws["A2"].font = f(9, color="44546A")

    for i, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = f(10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[3].height = 30

    for j, row in enumerate(rows):
        rr = 4 + j
        for i, v in enumerate(row, start=1):
            c = ws.cell(row=rr, column=i, value=v)
            c.font = f(9)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(i in (2, 7, 8, 9)))
        for i in (3, 4):
            ws.cell(row=rr, column=i).number_format = YEN
        ws.cell(row=rr, column=5).number_format = PCT

    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(HEADERS))}{3 + len(rows)}"
    return ws


def write_tsv(path, rows):
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\t".join(HEADERS) + "\n")
        for row in rows:
            out = []
            for i, v in enumerate(row):
                if v is None:
                    out.append("")
                elif i == 4:  # 粗利率だけ小数、金額は整数で出す
                    out.append(f"{v:.4f}")
                elif isinstance(v, float):
                    out.append(f"{v:.0f}")
                else:
                    out.append(str(v))
            fh.write("\t".join(out) + "\n")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    base = sys.argv[2] if len(sys.argv) > 2 else "CRM移行マスタ"

    recs, skipped = extract(src)
    wb = Workbook()
    wb.remove(wb.active)
    for kei in KEIS:
        rows = aggregate(recs, kei)
        write_sheet(wb, kei, rows)
        write_tsv(f"{base}_{kei}.tsv", rows)
        dual = sum(1 for r in rows if r[10])
        print(f"  {kei:6} {len(rows):>4}社  売上 {sum(r[2] for r in rows):>16,.0f}  "
              f"粗利 {sum(r[3] for r in rows):>14,.0f}  （両方持ち {dual}社）")
    wb.save(f"{base}.xlsx")
    print(f"書き出しました: {base}.xlsx / {base}_ストック.tsv / {base}_ショット.tsv")
    for sheet, miss in skipped:
        print(f"  ★スキップ {sheet}: {miss}")


if __name__ == "__main__":
    main()
