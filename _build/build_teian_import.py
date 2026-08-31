# -*- coding: utf-8 -*-
"""月次管理ブックの既存顧客を「LINEOA提案管理シート」と同じ列並びに組み替える。

  python _build/build_teian_import.py <月次管理ブック.xlsx> [出力先ベース名]
  オプション:
    --除外アカウント貸し   サービス詳細が「アカウント貸し」の案件を集計から外す

提案管理シートの見出し（15列）にそのまま合わせる。
  提案 / 結論 / ☑ / 追客 / 企業名 / Ac名 / 既存 / 商流 / 外注先 /
  提案商材 / 業界 / 担当① / 担当② / DYM売上 / LINE売上

埋める列と、空欄のままにする列
  企業名   = エンドクライアント名（1行=1社）
  既存     = ●（このリストは全部が既存顧客なので）
  商流     = 直案件が1件でもあれば「直」、無くて代理店経由があれば「代理店」、
             どちらも判別できなければ「不明」
  外注先   = 代理店経由の請求先。エンド名と違うものだけを連結する。
             ここを飛ばして直接エンドに電話すると代理店の頭越しになるので残す
  担当①   = アカウント担当（ASPは広告主担当①）
  担当②   = コンサル担当（ASPは媒体担当①）
  DYM売上  = 全媒体の実績売上の合計（税抜）
  LINE売上 = そのうち媒体「LINE公式アカウント」ぶんだけ
  提案 / 結論 / ☑ / 追客 / Ac名 / 提案商材 / 業界 = 空欄。
             これから決めることなので、実績データからは埋められない

売上は「提案時の見込み」ではなく実績。数字の意味が違うので、
入れてよいか迷ったら参考列（P列以降）だけ使って本体は空欄にすること。

出力
  <ベース名>.xlsx           15列＋参考列。貼る前の確認用
  <ベース名>_貼り付け用.tsv  15列ちょうど・見出し無し。そのまま行に貼れる
  <ベース名>_参考列.tsv      参考列だけ。P列以降に貼りたいとき用

元ブックは読むだけで一切変更しない。
"""
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from build_attack_list import extract  # 抽出処理は既存アタック用リストと共通

FONT = "メイリオ"
NAVY = "1F3864"
GRAY = "808080"
LIGHT = "F2F2F2"
YEN = "#,##0;-#,##0;\"-\""
PCT = "0.0%"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# 提案管理シートの見出し。ここを動かすと貼り付け位置がずれるので触らない
SHEET_COLS = ["提案", "結論", "☑", "追客", "企業名", "Ac名", "既存", "商流",
              "外注先", "提案商材", "業界", "担当①", "担当②", "DYM売上", "LINE売上"]
# 判断材料として右に足す列。貼るかどうかは任意
REF_COLS = ["契約形態", "案件数", "実績粗利", "粗利率", "媒体", "商材カテゴリ",
            "担当（全員）", "最新対象月", "アカウント貸し"]

WIDTHS = ([10, 10, 5, 8, 30, 16, 6, 8, 24, 14, 12, 12, 12, 14, 14] +
          [16, 8, 13, 9, 20, 24, 26, 12, 12])

LINE_MEDIA = "LINE公式アカウント"
KAIGASHI = "アカウント貸し"
KEIS = ["ストック", "ショット"]


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


def push(lst, v):
    if v and v not in lst:
        lst.append(v)


def aggregate(recs):
    """エンドクライアント単位で1行にまとめる。売上の大きい順＝アタックの優先順。"""
    g = defaultdict(lambda: {
        "売上": 0.0, "LINE売上": 0.0, "粗利": 0.0, "件数": 0, "貸し": 0,
        "直": False, "代理店": False, "外注先": [], "kei": set(),
        "媒体": [], "商材": [], "担当": [], "月": [], "acct": "", "cons": "",
    })

    for r in recs:
        end = r["エンドクライアント名"]
        if not end:
            continue
        d = g[end]
        d["件数"] += 1
        d["売上"] += r["売上"]
        d["粗利"] += r["粗利"]
        if r["媒体"] == LINE_MEDIA:
            d["LINE売上"] += r["売上"]
        if KAIGASHI in r["サービス詳細"]:
            d["貸し"] += 1

        shubetsu = r["種別"]
        if "直" in shubetsu:
            d["直"] = True
        elif "代理店" in shubetsu:
            d["代理店"] = True
        # 代理店経由の請求先はエンドと違う会社。頭越しを防ぐために残す
        if "代理店" in shubetsu and r["社名"] and r["社名"] != end:
            push(d["外注先"], r["社名"])

        if r["契約形態"] in KEIS:
            d["kei"].add(r["契約形態"])
        push(d["媒体"], r["媒体"])
        push(d["商材"], r["商材カテゴリ"])
        push(d["月"], r["対象月"])
        for t in r["担当"].split(" / "):
            push(d["担当"], t)
        # 担当①②は最初に出てきた案件のものを採用する（売上順に見ていないので
        # 空欄だったときだけ後の案件で埋める）
        if not d["acct"]:
            d["acct"] = r["アカウント担当"]
        if not d["cons"]:
            d["cons"] = r["コンサル担当"]

    def keiyaku(s):
        if len(s) == 2:
            return "ストック＋ショット"
        return next(iter(s)) if s else "（未設定）"

    rows = []
    for name in sorted(g, key=lambda k: -g[k]["売上"]):
        d = g[name]
        shoryu = "直" if d["直"] else ("代理店" if d["代理店"] else "不明")
        sheet = ["", "", "", "",                       # 提案 / 結論 / ☑ / 追客
                 name, "", "●", shoryu,                # 企業名 / Ac名 / 既存 / 商流
                 " / ".join(d["外注先"]), "", "",       # 外注先 / 提案商材 / 業界
                 d["acct"], d["cons"],                 # 担当① / 担当②
                 d["売上"], d["LINE売上"]]              # DYM売上 / LINE売上
        ref = [keiyaku(d["kei"]), d["件数"], d["粗利"],
               (d["粗利"] / d["売上"] if d["売上"] else None),
               " / ".join(d["媒体"]), " / ".join(d["商材"]),
               " / ".join(d["担当"]),
               max(d["月"], key=month_key) if d["月"] else "",
               f"●（{d['貸し']}件）" if d["貸し"] else ""]
        rows.append(sheet + ref)
    return rows


# ============================================================ 書き出し
def write_xlsx(path, rows, excluded):
    wb = Workbook()
    ws = wb.active
    ws.title = "提案管理シート用"

    ws["A1"] = "提案管理シート用｜列の並びを提案管理シートに合わせたもの"
    ws["A1"].font = f(14, bold=True, color=NAVY)
    ws["A2"] = (f"{len(rows)}社／A〜O列が提案管理シートと同じ並び。P列から右は判断材料で、"
                "貼るかどうかは任意。金額は実績（税抜）で、提案時の見込みではない。"
                + ("／アカウント貸しの案件は除外済み" if excluded else ""))
    ws["A2"].font = f(9, color="44546A")

    head = SHEET_COLS + REF_COLS
    for i, h in enumerate(head, start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = f(10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY if i <= len(SHEET_COLS) else GRAY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[3].height = 30

    money = (14, 15, 18)
    wrap = (9, 20, 21, 22)
    for j, row in enumerate(rows):
        rr = 4 + j
        for i, v in enumerate(row, start=1):
            c = ws.cell(row=rr, column=i, value=v)
            c.font = f(9)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(i in wrap))
            if i > len(SHEET_COLS):
                c.fill = PatternFill("solid", fgColor=LIGHT)
        for i in money:
            ws.cell(row=rr, column=i).number_format = YEN
        ws.cell(row=rr, column=19).number_format = PCT
        ws.cell(row=rr, column=7).alignment = Alignment(horizontal="center")

    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "F4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(head))}{3 + len(rows)}"
    wb.save(path)


def cell(v, pct=False):
    if v is None:
        return ""
    if pct:
        return f"{v:.4f}"
    if isinstance(v, float):
        return f"{v:.0f}"
    return str(v)


def write_tsv(path, rows, lo, hi, pct_at=None):
    """rows の [lo:hi] だけを書き出す。見出しは付けない（そのまま行に貼るため）。"""
    with open(path, "w", encoding="utf-8") as fh:
        for row in rows:
            fh.write("\t".join(cell(v, pct=(i == pct_at))
                               for i, v in enumerate(row[lo:hi], start=lo)) + "\n")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    excluded = "--除外アカウント貸し" in sys.argv
    if not args:
        print(__doc__)
        sys.exit(1)
    src = args[0]
    base = args[1] if len(args) > 1 else "提案管理シート_移管用"

    recs, skipped = extract(src)
    if excluded:
        before = len(recs)
        recs = [r for r in recs if KAIGASHI not in r["サービス詳細"]]
        print(f"  アカウント貸しを除外: {before - len(recs)}件")

    rows = aggregate(recs)
    n = len(SHEET_COLS)
    write_xlsx(f"{base}.xlsx", rows, excluded)
    write_tsv(f"{base}_貼り付け用.tsv", rows, 0, n)
    write_tsv(f"{base}_参考列.tsv", rows, n, len(rows[0]), pct_at=n + 3)

    line = sum(1 for r in rows if r[14])
    agency = sum(1 for r in rows if r[8])
    kashi = sum(1 for r in rows if r[len(SHEET_COLS) + 8])
    print(f"書き出しました: {base}.xlsx / {base}_貼り付け用.tsv / {base}_参考列.tsv")
    print(f"  社数              : {len(rows):,}")
    print(f"  DYM売上 合計      : {sum(r[13] for r in rows):,.0f}")
    print(f"  LINE売上 合計     : {sum(r[14] for r in rows):,.0f}（{line}社）")
    print(f"  外注先あり        : {agency}社（代理店の頭越しに注意）")
    print(f"  アカウント貸しあり: {kashi}社")
    print(f"  担当①が空欄      : {sum(1 for r in rows if not r[11])}社")
    for sheet, miss in skipped:
        print(f"  ★スキップ {sheet}: {miss}")


if __name__ == "__main__":
    main()
