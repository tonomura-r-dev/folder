# -*- coding: utf-8 -*-
"""Google スプレッドシートの新規タブに貼り付ける「AD集計」ブロックを書き出す。

  python _build/make_gsheet_block.py [出力先.tsv]

出力した TSV を新規タブの A1 に貼るだけで、AD シートを参照する集計タブができる。
AD シート側には一切書き込まないので、元シートは変更されない。

AD シートの列（2026-08 時点で確認済み）
  B=ヨミか請求か  H=請求額（税抜）  I=原価合計  J=利益
  K=計上種別      L=対象月          P=商材
  R,S,T,U,V       = アカウント / コンサル / 運用①②③
  AB〜AF          = 上記の按分売上
  AG〜AK          = 上記の按分利益
"""
import sys

SRC = "'AD'"
CATS = ["リスティング", "SNS", "DSP", "その他"]
KEIS = ["ストック", "ショット"]
# 参照する行数の上限。本番ブックは130シートあるので列全体（100万行）を見に行くと重い。
# ADの行数がこれを超えたら、数式内の数字をまとめて増やすこと。
LIMIT = 2000


def col(letter):
    """ADシートの1列。行数を LIMIT で打ち切る。"""
    return f"{SRC}!${letter}$2:${letter}${LIMIT}"


def rows():
    r = []
    A = r.append

    A(["AD集計｜商材カテゴリ × 契約形態 × 売上 × 粗利 × 担当"])
    A([f"出典：ADシート（全対象月）／金額は税抜／ADシートは参照のみ・変更しません"])
    A([])

    # ---------------- ① 検算
    A(["① 検算", "差分が0ならADシートと一致＝以下の集計も信用できます"])
    A(["項目", "件数", "売上", "原価", "粗利", "粗利率"])
    # 件数は COUNTA だと「空に見えるが数式が入っている」セルまで数えてしまうので、
    # 請求額が数値で入っている行だけを COUNT で数える。
    A(["ADシート全体",
       f"=COUNT({col('H')})",
       f"=SUM({col('H')})",
       f"=SUM({col('I')})",
       f"=SUM({col('J')})",
       '=IFERROR(E6/C6,"")'])
    # 按分列にエラー値が混ざっていても止まらないよう IFERROR で包む
    A(["担当按分の合計", "",
       f"=SUMPRODUCT(IFERROR({SRC}!$AB$2:$AF${LIMIT},0))", "",
       f"=SUMPRODUCT(IFERROR({SRC}!$AG$2:$AK${LIMIT},0))",
       '=IFERROR(E7/C7,"")'])
    A(["差分（0ならOK）", "", "=C6-C7", "", "=E6-E7", ""])
    A([])

    # ---------------- ② 商材カテゴリ × 契約形態
    A(["② 商材カテゴリ × 契約形態", "「件数・売上・粗利」は請求＋ヨミの合計。右側が内訳"])
    A(["商材カテゴリ", "契約形態", "件数", "売上", "粗利", "粗利率",
       "請求_売上", "請求_粗利率", "ヨミ_売上", "ヨミ_粗利率"])

    first = len(r) + 1  # 1始まりの行番号
    for cat in CATS:
        for kei in KEIS:
            n = len(r) + 1
            base = f"{col('P')},$A{n},{col('K')},$B{n}"
            A([cat, kei,
               f"=COUNTIFS({base})",
               f"=SUMIFS({col('H')},{base})",
               f"=SUMIFS({col('J')},{base})",
               f'=IFERROR(E{n}/D{n},"")',
               f'=SUMIFS({col("H")},{base},{col("B")},"*請求*")',
               f'=IFERROR(SUMIFS({col("J")},{base},{col("B")},"*請求*")/G{n},"")',
               f'=SUMIFS({col("H")},{base},{col("B")},"*ヨミ*")',
               f'=IFERROR(SUMIFS({col("J")},{base},{col("B")},"*ヨミ*")/I{n},"")'])
    last = len(r)

    n = len(r) + 1
    A(["合計", "",
       f"=COUNT({col('H')})",
       f"=SUM({col('H')})",
       f"=SUM({col('J')})",
       f'=IFERROR(E{n}/D{n},"")',
       f'=SUMIFS({col("H")},{col("B")},"*請求*")',
       f'=IFERROR(SUMIFS({col("J")},{col("B")},"*請求*")/G{n},"")',
       f'=SUMIFS({col("H")},{col("B")},"*ヨミ*")',
       f'=IFERROR(SUMIFS({col("J")},{col("B")},"*ヨミ*")/I{n},"")'])
    A(["うち未分類（商材か計上種別が空欄）",
       "", f"=C{n}-SUM(C{first}:C{last})", f"=D{n}-SUM(D{first}:D{last})",
       f"=E{n}-SUM(E{first}:E{last})"])
    A([])

    # ---------------- ③ 担当別
    A(["③ 担当別（按分後・二重計上なし）",
       "1案件を複数担当で分けた後の数字。合計は①の「担当按分の合計」と一致します"])
    A(["担当", "売上（按分）", "粗利（按分）", "粗利率"])
    A([let_formula()])
    A([])

    # ---------------- 注記
    A(["読み方・注意点"])
    for t in [
        "・商材カテゴリ＝ADシートの「商材」列。契約形態＝「計上種別」列（ストック／ショット）。",
        "・売上＝「請求額（税抜）」、粗利＝「利益」。粗利率＝粗利÷売上。",
        "・対象月は絞っていません（全月合算）。月で見たいときはADシートの「対象月」で絞った表を別途作ります。",
        "・請求＝確定、ヨミ＝見込。粗利率が大きく違うので②の右側で内訳を出しています。",
        "・③は按分後の数字です。1案件に複数担当がつくため、按分前で数えると売上が担当人数ぶん重複します。",
        "・このタブはADシートを参照しているだけです。ADシートには何も書き込んでいません。",
        f"・数式はADシートの2〜{LIMIT}行目を見ています。行数がこれを超えたら、数式内の{LIMIT}をまとめて増やしてください。",
        "・件数は「請求額（税抜）に数値が入っている行」を数えています。空欄の行は含みません。",
    ]:
        A([t])
    return r


def let_formula():
    """担当別テーブルを1セルで出す。売上の多い順に並ぶ。"""
    roles = [("R", "AB", "AG"), ("S", "AC", "AH"), ("T", "AD", "AI"),
             ("U", "AE", "AJ"), ("V", "AF", "AK")]
    names = ",".join(col(c) for c, _, _ in roles)
    uri = "+".join(f"SUMIF({col(c)},x,{col(s)})" for c, s, _ in roles)
    rieki = "+".join(f"SUMIF({col(c)},x,{col(p)})" for c, _, p in roles)
    return (
        "=LET("
        f"n,SORT(UNIQUE(TOCOL({{{names}}},1))),"
        f"u,MAP(n,LAMBDA(x,{uri})),"
        f"r,MAP(n,LAMBDA(x,{rieki})),"
        'SORT({n,u,r,MAP(u,r,LAMBDA(a,b,IFERROR(b/a,"")))},2,FALSE))'
    )


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "_build/ad_summary_gsheet.tsv"
    data = rows()
    with open(out, "w", encoding="utf-8") as fh:
        for r in data:
            fh.write("\t".join(str(c) for c in r) + "\n")
    print(f"書き出しました: {out}（{len(data)}行）")
    print("\n--- ③担当別の1セル数式（単独で貼る場合はこれ）---")
    print(let_formula())


if __name__ == "__main__":
    main()
