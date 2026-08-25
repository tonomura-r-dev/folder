# -*- coding: utf-8 -*-
"""Google スプレッドシートの新規タブに貼り付ける「提案管理集計」ブロックを書き出す。

  python _build/make_gsheet_block.py [出力先.tsv]

出力した TSV を新規タブの A1 に貼るだけで、AD と OA（LINE公式アカウント）を
まとめて集計するタブができる。元シートには一切書き込まない（参照するだけ）。

対象シートの列（2026-08-25 に実物で確認済み。ADとOAで並びは完全に同じ）
  B=ヨミか請求か（ADのみ／OAは記入日）  H=請求額（税抜）  I=原価合計  J=利益
  K=計上種別（ストック／ショット）      L=対象月          P=商材
  R,S,T,U,V  = アカウント / コンサル / 運用①②③
  AB〜AF     = 上記の按分売上
  AG〜AK     = 上記の按分利益
"""
import sys

# 参照する行数の上限。130シートあるブックなので列全体（100万行）を見に行くと重い。
# 行数がこれを超えたら、数式内の数字をまとめて増やすこと。
LIMIT = 2000

KEIS = ["ストック", "ショット"]

# 対象媒体。sheet はスプレッドシート上のタブ名そのまま。
# yomi=True のシートだけ B列で「ヨミ／請求」を判別できる。
MEDIA = [
    {"label": "AD", "sheet": "AD", "yomi": True,
     "cats": ["リスティング", "SNS", "DSP", "その他"]},
    {"label": "OA（商流のみ）", "sheet": "LINE公式アカウント", "yomi": False,
     "cats": ["LINE公式アカウント"]},
]

HEADERS = ["媒体", "商材カテゴリ", "契約形態", "件数", "売上", "粗利", "粗利率",
           "請求_売上", "請求_粗利率", "ヨミ_売上", "ヨミ_粗利率"]
NO_YOMI = "－"  # ヨミ／請求の区分を持たないシートの表示


def col(sheet, letter):
    """1シートの1列。行数を LIMIT で打ち切る。"""
    return f"'{sheet}'!${letter}$2:${letter}${LIMIT}"


def joined(fn):
    """全媒体ぶんを + でつないだ数式を返す。"""
    return "+".join(fn(m["sheet"]) for m in MEDIA)


def rows():
    r = []
    A = r.append

    A(["提案管理集計｜商材カテゴリ × 契約形態 × 売上 × 粗利 × 担当"])
    A(["対象：ADシート＋OA（LINE公式アカウント）シート／全対象月／金額は税抜／"
       "元シートは参照のみ・変更しません"])
    A([])

    # ---------------- ① 検算
    A(["① 検算", "各媒体の数字が元シートと合っているかの確認。差分が0なら以下の集計も信用できます"])
    A(["媒体", "件数", "売上", "原価", "粗利", "粗利率"])
    first_check = len(r) + 1
    for m in MEDIA:
        n = len(r) + 1
        s = m["sheet"]
        A([m["label"],
           f"=COUNT({col(s, 'H')})",
           f"=SUM({col(s, 'H')})",
           f"=SUM({col(s, 'I')})",
           f"=SUM({col(s, 'J')})",
           f'=IFERROR(E{n}/C{n},"")'])
    last_check = len(r)

    n = len(r) + 1
    A(["合計",
       f"=SUM(B{first_check}:B{last_check})",
       f"=SUM(C{first_check}:C{last_check})",
       f"=SUM(D{first_check}:D{last_check})",
       f"=SUM(E{first_check}:E{last_check})",
       f'=IFERROR(E{n}/C{n},"")'])
    total_check = n

    # 按分列にエラー値が混ざっていても止まらないよう IFERROR で包む
    n = len(r) + 1
    A(["担当按分の合計", "",
       "=" + joined(lambda s: f"SUMPRODUCT(IFERROR('{s}'!$AB$2:$AF${LIMIT},0))"),
       "",
       "=" + joined(lambda s: f"SUMPRODUCT(IFERROR('{s}'!$AG$2:$AK${LIMIT},0))"),
       ""])
    bunpai = n

    A(["差分（0ならOK）", "", f"=C{total_check}-C{bunpai}", "",
       f"=E{total_check}-E{bunpai}", ""])
    A([])

    # ---------------- ② 商材カテゴリ × 契約形態
    A(["② 商材カテゴリ × 契約形態",
       "「件数・売上・粗利」は請求＋ヨミの合計。右側がその内訳"])
    A(HEADERS)

    first = len(r) + 1
    for m in MEDIA:
        s, has_yomi = m["sheet"], m["yomi"]
        for cat in m["cats"]:
            for kei in KEIS:
                n = len(r) + 1
                base = f"{col(s, 'P')},$B{n},{col(s, 'K')},$C{n}"
                row = [m["label"], cat, kei,
                       f"=COUNTIFS({base})",
                       f"=SUMIFS({col(s, 'H')},{base})",
                       f"=SUMIFS({col(s, 'J')},{base})",
                       f'=IFERROR(F{n}/E{n},"")']
                if has_yomi:
                    b = col(s, "B")
                    row += [
                        f'=SUMIFS({col(s, "H")},{base},{b},"*請求*")',
                        f'=IFERROR(SUMIFS({col(s, "J")},{base},{b},"*請求*")/H{n},"")',
                        f'=SUMIFS({col(s, "H")},{base},{b},"*ヨミ*")',
                        f'=IFERROR(SUMIFS({col(s, "J")},{base},{b},"*ヨミ*")/J{n},"")',
                    ]
                else:
                    row += [NO_YOMI, NO_YOMI, NO_YOMI, NO_YOMI]
                A(row)
    last = len(r)

    n = len(r) + 1
    yomi_sheets = [m for m in MEDIA if m["yomi"]]
    q_uri = "+".join(f'SUMIFS({col(m["sheet"], "H")},{col(m["sheet"], "B")},"*請求*")'
                     for m in yomi_sheets)
    q_ri = "+".join(f'SUMIFS({col(m["sheet"], "J")},{col(m["sheet"], "B")},"*請求*")'
                    for m in yomi_sheets)
    y_uri = "+".join(f'SUMIFS({col(m["sheet"], "H")},{col(m["sheet"], "B")},"*ヨミ*")'
                     for m in yomi_sheets)
    y_ri = "+".join(f'SUMIFS({col(m["sheet"], "J")},{col(m["sheet"], "B")},"*ヨミ*")'
                    for m in yomi_sheets)
    A(["合計", "", "",
       f"=C{total_check}", f"=D{total_check}", f"=F{total_check}",
       f'=IFERROR(F{n}/E{n},"")',
       f"={q_uri}", f'=IFERROR(({q_ri})/H{n},"")',
       f"={y_uri}", f'=IFERROR(({y_ri})/J{n},"")'])
    A(["うち未分類（商材か計上種別が空欄）", "", "",
       f"=D{n}-SUM(D{first}:D{last})",
       f"=E{n}-SUM(E{first}:E{last})",
       f"=F{n}-SUM(F{first}:F{last})"])
    A([])

    # ---------------- ③ 担当別
    A(["③ 担当別（按分後・二重計上なし）",
       "AD・OAを合算。1案件を複数担当で分けた後の数字なので、合計は①の「担当按分の合計」と一致します"])
    A(["担当", "売上（按分）", "粗利（按分）", "粗利率"])
    A([let_formula()])
    A([])

    # ---------------- 注記
    A(["読み方・注意点"])
    for t in [
        "・商材カテゴリ＝元シートの「商材」列。契約形態＝「計上種別」列（ストック／ショット）。",
        "・売上＝「請求額（税抜）」、粗利＝「利益」。粗利率＝粗利÷売上。",
        "・対象月は絞っていません（全月合算）。月で見たいときは元シートの「対象月」で絞った表を別途作ります。",
        "・請求＝確定、ヨミ＝見込。粗利率が大きく違うので②の右側で内訳を出しています。",
        f"・OA（LINE公式アカウント）シートにはヨミ／請求の区分が無いため、該当欄は「{NO_YOMI}」としています。",
        "・③は按分後の数字です。1案件に複数担当がつくため、按分前で数えると売上が担当人数ぶん重複します。",
        "・このタブは元シートを参照しているだけです。AD・OAのシートには何も書き込んでいません。",
        f"・数式は各シートの2〜{LIMIT}行目を見ています。行数がこれを超えたら、数式内の{LIMIT}をまとめて増やしてください。",
        "・件数は「請求額（税抜）に数値が入っている行」を数えています。空欄の行は含みません。",
    ]:
        A([t])
    return r


def let_formula():
    """AD・OA を合算した担当別テーブルを1セルで出す。売上の多い順に並ぶ。"""
    roles = [("R", "AB", "AG"), ("S", "AC", "AH"), ("T", "AD", "AI"),
             ("U", "AE", "AJ"), ("V", "AF", "AK")]
    names = ",".join(col(m["sheet"], c)
                     for m in MEDIA for c, _, _ in roles)
    uri = "+".join(f"SUMIF({col(m['sheet'], c)},x,{col(m['sheet'], s)})"
                   for m in MEDIA for c, s, _ in roles)
    rieki = "+".join(f"SUMIF({col(m['sheet'], c)},x,{col(m['sheet'], p)})"
                     for m in MEDIA for c, _, p in roles)
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
        for row in data:
            fh.write("\t".join(str(c) for c in row) + "\n")
    print(f"書き出しました: {out}（{len(data)}行）")


if __name__ == "__main__":
    main()
