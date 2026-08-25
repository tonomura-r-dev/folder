# -*- coding: utf-8 -*-
"""Google スプレッドシートの新規タブに貼り付ける「全媒体 案件集計」ブロックを書き出す。

  python _build/make_gsheet_block.py [出力先.tsv]

出力した TSV を新規タブの A1 に貼るだけで、全13媒体シートを参照する集計タブができる。
元シートには一切書き込まない（参照するだけ）。

佐村さんの指示（2026-08-25）
  ・商材カテゴリ × 契約形態 × 売上 × 粗利 が1枚で分かること
  ・AD・OA だけでなく全媒体
  ・担当別の数字は不要。案件全体の数字だけでよい

列の並びは媒体ごとに違う（メディア・ベトナム・マス広告だけズレている）ので、
シートごとに列を指定している。2026-08-25 に実データで確認済み。
"""
import sys

# 参照する行数の上限。130シートあるブックなので列全体（100万行）を見に行くと重い。
# 行数がこれを超えたら、数式内の数字をまとめて増やすこと。
LIMIT = 2000

KEIS = ["ストック", "ショット"]

# sheet: タブ名 / uri: 請求額（税抜）/ genka: 原価合計 / rieki: 利益
# kei: 計上種別 / cat: 商材 / cats: その媒体に実在する商材
MEDIA = [
    {"sheet": "AD", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["リスティング", "SNS", "DSP", "その他"]},
    {"sheet": "アフィ", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["アフィリエイト"]},
    {"sheet": "CS", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["SEO固定", "SEO成果"]},
    {"sheet": "MEO", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["SEO固定"]},
    {"sheet": "制作", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["F-code", "Spider", "保守管理", "サイト修正", "バナー"]},
    {"sheet": "PR", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["PR", "タレントシェア"]},
    {"sheet": "風評", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["風評被害対策", "排他的RFO", "排他的SEO", "RFO", "排他的RFO(G)",
              "風評監視ツール", "2ch"]},
    {"sheet": "タレントシェア", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["タレントシェア"]},
    {"sheet": "LINE公式アカウント", "uri": "H", "genka": "I", "rieki": "J", "kei": "K",
     "cat": "P", "cats": ["LINE公式アカウント"]},
    {"sheet": "メディア", "uri": "K", "genka": "L", "rieki": "M", "kei": "N", "cat": "F",
     "cats": ["SEO固定掲載費", "紹介手数料", "AD成果報酬", "SEO成果報酬", "成約手数料",
              "ファクタリング", "リフォーム"]},
    {"sheet": "ベトナム", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "N",
     "cats": ["リスティング", "SNS", "BPO", "CS", "PR"]},
    {"sheet": "ASP", "uri": "H", "genka": "I", "rieki": "J", "kei": "K", "cat": "P",
     "cats": ["ASP"]},
    {"sheet": "マス広告(MA)・その他", "uri": "H", "genka": "I", "rieki": "J", "kei": "K",
     "cat": "N", "cats": ["CM", "アドトラック"]},
]


def col(m, key):
    """媒体 m の key 列。行数を LIMIT で打ち切る。"""
    letter = m[key]
    return f"'{m['sheet']}'!${letter}$2:${letter}${LIMIT}"


def rows():
    r = []
    A = r.append

    A(["全媒体 案件集計｜商材カテゴリ × 契約形態 × 売上 × 粗利"])
    A(["対象：全13媒体シート／全対象月／金額は税抜／元シートは参照のみ・変更しません"])
    A([])

    # ---------------- ① 媒体別サマリ 兼 検算
    A(["① 媒体別サマリ（検算兼用）",
       "各行がそのままの媒体シートの合計です。元シートと突き合わせて確認できます"])
    A(["媒体", "件数", "売上", "原価", "粗利", "粗利率"])
    first_m = len(r) + 1
    for m in MEDIA:
        n = len(r) + 1
        A([m["sheet"],
           f"=COUNT({col(m, 'uri')})",
           f"=SUM({col(m, 'uri')})",
           f"=SUM({col(m, 'genka')})",
           f"=SUM({col(m, 'rieki')})",
           f'=IFERROR(E{n}/C{n},"")'])
    last_m = len(r)

    n = len(r) + 1
    A(["合計",
       f"=SUM(B{first_m}:B{last_m})",
       f"=SUM(C{first_m}:C{last_m})",
       f"=SUM(D{first_m}:D{last_m})",
       f"=SUM(E{first_m}:E{last_m})",
       f'=IFERROR(E{n}/C{n},"")'])
    grand = n
    A([])

    # ---------------- ② 契約形態のまとめ
    # 13媒体ぶんの SUMIFS をつなぐと1セル3000文字を超えて扱いづらいので、
    # ③の明細を集計する形にする（結果は同じで、数式は50分の1の長さになる）。
    d_first = len(r) + 10  # ②(6行) + 空行 + ③見出し(2行) を足した先が明細の先頭
    d_last = d_first + sum(len(m["cats"]) for m in MEDIA) * len(KEIS) - 1
    dc = f"$C${d_first}:$C${d_last}"  # ③の契約形態列

    A(["② 契約形態別（全媒体）", "ストックとショットで粗利率がどう違うかを見るところ"])
    A(["契約形態", "件数", "売上", "粗利", "粗利率"])
    first_k = len(r) + 1
    for kei in KEIS:
        n = len(r) + 1
        A([kei,
           f"=SUMIF({dc},$A{n},$D${d_first}:$D${d_last})",
           f"=SUMIF({dc},$A{n},$E${d_first}:$E${d_last})",
           f"=SUMIF({dc},$A{n},$F${d_first}:$F${d_last})",
           f'=IFERROR(D{n}/C{n},"")'])
    last_k = len(r)
    n = len(r) + 1
    A(["合計", f"=B{grand}", f"=C{grand}", f"=E{grand}", f'=IFERROR(D{n}/C{n},"")'])
    A(["うち計上種別が空欄", f"=B{n}-SUM(B{first_k}:B{last_k})",
       f"=C{n}-SUM(C{first_k}:C{last_k})", f"=D{n}-SUM(D{first_k}:D{last_k})"])
    A([])

    # ---------------- ③ 媒体 × 商材カテゴリ × 契約形態
    A(["③ 媒体 × 商材カテゴリ × 契約形態", "これが本体。オートフィルタやピボットの元データにも使えます"])
    A(["媒体", "商材カテゴリ", "契約形態", "件数", "売上", "粗利", "粗利率"])
    first_d = len(r) + 1
    for m in MEDIA:
        for cat in m["cats"]:
            for kei in KEIS:
                n = len(r) + 1
                base = f"{col(m, 'cat')},$B{n},{col(m, 'kei')},$C{n}"
                A([m["sheet"], cat, kei,
                   f"=COUNTIFS({base})",
                   f"=SUMIFS({col(m, 'uri')},{base})",
                   f"=SUMIFS({col(m, 'rieki')},{base})",
                   f'=IFERROR(F{n}/E{n},"")'])
    last_d = len(r)

    n = len(r) + 1
    A(["合計", "", "", f"=B{grand}", f"=C{grand}", f"=E{grand}",
       f'=IFERROR(F{n}/E{n},"")'])
    A(["うち未分類（商材か計上種別が空欄）", "", "",
       f"=D{n}-SUM(D{first_d}:D{last_d})",
       f"=E{n}-SUM(E{first_d}:E{last_d})",
       f"=F{n}-SUM(F{first_d}:F{last_d})"])
    A([])

    # ---------------- 注記
    A(["読み方・注意点"])
    for t in [
        "・商材カテゴリ＝各媒体シートの「商材」列。契約形態＝「計上種別」列（ストック／ショット）。",
        "・売上＝「請求額（税抜）」、粗利＝「利益」。粗利率＝粗利÷売上。",
        "・①の各行は、その媒体シートをそのまま合計したものです。元シートと数字が合うか確認できます。",
        "・③の「うち未分類」は、商材または計上種別が空欄の案件です。0でなければ元シートの入力漏れです。",
        "・③に商材が新しく増えた場合、その行は表に出ず「うち未分類」に入ります。気づいたら行を足してください。",
        "・対象月は絞っていません（全月合算）。月で見たいときは元シートの「対象月」で絞った表を別途作ります。",
        "・ヨミ（見込）と請求（確定）はADシートのみ区分があります。全媒体では区別していません。",
        "・このタブは元シートを参照しているだけです。各媒体シートには何も書き込んでいません。",
        f"・数式は各シートの2〜{LIMIT}行目を見ています。行数がこれを超えたら、数式内の{LIMIT}をまとめて増やしてください。",
        "・件数は「請求額（税抜）に数値が入っている行」を数えています。空欄の行は含みません。",
    ]:
        A([t])
    return r


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "_build/zen_baitai_shukei.tsv"
    data = rows()
    with open(out, "w", encoding="utf-8") as fh:
        for row in data:
            fh.write("\t".join(str(c) for c in row) + "\n")
    print(f"書き出しました: {out}（{len(data)}行）")


if __name__ == "__main__":
    main()
