# -*- coding: utf-8 -*-
"""すなおコンサルティング SIM ver1.2：LPのUUを仮に3,000に置いた版（2026-10-19 殿村さん「大体はわかるやん」）。

    python _build/build_sunao_sim_v12.py <SIM ver1.1.xlsx> [出力.xlsx]

  ver1.1をベースに、XMLを直接書き換える（チェックボックスを消さない）。
  ・イベント月（1・4か月目）：LP UU 仮に3,000（広告クリック 約2,760＋自社のSNS等が少し。UTAGEのLPは検索から来ない）
    セミナー予約 81件（前回1回分の実績）
  ・平常期：少量の広告が続いている（殿村さん）→ LP UU 仮に月130（広告 月10万円÷8月のクリック単価765円）。サンクスLINEは0
  ・係数は定義テキスト：ポップアップ＝UU×表示65%×12.5%×12.5%／サンクスLINE＝予約×遷移80%×追加45%
  ・成果単価：セミナー予約 1.2万円／サロン入会 2.5万円（殿村さん確定）
  ・CV（目安。細かい件数は殿村さんが決める）：セミナー予約 1・0・0・1・0・0／サロン入会 0（クリックの0.3〜1.0%前後）
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import set_f, set_n, set_s  # noqa: E402

COLS = "EFGHIJ"
EVENT = "EH"                                                   # イベント月（1・4か月目）
UU_EVENT, UU_NORMAL, APPLY = 3000, 130, 81
CV_SEMINAR = [1, 0, 0, 1, 0, 0]


def fill(root):
    set_s(root, "A4", "株式会社すなおコンサルティング御中_LINEOA施策提案【SIM】ver1.2（ポップアップツール＋サンクスLINE誘導）")
    set_n(root, "D8", 12000)
    set_n(root, "D9", 25000)
    set_n(root, "H8", UU_EVENT)
    set_n(root, "H9", UU_EVENT)
    set_n(root, "H10", APPLY)
    for i, c in enumerate(COLS):
        if c not in EVENT:                                     # 平常期も少量の広告分のUU
            set_n(root, f"{c}14", UU_NORMAL)
            set_n(root, f"{c}15", UU_NORMAL)
        set_f(root, f"{c}23", f"IF({c}18=FALSE,0,{c}15*$P$18*0.125*0.125+({c}14-{c}15)*$P$19*0.12*0.1)")
        set_f(root, f"{c}24", f"IF({c}19=FALSE,0,{c}16*0.8*0.45)")
        set_n(root, f"{c}36", CV_SEMINAR[i])
        set_n(root, f"{c}37", 0)
    notes = {
        "N8": "CV①セミナー予約（成果単価 1.2万円）。LINEのリンク経由だけ。サンクスLINEで友だちになった方（予約済み）は数えない",
        "N9": "CV②サロン入会（成果単価 2.5万円）。LINEのリンク経由だけ",
        "N10": f"イベント月（1・4か月目）：LP UU 仮に{UU_EVENT:,}（広告クリック 約2,760＋自社のSNS等が少し）・セミナー予約 {APPLY}件（前回1回分の実績）。広告費はADの予算なので入れていない",
        "N14": f"平常期も少量の広告が続いている → LP UU 仮に月{UU_NORMAL}（広告 月10万円÷クリック単価765円）",
        "N23": "ポップアップツール＝LP UU×表示65%×クリック12.5%×追加12.5%（定義テキストの係数）",
        "N24": "サンクスLINE誘導＝セミナー予約×遷移80%×追加45%（定義テキストの係数）。平常期は予約が無いので0",
        "N26": "友だちの入口はポップアップツール＋サンクスLINE誘導だけ（資料の「友だちの入口」と同じ）",
        "N36": "セミナー予約＝クリックの0.3〜1.0%前後（目安。細かい件数は殿村さん）",
        "N37": "サロン入会＝目安では0（クリックが少ないため）。件数は殿村さん",
        "N67": "サンクスLINE誘導：初期15万・月3万〜",
    }
    for ref, text in notes.items():
        set_s(root, ref, text)


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    sheet = work / "xl/worksheets/sheet1.xml"                  # SIM_ポップアップ＋サンクスLINE
    tree = etree.parse(str(sheet))
    fill(tree.getroot())
    tree.write(str(sheet), xml_declaration=True, encoding="UTF-8", standalone=True)
    wb = work / "xl/workbook.xml"
    t = wb.read_text(encoding="utf-8")
    if "fullCalcOnLoad" not in t:
        t = t.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.write_text(t, encoding="utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(work / n, n)
    shutil.rmtree(work)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
