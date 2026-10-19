# -*- coding: utf-8 -*-
"""清水グループ SIM：ブロック率（27行目）の式と表記を直す（2026-10-19 殿村さん「治ったけど、表記ばぐった。直してもらえる？」）。

    python _build/patch_shimizu_sim_block.py <殿村さんのSIM.xlsx> [出力.xlsx]

  殿村さんのPC保存版をベースに、XMLを直接書き換える（チェックボックスを消さない）。
  ・E27（1か月目）は殿村さんの値のまま
  ・F27〜J27＝前の月×1.025（=+E$27*1.025 の形。殿村さんの書き方）。G27〜J27に残っていた古い数字も式にする
  ・F27だけ「0%」の表記になっていたので、ほかの月と同じ「0.00%」にそろえる（E27と同じスタイル）
  ・N27のメモを今の作り方に合わせる。開いたときにExcelが計算し直すよう fullCalcOnLoad を付ける
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import get_cell, set_f, set_s  # noqa: E402

COLS = "EFGHIJ"


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    sheet = work / "xl/worksheets/sheet1.xml"
    tree = etree.parse(str(sheet))
    root = tree.getroot()
    style = get_cell(root, "E27").get("s")                   # 1か月目と同じ「0.00%」のスタイル
    for prev, c in zip(COLS, COLS[1:]):
        set_f(root, f"{c}27", f"+{prev}$27*1.025")
        get_cell(root, f"{c}27").set("s", style)
    set_s(root, "N27", "ブロック率：1か月目（E27）から、2か月目以降は前の月×1.025（式）")
    tree.write(str(sheet), xml_declaration=True, encoding="UTF-8", standalone=True)
    wb = work / "xl/workbook.xml"
    t = wb.read_text(encoding="utf-8")
    if "fullCalcOnLoad" not in t:
        t = t.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.write_text(t, encoding="utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:                                      # 元の順番のまま詰め直す
            z.write(work / n, n)
    shutil.rmtree(work)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1])
