# -*- coding: utf-8 -*-
"""すなおコンサルティング SIM ver1.4：ステップ配信を5通 → 8通に（2026-10-27 殿村さん「直しましょう」）。

    python _build/fix_sunao_sim_v14.py <SIM ver1.3.xlsx> <出力.xlsx>

  資料ver7.3で、友だち追加から15日を2日おきに8通（0・2・4・6・8・10・12・14日目）にしたので、32行（ステップ配信本数）を8に。
  ほかの数字（友だちの入口・CVの件数・費用）は ver1.3 のまま。XMLを直接書き換える（チェックボックスを消さない）。
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import N, get_cell, set_n, set_s, strip_cache  # noqa: E402

COLS = "EFGHIJ"
STEPS = 8


def fill_sim(root):
    set_s(root, "A4", "株式会社すなおコンサルティング御中_LINEOA施策提案【SIM】ver1.4（ポップアップツール＋サンクスLINE誘導）")
    for c in COLS:
        set_n(root, f"{c}32", STEPS)
    set_s(root, "N32", "ステップ配信8通（友だち追加から15日、2日おき。その月の新しい友だち。比べる→確かめる→始める→予約の案内）")


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    p = work / "xl/worksheets/sheet1.xml"
    tree = etree.parse(str(p))
    root = tree.getroot()
    assert get_cell(root, "E32").find("m:v", N).text == "5", "32行が5通ではない"
    fill_sim(root)
    strip_cache(root)
    tree.write(str(p), xml_declaration=True, encoding="UTF-8", standalone=True)
    wb = work / "xl/workbook.xml"
    t = wb.read_text(encoding="utf-8")
    if "fullCalcOnLoad" not in t:
        t = t.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.write_text(t, encoding="utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(work / n, n)
    shutil.rmtree(work)
    with zipfile.ZipFile(out) as z:
        assert "xl/featurePropertyBag/featurePropertyBag.xml" in z.namelist(), "チェックボックスが消えた"
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
