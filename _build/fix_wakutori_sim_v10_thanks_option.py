# -*- coding: utf-8 -*-
"""ワクトリ SIM ver1.0：サンクスLINE誘導を両プランともオプションにする（2026-10-28 殿村さん）。SIM②は元から外れているので、SIM①だけ直す。

    python _build/fix_wakutori_sim_v10_thanks_option.py <SIM ver1.0.xlsx> [出力.xlsx]

  SIM①（sheet1）：E19〜J19＝FALSE（チェックを外す）／D67〜J67＝0／A4 の「＋サンクスLINE」を外す／N24・N67 の説明をSIM②と同じに。
  XMLを直接書き換える（チェックボックスを消さない）。出力を省くと上書き。
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import get_cell, set_b, set_n, set_s, strip_cache  # noqa: E402


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    p = work / "xl/worksheets/sheet1.xml"
    tree = etree.parse(str(p))
    root = tree.getroot()
    for c in "EFGHIJ":
        set_b(root, f"{c}19", False)
    for c in "DEFGHIJ":
        set_n(root, f"{c}67", 0)
    set_s(root, "A4", "株式会社ワクトリ御中_LINEOA施策提案【SIM】プラン①（月3投稿）：ポップアップ")
    set_s(root, "N24", "サンクスLINE誘導はオプション（このプランのSIMには入れていない）")
    set_s(root, "N67", "サンクスLINE誘導：オプション（初期15万円・月3万円〜）")
    strip_cache(root)
    tree.write(str(p), xml_declaration=True, encoding="UTF-8", standalone=True)
    q = work / "xl/worksheets/sheet2.xml"                      # 比較シートの古い計算結果も消す
    t = etree.parse(str(q))
    strip_cache(t.getroot())
    t.write(str(q), xml_declaration=True, encoding="UTF-8", standalone=True)
    wb = work / "xl/workbook.xml"
    s = wb.read_text(encoding="utf-8")
    if "fullCalcOnLoad" not in s:
        s = s.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.write_text(s, encoding="utf-8")
    tmp = Path(tempfile.mkdtemp()) / "o.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(work / n, n)
    shutil.move(str(tmp), out)
    shutil.rmtree(work)
    with zipfile.ZipFile(out) as z:
        assert "xl/featurePropertyBag/featurePropertyBag.xml" in z.namelist(), "チェックボックスが消えた"
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1])
