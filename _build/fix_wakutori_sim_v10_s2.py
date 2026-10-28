# -*- coding: utf-8 -*-
"""ワクトリ SIM ver1.0（殿村さんのPC保存版）：SIM②_月5投稿にも、殿村さんがSIM①に入れた直しを同じように入れる（2026-10-28 殿村さん）。

    python _build/fix_wakutori_sim_v10_s2.py <SIM ver1.0（PC保存版）.xlsx> <出力.xlsx>

  SIM①でやった直し：成果単価 D8＝0（成果報酬型は外す）／ブロック率 D27・E27＝0.341／応募 E36〜J36 の CVR を 0.4〜0.9%。
  SIM②（sheet3）に同じ値を入れる。LINE運用コンサル（69行）はSIM②のまま（金額は殿村さんが決める）。
  XMLを直接書き換える（チェックボックスを消さない）。
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import get_cell, set_f, set_n, strip_cache  # noqa: E402

CVR = dict(zip("EFGHIJ", ["0.004", "0.0045", "0.0055", "0.0065", "0.0075", "0.009"]))   # SIM①（殿村さんの値）と同じ


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    s1 = etree.parse(str(work / "xl/worksheets/sheet1.xml")).getroot()
    assert get_cell(s1, "E36").find("{*}f").text == "ROUND(E35*0.004,0)", "SIM①の値が想定と違う"
    p = work / "xl/worksheets/sheet3.xml"
    tree = etree.parse(str(p))
    root = tree.getroot()
    set_n(root, "D8", 0)
    set_n(root, "D27", 0.341)
    set_n(root, "E27", 0.341)
    for c, r in CVR.items():
        set_f(root, f"{c}36", f"ROUND({c}35*{r},0)")
    strip_cache(root)
    tree.write(str(p), xml_declaration=True, encoding="UTF-8", standalone=True)
    for n in ("xl/worksheets/sheet4.xml",):                 # 比較シートの古い計算結果も消す
        q = work / n
        t = etree.parse(str(q))
        strip_cache(t.getroot())
        t.write(str(q), xml_declaration=True, encoding="UTF-8", standalone=True)
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
