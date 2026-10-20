# -*- coding: utf-8 -*-
"""横浜シミズ SIM：離脱防止ポップアップの費用を入れる（2026-10-20 殿村さん「費用の修正を」）。

    python _build/patch_yokohama_shimizu_sim_cost.py <殿村さんのSIM.xlsx> [出力.xlsx]

  殿村さんのPC保存版（ver2.0で離脱防止をON・ステップ配信4通にした版）をベースに、XMLを直接書き換える（チェックボックスを消さない）。
  ・66行（離脱防止ポップアップ）：初期 D66＝15,000／月額 E66〜J66＝30,000
  ・古くなったメモ（A4・N26・N32・N66）を今の中身に合わせる
  CV数は殿村さんが決めるので触らない。出力を省くと上書き。
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import set_n, set_s, strip_cache  # noqa: E402


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    wb = (work / "xl/workbook.xml").read_text(encoding="utf-8")
    assert '<sheet name="SIM" sheetId="1" r:id="rId1"/>' in wb
    sheet = work / "xl/worksheets/sheet1.xml"
    tree = etree.parse(str(sheet))
    root = tree.getroot()
    set_n(root, "D66", 15000)
    for c in "EFGHIJ":
        set_n(root, f"{c}66", 30000)
    set_s(root, "N66", "離脱防止ポップアップ：初期1.5万円・月3万円")
    set_s(root, "A4", "株式会社横浜シミズ御中_LINEOA施策提案【SIM】（サイトのLINEボタン＋離脱防止＋サンクスLINE誘導）")
    set_s(root, "N26", "サイトのLINEボタン（UUの0.5%・FMTの係数）＋離脱防止ポップアップ＋サンクスLINE誘導")
    set_s(root, "N32", "ステップ配信 4通（あいさつメッセージの後、友だち追加から14日以内）")
    strip_cache(root)
    tree.write(str(sheet), xml_declaration=True, encoding="UTF-8", standalone=True)
    if "fullCalcOnLoad" not in wb:
        (work / "xl/workbook.xml").write_text(wb.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1), encoding="utf-8")
    tmp = Path(tempfile.mkdtemp()) / "out.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:                                      # 元の順番のまま詰め直す
            z.write(work / n, n)
    shutil.move(str(tmp), out)
    shutil.rmtree(work)
    with zipfile.ZipFile(out) as z:
        assert "xl/featurePropertyBag/featurePropertyBag.xml" in z.namelist(), "チェックボックスが消えた"
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1])
