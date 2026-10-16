# -*- coding: utf-8 -*-
"""タナカホーム SIM ver1.0（殿村さんのPC版）→ ver1.1：CVの並びを「①資料請求 → ②来場予約」に入れ替える。
資料請求の後に来場予約をするので、①を資料請求にする（2026-10-16 殿村さん）。件数は各行のまま（①5〜16件・②2〜9件）。
比較シートの16行（CV数）を①〜④の合計に（SIM本体の40行と揃える）。チェックボックスを消さないようXMLを直接書き換える。

    python _build/patch_tanaka_sim_v11.py <ver1.0.xlsx> <ver1.1.xlsx>
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import set_f, set_s, strip_cache, q  # noqa: E402

src, out = Path(sys.argv[1]), Path(sys.argv[2])
work = Path(tempfile.mkdtemp())
with zipfile.ZipFile(src) as z:
    names = z.namelist()
    z.extractall(work)

wb = etree.parse(str(work / "xl/workbook.xml"))
rels = etree.parse(str(work / "xl/_rels/workbook.xml.rels"))
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
target = {r.get("Id"): r.get("Target") for r in rels.getroot()}
sheets = {s.get("name"): "xl/" + target[s.get(R)].lstrip("/").replace("xl/", "") for s in wb.getroot().iter(q("sheet"))}
sim_path, cmp_path = sheets["SIM_離脱防止＋サンクスLINE"], sheets["SIM_比較"]

sim = etree.parse(str(work / sim_path))
r = sim.getroot()
set_s(r, "C8", "資料請求（LINE経由）")
set_s(r, "C9", "来場予約（LINE経由）")
set_s(r, "N8", "サイトUU 13,228／友だち417人（2026-10-15〜16 殿村さん）。CVは①資料請求→②来場予約の順（資料請求の後に来場予約をする）")
set_s(r, "N9", "合計CV・CVR・CPA（40・41・48・49行）は①資料請求＋②来場予約の合計で計算")
set_s(r, "N36", "CV①資料請求（LINE経由）")
set_s(r, "N37", "CV②来場予約（LINE経由）")
strip_cache(r)
sim.write(str(work / sim_path), xml_declaration=True, encoding="UTF-8", standalone=True)

cmpx = etree.parse(str(work / cmp_path))
c = cmpx.getroot()
set_s(c, "A16", "CV数（+LINEOA施策）")
for col in "CDE":
    set_f(c, f"{col}16", f"SUM({col}17:{col}20)")
set_s(c, "A17", "CV①_資料請求（LINE経由）")
set_s(c, "A18", "CV②_来場予約（LINE経由）")
strip_cache(c)
cmpx.write(str(work / cmp_path), xml_declaration=True, encoding="UTF-8", standalone=True)

root = wb.getroot()
cp = root.find(q("calcPr"))
if cp is None:
    cp = etree.SubElement(root, q("calcPr"))
cp.set("fullCalcOnLoad", "1")
wb.write(str(work / "xl/workbook.xml"), xml_declaration=True, encoding="UTF-8", standalone=True)

with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.write(work / n, n)
shutil.rmtree(work)
print("saved:", out)
