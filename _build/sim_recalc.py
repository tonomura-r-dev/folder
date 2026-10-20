# -*- coding: utf-8 -*-
"""SIM（xlsx）を検算用に全部計算し直したコピーを作る。

  ExcelやXML直接編集で保存したxlsxは、数式セルに前回の計算結果（キャッシュ値）が残っている。
  LibreOfficeで変換しても、そのキャッシュ値はそのまま使われ、計算し直されない（2026-10-20 すなおSIMで確認：
  45行の式を直しても、51行の総計は古い値 401,000 のままだった）。
  → コピーの数式セルからキャッシュ値と calcChain を消してから変換し、全部をLibreOfficeに計算させる。

    from _build.sim_recalc import recalc_copy
    path = recalc_copy(src, workdir)   # 計算し直した xlsx のパス（読むだけ。納品には使わない）
"""
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def recalc_copy(src, workdir=None):
    work = Path(workdir or tempfile.mkdtemp())
    ext = work / "x"
    with zipfile.ZipFile(src) as z:
        names = [n for n in z.namelist() if n != "xl/calcChain.xml"]
        z.extractall(ext)
    for p in (ext / "xl/worksheets").glob("sheet*.xml"):
        t = p.read_text(encoding="utf-8")
        p.write_text(re.sub(r"(<f[^>]*>[^<]*</f>|<f[^>]*/>)<v>[^<]*</v>", r"\1", t), encoding="utf-8")
    for rel, pat in (("[Content_Types].xml", r"<Override[^>]*calcChain[^>]*/>"),
                     ("xl/_rels/workbook.xml.rels", r"<Relationship[^>]*calcChain[^>]*/>")):
        p = ext / rel
        p.write_text(re.sub(pat, "", p.read_text(encoding="utf-8")), encoding="utf-8")
    tmp = work / "nocache.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(ext / n, n)
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(work / "c"), str(tmp)],
                   check=True, capture_output=True, timeout=180)
    return work / "c" / "nocache.xlsx"


if __name__ == "__main__":
    print(recalc_copy(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None))
