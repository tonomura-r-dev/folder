# -*- coding: utf-8 -*-
"""pptx を画像にする（LibreOffice。並行して動かせるよう、毎回別のプロファイルを使う）。

    python -m _build.arauchi_v3.render <pptx> <出力フォルダ> [--pages 3,5,7]

出力は <出力フォルダ>/p-NN.jpg（NN＝出来上がりのページ番号）。
※メイリオが無いので代替フォント。代替フォントはメイリオより少し幅が広い（ここで収まればPowerPointでも収まる側）。
"""
import argparse
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("outdir")
    ap.add_argument("--pages", default="")
    ap.add_argument("--dpi", type=int, default=110)
    a = ap.parse_args()
    out = Path(a.outdir).resolve()
    out.mkdir(parents=True, exist_ok=True)
    prof = tempfile.mkdtemp(prefix="lo_prof_")
    work = tempfile.mkdtemp(prefix="lo_work_")
    try:
        src = Path(work) / "deck.pptx"
        shutil.copy(a.pptx, src)
        subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless",
                        "--convert-to", "pdf", "--outdir", work, str(src)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
        import fitz
        doc = fitz.open(Path(work) / "deck.pdf")
        pages = [int(x) for x in a.pages.split(",") if x.strip()] or list(range(1, doc.page_count + 1))
        z = a.dpi / 72.0
        for n in pages:
            doc[n - 1].get_pixmap(matrix=fitz.Matrix(z, z)).save(str(out / f"p-{n:02d}.jpg"))
        print(f"{len(pages)} pages → {out}  (total {doc.page_count})")
    finally:
        shutil.rmtree(prof, ignore_errors=True)
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
