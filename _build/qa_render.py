# -*- coding: utf-8 -*-
"""pptxを1枚ずつPNG化してQAする。
Windows: PowerPoint COM（実物どおりの見た目）
Linux（クラウド）: LibreOffice経由（フォント代替で見た目は近似）

使い方:
  python _build/qa_render.py <pptxのパス> [出力先ディレクトリ]
"""
import shutil
import subprocess
import sys
from pathlib import Path


def render_windows(pptx: Path, outdir: Path) -> int:
    ps = f'''
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open("{pptx}", $true, $false, $false)
$n = $pres.Slides.Count
foreach ($i in 1..$n) {{
  $name = "{outdir}\\slide-{{0:D2}}.png" -f $i
  $pres.Slides.Item($i).Export($name, "PNG", 1200, 831)
}}
$pres.Close()
$pp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pp) | Out-Null
Write-Output $n
'''
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stderr, file=sys.stderr)
        return 0
    return len(list(outdir.glob("slide-*.png")))


def render_libreoffice(pptx: Path, outdir: Path) -> int:
    """LibreOffice → PDF → PNG化。
    pdftoppm があればそれを使い、無ければ PyMuPDF にフォールバックする
    （クラウドコンテナには poppler が入っていないことがある）。"""
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", str(outdir), str(pptx)], check=True)
    pdf = outdir / (pptx.stem + ".pdf")
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-jpeg", "-r", "110", str(pdf),
                        str(outdir / "slide")], check=True)
        return len(list(outdir.glob("slide-*.jpg")))

    import fitz  # PyMuPDF
    doc = fitz.open(pdf)
    zoom = 110 / 72.0
    for i, page in enumerate(doc, 1):
        page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(
            str(outdir / f"slide-{i:02d}.png"))
    n = doc.page_count
    doc.close()
    return n


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    pptx = Path(sys.argv[1]).resolve()
    outdir = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else pptx.parent / "_qa"
    outdir.mkdir(parents=True, exist_ok=True)
    for old in list(outdir.glob("slide-*.png")) + list(outdir.glob("slide-*.jpg")):
        old.unlink()

    if sys.platform == "win32":
        n = render_windows(pptx, outdir)
        note = "PowerPoint COM（実物どおり）"
    else:
        n = render_libreoffice(pptx, outdir)
        note = "LibreOffice（※メイリオ未導入だと代替フォント。文字あふれ判定はPCで最終確認すること）"
    print(f"{n} 枚を出力: {outdir}")
    print(f"レンダラ: {note}")


if __name__ == "__main__":
    main()
