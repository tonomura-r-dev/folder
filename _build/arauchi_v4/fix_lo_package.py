# -*- coding: utf-8 -*-
"""LibreOfficeで保存し直されたpptxの「ファイルのつなぎ」の不具合を直す（PowerPointの「修復」の原因になるもの）。

    python -m _build.arauchi_v4.fix_lo_package <入力.pptx> <出力.pptx>

直すもの
- SVG画像の種類の宣言が `image/svg`（誤り）→ `image/svg+xml`
- 中身がEMFなのに拡張子が .wmf（種類の宣言も x-wmf）→ .emf にして `image/x-emf`
中身（スライドのXML）は一切変えない。
"""
import re
import sys
import zipfile


def fix(src, dst):
    zi = zipfile.ZipFile(src)
    ct = zi.read("[Content_Types].xml").decode("utf-8")
    n_svg = ct.count('ContentType="image/svg"')
    ct = ct.replace('ContentType="image/svg"', 'ContentType="image/svg+xml"')
    renames = {}
    for name in zi.namelist():
        if name.startswith("ppt/media/") and name.endswith(".wmf"):
            data = zi.read(name)
            if data[40:44] == b" EMF":          # 中身はEMF
                renames[name] = name[:-4] + ".emf"
    for old, new in renames.items():            # 個別の宣言（Override）はやめて、PowerPointと同じ「拡張子ごとの宣言（Default）」にする
        ct, k = re.subn(r'<Override PartName="/%s" ContentType="image/x-wmf"/>' % re.escape(old), "", ct)
        assert k == 1, old
    defaults = ""
    if renames:
        defaults += '<Default Extension="emf" ContentType="image/x-emf"/>'
    if n_svg:
        ct = re.sub(r'<Override PartName="/ppt/media/[^"]+\.svg" ContentType="image/svg\+xml"/>', "", ct)
        defaults += '<Default Extension="svg" ContentType="image/svg+xml"/>'
    ct = ct.replace('<Default Extension="xml"', defaults + '<Default Extension="xml"', 1)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            name = info.filename
            data = zi.read(name)
            if name == "[Content_Types].xml":
                data = ct.encode("utf-8")
            elif name.endswith(".rels"):
                s = data.decode("utf-8")
                for old, new in renames.items():
                    s = s.replace("../media/" + old.split("/")[-1], "../media/" + new.split("/")[-1])
                data = s.encode("utf-8")
            zo.writestr(renames.get(name, name), data)
    return n_svg, renames


if __name__ == "__main__":
    n, r = fix(sys.argv[1], sys.argv[2])
    print("svgの宣言を直した数:", n, "／ 拡張子を直した画像:", r)
