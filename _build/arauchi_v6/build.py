# -*- coding: utf-8 -*-
"""あらうち提案資料 v6 をつくる（2026-10-13 殿村さん指示：成果報酬プランの追加）。

    python -m _build.arauchi_v6.build <v5.pptx> <出力.pptx>

1. v5 に pages.py の直しを当てる（python-pptx。S24を直し、新しい2ページを足す）
2. v5 のzipに、直したS24のXMLと新しい2ページだけを差し込む（それ以外のパーツは v5 のまま。
   [Content_Types].xml・presentation.xml・rels に新しい2ページ分を足す）
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation

from _build.arauchi_v4.build import c14n, normalize_decl, rel_set, slide_order
from _build.arauchi_v6.pages import apply

NEW_AT = (24, 25)     # 0始まりの位置（p25・p26）
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"


def repack(base_path, built_path, out_path):
    zb, zn = zipfile.ZipFile(base_path), zipfile.ZipFile(built_path)
    base_order, built_order = slide_order(zb), slide_order(zn)
    assert len(built_order) == len(base_order) + len(NEW_AT), (len(base_order), len(built_order))
    repl, new_slides = {}, []
    for i, n_name in enumerate(built_order):
        if i in NEW_AT:
            new_slides.append((i, n_name))
            continue
        j = i - sum(1 for k in NEW_AT if k < i)
        b_name = base_order[j]
        if c14n(zb.read(b_name)) != c14n(zn.read(n_name)):
            repl[b_name] = normalize_decl(zn.read(n_name))
        b_rels = b_name.replace("slides/", "slides/_rels/") + ".rels"
        n_rels = n_name.replace("slides/", "slides/_rels/") + ".rels"
        if rel_set(zb.read(b_rels)) != rel_set(zn.read(n_rels)):
            raise SystemExit(f"想定外のrelsの違い: {b_rels}")
    existing = {n for n in zb.namelist()}
    nums = [int(m.group(1)) for n in existing for m in [re.match(r"ppt/slides/slide(\d+)\.xml$", n)] if m]
    nxt = max(nums) + 1
    prels = zb.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    pres = zb.read("ppt/presentation.xml").decode("utf-8")
    ct = zb.read("[Content_Types].xml").decode("utf-8")
    rid_n = max(int(x) for x in re.findall(r'Id="rId(\d+)"', prels)) + 1
    sld_ids = [int(x) for x in re.findall(r'<p:sldId id="(\d+)"', pres)]
    sid_n = max(sld_ids) + 1
    entries = re.findall(r"<p:sldId [^>]*/>", pres)
    added_parts = {}
    new_entries = []
    for k, (idx, n_name) in enumerate(new_slides):
        part = f"ppt/slides/slide{nxt + k}.xml"
        added_parts[part] = normalize_decl(zn.read(n_name))
        added_parts[part.replace("slides/", "slides/_rels/") + ".rels"] = normalize_decl(zn.read(n_name.replace("slides/", "slides/_rels/") + ".rels"))
        ct = ct.replace("</Types>", f'<Override PartName="/{part}" ContentType="{SLIDE_CT}"/></Types>')
        rid = f"rId{rid_n + k}"
        prels = prels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{SLIDE_REL}" Target="slides/slide{nxt + k}.xml"/></Relationships>')
        new_entries.append((idx, f'<p:sldId id="{sid_n + k}" r:id="{rid}"/>'))
    for idx, e in new_entries:                      # 小さい位置から順に差し込む（indexは出来上がりの位置）
        entries.insert(idx, e)
    pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", lambda m: "<p:sldIdLst>" + "".join(entries) + "</p:sldIdLst>", pres, flags=re.S)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zb.infolist():
            name = info.filename
            if name == "[Content_Types].xml":
                data = ct.encode("utf-8")
            elif name == "ppt/presentation.xml":
                data = pres.encode("utf-8")
            elif name == "ppt/_rels/presentation.xml.rels":
                data = prels.encode("utf-8")
            elif name in repl:
                data = repl.pop(name)
            else:
                data = zb.read(name)
            zo.writestr(name, data)
        for name, data in added_parts.items():
            zo.writestr(name, data)
    return sorted(repl), sorted(added_parts)


def main(base, out):
    tmp = out + ".tmp.pptx"
    prs = Presentation(base)
    apply(prs)
    prs.save(tmp)
    left, added = repack(base, tmp, out)
    Path(tmp).unlink()
    print("saved:", out)
    print("追加したパーツ:", added)


if __name__ == "__main__":
    main(*sys.argv[1:3])
