# -*- coding: utf-8 -*-
"""あらうち提案資料 v8 をつくる（2026-10-13 殿村さん指示）。

    python -m _build.arauchi_v8.build <v7.pptx（殿村さんがPowerPointで保存した48枚）> <出力.pptx>

1. v7 に pages.py の直しを当てる（python-pptx）
2. v7 のzipに、直したスライドのXMLだけを差し込む（それ以外のパーツは PowerPoint が書いたまま）
3. プラン④の詳細（v7のS17）を zip から外す（presentation.xml・rels・[Content_Types].xml・docProps/app.xml も合わせる）
   ※PowerPointで「修復」が出る件の対策。python-pptxでzipごと保存し直さない（CLAUDE.md 2026-10-13）
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation

from _build.arauchi_v4.build import c14n, normalize_decl, rel_set, slide_order
from _build.arauchi_v8.pages import apply

DROP_INDEX = 16          # v7のS17（0始まりで16）＝プラン④｜成果報酬型


def repack(base_path, built_path, out_path):
    zb, zn = zipfile.ZipFile(base_path), zipfile.ZipFile(built_path)
    base_order, built_order = slide_order(zb), slide_order(zn)
    assert len(base_order) == len(built_order) == 48, (len(base_order), len(built_order))
    drop = base_order[DROP_INDEX]
    drop_rels = drop.replace("slides/", "slides/_rels/") + ".rels"

    # ---- 削除するパーツ（スライドと、そのノート）
    gone = {drop, drop_rels}
    for r in etree.fromstring(zb.read(drop_rels)):
        if r.get("Type").endswith("/notesSlide"):
            nt = "ppt/" + r.get("Target").replace("../", "")
            gone |= {nt, nt.replace("notesSlides/", "notesSlides/_rels/") + ".rels"}
    # 他のどこからも参照されていないこと
    for name in zb.namelist():
        if name.endswith(".rels") and name not in gone and name != "ppt/_rels/presentation.xml.rels":
            assert drop.split("/")[-1] not in zb.read(name).decode("utf-8"), name

    # ---- presentation.xml / rels から外す
    slide_file = drop.split("/")[-1]
    prels = zb.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    m = re.search(r'<Relationship [^>]*Target="slides/%s"[^>]*/>' % re.escape(slide_file), prels)
    rid = re.search(r'Id="(rId\d+)"', m.group(0)).group(1)
    prels = prels.replace(m.group(0), "")
    pres = zb.read("ppt/presentation.xml").decode("utf-8")
    m2 = re.search(r'<p:sldId [^>]*r:id="%s"[^>]*/>' % rid, pres)
    pres = pres.replace(m2.group(0), "")

    # ---- [Content_Types].xml
    ct = zb.read("[Content_Types].xml").decode("utf-8")
    for g in gone:
        if not g.endswith(".rels"):
            ct, n = re.subn(r'<Override [^>]*PartName="/%s"[^>]*/>' % re.escape(g), "", ct)
            assert n == 1, g

    # ---- docProps/app.xml（スライド数・タイトル一覧）
    app = zb.read("docProps/app.xml").decode("utf-8")
    assert "<Slides>48</Slides>" in app and '<vt:vector size="56" baseType="lpstr">' in app
    app = app.replace("<Slides>48</Slides>", "<Slides>47</Slides>")
    app = app.replace("<vt:lpstr>スライド タイトル</vt:lpstr></vt:variant><vt:variant><vt:i4>48</vt:i4>",
                      "<vt:lpstr>スライド タイトル</vt:lpstr></vt:variant><vt:variant><vt:i4>47</vt:i4>")
    app = app.replace('<vt:vector size="56" baseType="lpstr">', '<vt:vector size="55" baseType="lpstr">')
    app = app.replace("<vt:lpstr>PowerPoint プレゼンテーション</vt:lpstr>", "", 1)
    if any(g.startswith("ppt/notesSlides/notesSlide") and not g.endswith(".rels") for g in gone):
        app = re.sub(r"<Notes>(\d+)</Notes>", lambda mm: "<Notes>%d</Notes>" % (int(mm.group(1)) - 1), app)

    # ---- スライド：並び順で対応づけ、中身が変わったものだけ差し替え
    replaced = {}
    for i, (b_name, n_name) in enumerate(zip(base_order, built_order)):
        if i == DROP_INDEX:
            assert c14n(zb.read(b_name)) == c14n(zn.read(n_name)), "削除するスライドを触ってしまった"
            continue
        if c14n(zb.read(b_name)) != c14n(zn.read(n_name)):
            replaced[b_name] = normalize_decl(zn.read(n_name))
        b_rels = b_name.replace("slides/", "slides/_rels/") + ".rels"
        n_rels = n_name.replace("slides/", "slides/_rels/") + ".rels"
        if rel_set(zb.read(b_rels)) != rel_set(zn.read(n_rels)):
            raise SystemExit(f"想定外のrelsの違い: {b_rels}")

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zb.infolist():
            name = info.filename
            if name in gone:
                continue
            if name == "ppt/presentation.xml":
                data = pres.encode("utf-8")
            elif name == "ppt/_rels/presentation.xml.rels":
                data = prels.encode("utf-8")
            elif name == "[Content_Types].xml":
                data = ct.encode("utf-8")
            elif name == "docProps/app.xml":
                data = app.encode("utf-8")
            elif name in replaced:
                data = replaced.pop(name)
            else:
                data = zb.read(name)
            zo.writestr(name, data)
    assert not replaced, replaced
    return sorted(n for i, n in enumerate(base_order) if i != DROP_INDEX
                  and c14n(zb.read(n)) != c14n(zn.read(built_order[i]))), sorted(gone)


def main(base, out):
    tmp = out + ".tmp.pptx"
    prs = Presentation(base)
    apply(prs)
    prs.save(tmp)
    changed, gone = repack(base, tmp, out)
    Path(tmp).unlink()
    print("saved:", out)
    print("差し替えたスライドのXML:", changed)
    print("削除したパーツ:", gone)


if __name__ == "__main__":
    main(*sys.argv[1:3])
