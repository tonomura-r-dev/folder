# -*- coding: utf-8 -*-
"""あらうち提案資料 v4 をつくる（2026-10-13 殿村さん指示）。

    python -m _build.arauchi_v4.build <v3.pptx> <PowerPointで保存した47枚.pptx> <出力.pptx> [--only 3,5,24]

1. v3（46枚）に pages.py の直しを当てる（python-pptx）
2. 出来上がったスライドのXMLだけを、PowerPointで保存した元の47枚のzipに差し込む
   （[Content_Types].xml・rels・presentation.xml などは PowerPoint が書いたまま。元S16の削除だけ手で反映）
   ※PowerPointで「修復」が出る件の対策。python-pptxでzipごと保存し直すのをやめた。
"""
import argparse
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation

from _build.arauchi_v4.pages import APPLY

DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")


def c14n(data):
    return etree.tostring(etree.fromstring(data), method="c14n")


def rel_set(data):
    return {(r.get("Id"), r.get("Type"), r.get("Target"), r.get("TargetMode")) for r in etree.fromstring(data)}


def slide_order(z):
    """presentation.xml の並び順でスライドのパーツ名を返す（python-pptx は保存時にスライドの番号を振り直すので、名前では対応が取れない）"""
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    rid_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    rels = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    return ["ppt/" + rels[s.get(rid_attr)] for s in pres.findall(".//p:sldId", ns)]


def normalize_decl(data):
    root = etree.fromstring(data)
    body = etree.tostring(root, encoding="UTF-8", xml_declaration=False)
    return DECL + body


def apply_pages(src, tmp, only=None):
    prs = Presentation(src)
    for no, fn in APPLY.items():
        if only and no not in only:
            continue
        fn(prs.slides[no - 1])
    prs.save(tmp)


def repack(base_path, built_path, out_path, drop="ppt/slides/slide16.xml"):
    zb = zipfile.ZipFile(base_path)
    zn = zipfile.ZipFile(built_path)
    base_names = zb.namelist()
    drop_rels = drop.replace("slides/", "slides/_rels/") + ".rels"

    # ---- 削除するパーツ（元S16と、そのノート）
    gone = {drop, drop_rels}
    notes_target = None
    for r in etree.fromstring(zb.read(drop_rels)):
        if r.get("Type").endswith("/notesSlide"):
            notes_target = "ppt/" + r.get("Target").replace("../", "")
    if notes_target:
        gone.add(notes_target)
        gone.add(notes_target.replace("notesSlides/", "notesSlides/_rels/") + ".rels")

    # ---- presentation.xml / rels から元S16を外す
    slide_file = drop.split("/")[-1]
    prels = zb.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    m = re.search(r'<Relationship [^>]*Target="slides/%s"[^>]*/>' % re.escape(slide_file), prels)
    rid = re.search(r'Id="(rId\d+)"', m.group(0)).group(1)
    prels = prels.replace(m.group(0), "")
    pres = zb.read("ppt/presentation.xml").decode("utf-8")
    m2 = re.search(r'<p:sldId [^>]*r:id="%s"[^>]*/>' % rid, pres)
    pres = pres.replace(m2.group(0), "")

    # ---- [Content_Types].xml から外す
    ct = zb.read("[Content_Types].xml").decode("utf-8")
    for g in gone:
        if g.endswith(".rels"):
            continue
        ct, n = re.subn(r'<Override [^>]*PartName="/%s"[^>]*/>' % re.escape(g), "", ct)
        assert n == 1, g

    # ---- docProps/app.xml（スライド数・タイトル一覧）
    app = zb.read("docProps/app.xml").decode("utf-8")
    app = app.replace("<Slides>47</Slides>", "<Slides>46</Slides>")
    if notes_target:
        app = re.sub(r"<Notes>(\d+)</Notes>", lambda mm: "<Notes>%d</Notes>" % (int(mm.group(1)) - 1), app)
    app = app.replace("<vt:lpstr>スライド タイトル</vt:lpstr></vt:variant><vt:variant><vt:i4>47</vt:i4>",
                      "<vt:lpstr>スライド タイトル</vt:lpstr></vt:variant><vt:variant><vt:i4>46</vt:i4>")
    app = app.replace('<vt:vector size="54" baseType="lpstr">', '<vt:vector size="53" baseType="lpstr">')
    app = app.replace("<vt:lpstr>PowerPoint プレゼンテーション</vt:lpstr>", "", 1)

    # ---- スライド：並び順で対応づけ、中身が変わったものだけ差し替え
    base_order = slide_order(zb)
    built_order = slide_order(zn)
    assert base_order[15] == drop, (base_order[15], drop)
    assert len(base_order) - 1 == len(built_order), (len(base_order), len(built_order))
    pairs = list(zip([n for i, n in enumerate(base_order) if i != 15], built_order))
    replaced = []
    new_parts = {}
    for b_name, n_name in pairs:
        if c14n(zb.read(b_name)) != c14n(zn.read(n_name)):
            new_parts[b_name] = normalize_decl(zn.read(n_name))
            replaced.append(b_name)
        b_rels = b_name.replace("slides/", "slides/_rels/") + ".rels"
        n_rels = n_name.replace("slides/", "slides/_rels/") + ".rels"
        rb = zb.read(b_rels)
        rn = zn.read(n_rels)
        if rel_set(rb) != rel_set(rn):
            # 画像・ノート・レイアウトの宛先は変わらないはず。違うのはリンク（新しく足した分）だけ
            only_b = rel_set(rb) - rel_set(rn)
            only_n = rel_set(rn) - rel_set(rb)
            assert all(r[1].endswith("/hyperlink") for r in only_n) and not only_b, (b_rels, only_b, only_n)
            new_parts[b_rels] = normalize_decl(rn)
            replaced.append(b_rels)

    # ---- 出力（元のzipの並びのまま）
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
            elif name in new_parts:
                data = new_parts.pop(name)
            else:
                data = zb.read(name)
            zo.writestr(name, data)
        for name, data in new_parts.items():       # 元に無かったrels（リンクを足したスライド）
            zo.writestr(name, data)
    return sorted(replaced), sorted(gone)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("v3")
    ap.add_argument("base47")
    ap.add_argument("out")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    only = {int(x) for x in a.only.split(",") if x.strip()} or None
    tmp = a.out + ".tmp.pptx"
    apply_pages(a.v3, tmp, only)
    replaced, gone = repack(a.base47, tmp, a.out)
    Path(tmp).unlink()
    print("saved:", a.out)
    print("差し替えたパーツ:", len(replaced), "／ 削除:", gone)


if __name__ == "__main__":
    main()
