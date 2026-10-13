# -*- coding: utf-8 -*-
"""PowerPointの「修復」が出るスライドを切り分けるための、1枚だけ入ったファイルを作る。

    python -m _build.arauchi_v4.bisect <PowerPointで保存した47枚.pptx> <v4.pptx> <出力フォルダ>

- 元の47枚のzipをもとに、1枚のスライドだけ残す（ほかのスライド・ノート・使わない画像は外す）
- 残す1枚の中身は v4 の版（対照用に、元のままの版も作る）
- ファイル名＝切り分け_S○○.pptx（○○＝v4でのスライド番号）
"""
import posixpath
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

from _build.arauchi_v4.build import slide_order, normalize_decl, c14n, rel_set

PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def read_all(path):
    z = zipfile.ZipFile(path)
    return {i.filename: z.read(i.filename) for i in z.infolist()}, [i.filename for i in z.infolist()]


def targets(parts, rels_name):
    base = posixpath.dirname(posixpath.dirname(rels_name))
    out = []
    for r in etree.fromstring(parts[rels_name]):
        if r.get("TargetMode") == "External":
            continue
        t = r.get("Target")
        out.append(posixpath.normpath(t[1:] if t.startswith("/") else posixpath.join(base, t)))
    return out


def reachable(parts):
    seen, todo = set(), ["ppt/presentation.xml", "docProps/app.xml", "docProps/core.xml"]
    seen.update(todo)
    seen.update({"_rels/.rels", "[Content_Types].xml"})
    while todo:
        p = todo.pop()
        rn = posixpath.join(posixpath.dirname(p), "_rels", posixpath.basename(p) + ".rels")
        if rn in parts:
            seen.add(rn)
            for t in targets(parts, rn):
                if t not in seen and t in parts:
                    seen.add(t)
                    todo.append(t)
    return seen


def make_single(base_parts, order, keep_idx, new_xml, new_rels, out_path, order_names):
    parts = dict(base_parts)
    keep = order[keep_idx]
    if new_xml is not None:
        parts[keep] = new_xml
    if new_rels is not None:
        parts[keep.replace("slides/", "slides/_rels/") + ".rels"] = new_rels
    # presentation.xml / rels：残す1枚だけにする
    prels = parts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    pres = parts["ppt/presentation.xml"].decode("utf-8")
    rid_keep = re.search(r'<Relationship [^>]*Id="(rId\d+)"[^>]*Target="%s"' % re.escape(keep.replace("ppt/", "")), prels)
    if rid_keep is None:
        rid_keep = re.search(r'<Relationship [^>]*Target="%s"[^>]*Id="(rId\d+)"' % re.escape(keep.replace("ppt/", "")), prels)
    rid_keep = rid_keep.group(1)
    for name in order:
        if name == keep:
            continue
        rel_target = name.replace("ppt/", "")
        m = re.search(r'<Relationship [^>]*Target="%s"[^>]*/>' % re.escape(rel_target), prels)
        rid = re.search(r'Id="(rId\d+)"', m.group(0)).group(1)
        prels = prels.replace(m.group(0), "")
        pres = re.sub(r'<p:sldId [^>]*r:id="%s"[^>]*/>' % rid, "", pres)
    parts["ppt/_rels/presentation.xml.rels"] = prels.encode("utf-8")
    parts["ppt/presentation.xml"] = pres.encode("utf-8")
    # 使わないパーツを外す
    keep_set = reachable(parts)
    for n in list(parts):
        if n not in keep_set:
            del parts[n]
    ct = parts["[Content_Types].xml"].decode("utf-8")
    for m in re.findall(r'<Override [^>]*PartName="/([^"]+)"[^>]*/>', ct):
        if m not in parts:
            ct = re.sub(r'<Override [^>]*PartName="/%s"[^>]*/>' % re.escape(m), "", ct)
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    # app.xml（スライド数・タイトル一覧）
    app = parts["docProps/app.xml"].decode("utf-8")
    n_notes = sum(1 for n in parts if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n))
    app = re.sub(r"<Slides>\d+</Slides>", "<Slides>1</Slides>", app)
    app = re.sub(r"<Notes>\d+</Notes>", "<Notes>%d</Notes>" % n_notes, app)
    app = re.sub(r"(<vt:lpstr>スライド タイトル</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+(</vt:i4>)", r"\g<1>1\g<2>", app)
    titles = re.search(r'<vt:vector size="(\d+)" baseType="lpstr">(.*?)</vt:vector>', app, re.S)
    items = re.findall(r"<vt:lpstr>.*?</vt:lpstr>", titles.group(2), re.S)
    head, slides_t = items[:7], items[7:]
    new_items = head + slides_t[:1]
    app = app.replace(titles.group(0), '<vt:vector size="%d" baseType="lpstr">%s</vt:vector>' % (len(new_items), "".join(new_items)))
    parts["docProps/app.xml"] = app.encode("utf-8")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        names = [n for n in order_names if n in parts] + [n for n in parts if n not in order_names]
        for n in names:
            zo.writestr(n, parts[n])


def main(base_path, v4_path, outdir):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    base_parts, order_names = read_all(base_path)
    zb, zn = zipfile.ZipFile(base_path), zipfile.ZipFile(v4_path)
    base_order, v4_order = slide_order(zb), slide_order(zn)
    drop = base_order[15]
    pairs = [(b, n) for i, (b, n) in enumerate(zip([x for j, x in enumerate(base_order) if j != 15], v4_order))]
    changed = []
    for page, (b, n) in enumerate(pairs, 1):
        nx, nr = zn.read(n), zn.read(n.replace("slides/", "slides/_rels/") + ".rels")
        rb = zb.read(b.replace("slides/", "slides/_rels/") + ".rels")
        xml_changed = c14n(zb.read(b)) != c14n(nx)
        rel_changed = rel_set(rb) != rel_set(nr)
        if xml_changed or rel_changed:
            changed.append((page, b, normalize_decl(nx), normalize_decl(nr) if rel_changed else None))
    for page, b, nx, nr in changed:
        make_single(base_parts, base_order, base_order.index(b), nx, nr, outdir / f"切り分け_S{page:02d}.pptx", order_names)
    # 対照：元のまま（PowerPointが保存したまま）の1枚。S10（直していないページ）
    b10 = pairs[9][0]
    make_single(base_parts, base_order, base_order.index(b10), None, None, outdir / "切り分け_対照_元のS10.pptx", order_names)
    print("作成:", [p for p, *_ in changed], "＋対照")


if __name__ == "__main__":
    main(*sys.argv[1:4])
