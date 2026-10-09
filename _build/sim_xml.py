# -*- coding: utf-8 -*-
"""SIM（FMT ver3.00）をXML直接編集で組み立てる共通部品。

FMTはセル内チェックボックス（xl/featurePropertyBag/）を使っているので、
openpyxlで保存すると消える（CLAUDE.md 2026-09-29）。xlsxを展開して sheet XML を直接書き換える。

使い方（例）:
    from sim_xml import build_workbook, set_f, set_n, set_b, set_s
    build_workbook(SRC, OUT, [(sim_name, cmp_name, fill_sim, fill_cmp), ...])

    fill_sim(root) / fill_cmp(root, sim_name) で各シートを書き換える。
    SIMシートはFMTの sheet1（SIM1_）、比較シートは sheet2（SIM1_比較）をコピーして使う。
"""
import re
import shutil
import tempfile
import uuid
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = {"m": NS}
XR_UID = "{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid"


def q(tag):
    return f"{{{NS}}}{tag}"


# ---------------- セル操作 ----------------
def col_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def get_cell(root, ref):
    col, r = split_ref(ref)
    sd = root.find("m:sheetData", N)
    row = sd.find(f"m:row[@r='{r}']", N)
    if row is None:
        row = etree.SubElement(sd, q("row"), r=str(r))
        rows = sorted(sd.findall("m:row", N), key=lambda e: int(e.get("r")))
        for e in rows:
            sd.remove(e)
            sd.append(e)
    c = row.find(f"m:c[@r='{ref}']", N)
    if c is None:
        c = etree.Element(q("c"), r=ref)
        after = None
        for e in row.findall("m:c", N):
            if col_idx(split_ref(e.get("r"))[0]) < col_idx(col):
                after = e
        if after is None:
            row.insert(0, c)
        else:
            after.addnext(c)
    return c


def _clear(c):
    for ch in list(c):
        c.remove(ch)
    if "t" in c.attrib:
        del c.attrib["t"]


def set_f(root, ref, formula):
    """数式（先頭の = は付けても付けなくてもよい）"""
    c = get_cell(root, ref)
    _clear(c)
    etree.SubElement(c, q("f")).text = formula.lstrip("=")


def set_n(root, ref, value):
    c = get_cell(root, ref)
    _clear(c)
    etree.SubElement(c, q("v")).text = repr(value) if isinstance(value, float) else str(value)


def set_b(root, ref, value):
    """TRUE/FALSE（チェックボックスのセルはスタイルが残るので、チェックボックスのまま）"""
    c = get_cell(root, ref)
    _clear(c)
    c.set("t", "b")
    etree.SubElement(c, q("v")).text = "1" if value else "0"


def set_s(root, ref, text):
    c = get_cell(root, ref)
    _clear(c)
    c.set("t", "inlineStr")
    t = etree.SubElement(etree.SubElement(c, q("is")), q("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def check_shared(root, name):
    """共有数式の親を消していないか（子だけ残ると壊れる）"""
    masters = {f.get("si") for f in root.iter(q("f")) if f.get("t") == "shared" and f.get("ref")}
    orphans = [f.getparent().get("r") for f in root.iter(q("f"))
               if f.get("t") == "shared" and not f.get("ref") and f.get("si") not in masters]
    assert not orphans, f"{name}: 共有数式の親が無いセル {orphans}"


def strip_cache(root):
    """数式セルの古い計算結果を消す（FMTの値が残るとプレビューやLibreOfficeで古い数字が出る）。
    開いたときに fullCalcOnLoad でExcelが計算し直す。"""
    for c in root.iter(q("c")):
        if c.find(q("f")) is not None:
            v = c.find(q("v"))
            if v is not None:
                c.remove(v)
            if c.get("t") in ("str", "e", "b", "n"):
                del c.attrib["t"]


def replace_sheet_ref(root, old, new):
    """比較シートの参照（SIM1_!）を新しいシート名に付け替える"""
    for f in root.iter(q("f")):
        if f.text and old in f.text:
            f.text = f.text.replace(old, new)


# ---------------- 組み立て ----------------
def build_workbook(src, out, plans, uid_seed="sim"):
    """plans = [(sim_name, cmp_name, fill_sim(root), fill_cmp(root, sim_name)), ...]（並び＝シートの並び）"""
    src, out = Path(src), Path(out)
    work = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(src) as z:
            z.extractall(work)
        ws_dir = work / "xl" / "worksheets"
        sim_xml = etree.parse(str(ws_dir / "sheet1.xml"))
        cmp_xml = etree.parse(str(ws_dir / "sheet2.xml"))

        files, sheet_list, next_file = {}, [], 4
        for i, (sim_name, cmp_name, fill_sim, fill_cmp) in enumerate(plans):
            s, c = deepcopy(sim_xml), deepcopy(cmp_xml)
            fill_sim(s.getroot())
            replace_sheet_ref(c.getroot(), "SIM1_!", f"'{sim_name}'!")
            fill_cmp(c.getroot(), sim_name)
            if i == 0:
                sf, cf = "sheet1.xml", "sheet2.xml"
            else:
                sf, cf = f"sheet{next_file}.xml", f"sheet{next_file + 1}.xml"
                next_file += 2
                for sv in s.getroot().iter(q("sheetView")):
                    sv.attrib.pop("tabSelected", None)
                for k, tree in enumerate((s, c)):  # 同じIDが2枚あるとExcelが修復をかけることがある
                    tree.getroot().set(XR_UID, "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"{uid_seed}-{i}-{k}")).upper() + "}")
            check_shared(s.getroot(), sim_name)
            check_shared(c.getroot(), cmp_name)
            files[sf], files[cf] = s, c
            sheet_list += [(sim_name, sf), (cmp_name, cf)]
        sheet_list.append(("SIM考え方", "sheet3.xml"))
        files["sheet3.xml"] = etree.parse(str(ws_dir / "sheet3.xml"))

        for fn, tree in files.items():
            strip_cache(tree.getroot())
            tree.write(str(ws_dir / fn), xml_declaration=True, encoding="UTF-8", standalone=True)

        # workbook.xml.rels：シートの関係を足し、calcChain を外す
        rels_path = work / "xl" / "_rels" / "workbook.xml.rels"
        rels = rels_path.read_text(encoding="utf-8")
        rel_ids, n_rel = {}, 100
        for name, fn in sheet_list:
            m = re.search(r'Id="(rId\d+)"[^>]*Target="worksheets/%s"' % fn, rels) or \
                re.search(r'Target="worksheets/%s"[^>]*Id="(rId\d+)"' % fn, rels)
            if m:
                rel_ids[fn] = m.group(1)
            else:
                rid = f"rId{n_rel}"
                n_rel += 1
                rels = rels.replace("</Relationships>",
                                    f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{fn}"/></Relationships>')
                rel_ids[fn] = rid
        rels = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', "", rels)
        rels_path.write_text(rels, encoding="utf-8")
        (work / "xl" / "calcChain.xml").unlink(missing_ok=True)

        wb_path = work / "xl" / "workbook.xml"
        wb = wb_path.read_text(encoding="utf-8")
        sheets_xml = "<sheets>" + "".join(
            f'<sheet name="{name}" sheetId="{i + 1}" r:id="{rel_ids[fn]}"/>' for i, (name, fn) in enumerate(sheet_list)
        ) + "</sheets>"
        wb = re.sub(r"<sheets>.*?</sheets>", sheets_xml, wb, flags=re.S)
        wb = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wb)
        wb_path.write_text(wb, encoding="utf-8")

        ct_path = work / "[Content_Types].xml"
        ct = ct_path.read_text(encoding="utf-8")
        ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
        for _, fn in sheet_list:
            part = f"/xl/worksheets/{fn}"
            if part not in ct:
                ct = ct.replace("</Types>",
                                f'<Override PartName="{part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        ct_path.write_text(ct, encoding="utf-8")

        app_path = work / "docProps" / "app.xml"
        app = app_path.read_text(encoding="utf-8")
        app = re.sub(r"(<vt:variant><vt:i4>)\d+(</vt:i4>)", rf"\g<1>{len(sheet_list)}\g<2>", app)
        titles = "".join(f"<vt:lpstr>{name}</vt:lpstr>" for name, _ in sheet_list)
        app = re.sub(r'<TitlesOfParts><vt:vector size="\d+" baseType="lpstr">.*?</vt:vector></TitlesOfParts>',
                     f'<TitlesOfParts><vt:vector size="{len(sheet_list)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>',
                     app, flags=re.S)
        app_path.write_text(app, encoding="utf-8")

        out.unlink(missing_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(ct_path, "[Content_Types].xml")
            for p in sorted(work.rglob("*")):
                if p.is_file() and p.name != "[Content_Types].xml":
                    z.write(p, p.relative_to(work).as_posix())
    finally:
        shutil.rmtree(work)

    with zipfile.ZipFile(out) as z:
        assert "xl/featurePropertyBag/featurePropertyBag.xml" in z.namelist(), "チェックボックスが消えた"
        assert "xfpb" in z.read("xl/styles.xml").decode("utf-8"), "チェックボックスのスタイルが消えた"
    return [n for n, _ in sheet_list]
