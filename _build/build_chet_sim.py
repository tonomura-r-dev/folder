# -*- coding: utf-8 -*-
"""CHET Group（Qooo!!）SIM 予算別3パターン（200万・100万・50万）

ベースは _templates/DYM_LINEOA_SIM_FMT_ver3.00.xlsx。
FMTはセル内チェックボックス（xl/featurePropertyBag/）を使っているので、
openpyxlで保存せず、xlsxを展開して sheet XML を直接書き換える（CLAUDE.md 2026-09-29）。

  SIM1_予算200万 / SIM1_比較 / SIM2_予算100万 / SIM2_比較 / SIM3_予算50万 / SIM3_比較 / SIM考え方
  （番号順＝強い順）

使い方:
    python _build/build_chet_sim.py [YYYYMMDD]
"""
import re
import shutil
import sys
import tempfile
import uuid
import zipfile
from copy import deepcopy
from datetime import date
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_templates" / "DYM_LINEOA_SIM_FMT_ver3.00.xlsx"
STAMP = sys.argv[1] if len(sys.argv) > 1 else date.today().strftime("%Y%m%d")
OUT = ROOT / f"{STAMP}_株式会社CHET Group御中_LINEOA施策提案【SIM】ver1.0.xlsx"

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = {"m": NS}
q = lambda tag: f"{{{NS}}}{tag}"
XR_UID = "{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid"

# ---------------- 与件 ----------------
CLIENT = "株式会社CHET Group御中_LINEOA施策提案【SIM】"
UU = 15661                 # サイトUU（先方から受領）
FRIENDS_NOW = 12566        # wakaba projectのLINE友だち数（page.line.me・2026/10/1）
UNIT = 1000                # 友だち1人あたり（先方の良い時の実績）
POSTS = 4                  # 企画配信 月4本
STEPS = 10                 # ステップ配信 10本
BLOCK0, BLOCK_K = 0.406, 0.98          # ブロック率：40.6%から出し分け配信で徐々に低下
CVR = [0.005, 0.006, 0.007, 0.008, 0.009, 0.010]   # 申込÷想定Click（0.3〜1.0%の目安内で月ごとに上げる）
INIT_BUILD = 200000        # 初期構築（仮）
PLANS = [  # (シート番号, 予算, シート名)
    (1, 2_000_000, "SIM1_予算200万"),
    (2, 1_000_000, "SIM2_予算100万"),
    (3, 500_000, "SIM3_予算50万"),
]
MONTHS = "EFGHIJ"


# ---------------- XMLセル操作 ----------------
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
    c = get_cell(root, ref)
    _clear(c)
    f = etree.SubElement(c, q("f"))
    f.text = formula.lstrip("=")


def set_n(root, ref, value):
    c = get_cell(root, ref)
    _clear(c)
    etree.SubElement(c, q("v")).text = repr(value) if isinstance(value, float) else str(value)


def set_b(root, ref, value):
    c = get_cell(root, ref)
    _clear(c)
    c.set("t", "b")
    etree.SubElement(c, q("v")).text = "1" if value else "0"


def set_s(root, ref, text):
    c = get_cell(root, ref)
    _clear(c)
    c.set("t", "inlineStr")
    is_ = etree.SubElement(c, q("is"))
    t = etree.SubElement(is_, q("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def check_shared(root, name):
    """共有数式の親を消していないか（子だけ残ると壊れる）"""
    masters = {f.get("si") for f in root.iter(q("f")) if f.get("t") == "shared" and f.get("ref")}
    orphans = [f.getparent().get("r") for f in root.iter(q("f"))
               if f.get("t") == "shared" and not f.get("ref") and f.get("si") not in masters]
    assert not orphans, f"{name}: 共有数式の親が無いセル {orphans}"


# ---------------- SIMシートの中身 ----------------
def build_sim(root, budget):
    man = budget // 10000
    set_s(root, "A4", f"{CLIENT}（予算{man}万／友だち1人あたり{UNIT:,}円）")
    # 要件定義
    set_s(root, "C8", "LINE追加（新しい友だち）")
    set_n(root, "D8", 0)
    set_s(root, "C9", "申込（LINE上の申込フォーム）")
    set_n(root, "D9", 0)
    set_s(root, "C10", "-")
    set_n(root, "D10", 0)
    set_n(root, "H8", UU)
    set_n(root, "H9", 0)
    set_s(root, "G10", "CV（別施策）")
    set_n(root, "H10", 0)
    set_n(root, "K8", 0)
    set_b(root, "P9", False)
    # 施策トグル：広告→LP→LINE（CPF行）だけ使う
    set_s(root, "B17", "友だち追加（広告→LP→LINE）")
    for col in MONTHS:
        set_b(root, f"{col}17", True)
        set_b(root, f"{col}18", False)
    # 友だち追加（広告）＝ 予算 ÷ 友だち1人あたり
    set_s(root, "B22", "友だち追加（広告→LP→LINE）")
    for col in MONTHS:
        set_f(root, f"{col}22", f"IF({col}17=FALSE,0,ROUND({col}71/$P$30,0))")
    set_s(root, "N22", f"予算（71行）÷友だち1人あたり{UNIT:,}円（P30）。単価は先方の良い時の実績（広告・運用込みの見込み）")
    set_s(root, "O30", "友だち1人あたり（広告・運用込み）")
    set_n(root, "P30", UNIT)
    set_s(root, "N26", "＋サイトからの自然な追加：UU×0.5%（FMT標準のP15）")
    # ブロック率
    set_n(root, "E27", BLOCK0)
    for prev, col in zip(MONTHS, MONTHS[1:]):
        set_f(root, f"{col}27", f"+{prev}27*{BLOCK_K}")
    set_s(root, "N27", "ブロック率：40.6%から、出し分け配信で毎月2%ずつ低下（仮）")
    # 既存の友だち
    set_n(root, "D28", FRIENDS_NOW)
    set_f(root, "D29", "+D28*(1-D27)")
    set_s(root, "N28", f"D28：wakaba projectのLINEの友だち{FRIENDS_NOW:,}人（page.line.me・2026/10/1）。他のLINEは非公開で含めていない")
    # 配信本数
    for col in MONTHS:
        set_n(root, f"{col}30", POSTS)
        set_n(root, f"{col}32", STEPS)
    # 想定Click：配信の流入＋リッチメニューのタップ（有効な友だち×月に使う率30%×2回）
    for col in MONTHS:
        set_f(root, f"{col}35", f"+{col}31+{col}33+1.05*({col}31+{col}33)+{col}29*0.3*2")
    set_s(root, "N35", "＋リッチメニューのタップ：有効な友だち×月に使う率30%×2回")
    # CV①＝LINE追加（新しい友だち）。CPA・CVRには含めない
    for col in MONTHS:
        set_f(root, f"{col}36", f"ROUND({col}26,0)")
    set_s(root, "N36", "CV①LINE追加＝その月の新しい友だち（広告＋サイト）。CPA・CVRには含めない")
    # CV②＝申込。想定Click×CVR（0.5%→1.0%）
    for col, r in zip(MONTHS, CVR):
        set_f(root, f"{col}37", f"ROUND({col}35*{r},0)")
    set_s(root, "N37", "CV②申込：想定Click×0.5%→1.0%（運用が育つ想定で月ごとに上げる）")
    # 合計CV＝CV②のみ
    for col in "D" + MONTHS:
        set_f(root, f"{col}40", f"+{col}37")
    set_s(root, "N40", "合計CVはCV②（申込）のみ。CPA・CVRは申込で計算")
    # 費用
    for col in "D" + MONTHS:
        set_n(root, f"{col}65", 0)
        set_n(root, f"{col}66", 0)
    set_s(root, "N65", "従量課金は54行（新料金テーブル）で計算。65行の旧式（2.75円）は二重計上になるため0")
    set_s(root, "N66", "離脱防止：今回のSIMには入れない")
    set_n(root, "D69", INIT_BUILD)
    for col in MONTHS:
        set_n(root, f"{col}69", 0)
    set_s(root, "N69", "初期構築（仮）。月々の運用費は71行の予算に含む")
    set_s(root, "B71", "予算（友だち追加・運用込み）")
    set_n(root, "D71", 0)
    for col in MONTHS:
        set_n(root, f"{col}71", budget)
    set_s(root, "N71", f"予算型コンサル：月{man}万（友だち追加の広告と運用を含む）")


def build_compare(root, sheet_name):
    for f in root.iter(q("f")):
        if f.text and "SIM1_!" in f.text:
            f.text = f.text.replace("SIM1_!", f"'{sheet_name}'!")
    # CV数は申込（CV②）のみ（チェングロウス ver2.6 と同じ）
    set_s(root, "A16", "CV数（+LINEOA施策）※申込のみ")
    for col in "CDE":
        set_f(root, f"{col}16", f"+{col}18")
    set_s(root, "A17", "CV①_LINE追加（参考・CPAに含めない）")
    set_s(root, "A18", "CV②_申込（LINE上の申込フォーム）")
    # 先方のKPI：友だち1人あたり（月の総額÷その月の新しい友だち）
    set_s(root, "A24", "友だち1人あたり（月の総額÷新しい友だち）")
    for col in "CDE":
        c = get_cell(root, f"{col}24")
        c.set("s", get_cell(root, f"{col}23").get("s", "0"))
    get_cell(root, "A24").set("s", get_cell(root, "A23").get("s", "0"))
    set_s(root, "C24", "--")
    for col in "DE":
        set_f(root, f"{col}24", f'+IFERROR({col}5/{col}17,"--")')


# ---------------- 組み立て ----------------
work = Path(tempfile.mkdtemp())
with zipfile.ZipFile(SRC) as z:
    z.extractall(work)

ws_dir = work / "xl" / "worksheets"
sim_xml = etree.parse(str(ws_dir / "sheet1.xml"))
cmp_xml = etree.parse(str(ws_dir / "sheet2.xml"))

files = {}  # 出力ファイル名 -> tree
sheet_list = []  # (name, file)
next_file = 4
for i, (no, budget, name) in enumerate(PLANS):
    s = deepcopy(sim_xml)
    c = deepcopy(cmp_xml)
    build_sim(s.getroot(), budget)
    build_compare(c.getroot(), name)
    if i == 0:
        sf, cf = "sheet1.xml", "sheet2.xml"
    else:
        sf, cf = f"sheet{next_file}.xml", f"sheet{next_file + 1}.xml"
        next_file += 2
        for sv in s.getroot().iter(q("sheetView")):
            sv.attrib.pop("tabSelected", None)
        # シートごとに別のID（同じIDが2枚あるとExcelが修復をかけることがある）
        for k, tree in enumerate((s, c)):
            uid = uuid.uuid5(uuid.NAMESPACE_URL, f"chet-sim-{no}-{k}")
            tree.getroot().set(XR_UID, "{" + str(uid).upper() + "}")
    for tree, label in ((s, name), (c, f"SIM{no}_比較")):
        check_shared(tree.getroot(), label)
    files[sf], files[cf] = s, c
    sheet_list += [(name, sf), (f"SIM{no}_比較", cf)]
sheet_list.append(("SIM考え方", "sheet3.xml"))

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


files["sheet3.xml"] = etree.parse(str(ws_dir / "sheet3.xml"))
for fn, tree in files.items():
    strip_cache(tree.getroot())
    tree.write(str(ws_dir / fn), xml_declaration=True, encoding="UTF-8", standalone=True)

# workbook.xml：シート一覧と再計算
wb_path = work / "xl" / "workbook.xml"
wb = wb_path.read_text(encoding="utf-8")
rels_path = work / "xl" / "_rels" / "workbook.xml.rels"
rels = rels_path.read_text(encoding="utf-8")
rel_ids = {}
for name, fn in sheet_list:
    m = re.search(r'Id="(rId\d+)"[^>]*Target="worksheets/%s"' % fn, rels) or \
        re.search(r'Target="worksheets/%s"[^>]*Id="(rId\d+)"' % fn, rels)
    if m:
        rel_ids[fn] = m.group(1)
n_rel = 100
for name, fn in sheet_list:
    if fn not in rel_ids:
        rid = f"rId{n_rel}"
        n_rel += 1
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{fn}"/></Relationships>')
        rel_ids[fn] = rid
# calcChain は消す（再計算時にExcelが作り直す）
rels = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', "", rels)
rels_path.write_text(rels, encoding="utf-8")
(work / "xl" / "calcChain.xml").unlink(missing_ok=True)

sheets_xml = "<sheets>" + "".join(
    f'<sheet name="{name}" sheetId="{i + 1}" r:id="{rel_ids[fn]}"/>' for i, (name, fn) in enumerate(sheet_list)
) + "</sheets>"
wb = re.sub(r"<sheets>.*?</sheets>", sheets_xml, wb, flags=re.S)
wb = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wb)
wb_path.write_text(wb, encoding="utf-8")

ct_path = work / "[Content_Types].xml"
ct = ct_path.read_text(encoding="utf-8")
ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
for name, fn in sheet_list:
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

# zip（[Content_Types].xml を先頭に）
OUT.unlink(missing_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(ct_path, "[Content_Types].xml")
    for p in sorted(work.rglob("*")):
        if p.is_file() and p.name != "[Content_Types].xml":
            z.write(p, p.relative_to(work).as_posix())
shutil.rmtree(work)

# チェックボックス（featurePropertyBag）が残っているか
with zipfile.ZipFile(OUT) as z:
    assert "xl/featurePropertyBag/featurePropertyBag.xml" in z.namelist(), "チェックボックスが消えた"
print("saved:", OUT.name)
print("sheets:", [n for n, _ in sheet_list])
