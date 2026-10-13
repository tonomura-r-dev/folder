# -*- coding: utf-8 -*-
"""株式会社あらうち LINEOA施策SIM ver1.5 ＝ ver1.4（殿村さんが手で直した版＝正）に、ADの提案（2026/9/14）の流入を足す。

ver1.4 のシートは1文字も変えない。足すのは新しい3枚だけ：
  ・AD提案0914   ＝ ADの提案シミュレーション（Meta広告）の数字（月）と、LINEへの影響のまとめ
  ・SIM_AD30万   ＝ ver1.4 のSIMシートのコピー。広告流入（H9）にAD30万円案のクリック数を足した場合
  ・SIM_AD100万  ＝ 同じく100万円案
問い合わせ①②は、ver1.4（AD無し）の「新しい友だちに対する率」「クリックに対するCVR」をそのまま使う式にした
（ver1.4の①②を殿村さんが直せば、新しい3枚もそれに合わせて動く）。
広告から来た方の友だち追加率は、FMTの広告用の値（離脱防止 表示65%×8%×8%／サイトのボタン0.25%）をそのまま使う
（広告以外の約半分）。ADのクリック数はサイトへの流入数として足す（UUより多めになる）。

ver1.4 はセル内チェックボックスを使っているので、openpyxlで保存せず、xlsxを展開してXMLを直接書き換える。
LibreOfficeで変換したファイルは納品しない（検算にだけ使う）。

使い方:
    python _build/build_arauchi_sim_v15.py [出力.xlsx]
"""
import re
import shutil
import sys
import tempfile
import uuid
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import N, XR_UID, check_shared, get_cell, q, set_f, set_n, set_s, strip_cache  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "株式会社あらうち御中_LINEOA施策提案【SIM】ver1.4.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "株式会社あらうち御中_LINEOA施策提案【SIM】ver1.5.xlsx"

BASE = "SIM_離脱防止＋サイトのボタン"
AD = "AD提案0914"
SC = [("SIM_AD30万", "B", "AD30万円案"), ("SIM_AD100万", "C", "AD100万円案")]   # (シート名, ADシートの列, 呼び名)
B_ = f"'{BASE}'!"
A_ = f"'{AD}'!"

# 書式（ver1.4で使っているスタイルをそのまま借りる）
S_TITLE, S_SEC, S_HEAD, S_LABEL, S_INT, S_PCT, S_YEN, S_NOTE, S_INPUT = 56, 29, 139, 140, 10, 18, 28, 69, 122


def put(root, ref, value, style=None):
    if isinstance(value, str) and value.startswith("="):
        set_f(root, ref, value)
    elif isinstance(value, str):
        set_s(root, ref, value)
    else:
        set_n(root, ref, value)
    if style is not None:
        get_cell(root, ref).set("s", str(style))


# ---------------------------------------------------------------- ADの提案シート
AD_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<dimension ref="A1:K32"/>'
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="18.5"/>'
    '<cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="26" customWidth="1"/>'
    '<col min="3" max="3" width="22" customWidth="1"/><col min="4" max="9" width="13" customWidth="1"/>'
    '<col min="10" max="10" width="18" customWidth="1"/><col min="11" max="11" width="26" customWidth="1"/></cols>'
    '<sheetData/>'
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    '</worksheet>'
)


def fill_ad(root):
    put(root, "A1", "AD側の提案（2026/9/14作成）と、LINEへの影響", S_TITLE)
    put(root, "A2", "ADの提案シミュレーション（Meta広告）の数字。他社実績をもとにした想定値で、成果の保証ではない。"
                    "CV地点＝お問い合わせ（タップ込み）／地域＝東京都／25〜65歳／男女／金額は税抜き・グロス", S_NOTE)
    put(root, "A4", "■ADの提案（月）", S_SEC)
    for ref, t in (("A5", "項目"), ("B5", "30万円案"), ("C5", "100万円案")):
        put(root, ref, t, S_HEAD)
    rows = [
        ("配信予算（円）", 300000, 1000000, S_YEN),
        ("表示回数", 90625, 306250, S_INT),
        ("クリック数", 5811, 18030, S_INT),
        ("クリック率", 0.0641, 0.0589, S_PCT),
        ("クリック単価（円）", 52, 55, S_YEN),
        ("CV（ADの想定・タップ込み）", 340, 1000, S_INT),
        ("CVR（クリックに対するCV）", 0.0585, 0.0555, S_PCT),
    ]
    for i, (lab, a, b, st) in enumerate(rows):
        r = 6 + i
        put(root, f"A{r}", lab, S_LABEL)
        put(root, f"B{r}", a, st)
        put(root, f"C{r}", b, st)
    put(root, "A13", "CPA（円）＝予算÷CV", S_LABEL)
    put(root, "B13", "=B6/B11", S_YEN)
    put(root, "C13", "=C6/C11", S_YEN)
    put(root, "A14", "内訳：30万＝meta_オーディエンス25万＋meta_RM5万／100万＝オーディエンス90万＋RM10万（RM＝リマーケティング）", S_NOTE)

    put(root, "A16", "■LINEへの影響（プラン①｜総合運用・月13.5万。LINEの費用はADの有無で変わらない）", S_SEC)
    heads = ["パターン", "指標", "1か月目", "2か月目", "3か月目", "4か月目", "5か月目", "6か月目", "6か月累計", "CPA（6か月目）", "CPA（累計・初期込み）"]
    for j, t in enumerate(heads):
        put(root, f"{'ABCDEFGHIJK'[j]}17", t, S_HEAD)
    blocks = [("AD無し（今のSIM＝ver1.4）", BASE), ("AD30万円案の流入を足す", SC[0][0]), ("AD100万円案の流入を足す", SC[1][0])]
    for k, (lab, sheet) in enumerate(blocks):
        r1, r2 = 18 + 2 * k, 19 + 2 * k
        sh = f"'{sheet}'!"
        put(root, f"A{r1}", lab, S_LABEL)
        put(root, f"A{r2}", "", S_LABEL)
        put(root, f"B{r1}", "新しい友だち（人／月）", S_LABEL)
        put(root, f"B{r2}", "問い合わせ合計（件／月）", S_LABEL)
        for j, col in enumerate("EFGHIJ"):      # 元のシートのE〜J＝1〜6か月目
            put(root, f"{'CDEFGH'[j]}{r1}", f"={sh}{col}26", S_INT)
            put(root, f"{'CDEFGH'[j]}{r2}", f"={sh}{col}40", S_INT)
        put(root, f"I{r1}", f"=SUM(C{r1}:H{r1})", S_INT)
        put(root, f"I{r2}", f"=SUM(C{r2}:H{r2})", S_INT)
        put(root, f"J{r2}", f"={sh}J49", S_YEN)
        put(root, f"K{r2}", f"=({sh}D51+SUM({sh}E51:J51))/I{r2}", S_YEN)

    put(root, "A25", "■ADと合わせた目安（6か月目の月）", S_SEC)
    for ref, t in (("A26", "パターン"), ("B26", "AD費用＋LINE費用（円）"), ("C26", "AD CV＋LINE CV（件）"), ("D26", "合算CPA（円）")):
        put(root, ref, t, S_HEAD)
    for k, (sheet, col, name) in enumerate(SC):
        r = 27 + k
        put(root, f"A{r}", name, S_LABEL)
        put(root, f"B{r}", f"={col}6+'{sheet}'!J51", S_YEN)
        put(root, f"C{r}", f"={col}11+'{sheet}'!J40", S_INT)
        put(root, f"D{r}", f"=B{r}/C{r}", S_YEN)
    put(root, "A30", "※ADのCVは「お問い合わせ（タップ込み）」、LINEのCVは「物件情報の送信」で数え方が違う。ADのCVにLINEボタンのタップが含まれる場合は重なるので、合算は参考値。", S_NOTE)
    put(root, "A31", "※ADのクリック数を、サイトへの流入数としてそのまま足している（UUより多めになる）。広告から来た方の友だち追加率は、FMTの広告用の値"
                     "（離脱防止＝表示65%×8%×8%／サイトのボタン0.25%）で、広告以外（55%×12%×12%／0.5%）の約半分。", S_NOTE)
    put(root, "A32", "※問い合わせ①②は、ver1.4（AD無し）の「新しい友だちに対する率」「クリックに対するCVR」を、そのまま広告流入にも当てている（仮置き）。", S_NOTE)


# ---------------------------------------------------------------- AD流入を足したSIMシート（ver1.4のコピー）
def fill_scenario(root, col, name):
    put(root, "A4", f"株式会社あらうち御中_LINEOA施策提案【SIM】（離脱防止＋サイトのLINEボタン）＋ {name}の広告流入")
    put(root, "H9", f"={A_}{col}8")                 # 広告・アフィUU＝ADのクリック数
    put(root, "H8", f"={B_}H8+H9")                   # UU全体＝今のUU＋広告流入
    for c in "EFGHIJ":
        put(root, f"{c}36", f"=ROUND({c}26*{B_}{c}36/{B_}{c}26,0)")
        put(root, f"{c}37", f"=ROUND({c}35*{B_}{c}37/{B_}{c}35,0)")
    put(root, "N8", f"UU＝今の1,737＋ADの提案（9/14）の月間クリック数（{name}）。ADの数字はシート「{AD}」")
    put(root, "N9", "広告流入。友だち追加率はFMTの広告用の値（離脱防止 表示65%×8%×8%／サイトのボタン0.25%）")
    put(root, "N36", "問い合わせ①＝新規友だち×ver1.4（AD無し）の率（22→57%）。広告流入にも同じ率を当てている（仮置き）")
    put(root, "N37", "問い合わせ②＝配信クリック×ver1.4（AD無し）のCVR（0.3→2.0%）。広告流入にも同じ率を当てている（仮置き）")
    for sv in root.iter(q("sheetView")):
        sv.attrib.pop("tabSelected", None)


def build(src=SRC, out=OUT):
    work = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(src) as z:
            z.extractall(work)
        ws_dir = work / "xl" / "worksheets"
        base = etree.parse(str(ws_dir / "sheet1.xml"))

        new_sheets = []          # (シート名, ファイル名, tree)
        ad = etree.ElementTree(etree.fromstring(AD_XML.encode("utf-8")))
        fill_ad(ad.getroot())
        new_sheets.append((AD, "sheet4.xml", ad))
        for i, (name, col, label) in enumerate(SC):
            s = deepcopy(base)
            fill_scenario(s.getroot(), col, label)
            s.getroot().set(XR_UID, "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"arauchi-sim-v15-{name}")).upper() + "}")
            check_shared(s.getroot(), name)
            strip_cache(s.getroot())
            new_sheets.append((name, f"sheet{5 + i}.xml", s))
        for _, fn, tree in new_sheets:
            tree.write(str(ws_dir / fn), xml_declaration=True, encoding="UTF-8", standalone=True)

        # workbook.xml.rels：新しいシートを足し、calcChain（古い計算順）を外す
        rels_path = work / "xl" / "_rels" / "workbook.xml.rels"
        rels = rels_path.read_text(encoding="utf-8")
        rid = {}
        for k, (name, fn, _) in enumerate(new_sheets):
            rid[fn] = f"rId{101 + k}"
            rels = rels.replace("</Relationships>",
                                f'<Relationship Id="{rid[fn]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{fn}"/></Relationships>')
        rels = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', "", rels)
        rels_path.write_text(rels, encoding="utf-8")
        (work / "xl" / "calcChain.xml").unlink(missing_ok=True)

        wb_path = work / "xl" / "workbook.xml"
        wb = wb_path.read_text(encoding="utf-8")
        add = "".join(f'<sheet name="{name}" sheetId="{4 + k}" r:id="{rid[fn]}"/>' for k, (name, fn, _) in enumerate(new_sheets))
        assert "</sheets>" in wb
        wb = wb.replace("</sheets>", add + "</sheets>")
        wb = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wb)
        wb_path.write_text(wb, encoding="utf-8")

        ct_path = work / "[Content_Types].xml"
        ct = ct_path.read_text(encoding="utf-8")
        ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
        for _, fn, _ in new_sheets:
            ct = ct.replace("</Types>", f'<Override PartName="/xl/worksheets/{fn}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        ct_path.write_text(ct, encoding="utf-8")

        app_path = work / "docProps" / "app.xml"
        if app_path.exists():
            app = app_path.read_text(encoding="utf-8")
            names = re.findall(r"<vt:lpstr>([^<]*)</vt:lpstr>", re.search(r"<TitlesOfParts>.*?</TitlesOfParts>", app, re.S).group(0))
            allnames = names + [n for n, _, _ in new_sheets]
            app = re.sub(r"(<vt:variant><vt:i4>)\d+(</vt:i4>)", rf"\g<1>{len(allnames)}\g<2>", app, count=1)
            titles = "".join(f"<vt:lpstr>{n}</vt:lpstr>" for n in allnames)
            app = re.sub(r"<TitlesOfParts>.*?</TitlesOfParts>",
                         f'<TitlesOfParts><vt:vector size="{len(allnames)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>', app, flags=re.S)
            app_path.write_text(app, encoding="utf-8")

        out = Path(out)
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
    return out


if __name__ == "__main__":
    print("saved:", build())
