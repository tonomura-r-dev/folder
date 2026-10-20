# -*- coding: utf-8 -*-
"""横浜シミズ SIM ver2.1：清水グループ SIM ver1.0（殿村さんのPC保存版）を、20261020の資料（16枚）に合わせて組み直す。

    python _build/build_yokohama_shimizu_sim_v2.py <清水グループSIM ver1.0.xlsx> <出力.xlsx>

  2026-10-20 殿村さんと認識合わせした内容：
  ・LINE公式アカウントは新規開設（今の友だち 0人から）。UU 8,947 は1か月分。月の応募数は不明 → UU×1.0%＝89件（資料と同じ）
  ・友だちの入口＝サイトのLINEボタン（UUの0.5%・FMTの係数）＋離脱防止ポップアップ＋サンクスLINE誘導（応募×遷移80%×追加40%）
    （ver2.0は離脱防止なし → 殿村さん「さすがに離脱防止も入れるか」で ver2.1 に入れた）
  ・CV①＝新規応募（サイトのLINEボタン・離脱防止で友だちになった応募前の方）／CV③＝再応募（一度勤務した方の、次回イベントへの勤務希望）
  ・ブロック率は32%から、2か月目以降は前の月×1.025（式はそのまま）
  ・企画配信 月4回。ステップ配信 4通（あいさつメッセージの後、友だち追加から14日以内。殿村さん「絶対いるでしょ」）
  ・費用は資料の金額（初期構築10万円・月10万円、サンクスLINE誘導 初期15万円・月3万円）。資料の費用は、このSIMで差し替える
  XMLを直接書き換える（チェックボックスを消さない）。シート名の「竹」は外して「SIM」にする。
"""
import datetime
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import check_shared, get_cell, set_b, set_f, set_n, set_s, strip_cache  # noqa: E402

OLD_NAME = "SIM_竹（コンサル基本）"
NEW_NAME = "SIM"
MONTHS = "EFGHIJ"
NEW_CV = [0, 1, 2, 1, 2, 3]       # 新規応募（手入力。合計CVでクリックの0.3〜1.0%前後）
RE_CV = [0, 0, 0, 1, 1, 2]        # 再応募（一度勤務した方から。4か月目以降）


def fill_sim(root):
    base = (datetime.date(2026, 10, 20) - datetime.date(1899, 12, 30)).days
    set_n(root, "B2", base)
    set_s(root, "A4", "株式会社横浜シミズ御中_LINEOA施策提案【SIM】（サイトのLINEボタン＋離脱防止＋サンクスLINE誘導）")
    # CV地点
    set_s(root, "C8", "新規応募")
    set_s(root, "C10", "再応募")
    set_s(root, "N8", "CV①新規応募：サイトのLINEボタン・離脱防止で友だちになった方（応募前）が、配信から応募した数。"
                      "サンクスLINE誘導で友だちになった方は応募済みなので、新規応募には数えない")
    set_n(root, "H10", 89)
    set_s(root, "N10", "サイトからの応募は与件に無いので、仮にUUの1.0%（月89件・資料と同じ）")
    # 施策：離脱防止（FMTの式のまま：UU×表示65%×クリック12.5%×追加12.5%）・サンクスLINE誘導
    for c in MONTHS:
        set_b(root, f"{c}18", True)
        set_f(root, f"{c}24", f"IF({c}19=FALSE,0,{c}16*0.8*0.4)")
    set_s(root, "N24", "サンクスLINE誘導＝応募×遷移80%×追加40%（資料と同じ・定義テキストの係数）")
    set_s(root, "N26", "サイトのLINEボタン（UUの0.5%・FMTの係数）＋離脱防止ポップアップ＋サンクスLINE誘導")
    # ブロック率・配信
    set_n(root, "E27", 0.32)
    for c in MONTHS:
        set_n(root, f"{c}30", 4)
        set_n(root, f"{c}32", 4)
    set_s(root, "N30", "企画配信 月4回（週1回。次回イベントの募集・会場の案内）")
    set_s(root, "N32", "ステップ配信 4通（あいさつメッセージの後、友だち追加から14日以内）")
    # CV（手入力）
    for c, a, b in zip(MONTHS, NEW_CV, RE_CV):
        set_n(root, f"{c}36", a)
        set_n(root, f"{c}38", b)
    set_s(root, "N36", "新規応募：合計CVでクリックの0.3〜1.0%前後になるように置く")
    set_s(root, "N38", "再応募：一度勤務した方が、次回イベントの募集から勤務希望を出した数（応募→勤務を経てからなので4か月目から）")
    # 費用
    set_n(root, "D66", 15000)
    for c in MONTHS:
        set_n(root, f"{c}66", 30000)
    set_s(root, "N66", "離脱防止ポップアップ：初期1.5万円・月3万円")
    set_s(root, "B69", "LINE運用（運用代行・効率改善）")
    for c in "D" + MONTHS:
        set_n(root, f"{c}69", 100000)
    set_s(root, "N69", "LINE運用：資料の金額（初期構築10万円・月10万円）")


def fill_cmp(root):
    for f in root.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f"):
        if f.text and OLD_NAME in f.text:
            f.text = f.text.replace(f"'{OLD_NAME}'!", f"{NEW_NAME}!")
    set_s(root, "A17", "CV①_新規応募")
    set_s(root, "A19", "CV③_再応募")


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    ws = work / "xl/worksheets"
    for fn, fill in (("sheet1.xml", fill_sim), ("sheet2.xml", fill_cmp)):
        tree = etree.parse(str(ws / fn))
        fill(tree.getroot())
        strip_cache(tree.getroot())
        check_shared(tree.getroot(), fn)
        tree.write(str(ws / fn), xml_declaration=True, encoding="UTF-8", standalone=True)
    for rel in ("xl/workbook.xml", "docProps/app.xml"):
        p = work / rel
        t = p.read_text(encoding="utf-8")
        assert OLD_NAME in t, rel
        p.write_text(t.replace(OLD_NAME, NEW_NAME), encoding="utf-8")
    wb = work / "xl/workbook.xml"
    t = wb.read_text(encoding="utf-8")
    if "fullCalcOnLoad" not in t:
        t = t.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.write_text(t, encoding="utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:                                      # 元の順番のまま詰め直す
            z.write(work / n, n)
    shutil.rmtree(work)
    with zipfile.ZipFile(out) as z:
        assert "xl/featurePropertyBag/featurePropertyBag.xml" in z.namelist(), "チェックボックスが消えた"
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
