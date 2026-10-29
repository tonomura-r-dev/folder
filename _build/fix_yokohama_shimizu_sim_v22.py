# -*- coding: utf-8 -*-
"""横浜シミズ SIM ver2.1 → ver2.2：成果報酬版（SIM②）をセントスタッフの形に寄せる（2026-10-29 殿村さん）。

    python _build/fix_yokohama_shimizu_sim_v22.py <ver2.1.xlsx> <ver2.2.xlsx>

  考え方：応募数はコンサル版と同じなので、単価を上げても先方は選ばない。単価は相場の 8,000円 に下げ、その分を応募数で取る
  （コンサル版 月21.5万円 と並ぶのは (21.5万-6.5万)÷8,000円＝月19件ほど）。応募を増やす根拠＝通知メッセージで登録済みのスタッフを友だちにする。
  SIM②（sheet4）：成果単価 D8・D10＝8,000／通知メッセージ_既存リスト E21〜J21＝TRUE／想定顧客リスト K8＝仮に5,000人（2026-10-29 殿村さん。マイナビの求人10件×30人×13回＝年約3,900人＋前からの方）
  （式はFMTのまま＝リスト×15%×15%を毎月）。サイトからの応募 H10＝仮に月150件（SIM①・SIM②とも。サンクスLINE誘導の母数）。CV（36・38行）は殿村さんが決めるので触らない。SIM①は触らない。
  XMLを直接書き換える（チェックボックスを消さない）。
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import get_cell, set_b, set_n, set_s, strip_cache  # noqa: E402

UNIT = 8000
SITE_APPLY = 150
N10 = ("サイトからの応募は、仮に月150件（新規登録は月約300人＝マイナビの求人10件×1件30人×4週ごと。"
       "そのうち自社サイトからの応募を半分と置いた。マイナビ経由はマイナビの完了画面で終わるので、サンクスLINE誘導の母数に入れない）")
LIST = 5000


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    for sheet in ("sheet1.xml", "sheet4.xml"):              # サイトからの応募（H10）は会社の事実なので、SIM①・SIM②の両方に入れる
        sp = work / "xl/worksheets" / sheet
        st = etree.parse(str(sp))
        set_n(st.getroot(), "H10", SITE_APPLY)
        set_s(st.getroot(), "N10", N10)
        strip_cache(st.getroot())
        st.write(str(sp), xml_declaration=True, encoding="UTF-8", standalone=True)
    p = work / "xl/worksheets/sheet4.xml"
    tree = etree.parse(str(p))
    root = tree.getroot()
    assert "$K$8*0.15*0.15" in "".join(t.text or "" for t in get_cell(root, "E25").iter("{*}f"))
    set_n(root, "D8", UNIT)
    set_n(root, "D10", UNIT)
    for c in "EFGHIJ":
        set_b(root, f"{c}21", True)
    set_n(root, "K8", LIST)
    set_s(root, "A4", f"株式会社横浜シミズ御中_LINEOA施策提案【SIM】成果報酬（応募1件 {UNIT:,}円）"
                      "（サイトのLINEボタン＋離脱防止＋サンクスLINE誘導＋通知メッセージ）")
    set_s(root, "N21", f"通知メッセージ（既存リスト）：登録済みのスタッフに配信して友だち追加。リストは仮に{LIST:,}人"
                       "（マイナビの求人10件×1件30人×4週ごと＝年約3,900人の新規登録＋それ以前から続けている方）。"
                       "毎月リスト×15%×15%が友だちに（FMTの式）")
    strip_cache(root)
    tree.write(str(p), xml_declaration=True, encoding="UTF-8", standalone=True)
    for cmp_ in ("sheet2.xml", "sheet5.xml"):
        q = work / "xl/worksheets" / cmp_
        t = etree.parse(str(q))
        strip_cache(t.getroot())
        t.write(str(q), xml_declaration=True, encoding="UTF-8", standalone=True)
    wb = work / "xl/workbook.xml"
    s = wb.read_text(encoding="utf-8")
    if "fullCalcOnLoad" not in s:
        s = s.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.write_text(s, encoding="utf-8")
    tmp = Path(tempfile.mkdtemp()) / "o.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(work / n, n)
    shutil.move(str(tmp), out)
    shutil.rmtree(work)
    with zipfile.ZipFile(out) as z:
        assert any("featurePropertyBag" in n for n in z.namelist()), "チェックボックスが消えた"
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
