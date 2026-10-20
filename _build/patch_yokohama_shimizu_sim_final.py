# -*- coding: utf-8 -*-
"""横浜シミズ SIM：殿村さんがCV数を入れた版の、古くなった説明を今の中身に合わせて仕上げる（2026-10-20 殿村さん「資料より先にSIMを完成してほしい」）。

    python _build/patch_yokohama_shimizu_sim_final.py <殿村さんのSIM.xlsx> [出力.xlsx]

  数字（CV数・費用・係数）は殿村さんの値のまま。文字のセルだけを書き換える（XML直接・チェックボックスを消さない）。
  ・A4（タイトル）に離脱防止を足す
  ・N8・N36（新規応募）：離脱防止で友だちになった方も入る
  ・N26（友だちの増え方）・N66（離脱防止の費用）：「入れていない」→ 入れた中身に
  ・N32（ステップ配信）：「入れていない」→ 4通
  ・N38（再応募）：「4か月目から」→ 殿村さんの数字どおり「3か月目から」
  出力を省くと上書き。
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import set_s  # noqa: E402

TEXTS = {
    "A4": "株式会社横浜シミズ御中_LINEOA施策提案【SIM】（サイトのLINEボタン＋離脱防止＋サンクスLINE誘導）",
    "N8": "CV①新規応募：サイトのLINEボタン・離脱防止で友だちになった方（応募前）が、配信から応募した数。"
          "サンクスLINE誘導で友だちになった方は応募済みなので、新規応募には数えない",
    "N26": "サイトのLINEボタン（UUの0.5%・FMTの係数）＋離脱防止ポップアップ＋サンクスLINE誘導",
    "N32": "ステップ配信 4通（あいさつメッセージの後、友だち追加から14日以内）",
    "N36": "新規応募：サイトのLINEボタン・離脱防止で友だちになった方が、企画配信・ステップ配信から応募した数",
    "N38": "再応募：一度勤務した方が、次回イベントの募集から勤務希望を出した数（応募→勤務を経てからなので3か月目から）",
    "N66": "離脱防止ポップアップ：初期1.5万円・月3万円",
}


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    assert '<sheet name="SIM" sheetId="1" r:id="rId1"/>' in (work / "xl/workbook.xml").read_text(encoding="utf-8")
    sheet = work / "xl/worksheets/sheet1.xml"
    tree = etree.parse(str(sheet))
    for ref, text in TEXTS.items():
        set_s(tree.getroot(), ref, text)
    tree.write(str(sheet), xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = Path(tempfile.mkdtemp()) / "out.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:                                      # 元の順番のまま詰め直す
            z.write(work / n, n)
    shutil.move(str(tmp), out)
    shutil.rmtree(work)
    with zipfile.ZipFile(out) as z:
        assert "xl/featurePropertyBag/featurePropertyBag.xml" in z.namelist(), "チェックボックスが消えた"
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1])
