# -*- coding: utf-8 -*-
"""すなおコンサルティング SIM ver1.3：殿村さんのPC保存版（ver1.2・リピートCV入り）の成果報酬の固定費を直す（2026-10-20）。

    python _build/fix_sunao_sim_v13.py <殿村さんのSIM ver1.2.xlsx> <出力.xlsx>

  ・45行（成果報酬プランの固定費）：FMTの式「SUM(65:69)+SUM(71:72)」は、69行＝初期構築費（月額0）の前提。
    すなおでは69行を「LINE運用コンサル（月額30万／10万）」に使っているので、成果報酬プランの月額にコンサル費が入っていた。
    → E45:J45 を「SUM(65:68)+SUM(71:72)」に（月額＝ツール費＋アカウント費。コンサルの月額は成果報酬に置き換わる）。初期（D45）はそのまま。
  ・成果報酬はセミナー予約（新規・リピート）だけ。サロン入会は0円（計測できないので成果報酬の対象外。資料内のCVとしては数える）。
  ・N列のメモと、比較シートのCV③の名前を今の形に合わせる。
  XMLを直接書き換える（チェックボックスを消さない）。殿村さんが入れた数字・名前は触らない。
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sim_xml import set_f, set_s  # noqa: E402

COLS = "EFGHIJ"


def fill_sim(root):
    set_s(root, "A4", "株式会社すなおコンサルティング御中_LINEOA施策提案【SIM】ver1.3（ポップアップツール＋サンクスLINE誘導）")
    for c in COLS:
        set_f(root, f"{c}45", f"SUM({c}65:{c}68)+SUM({c}71:{c}72)")
    notes = {
        "N9": "CV②サロン入会：成果報酬の対象外（入会はLINEから計測できないため。成果単価0円）。資料内のCVとしては数える",
        "N10": "イベント月（1・4か月目）：LP UU 仮に3,000（広告クリック 約2,760＋自社のSNS等が少し）・セミナー予約 80件（前回1回分の実績）。広告費はADの予算なので入れていない",
        "N37": "サロン入会＝資料内のCVとして数える（成果報酬の対象外）",
        "N38": "リピート予約＝2回目以降のイベントのセミナー予約（LINEのリンク経由・過去に予約がある方）。成果単価1.2万円",
        "N45": "成果報酬プランの月額＝ツール費＋アカウント費（65〜68・71〜72行）。69行のLINE運用コンサルの月額は成果報酬に置き換わるので入れない。初期は全部入れる",
        "N50": "成果報酬＝新規のセミナー予約（36行）＋リピート予約（38行）×1.2万円。サロン入会（37行）は入れない",
    }
    for ref, text in notes.items():
        set_s(root, ref, text)


def fill_cmp(root):
    set_s(root, "A19", "CV③_セミナー予約（リピート）")


def main(src, out):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    for sheet, fn in (("sheet1.xml", fill_sim), ("sheet2.xml", fill_cmp)):
        p = work / "xl/worksheets" / sheet
        tree = etree.parse(str(p))
        fn(tree.getroot())
        tree.write(str(p), xml_declaration=True, encoding="UTF-8", standalone=True)
    wb = work / "xl/workbook.xml"
    t = wb.read_text(encoding="utf-8")
    if "fullCalcOnLoad" not in t:
        t = t.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.write_text(t, encoding="utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(work / n, n)
    shutil.rmtree(work)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
