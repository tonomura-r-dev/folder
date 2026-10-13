# -*- coding: utf-8 -*-
"""v6 の検品：既存ページが変わっていないこと、プラン④の必須項目、数字、禁止語。

    python -m _build.arauchi_v6.checks <v5.pptx> <v6.pptx>
"""
import sys
import zipfile

from pptx import Presentation

from _build.arauchi_v4.build import c14n, slide_order

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def page_text(slide):
    return "\n".join("".join((t.text or "") for t in p.iter(A + "t")) for p in slide._element.iter(A + "p"))


def main(v5, v6):
    ok = True

    def check(cond, msg):
        nonlocal ok
        print(("OK  " if cond else "NG  ") + msg)
        ok = ok and cond

    z5, z6 = zipfile.ZipFile(v5), zipfile.ZipFile(v6)
    o5, o6 = slide_order(z5), slide_order(z6)
    check(len(o5) == 46 and len(o6) == 48, f"枚数 v5={len(o5)} → v6={len(o6)}（2ページ追加）")
    same, diff = [], []
    for i, n6 in enumerate(o6, 1):
        if i in (25, 26):
            continue
        j = i if i < 25 else i - 2
        (same if c14n(z5.read(o5[j - 1])) == c14n(z6.read(n6)) else diff).append(i)
    check(diff == [24], f"v5と中身が違う既存ページ＝{diff}（期待＝[24]だけ。ほかの{len(same)}枚は同じ）")

    prs = Presentation(v6)
    t = {i + 1: page_text(s) for i, s in enumerate(prs.slides)}
    p24, p25, p26 = t[24], t[25], t[26]
    for k in ("プラン①", "プラン②", "プラン③", "プラン④"):
        check(all(k in x for x in (p24, p25)), f"p24・p25に『{k}』")
    check(p24.count("21.5万円") >= 5, "p24：初期費用21.5万円（4プラン分＋注記）")
    for s_ in ("13.5万円", "8.5万円", "ツール利用料", "手数料", "8,000円×", "有効CV件数", "成果単価", "コンサル型", "成果報酬型", "別契約"):
        check(s_ in p24, f"p24：『{s_}』")
    check("固定のコンサル費と成果報酬を併せて請求することはありません" in p24, "p24：二重請求しない旨")
    check("確認中" in p24 and "確定次第ご案内" in p24, "p24：フォームの技術要件・ツール費用は確認中（確定次第ご案内）")
    for s_ in ("215,000円", "8,000円／有効CV", "有効CV件数 × 8,000円", "23件（仮）", "184,000円（試算）", "新規・既存", "専用フォーム",
               "物件情報", "重複・テスト送信", "無効CV", "成果保証ではありません", "実際の有効CV件数に応じた", "確認中", "別契約", "CV地点",
               "STEP 1", "STEP 2", "STEP 3", "STEP 4", "月額固定コンサル費", "初期費用"):
        check(s_ in p26, f"p26：『{s_}』")
    check(23 * 8000 == 184000, "計算：23件 × 8,000円 ＝ 184,000円")
    check("成果対象" in p26 or "成果の対象" in p26, "p26：成果の対象（新規・既存の両方）の明記")
    mine = p24 + p25 + p26
    for bad in ("松竹梅", "3つのプラン", "3プラン", "予算", "ヒアリング", "CPA", "スコープ", "∞"):
        check(bad not in mine, f"p24〜26に『{bad}』が無い")
    check("プラン①｜総合運用の基本施策を対象とした試算" in t[13] and "184,000" not in t[13], "p13（費用対効果）：プラン①の試算のまま（184,000円などに置き換えていない）")
    # プラン③の既存の値
    for s_ in ("月4〜5本", "月1〜2本", "10通", "初期導入支援", "限定的", "必要に応じ別途見積"):
        check(s_ in p25, f"p25：既存の運用内容『{s_}』")
    print("\n総合:", "全部OK" if ok else "NGあり")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:3]))
