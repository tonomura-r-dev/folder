# -*- coding: utf-8 -*-
"""タナカホーム「LINE公式アカウント運用のご提案」：本文の文字を、作り直したページ（配信・概算お見積り・料金プラン・運用内容＝12〜15pt）
くらいの大きさにそろえる（2026-10-16 殿村さん「v1の全ページを大きく」）。

    python -m _build.patch_tanaka_pc1 <PC保存版（10/16昼）> <中間.pptx>
    python -m _build.patch_tanaka_font_up <中間.pptx> <出力.pptx>
  （殿村さんのPC保存版v1＝_data/tanaka/PC保存版_20261016b_…v1.pptx とは、現状と打ち手の「別紙」を外したところだけ違っていた。pc1に同じ直しを入れてある）

  対象：表紙・裏表紙・作り直した4ページ・Appendix以外の本編（提案の流れ〜来場前のお知らせ、今後の展開、スケジュール）。
  タイトル・リード（14pt以上）と注記はそのまま。大きくした後にはみ出す図形は、ページごとに位置・高さを直す（LAYOUT）。
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

TARGET = ["提案の流れ", "現状と打ち手", "LINE公式アカウントの現状", "競合の強み", "競合の取り組みに対して", "課題",
          "前後検索「注文住宅」", "解決策：資料請求以前から", "友だちの入口", "来場前のお知らせ", "今後の展開", "スケジュール"]
KEEP_NAMES = ("注記", "画像の説明", "凡例", "分析", "LINE追加", "示唆")             # 前後検索の右の列と帯はpc1で大きさを決めている


def new_size(pt):
    if pt >= 14:
        return pt
    if pt <= 10.5:
        return 12
    if pt <= 11:
        return 12.5
    if pt <= 12:
        return 13
    return 14


def title_of(s):
    for sh in s.shapes:
        if sh.shape_id == 2 and sh.has_text_frame:
            return sh.text_frame.text
    return ""


def bump(el):
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for rpr in el.iter(qn(tag)):
            sz = rpr.get("sz")
            if sz:
                rpr.set("sz", str(int(round(new_size(int(sz) / 100) * 100))))


def main(base, out):
    prs = Presentation(base)
    done = []
    for s in prs.slides:
        t = title_of(s)
        if not any(t.startswith(k) for k in TARGET):
            continue
        for sh in s.shapes:
            if sh.shape_id in (2, 6) or any(k in sh.name for k in KEEP_NAMES):
                continue
            bump(sh._element)
        if t.startswith("今後の展開"):                      # 型の箱：文字が3行になるので高さを広げる
            for k, sh in enumerate(sorted([x for x in s.shapes if x.name.startswith("型")], key=lambda x: x.top)):
                sh.top, sh.height = Inches(2.3 + k * 0.92), Inches(0.85)
        done.append(t[:20])
    prs.save(out)
    print("saved:", out, len(done), "枚", done)


if __name__ == "__main__":
    main(*sys.argv[1:3])
