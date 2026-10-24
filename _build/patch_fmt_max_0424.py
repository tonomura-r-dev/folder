# -*- coding: utf-8 -*-
"""提案FMT（MAX版）殿村さん作業中の版への直し（2026-10-24）。

    python -m _build.patch_fmt_max_0424 <殿村さんPC保存版.pptx> <出力.pptx>

  ① S39〜S42 の上の一文（要旨）：13pt → 16pt、1.40 → 1.05インチに上げる（塾の新FMT版 P53〜P56 と同じ直し）
  ② S40 の一文が、S42（その他）用の文になっていたので元の文に戻し、S42 にその文を入れる（S42 は S41 と同じ文だった）
  ③ S16 の一文：1行目が枠からはみ出して3行になるので、改行の位置だけ「段階ごとに」の後ろへ（言葉は変えない）
  ほかのページは変えない。
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

S40 = ("配信の体制や社内の確認の流れなど、運用の進め方に関わる確認事項", "流入量・CV・CV地点から、LINEで拾えるユーザーの規模を見積もる")
S42 = ("現状の規模・活用度・費用から、改善余地を見極める", "配信の体制や社内の確認の流れなど、運用の進め方に関わる確認事項")
S16 = ["検討が始まったタイミング（業界キーワードの検索）から段階ごとに",
       "不安を解消し、信頼を積み上げて15日以内に、CVへつなげる配信を構築。"]


def set_text(p, text):
    runs = p.findall(qn("a:r"))
    for r in runs[1:]:
        p.remove(r)
    runs[0].find(qn("a:t")).text = text


def lead(s):
    return next(sh for sh in s.shapes if sh.name == "Rectangle 2" and sh.has_text_frame)


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    assert len(slides) == 43
    for no, (old, new) in ((40, S40), (42, S42)):
        sh = lead(slides[no - 1])
        assert sh.text_frame.text == old, (no, sh.text_frame.text)
        set_text(sh.text_frame.paragraphs[0]._p, new)
    for no in (39, 40, 41, 42):
        sh = lead(slides[no - 1])
        for p in sh.text_frame._txBody.iter(qn("a:p")):
            for r in p.findall(qn("a:r")):
                r.find(qn("a:rPr")).set("sz", "1600")
            end = p.find(qn("a:endParaRPr"))
            if end is not None:
                end.set("sz", "1600")
        sh.top = Inches(1.05)
    sh = next(x for x in slides[15].shapes if x.has_text_frame and x.text_frame.text.startswith("検討が始まった"))
    ps = sh.text_frame._txBody.findall(qn("a:p"))
    assert len(ps) == 2 and "".join(S16) == sh.text_frame.text.replace("\n", "").replace("\x0b", "")
    for p, t in zip(ps, S16):
        set_text(p, t)
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
