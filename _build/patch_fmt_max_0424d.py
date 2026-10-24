# -*- coding: utf-8 -*-
"""提案FMT（MAX版）の直し（2026-10-24 殿村さん・4回目）。

    python -m _build.patch_fmt_max_0424d <FMT_20261024c.pptx> <出力.pptx>

  ① S16：5枚目のカードが「Day8〜9／他社比較・選び方」で4枚目と重複 → 「Day12〜15／相談オファー」（下の説明文の Day12〜15 に合わせる）。
     転職の例なのに買取の言葉（査定）が残っていたので、下の説明文「査定オファー」→「相談オファー」、3枚目「査定〜面談」→「応募〜面接」
  ② 呼び方を「ポップアップ」にそろえる（S12「ポップアップバナー」、S30「離脱防止バナー」、S31・S32「ポップアップ / 離脱防止」「ポップアップ(離脱防止 など)」）
  S35 は前回のまま（殿村さん）。ほかのページは変えない。
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches


def paras(s):
    for p in s._element.iter(qn("a:p")):
        rs = p.findall(qn("a:r"))
        if rs:
            yield p, rs, "".join(r.find(qn("a:t")).text or "" for r in rs)


def whole(s, old, new, count=None):
    """段落の文字全体が old のものを new に（a:r が分かれていたら最初の書式で1つに）"""
    n = 0
    for p, rs, t in paras(s):
        if t == old:
            for r in rs[1:]:
                p.remove(r)
            rs[0].find(qn("a:t")).text = new
            n += 1
    assert n and (count is None or n == count), (old, n)


def part(s, old, new):
    """a:t の中の old を new に（1か所）"""
    for t in s._element.iter(qn("a:t")):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            return
    raise KeyError(old)


def card5(s):
    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                yield from walk(sh.shapes)
            elif sh.has_text_frame:
                yield sh
    hit = 0
    for sh in walk(s.shapes):
        if abs(sh.left - Inches(8.57)) < Inches(0.03):
            t = sh.text_frame.text
            if t in ("Day8〜9", "他社比較・選び方"):
                p = sh.text_frame.paragraphs[0]
                p.runs[0].text = "Day12〜15" if t == "Day8〜9" else "相談オファー"
                for r in p.runs[1:]:
                    r._r.getparent().remove(r._r)
                hit += 1
    assert hit == 2, hit


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    assert len(slides) == 43
    s16 = slides[15]
    card5(s16)
    part(s16, "査定オファ", "相談オファ")
    part(s16, "査定〜面談の流れを", "応募〜面接の流れを")
    whole(slides[11], "ポップアップバナー", "ポップアップ", 1)
    whole(slides[29], "離脱防止バナー", "ポップアップ", 1)
    for n in (31, 32):
        whole(slides[n - 1], "新規友だちの追加（ポップアップ / 離脱防止）", "新規友だちの追加（ポップアップ）", 1)
        whole(slides[n - 1], "ポップアップ(離脱防止 など)", "ポップアップ", 1)
    for no, s in enumerate(slides, 1):
        t = "".join(x.text or "" for x in s._element.iter(qn("a:t")))
        assert "離脱防止" not in t and "ポップアップバナー" not in t, no
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
