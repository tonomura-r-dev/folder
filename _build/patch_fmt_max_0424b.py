# -*- coding: utf-8 -*-
"""提案FMT（MAX版）の直し（2026-10-24 殿村さん・2回目）。

    python -m _build.patch_fmt_max_0424b <FMT.pptx> <サンクスLINE_Appendix_ver2.pptx> <出力.pptx>

  ① S42（その他）：表（項目／確認内容）。S41 の表の図形を写して5行。「数値・内容」の列は入れず、表をページの中央に寄せる
  ② S35（サンクスLINE誘導）：殿村さんのサンクスLINE Appendix ver2 の1枚目（動画のQR・再生リンク付き）に差し替え。
     タイトルは S31〜S34 とそろえて「新規友だちの追加（サンクスLINE誘導）」、金額の札（初期15万円／月額3万円〜＋設定地点追加）を足す
  ③ S39〜S42 の本文（表・下の帯）を14ptに。上の一文（16pt）と小さい注記（8pt）はそのまま
"""
import copy
import sys

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches

from _build.patch_fmt_max_s42 import ROWS, set_text
from _build.patch_tanaka_v5 import copy_slide
from _build.pptx_parts import drop_slide, rect

NAVY, WHITE = "1F2A5C", "FFFFFF"
IN = 2.54
TITLE35 = "新規友だちの追加（サンクスLINE誘導）"
SKIP = {"テキスト ボックス 47", "Rectangle 2"}                     # タイトル・上の一文


def s42_table(s, src):
    by = {sh.name: sh for sh in src.shapes}
    band = next(sh for sh in s.shapes if sh.name == "Rectangle 24")
    left, right = by["Rectangle 3"], by["Rectangle 4"]
    width = right.left + right.width - left.left
    dx = int(band.left + (band.width - width) / 2 - left.left)       # 下の帯の中央に、表の中央を合わせる
    step = Inches(0.665)
    new = [copy.deepcopy(by[f"Rectangle {k}"]._element) for k in (3, 4)]
    for r in range(len(ROWS)):
        for k in (6, 7):
            e = copy.deepcopy(by[f"Rectangle {k}"]._element)
            off = e.find(".//" + qn("a:off"))
            off.set("y", str(int(off.get("y")) + step * r))
            new.append(e)
    for e in new:
        off = e.find(".//" + qn("a:off"))
        off.set("x", str(int(off.get("x")) + dx))
        band._element.addprevious(e)
    shapes = [sh for sh in s.shapes if sh._element in new]
    for sh in shapes[:2]:
        set_text(sh, sh.text_frame.text)
    for r, (item, what) in enumerate(ROWS):
        set_text(shapes[2 + r * 2], item)
        set_text(shapes[3 + r * 2], what)


def body_14(s):
    for sh in s.shapes:
        if sh.name in SKIP:
            continue
        for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
            for r in sh._element.iter(qn(tag)):
                if r.get("sz") and int(r.get("sz")) >= 1000:              # 8pt の注記は変えない
                    r.set("sz", "1400")


BREAKS = {  # 14ptにすると1〜3文字だけ次の行に落ちるセルは、意味の切れ目で改行する（ページ: [(改行の前まで, 全文)]）
    39: [("友だち追加の入口を増やす", "友だち追加の入口を増やす（ポップアップ・広告・店頭）"),
         ("（例：CV数・", "（例：CV数・リピート率）"),
         ("通数費・運用工数の削減", "通数費・運用工数の削減（セグメント・自動化）")],
    41: [("アカウント費・ツール費・運用費・", "アカウント費・ツール費・運用費・通数費の月額合計")],
}
IIKIRI = {  # 下の帯の文末を言い切りに（です・ますを付けない約束。S42 は「です」があると14ptで1行に収まらない）
    39: ("施策の配分と予算の目安に反映します", "施策の配分と予算の目安に反映する"),
    41: ("改善できる範囲を提案に反映します", "改善できる範囲を提案に反映する"),
    42: ("詳細なご提案が可能です。", "詳細なご提案が可能"),
}


def split_br(s, head, whole):
    """whole の段落を、head の後ろで a:br を入れて2行にする（a:r が分かれていたら最初の書式で1つにまとめる）"""
    for p in s._element.iter(qn("a:p")):
        rs = p.findall(qn("a:r"))
        if rs and "".join(r.find(qn("a:t")).text or "" for r in rs) == whole:
            for r in rs[1:]:
                p.remove(r)
            r2 = copy.deepcopy(rs[0])
            rs[0].find(qn("a:t")).text = head
            r2.find(qn("a:t")).text = whole[len(head):]
            br = p.makeelement(qn("a:br"), {})
            br.append(copy.deepcopy(rs[0].find(qn("a:rPr"))))
            rs[0].addnext(br)
            br.addnext(r2)
            return
    raise KeyError(whole)


def iikiri(s, old, new):
    for t in s._element.iter(qn("a:t")):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            return
    raise KeyError(old)


def s35_replace(prs, apx):
    slides = list(prs.slides)
    old = slides[34]
    assert "サンクスLINE誘導" in "".join(sh.text_frame.text for sh in old.shapes if sh.has_text_frame)
    src = list(apx.slides)[0]
    new = copy_slide(prs, src)
    title = next(sh for sh in new.shapes if sh.has_text_frame and sh.text_frame.text == "サンクスLINE｜機能・費用")
    p = title.text_frame.paragraphs[0]
    p.runs[0].text = TITLE35
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)
    rect(new, 4.57 * IN, 1.50 * IN, 1.25 * IN, 0.28 * IN, NAVY, [("初期：15万円", 11, True, WHITE, 0)], align=PP_ALIGN.CENTER,
         margins=(0.05, 0.0, 0.05, 0.0))
    rect(new, 5.90 * IN, 1.50 * IN, 2.30 * IN, 0.28 * IN, NAVY, [("月額3万円〜＋設定地点追加", 11, True, WHITE, 0)],
         align=PP_ALIGN.CENTER, margins=(0.05, 0.0, 0.05, 0.0))
    lst = prs.slides._sldIdLst
    e = list(lst)[-1]
    lst.remove(e)
    list(lst)[34].addprevious(e)                                           # いまの S35 の前へ
    drop_slide(prs, old)


def main(src, apx_path, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    assert len(slides) == 43
    s42_table(slides[41], slides[40])
    for n in (39, 40, 41, 42):
        body_14(slides[n - 1])
    for n, items in BREAKS.items():
        for head, whole in items:
            split_br(slides[n - 1], head, whole)
    for n, (old, new) in IIKIRI.items():
        iikiri(slides[n - 1], old, new)
    s35_replace(prs, Presentation(apx_path))
    assert len(prs.slides) == 43
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
