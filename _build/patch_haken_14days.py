# -*- coding: utf-8 -*-
"""人材業界（派遣・バイト・単発）：ステップ配信を7日間 → 14日間に（2026-10-21 殿村さん「何日目を変えるだけでいい」）。

    python -m _build.patch_haken_14days <殿村さんのPC保存版.pptx> <出力.pptx>

  何日目を2倍にした：Day0／1→2／3→6／5→10／7→14。中身（役割・配信内容・根拠）は変えていない。
  Day2 は根拠の「+2日「派遣 時給 相場」」とちょうど合う。
  画面イメージ（実文面の2枚）の中の「Day3」「Day7」「※ Day3は求人を送りません」は、
  画像を同じ字体（Noto Sans JP）で書き換えた版 _images/haken_day_s24.png / _images/haken_day_s25.png に差し替える。
  ほかのページ・文言は変えない。
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

SUBS = {
    2: [("7日間のステップ配信", "14日間のステップ配信")],
    6: [("短期7日", "短期14日")],
    16: [("追加から7日間", "追加から14日間")],
    20: [("Day1「時給", "Day2「時給")],
    22: [("応募が決まる最初の7日間", "最初の14日間")],
    23: [("7日ステップ", "14日ステップ"),
         ("応募は数日で決まるため、配信は7日間で組む。", "配信は、反応が高い友だち追加後の14日間で組む。"),
         ("Day0とDay7だけ", "Day0とDay14だけ"),
         ("Day3は求人", "Day6は求人"),
         ("Day7の新着", "Day14の新着")],
    24: [("Day0・Day3", "Day0・Day6"),
         ("Day3は求人をお届けしない", "Day6は求人をお届けしない"),
         ("Day3｜", "Day6｜"),
         ("Day7の新着", "Day14の新着")],
    25: [("Day7・契約満了前", "Day14・契約満了前"),
         ("Day7は新着", "Day14は新着"),
         ("Day7｜", "Day14｜")],
}
# 全体像の表の「Day」の列（数字だけの枠）
DAY_CELLS = {"TextBox 13": ("1", "2"), "TextBox 17": ("3", "6"), "TextBox 22": ("5", "10"), "TextBox 26": ("7", "14")}
IMAGES = {24: "_images/haken_day_s24.png", 25: "_images/haken_day_s25.png"}


def replace_in_runs(slide, old, new):
    n = 0
    for p in slide._element.iter(qn("a:p")):
        for t in p.iter(qn("a:t")):
            if t.text and old in t.text:
                t.text = t.text.replace(old, new)
                n += 1
        if n == 0:
            joined = "".join(t.text or "" for t in p.iter(qn("a:t")))
            assert old not in joined, f"文字が複数の書式に分かれている: {old}"
    return n


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    assert len(slides) == 33
    for no, pairs in SUBS.items():
        for old, new in pairs:
            n = replace_in_runs(slides[no - 1], old, new)
            assert n >= 1, (no, old)
    for sh in slides[22].shapes:
        if sh.name in DAY_CELLS:
            old, new = DAY_CELLS[sh.name]
            assert sh.text_frame.text == old, (sh.name, sh.text_frame.text)
            sh.text_frame.paragraphs[0].runs[0].text = new
    for no, path in IMAGES.items():
        pic = [s for s in slides[no - 1].shapes if s.shape_type == 13][0]
        _, rid = slides[no - 1].part.get_or_add_image_part(path)
        pic._element.blipFill.find(qn("a:blip")).set(qn("r:embed"), rid)
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
