# -*- coding: utf-8 -*-
"""提案FMT（MAX版）の直し（2026-10-24 殿村さん・3回目）。

    python -m _build.patch_fmt_max_0424c <FMT_20261024b.pptx> <出力.pptx>

  ① S39〜S41：表の一番右の列（S39＝度合い・目標、S40・S41＝数値・内容＝記入欄）を消し、S42 と同じく表をページの中央に寄せる
  ② S35：赤い文字の注記（※既存APIツールとの併用…）を消し、金額の札を少し上に（1.50 → 1.33インチ）
"""
import sys

from pptx import Presentation
from pptx.util import Inches

HEAD = "1F2A5C"


def fill_of(sh):
    try:
        return str(sh.fill.fore_color.rgb) if sh.fill.type == 1 else None
    except Exception:
        return None


def drop_right_and_center(s):
    heads = [sh for sh in s.shapes if fill_of(sh) == HEAD and sh.top < Inches(2.1)]
    right = max(heads, key=lambda sh: sh.left)
    top = min(h.top for h in heads)
    cells = [sh for sh in s.shapes if sh.top >= top - Inches(0.02) and sh.width < Inches(9) and fill_of(sh)]
    gone = [sh for sh in cells if abs(sh.left - right.left) < Inches(0.02)]
    for sh in gone:
        sh._element.getparent().remove(sh._element)
    keep = [sh for sh in cells if sh not in gone]
    left = min(sh.left for sh in keep)
    width = max(sh.left + sh.width for sh in keep) - left
    center = Inches(0.42) + Inches(10.0) / 2                       # 下の帯（0.42〜10.42）の中央
    dx = int(center - width / 2 - left)
    for sh in keep:
        sh.left = sh.left + dx
    return len(gone), dx / 914400


def s35(s):
    note = next(sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith("※既存APIツールとの併用"))
    note._element.getparent().remove(note._element)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text in ("初期：15万円", "月額3万円〜＋設定地点追加"):
            sh.top = Inches(1.33)


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    assert len(slides) == 43
    for n in (39, 40, 41):
        print(n, drop_right_and_center(slides[n - 1]))
    s35(slides[34])
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
