# -*- coding: utf-8 -*-
"""すなお ver4：資料全体から黄色・オレンジ・赤枠を外す（2026-10-20 殿村さん「黄色とオレンジはあんま使うなって言ってなかったっけ」「赤枠とかも辞めてほしい」）。

    python -m _build.patch_sunao_v4_colors <ver4.pptx> [出力.pptx]

  FMT（提案FMT_202607）にもともと入っている色も含めて、全ページを紺・白・灰色に寄せる。
  ・黄色・オレンジの塗り：文字が白なら紺、文字が黒っぽい（または文字なし）なら薄い灰色
  ・赤の塗り：薄い赤（明るさ・透明度の指定あり）は同じ明るさの紺、文字が白なら紺、それ以外は薄い灰色
  ・赤・黄色・オレンジの線（枠）：塗りなし・文字なしの四角（強調の赤枠）は消す。それ以外は紺の線に
  ・黄色の蛍光ペン（文字の背景）は外す。赤い文字は紺に、黄色・オレンジの文字は白に
  ・表紙の「貴社ご提案用資料」（ピンクの札）は外す（表紙はタイトル・サブタイトル・御中だけ）
  ・殿村さん「紺か黒か白っていう話をしてて、黄色・オレンジ色は使わないっていう約束は？」→ 緑はLINEの部分だけ残す
  画像の中の色は変えられない。出力を省くと上書き。
"""
import copy
import sys

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

NAVY = "1F2A5C"
LGRAY = "F2F2F2"
WARM = {"FFC000", "FFFF00", "FFFF99", "FBE2D5", "D9A441", "F4B183", "FFE699", "FFD966", "FFF2CC", "F8CBAD", "FCE4D6",
        "F59B21", "ED7D31"}
RED = {"FF0000", "C8102E", "FF3300", "C00000", "E60012"}
WARM_SCHEME = {"accent6"}
RED_SCHEME = {"accent2"}
LIGHT_TEXT = {"FFFFFF", "bg1", "lt1"}
MODS = {"lumMod", "lumOff", "alpha", "tint", "shade", "satMod"}


def kind(clr):
    """色の要素 → 'warm' / 'red' / None"""
    if clr.tag == qn("a:srgbClr"):
        v = clr.get("val", "").upper()
        return "warm" if v in WARM else "red" if v in RED else None
    if clr.tag == qn("a:schemeClr"):
        v = clr.get("val")
        return "warm" if v in WARM_SCHEME else "red" if v in RED_SCHEME else None
    return None


def has_mods(clr):
    return any(c.tag.split("}")[1] in MODS for c in clr)


def recolor(clr, val, keep_mods):
    new = etree.Element(qn("a:srgbClr"), val=val)
    if keep_mods:
        for c in clr:
            new.append(copy.deepcopy(c))
    clr.getparent().replace(clr, new)


def light_text(el):
    """図形・セルの文字の色が白っぽいか（文字が無ければ False）"""
    vals = []
    for run in el.iter(qn("a:r")):                           # 実際の文字の色（lstStyle の既定の色は見ない）
        r = run.find(qn("a:rPr"))
        c = r.find(qn("a:solidFill")) if r is not None else None
        if c is not None and len(c):
            vals.append(c[0].get("val"))
    texts = "".join(t.text or "" for t in el.iter(qn("a:t"))).strip()
    if not texts:
        return False
    if not vals:                                             # 文字の色の指定が無いときは、図形のスタイルの文字色を見る
        ref = el.find(qn("p:style") + "/" + qn("a:fontRef"))
        c = ref[0] if ref is not None and len(ref) else None
        return c is not None and c.get("val") in LIGHT_TEXT
    return sum(v in LIGHT_TEXT for v in vals) * 2 >= len(vals)


def fix_unreadable(s):
    """薄い灰色の塗りに白い文字（読めない）になった札を紺に戻す"""
    n = 0
    for sp in s.shapes._spTree.iter(qn("p:sp")):
        sppr = sp.find(qn("p:spPr"))
        f = sppr.find(qn("a:solidFill")) if sppr is not None else None
        if f is None or not len(f) or f[0].get("val") != LGRAY:
            continue
        if light_text(sp):
            recolor(f[0], NAVY, keep_mods=False)
            n += 1
    return n


def owner(el):
    """色の要素から、それを持つ図形（p:sp・p:pic・p:cxnSp）か表のセル（a:tc）までさかのぼる"""
    e = el
    while e is not None and e.tag not in (qn("p:sp"), qn("p:pic"), qn("p:cxnSp"), qn("a:tc"), qn("p:graphicFrame")):
        e = e.getparent()
    return e


def is_empty_frame(sp):
    """塗りなし・文字なしの四角（強調のための枠）"""
    if sp is None or sp.tag != qn("p:sp"):
        return False
    sppr = sp.find(qn("p:spPr"))
    if sppr is None or sppr.find(qn("a:noFill")) is None:
        return False
    return not "".join(t.text or "" for t in sp.iter(qn("a:t"))).strip()


def fix_slide(s):
    tree = s.shapes._spTree
    stats = dict(fill=0, line=0, frame=0, hl=0, text=0)
    for hl in list(tree.iter(qn("a:highlight"))):
        hl.getparent().remove(hl)
        stats["hl"] += 1
    frames = []
    for clr in list(tree.iter(qn("a:srgbClr"), qn("a:schemeClr"))):
        k = kind(clr)
        if k is None or clr.getparent() is None:
            continue
        par = clr.getparent()
        gp = par.getparent()
        ptag = par.tag.split("}")[1]
        gtag = gp.tag.split("}")[1] if gp is not None else ""
        if ptag == "solidFill" and gtag in ("spPr", "tcPr", "grpSpPr"):
            own = owner(clr)
            lt = light_text(own) if own is not None else False
            if k == "warm":
                recolor(clr, NAVY if lt else LGRAY, keep_mods=lt)
            else:
                if has_mods(clr) or lt:
                    recolor(clr, NAVY, keep_mods=True)
                else:
                    recolor(clr, LGRAY, keep_mods=False)
            stats["fill"] += 1
        elif ptag == "buClr":                                # 箇条書きの記号の色
            recolor(clr, NAVY, keep_mods=False)
            stats["text"] += 1
        elif ptag == "gs":
            recolor(clr, NAVY, keep_mods=True)
            stats["fill"] += 1
        elif ptag == "solidFill" and gtag in ("ln", "lnL", "lnR", "lnT", "lnB"):
            own = owner(clr)
            if gtag == "ln" and is_empty_frame(own):
                frames.append(own)
            else:
                recolor(clr, NAVY, keep_mods=True)
                stats["line"] += 1
        elif ptag == "solidFill" and gtag in ("rPr", "defRPr", "endParaRPr"):
            v = clr.get("val", "").upper()
            if k == "red":                                   # 赤い文字も紺に（紺・黒・白・灰色だけで組む）
                recolor(clr, NAVY, keep_mods=False)
                stats["text"] += 1
            elif k == "warm":
                recolor(clr, "FFFFFF", keep_mods=False)
                stats["text"] += 1
    for sp in frames:
        if sp.getparent() is not None:
            sp.getparent().remove(sp)
            stats["frame"] += 1
    return stats


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    cover = slides[0]
    for sh in list(cover.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip() == "貴社ご提案用資料":
            sh._element.getparent().remove(sh._element)
    total = dict(fill=0, line=0, frame=0, hl=0, text=0)
    for s in slides:
        for k, v in fix_slide(s).items():
            total[k] += v
        total["fixed"] = total.get("fixed", 0) + fix_unreadable(s)
    prs.save(out)
    print("saved:", out, total)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1])
