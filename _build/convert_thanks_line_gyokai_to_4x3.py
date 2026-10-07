# -*- coding: utf-8 -*-
"""サンクスLINE誘導・業界別（16:9＝33.87×19.05cm）を、やる気スイッチ資料と同じ 4:3（27.52×19.05cm）に変換する（2026-10-07 殿村さん指示）。
やる気スイッチのPPTX（4:3のDYMマスター）を器にして、業界別の6枚の図形を横方向だけ 27.52/33.87≒0.8125 倍に縮めて移す。
  - 図形・テキスト枠・線：x と幅を0.8125倍（高さ・yはそのまま）。文字サイズも同じ倍率（最小7pt）＝行の折り返しが変わらない。
  - 円（w=h）は円のまま（中心だけ移動）。
  - 画像：縦横比を保って0.8125倍に縮め、元の位置の中央に置く。画像の上に乗っている小さな図形（伏せ字パッチ・赤枠・「ロゴ」）は
    その画像と同じ変換をかけて、ずれないようにする。
  - タイトル枠は、やる気スイッチと同じ位置・幅（1.52, 0.38, 18.1, 0.94）・16ptのまま。
  python3 _build/convert_thanks_line_gyokai_to_4x3.py <業界別16:9.pptx> <やる気スイッチ4:3.pptx> [出力]
"""
import copy
import io
import re
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Cm, Emu, Pt

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1])
SHELL = Path(sys.argv[2])
OUT = Path(sys.argv[3]) if len(sys.argv) > 3 else ROOT / "20261007_サンクスLINE誘導のご提案_業界別ver3.7.pptx"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
MIN_PT = 7.0

src = Presentation(str(SRC))
shell = Presentation(str(SHELL))
W0, W1 = src.slide_width, shell.slide_width
assert src.slide_height == shell.slide_height, "高さが違う"
F = W1 / W0
print(f"幅 {W0/360000:.2f} → {W1/360000:.2f} cm（×{F:.4f}）")

# 器：やる気スイッチの全スライドを落とす（マスター・レイアウトだけ使う）
lst = shell.slides._sldIdLst
for el in list(lst):
    shell.part.drop_rel(el.rId)
    lst.remove(el)
layout = next(l for l in shell.slide_layouts if l.name == "4_タイトルとコンテンツ")


def xfrm_of(el):
    """図形の p:spPr/a:xfrm（グループは p:grpSpPr/a:xfrm、画像は p:spPr/a:xfrm、線も spPr）"""
    for path in ("p:spPr/a:xfrm", "p:grpSpPr/a:xfrm", "p:xfrm"):
        x = el.find(path, NS)
        if x is not None:
            return x
    return None


def geo(el):
    x = xfrm_of(el)
    off, ext = x.find("a:off", NS), x.find("a:ext", NS)
    return int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))


def set_geo(el, x, y, w, h):
    xf = xfrm_of(el)
    off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
    off.set("x", str(int(round(x))))
    off.set("y", str(int(round(y))))
    ext.set("cx", str(int(round(w))))
    ext.set("cy", str(int(round(h))))


def is_pic(el):
    return etree.QName(el).localname == "pic"


def scale_fonts(el, min_pt=MIN_PT):
    has_sz = False
    for tag in ("rPr", "endParaRPr", "defRPr"):
        for r in el.iter("{%s}%s" % (NS["a"], tag)):
            sz = r.get("sz")
            if sz:
                has_sz = True
                new = max(min_pt * 100, int(int(sz) * F / 50) * 50)  # 0.5pt刻みで切り捨て（折り返しを増やさない）
                r.set("sz", str(int(new)))
    if not has_sz:  # サイズ指定なし＝既定18pt。縮めた枠に合わせて明示する
        for r in el.iter("{%s}rPr" % NS["a"]):
            r.set("sz", str(int(int(1800 * F / 50) * 50)))
    # 文字枠の左右の余白も同じ倍率（未指定＝0.25cm）
    for bp in el.iter("{%s}bodyPr" % NS["a"]):
        for k in ("lIns", "rIns"):
            v = int(bp.get(k) or 91440)
            bp.set(k, str(int(v * F)))


for si, s in enumerate(src.slides):
    ns = shell.slides.add_slide(layout)
    for sh in list(ns.shapes):
        sh._element.getparent().remove(sh._element)
    tree = ns.shapes._spTree
    elems = [el for el in s.shapes._spTree if etree.QName(el).localname in ("sp", "pic", "grpSp", "cxnSp", "graphicFrame")]
    # 画像の変換（縦横比を保って縮小・元の枠の中央）を先に決める
    pics = []
    for el in elems:
        if is_pic(el):
            x, y, w, h = geo(el)
            nw, nh = w * F, h * F
            nx, ny = x * F + (w * F - nw) / 2, y + (h - nh) / 2
            pics.append((el, (x, y, w, h), (nx, ny, nw, nh)))

    def container_of(el):
        """この図形の中心が、どれかの（自分より大きい）画像の中に入っていればその画像の変換を返す"""
        x, y, w, h = geo(el)
        cx, cy = x + w / 2, y + h / 2
        best = None
        for pel, (px, py, pw, ph), new in pics:
            if pel is el or pw * ph <= w * h:
                continue
            if px <= cx <= px + pw and py <= cy <= py + ph:
                if best is None or pw * ph < best[0][2] * best[0][3]:
                    best = ((px, py, pw, ph), new)
        return best

    for el in elems:
        new = copy.deepcopy(el)
        if is_pic(new):
            # 画像の実体を新しいスライドに持ち込み、rId を付け替える
            blip = new.find(".//a:blip", NS)
            rid = blip.get("{%s}embed" % NS["r"])
            image_part = s.part.related_part(rid)
            _, new_rid = ns.part.get_or_add_image_part(io.BytesIO(image_part.blob))
            blip.set("{%s}embed" % NS["r"], new_rid)
        tree.append(new)
        x, y, w, h = geo(new)
        cont = container_of(el)
        if cont is not None:
            (px, py, pw, ph), (nx, ny, nw, nh) = cont
            k = nw / pw
            set_geo(new, nx + (x - px) * k, ny + (y - py) * k, w * k, h * k)
        elif is_pic(new):
            set_geo(new, *next(n for pel, _, n in pics if pel is el))
        else:
            name = new.find(".//p:cNvPr", NS).get("name")
            geom = new.find(".//a:prstGeom", NS)
            is_circle = geom is not None and geom.get("prst") == "ellipse" and abs(w - h) < Cm(0.02)
            if y < Cm(1.0) and (name == "TextBox 1" or (si == 1 and name.startswith("Google Shape;752"))):
                set_geo(new, Cm(1.52), Cm(0.38), Cm(18.1), Cm(0.94))  # タイトル枠はやる気スイッチと同じ位置・幅（2枚目も他と揃える）
                continue  # 文字サイズも据え置き（16pt）
            txt = "".join(t.text or "" for t in new.iter("{%s}t" % NS["a"]))
            if si == 2 and re.fullmatch(r"-?\d+日|0日\(検索起点\)", txt):  # 前後検索の目盛り：6ptに下げ、起点は2行に
                if txt.startswith("0日"):
                    ps = new.findall(".//a:p", NS)
                    p2 = copy.deepcopy(ps[0])
                    ps[0].find(".//a:t", NS).text = "0日"
                    p2.find(".//a:t", NS).text = "(検索起点)"
                    ps[0].addnext(p2)
                set_geo(new, x * F, y, w * F, h * 2)
                scale_fonts(new, min_pt=6.0)
                for r in new.iter("{%s}rPr" % NS["a"]):
                    r.set("sz", "600")
                continue
            if is_circle:
                set_geo(new, (x + w / 2) * F - h / 2, y, h, h)
            else:
                set_geo(new, x * F, y, w * F, h)
        scale_fonts(new)
    print(f"slide {si+1}: {len(elems)} shapes（画像 {len(pics)}）")

shell.save(str(OUT))
print("saved:", OUT.name, len(shell.slides), "枚")
