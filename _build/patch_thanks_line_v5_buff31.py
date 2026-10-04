# -*- coding: utf-8 -*-
"""サンクスLINEのご提案（4枚）1枚目を、DYM共通資料（LINEOA_BUFFF_3）31枚目「新規友だちの追加（サンクスLINE誘導）」に差し替える（2026-10-02 殿村さん指示）
動作画面（フォーム→CV→LINE友だち追加／トーク画面）がそろっているページ。
直した点：費用を 初期15万円 に、社内向けの「2026年7月～の価格改定」の注記を削除、タイトルを「サンクスLINE｜機能」に、字体をメイリオに。
  python3 _build/patch_thanks_line_v5_buff31.py <20261002_サンクスLINEのご提案.pptx> <LINEOA_BUFFF_3.pptx>
"""
import copy
import io
import sys
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Cm

dst_path, src_path = sys.argv[1], sys.argv[2]
prs = Presentation(dst_path)
src = Presentation(src_path).slides[30]
assert "サンクスLINE誘導" in "".join(sh.text_frame.text for sh in src.shapes if sh.has_text_frame)

new = prs.slides.add_slide(prs.slides[1].slide_layout)
for sh in list(new.shapes):
    sh._element.getparent().remove(sh._element)
for el in src.shapes._spTree:
    if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
        continue
    el = copy.deepcopy(el)
    for blip in el.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid:
            img = src.part.related_part(rid)
            _, new_rid = new.part.get_or_add_image_part(io.BytesIO(img.blob))
            blip.set(qn("r:embed"), new_rid)
    new.shapes._spTree.append(el)

# 旧1枚目を外し、新しいページを先頭に
lst = prs.slides._sldIdLst
old = lst[0]
prs.part.drop_rel(old.get(qn("r:id")))
lst.remove(old)
item = lst[-1]
lst.remove(item)
lst.insert(0, item)


def by(name):
    return [sh for sh in prs.slides[0].shapes if sh.name == name]


def set_par(sh, k, text):
    p = sh.text_frame.paragraphs[k]
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)


# タイトル
set_par(by("Google Shape;752;p28")[0], 0, "サンクスLINE｜機能")
# 費用：初期15万円
for sh in by("Google Shape;51;p19"):
    t = sh.text_frame.text
    if "初期" in t:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.text = r.text.replace("10", "15")
# 社内向けの注記（価格改定）を削除し、API併用の注記を上へ
for sh in by("テキスト ボックス 9"):
    sh._element.getparent().remove(sh._element)
for sh in by("テキスト ボックス 13"):
    sh.top = Cm(3.9)
# 実績の文：言い切りにそろえ、良い数字は赤にしない（紺）
from pptx.dml.color import RGBColor
for sh in prs.slides[0].shapes:
    if sh.has_text_frame and "40-50%" in sh.text_frame.text:
        for p in sh.text_frame.paragraphs:
            if "40-50%" in p.text:
                runs = p.runs
                runs[0].text = "美容クリニックでは、予約後来院率が40〜50%改善"
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
                runs[0].font.color.rgb = RGBColor(0x1F, 0x28, 0x5A)
                runs[0].font.bold = True
# 薄いグレーの注記が「CV」の文字と重なるので、左のブロックの中に寄せる
for sh in by("テキスト ボックス 1042"):
    sh.left, sh.top = Cm(1.0), Cm(15.0)

# 図形ID・字体
AFTER = {qn(x) for x in ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")}
sl = prs.slides[0]
for r in sl.shapes._spTree.iter(qn("a:r")):
    if r.find(qn("a:rPr")) is None:
        r.insert(0, r.makeelement(qn("a:rPr"), {"lang": "ja-JP"}))
for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
    for rpr in sl.shapes._spTree.iter(qn(tag)):
        if rpr.get("sz") and int(rpr.get("sz")) < 800:
            rpr.set("sz", "800")
        fs = []
        for ft in ("a:latin", "a:ea", "a:cs"):
            e = rpr.find(qn(ft))
            if e is None:
                e = rpr.makeelement(qn(ft), {})
            else:
                rpr.remove(e)
            e.set("typeface", "メイリオ")
            fs.append(e)
        nxt = next((c for c in rpr if c.tag in AFTER), None)
        for e in fs:
            (nxt.addprevious(e) if nxt is not None else rpr.append(e))
prs.save(dst_path)
print("saved", dst_path, len(prs.slides), "枚")
