# -*- coding: utf-8 -*-
"""サンクスLINEのご提案（4枚：機能／重要性／シーン1／シーン2）  2026-10-02
文面は `_drafts/サンクスLINE_資料化テキスト（2026-10-01）.md`（殿村さん確定版）のとおり。
土台は ver1.4 の5枚（先頭4枚の中身を作り直し、5枚目を削除）。
画像は使わない（殿村さん指示 2026-10-02）。すべて図形で図解
  python3 _build/build_thanks_line_v3.py <20260930_サンクスLINEのご提案ver1.4.pptx>
"""
import copy
import sys
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
IMG = ROOT / "_images"
OUT = ROOT / "20261002_サンクスLINEのご提案.pptx"
NAVY, INK, GRAY, WHITE = RGBColor(0x1F, 0x28, 0x5A), RGBColor(0x33, 0x33, 0x33), RGBColor(0x7F, 0x7F, 0x7F), RGBColor(0xFF, 0xFF, 0xFF)
TITLE = RGBColor(0x00, 0x20, 0x60)
GREENTXT = RGBColor(0x0B, 0x7A, 0x3B)

prs = Presentation(sys.argv[1])
SW = prs.slide_width / 360000


def shp(si, name):
    return next(x for x in prs.slides[si].shapes if x.name == name)


# ---- 図形の見本（ver1.4から複製して使う）----
TPL = {
    "card": copy.deepcopy(shp(1, "Rounded Rectangle 4")._element),        # 薄い青・中央
    "bullet": copy.deepcopy(shp(1, "Rounded Rectangle 12")._element),     # 薄い青・紺枠・左揃え
    "green": copy.deepcopy(shp(2, "Rounded Rectangle 8")._element),       # LINEの緑
    "navy": copy.deepcopy(shp(2, "Rounded Rectangle 4")._element),        # 紺
    "lgreen": copy.deepcopy(shp(1, "Rounded Rectangle 10")._element),     # 薄い緑・緑枠
    "arrow_navy": copy.deepcopy(shp(1, "Right Arrow 5")._element),
    "arrow_green": copy.deepcopy(shp(1, "Right Arrow 7")._element),
    "note": copy.deepcopy(shp(3, "TextBox 19")._element),                 # 出典の注記
}


def set_runs(sh, lines, align=None):
    """lines = [(text, size, bold, color), ...] 1要素＝1段落。書式は元の最初のrunを引き継ぐ。"""
    tf = sh.text_frame
    base_p = tf.paragraphs[0]._p
    base_r = base_p.find(qn("a:r"))
    rpr0 = copy.deepcopy(base_r.find(qn("a:rPr"))) if base_r is not None else None
    for p in list(tf._txBody.findall(qn("a:p")))[1:]:
        tf._txBody.remove(p)
    for r in list(base_p.findall(qn("a:r"))) + list(base_p.findall(qn("a:br"))):
        base_p.remove(r)
    empty_p = copy.deepcopy(base_p)
    for i, (text, size, bold, color) in enumerate(lines):
        p = base_p if i == 0 else copy.deepcopy(empty_p)
        if i:
            tf._txBody.append(p)
        r = p.makeelement(qn("a:r"), {})
        if rpr0 is not None:
            r.append(copy.deepcopy(rpr0))
        t = r.makeelement(qn("a:t"), {})
        t.text = text
        r.append(t)
        end = p.find(qn("a:endParaRPr"))
        (end.addprevious(r) if end is not None else p.append(r))
    for para, (text, size, bold, color) in zip(tf.paragraphs, lines):
        if align is not None:
            para.alignment = align
        for run in para.runs:
            run.font.size, run.font.bold = Pt(size), bold
            run.font.color.rgb = color


def add(s, kind, x, y, w, h, lines=None, align=None):
    el = copy.deepcopy(TPL[kind])
    s.shapes._spTree.append(el)
    sh = s.shapes[-1]
    sh.left, sh.top, sh.width, sh.height = Cm(x), Cm(y), Cm(w), Cm(h)
    if lines:
        set_runs(sh, lines, align)
    return sh


def source(s, y, text):
    sh = add(s, "note", 1.46, y, 24.6, 0.6)
    set_runs(sh, [(text, 10, False, GRAY)])


def head(si, title, lead):
    s = prs.slides[si]
    for sh in list(s.shapes):
        if sh.name not in ("TextBox 1", "TextBox 2", "Connector 3"):
            sh._element.getparent().remove(sh._element)
    set_runs(shp(si, "TextBox 1"), [(title, 16, True, TITLE)])
    set_runs(shp(si, "TextBox 2"), [(lead, 16, False, INK)])
    t2 = shp(si, "TextBox 2")
    t2.top, t2.height = Cm(1.9), Cm(1.2)
    return s


def picture(s, name, x_w, y, crop_bottom=0.0):
    W = x_w
    H = W * 941 * (1 - crop_bottom) / 1672
    pic = s.shapes.add_picture(str(IMG / name), Cm((SW - W) / 2), Cm(y), Cm(W), Cm(H))
    if crop_bottom:
        pic.crop_bottom = crop_bottom
    return pic, H


# ================= 1枚目：機能 =================
s = head(0, "サンクスLINE｜機能", "申込み完了画面から、LINEへ移行")
Y, H = 4.5, 4.8
add(s, "card", 1.46, Y, 5.8, H, [("申込み完了画面", 20, True, NAVY)])
add(s, "arrow_navy", 7.4, Y + H / 2 - 0.6, 1.0, 1.2)
add(s, "green", 8.55, Y, 5.4, H, [("LINEへ移行", 20, True, WHITE)])
add(s, "arrow_green", 14.1, Y + H / 2 - 0.6, 1.0, 1.2)
add(s, "lgreen", 15.25, Y, 10.81, (H - 0.3) / 2, [("新規 → 友だち追加", 17, True, GREENTXT)])
add(s, "lgreen", 15.25, Y + (H + 0.3) / 2, 10.81, (H - 0.3) / 2, [("既存の友だち → トーク画面", 17, True, GREENTXT)])
add(s, "bullet", 1.46, 9.8, 12.2, 3.9, [("入力内容を引き継ぎ、配信の出し分けに活用", 16, True, NAVY)], align=2)
add(s, "bullet", 13.86, 9.8, 12.2, 3.9, [("今のアカウントにそのまま導入、", 16, True, NAVY), ("APIツールとも併用可", 16, True, NAVY)], align=2)
add(s, "navy", 1.46, 14.1, 24.6, 2.6, [("費用：初期15万円／月額3万円〜", 24, True, WHITE)])

# ================= 2枚目：重要性（図形だけで図解）=================
s = head(1, "サンクスLINE｜重要性", "申込みの直後に、LINEでつながる")
BX, BW_ = 6.9, 19.16                                   # 棒グラフの左端・全幅
add(s, "card", 1.46, 4.5, 5.2, 2.3, [("申込み・予約", 16, True, NAVY)])
add(s, "navy", BX, 4.5, BW_, 2.3, [("予約 100件", 18, True, WHITE)])
add(s, "card", 1.46, 7.1, 5.2, 2.3, [("来店", 16, True, NAVY)])
add(s, "green", BX, 7.1, BW_ * 0.75, 2.3, [("来店", 18, True, WHITE)])
add(s, "card", BX + BW_ * 0.75 + 0.15, 7.1, BW_ * 0.25 - 0.15, 2.3, [("来店せず", 13, True, GRAY), ("2〜3割", 16, True, GRAY)])
CW, GAP, CY, CH = 7.4, 1.2, 10.0, 5.2
add(s, "card", 1.46, CY, CW, CH, [("申込み⇒来店までに", 15, True, NAVY), ("一定数が離脱する", 15, True, NAVY), ("（予約の2〜3割が来店せず／", 11, False, INK), ("保険見直し本舗）", 11, False, INK)])
add(s, "arrow_navy", 1.46 + CW + 0.1, CY + CH / 2 - 0.6, 1.0, 1.2)
add(s, "card", 1.46 + CW + GAP, CY, CW, CH, [("広告費を増やしても、", 15, True, NAVY), ("この離脱層は減らない", 15, True, NAVY)])
add(s, "arrow_green", 1.46 + 2 * CW + GAP + 0.1, CY + CH / 2 - 0.6, 1.0, 1.2)
add(s, "green", 1.46 + 2 * (CW + GAP), CY, CW, CH, [("関心が一番高い", 15, True, WHITE), ("申込み直後に、", 15, True, WHITE), ("LINEでつながるのが最適", 15, True, WHITE), ("（開封率 約70%／琴平バス）", 11, False, WHITE)])
source(s, CY + CH + 0.3, "出典：LINEヤフー for Business 導入事例（保険見直し本舗・琴平バス）")

# ================= 3枚目：シーン1 =================
s = head(2, "シーン1｜来店・来院の予約", "予約を、来店・来院につなげる")
Y, H, BW, AW = 4.5, 4.4, 5.25, 1.2
steps = [("card", "予約完了", NAVY), ("green", "LINEで予約確認", WHITE), ("green", "前日にお知らせ", WHITE), ("lgreen", "来店", GREENTXT)]
x = 1.46
for i, (kind, text, col) in enumerate(steps):
    add(s, kind, x, Y, BW, H, [(tx, 17, True, col) for tx in text.split("\n")])
    x += BW
    if i < 3:
        add(s, "arrow_navy" if i == 0 else "arrow_green", x + 0.1, Y + H / 2 - 0.6, 1.0, 1.2)
        x += AW
add(s, "navy", 1.46, 9.3, 24.6, 1.8, [("無断キャンセルを防ぎ、来店につなげる", 19, True, WHITE)])
add(s, "green", 1.46, 11.5, 12.2, 4.6, [("美容クリニック（弊社）", 15, False, WHITE), ("来院率 40〜50%改善", 21, True, WHITE)])
add(s, "navy", 13.86, 11.5, 12.2, 4.6, [("保険見直し本舗", 15, False, WHITE), ("予約100件あたり 面談＋5件", 21, True, WHITE)])
source(s, 16.4, "出典：弊社運用実績／LINEヤフー for Business 導入事例（保険見直し本舗）")

# ================= 4枚目：シーン2 =================
s = head(3, "シーン2｜資料請求", "資料請求を、予約・商談につなげる")
Y, H, BW, AW = 4.5, 4.4, 5.25, 1.2
steps = [("card", "資料請求", NAVY), ("green", "LINEへ案内", WHITE), ("green", "事例・動画を\n配信", WHITE), ("lgreen", "見学・商談の\n予約", GREENTXT)]
x = 1.46
for i, (kind, text, col) in enumerate(steps):
    add(s, kind, x, Y, BW, H, [(tx, 17, True, col) for tx in text.split("\n")])
    x += BW
    if i < 3:
        add(s, "arrow_navy" if i == 0 else "arrow_green", x + 0.1, Y + H / 2 - 0.6, 1.0, 1.2)
        x += AW
add(s, "navy", 1.46, 9.3, 24.6, 1.8, [("資料請求のみで終わらせず、予約・商談につなげる", 19, True, WHITE)])
add(s, "card", 1.46, 11.5, 12.2, 4.6, [("皮膚科クリニック", 15, False, NAVY), ("予約数 約1.2倍", 26, True, NAVY)])
add(s, "card", 13.86, 11.5, 12.2, 4.6, [("就職支援（UZUZ）", 15, False, NAVY), ("面談予約率 55%", 26, True, NAVY)])
source(s, 16.4, "出典：LINEヤフー for Business 導入事例（アクネクリニック・UZUZ）")

# ---- 5枚目（進め方）を削除 ----
lst = prs.slides._sldIdLst
sid = lst[4]
prs.part.drop_rel(sid.get(qn("r:id")))
lst.remove(sid)

# ---- 図形IDの振り直し ----
for sl in prs.slides:
    for n, c in enumerate(sl.shapes._spTree.iter(qn("p:cNvPr")), start=2):
        c.set("id", str(n))

# ---- 字体をメイリオに統一（latin→ea→cs の順）----
AFTER = {qn(x) for x in ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")}
for sl in prs.slides:
    for r in sl.shapes._spTree.iter(qn("a:r")):
        if r.find(qn("a:rPr")) is None:
            r.insert(0, r.makeelement(qn("a:rPr"), {"lang": "ja-JP"}))
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for rpr in sl.shapes._spTree.iter(qn(tag)):
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
prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
