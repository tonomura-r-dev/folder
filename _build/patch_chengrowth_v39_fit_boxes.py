# -*- coding: utf-8 -*-
"""チェングロウス ver3.8 → ver3.9：枠の大きさを文字の量に合わせる（2026-10-02 殿村さん指示）
文字数・文字サイズ・枠の幅から必要な高さを見積もり、枠の高さを合わせる。余った空きは上下の余白に回し、全体を中央に寄せる。
対象：11・12・15（行間のみ）・20・21（表）・23・24枚目（ver3.8の番号）
  python3 _build/patch_chengrowth_v39_fit_boxes.py <ver3.8.pptx>
"""
import math
import sys
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20260929_チェングロウス_LINE公式アカウント運用のご提案ver3.9.pptx"
E = 360000
prs = Presentation(sys.argv[1])
TOP, BOTTOM, GAP = 4.7, 16.7, 0.45      # 本文領域（cm）、ブロック間の隙間


def char_em(c):
    o = ord(c)
    if c == " ":
        return 0.3
    if o < 0x80:
        return 0.6
    return 1.0


def est_h(sh, min_h=0.0):
    """文字の量から必要な高さ（cm）を見積もる"""
    bp = sh._element.find(".//" + qn("a:bodyPr"))
    li, ti, ri, bi = (int(bp.get(a, d)) for a, d in (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720)))
    inner = (sh.width - li - ri) / E * 0.88
    total = 0.0
    for p in sh.text_frame.paragraphs:
        szs = [int(r.get("sz")) for r in p._p.iter(qn("a:rPr")) if r.get("sz")]
        end = p._p.find(qn("a:endParaRPr"))
        if end is not None and end.get("sz"):
            szs.append(int(end.get("sz")))
        size = (max(szs) if szs else 1800) / 100.0
        text = p.text
        w = sum(char_em(c) for c in text) * size * 0.03528
        lines = max(1, math.ceil(w / inner)) if text else 1
        total += lines * size * 1.5 * 0.03528
    return max(min_h, total + (ti + bi) / E + 0.35)


def by(n):
    return {sh.name: sh for sh in prs.slides[n - 1].shapes}


def place(sh, top=None, height=None):
    if top is not None:
        sh.top = int(top * E)
    if height is not None:
        sh.height = int(height * E)


def center_start(total, bottom=BOTTOM):
    avail = bottom - TOP
    return TOP + max(0.0, (avail - total) / 2)


# カード本文の先頭にある空の段落を消す（枠に無駄な空きが出るため）
for n, names in ((11, ("Rounded Rectangle 5", "Rounded Rectangle 7", "Rounded Rectangle 9")), (24, ("Rounded Rectangle 5", "Rounded Rectangle 7", "Rounded Rectangle 9"))):
    dd = by(n)
    for nm in names:
        ps = dd[nm].text_frame.paragraphs
        if ps and not ps[0].text.strip() and len(ps) > 1:
            ps[0]._p.getparent().remove(ps[0]._p)

# ---------- 11枚目 ----------
d = by(11)
body = max(est_h(d[k]) for k in ("Rounded Rectangle 5", "Rounded Rectangle 7", "Rounded Rectangle 9"))
site = max(est_h(d["Rounded Rectangle 11"]), est_h(d["Rounded Rectangle 12"]))
ban = est_h(d["Rounded Rectangle 13"], 1.2)
lab = 0.8
total = 1.0 + 0.1 + body + GAP + lab + 0.1 + site + GAP + ban
over = total - (16.9 - TOP)
if over > 0:                      # 入り切らないときは、カード本文の高さを詰める
    body -= over
    total -= over
y = center_start(total, 16.9)
for h, b in (("Rounded Rectangle 4", "Rounded Rectangle 5"), ("Rounded Rectangle 6", "Rounded Rectangle 7"), ("Rounded Rectangle 8", "Rounded Rectangle 9")):
    place(d[h], y, 1.0)
    place(d[b], y + 1.1, body)
y2 = y + 1.1 + body + GAP
place(d["TextBox 10"], y2, lab)
for k in ("Rounded Rectangle 11", "Rounded Rectangle 12"):
    place(d[k], y2 + lab + 0.1, site)
place(d["Rounded Rectangle 13"], y2 + lab + 0.1 + site + GAP, ban)

# ---------- 12枚目 ----------
d = by(12)
body = max(est_h(d["Rounded Rectangle 5"]), est_h(d["Rounded Rectangle 8"]))
ban = est_h(d["Rounded Rectangle 10"], 2.0)
total = 1.1 + body + 1.3 + ban
y = center_start(total, 16.6)
for h, b in (("Rounded Rectangle 4", "Rounded Rectangle 5"), ("Rounded Rectangle 7", "Rounded Rectangle 8")):
    place(d[h], y, 1.1)
    place(d[b], y + 1.1, body)
place(d["Right Arrow 6"], y + 1.1 + body / 2 - 0.4)
place(d["Down Arrow 9"], y + 1.1 + body + 0.3)
place(d["Rounded Rectangle 10"], y + 1.1 + body + 1.3, ban)

# ---------- 20・21枚目（動線×列の表）----------
for n in (20, 21):
    d = by(n)
    rows = {}
    for nm, sh in d.items():
        if nm.startswith("Rounded Rectangle"):
            rows.setdefault(round(sh.top / E, 1), []).append(sh)
    tops = sorted(rows)
    hs = [max(est_h(s, 2.4) for s in rows[t]) for t in tops]
    gap = 0.3
    hdr = 1.0 if "Table 4" in d else 0
    total = hdr + 0.2 + sum(hs) + gap * (len(hs) - 1)
    y = center_start(total, 16.7)
    if "Table 4" in d:
        place(d["Table 4"], y)
    yy = y + hdr + 0.2
    for t, h in zip(tops, hs):
        for s in rows[t]:
            place(s, yy, h)
        yy += h + gap

# ---------- 23枚目 ----------
d = by(23)
cards = max(est_h(d[k]) for k in ("Rounded Rectangle 4", "Rounded Rectangle 5", "Rounded Rectangle 6"))
ban = est_h(d["Rounded Rectangle 14"], 1.2)
tl_names = ("Rounded Rectangle 7", "Rounded Rectangle 8", "TextBox 9", "Rounded Rectangle 10", "Rounded Rectangle 11", "Rounded Rectangle 12", "TextBox 13")
tl_top = min(d[k].top for k in tl_names) / E
tl_h = max((d[k].top + d[k].height) for k in tl_names) / E - tl_top
total = cards + GAP + tl_h + GAP + ban
y = center_start(total, 16.9)
for k in ("Rounded Rectangle 4", "Rounded Rectangle 5", "Rounded Rectangle 6"):
    place(d[k], y, cards)
off = y + cards + GAP - tl_top
for k in tl_names:
    place(d[k], d[k].top / E + off)
place(d["Rounded Rectangle 14"], y + cards + GAP + tl_h + GAP, ban)

# ---------- 24枚目 ----------
d = by(24)
body = max(est_h(d[k]) for k in ("Rounded Rectangle 5", "Rounded Rectangle 7", "Rounded Rectangle 9"))
dev = est_h(d["Rounded Rectangle 10"], 1.8)
ban = est_h(d["Rounded Rectangle 11"], 1.2)
total = 1.0 + 0.1 + body + GAP + dev + GAP + ban
y = center_start(total, 16.9)
for h, b in (("Rounded Rectangle 4", "Rounded Rectangle 5"), ("Rounded Rectangle 6", "Rounded Rectangle 7"), ("Rounded Rectangle 8", "Rounded Rectangle 9")):
    place(d[h], y, 1.0)
    place(d[b], y + 1.1, body)
place(d["Rounded Rectangle 10"], y + 1.1 + body + GAP, dev)
place(d["Rounded Rectangle 11"], y + 1.1 + body + GAP + dev + GAP, ban)

prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
