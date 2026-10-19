# -*- coding: utf-8 -*-
"""塾ver3：全体的に文字を大きくする。はみ出る所は大きくしない（2026-10-19 殿村さん）。

    python -m _build.juku_fmt.enlarge <塾ver3（56枚）> <出力.pptx>

  - 対象：表示しているページの、14pt以下の文字（見出し16pt・大きな数字はそのまま）。表紙・最後のページ・非表示は除く。
  - 箱（図形・表のセル）ごとに、文字を ×1.2 → ×1.16 → ×1.12 → ×1.09 → ×1.06 の順に試し、メイリオで箱に収まる一番大きい倍率にする。
    どれも収まらなければ、その箱は変えない。同じ形の箱（同じ大きさ・同じ文字の大きさ）と同じ表のセルは、一番小さい倍率にそろえる
    （1つでも大きくできない箱があれば、そのまとまりは全部そのまま）。
  - 収まるかの計算（メイリオ）：全角＝1文字分、半角＝DejaVu Sansで測った幅（メイリオの英数字に近い）、行の高さ＝1.5倍×行間の指定。左右・上下の余白を引いた箱の中で、
    折り返しの行数×行の高さ＋段落の前後の空き が、箱の高さ以内か。折り返しなしの箱は、1行が箱の幅に収まるか。
    さらに、大きくして折り返しが増える（1行の帯が2行になる等）箱は大きくしない。
  ほかは1バイトも変えない（直したスライドのXMLだけ差し込む）。
"""
import math
import sys
import unicodedata
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

import _build.juku_fmt.build as jb

EMU_PT = 12700
MAX_PT = 14.0
STEPS = (1.2, 1.16, 1.12, 1.09, 1.06)
LINE = 1.5                    # メイリオの行の高さ（文字の大きさに対して）


from functools import lru_cache

from PIL import ImageFont

_FONTS = {False: ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 100),
          True: ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 100)}


@lru_cache(maxsize=None)
def char_w(ch, bold=False):
    """1文字の幅（文字の大きさに対する割合）。全角＝1。半角（英数字・記号）はメイリオに近い幅の DejaVu Sans で測る"""
    if unicodedata.east_asian_width(ch) in "WF":
        return 1.0
    w = _FONTS[bold].getlength(ch) / 100
    return w if w > 0 else 1.0


def runs_of(p):
    """段落の中の (文字, 大きさpt) の並び。改行は '\n'。大きさの指定が無い文字は None"""
    out = []
    for el in p:
        tag = el.tag.split("}")[1]
        if tag in ("r", "fld"):
            rpr = el.find(qn("a:rPr"))
            sz = rpr.get("sz") if rpr is not None else None
            bold = rpr is not None and rpr.get("b") == "1"
            t = el.find(qn("a:t"))
            out.append(((t.text or "") if t is not None else "", int(sz) / 100 if sz else None, bold))
        elif tag == "br":
            out.append(("\n", None, False))
    return out


def para_height(p, width_pt, scale, wrap=True):
    """段落の高さ（pt）。折り返しなしなら (1行の幅, 高さ)"""
    runs = runs_of(p)
    ppr = p.find(qn("a:pPr"))
    sizes = [s for _, s, _ in runs if s]
    if not sizes:
        end = p.find(qn("a:endParaRPr"))
        sizes = [int(end.get("sz")) / 100] if end is not None and end.get("sz") else [12.0]
    big = max(sizes) * scale
    pct, pts = 1.0, None
    if ppr is not None:
        ln = ppr.find(qn("a:lnSpc"))
        if ln is not None:
            if ln.find(qn("a:spcPct")) is not None:
                pct = int(ln.find(qn("a:spcPct")).get("val")) / 100000
            elif ln.find(qn("a:spcPts")) is not None:
                pts = int(ln.find(qn("a:spcPts")).get("val")) / 100 * scale
    line_h = pts if pts else big * LINE * pct
    extra = 0.0
    if ppr is not None:
        for tag in ("a:spcBef", "a:spcAft"):
            sp = ppr.find(qn(tag))
            if sp is not None and sp.find(qn("a:spcPts")) is not None:
                extra += int(sp.find(qn("a:spcPts")).get("val")) / 100
            elif sp is not None and sp.find(qn("a:spcPct")) is not None:
                extra += int(sp.find(qn("a:spcPct")).get("val")) / 100000 * big
    indent = 0.0
    if ppr is not None and ppr.get("marL"):
        indent = int(ppr.get("marL")) / EMU_PT
    lines, cur, widest = 1, 0.0, 0.0
    avail = max(width_pt - indent, 1.0) * 0.97
    prev = None
    for text, sz, bold in runs:
        s = (sz or max(sizes)) * scale
        for ch in text:
            if ch == "\n":
                widest = max(widest, cur)
                lines += 1
                cur = 0.0
                prev = None
                continue
            kind = "cjk" if unicodedata.east_asian_width(ch) in "WF" else ("alnum" if ch.isascii() and ch.isalnum() else None)
            w = char_w(ch, bold) * s
            if prev and kind and prev != kind:
                w += 0.25 * s                                # 日本語と英数字の間の自動の空き（PowerPointも空ける）
            prev = kind or prev
            if wrap and cur + w > avail and cur > 0:
                lines += 1
                widest = max(widest, cur)
                cur = 0.0
            cur += w
    widest = max(widest, cur)
    return widest + indent, lines * line_h + extra, lines


def layout(txbody, w_emu, h_emu, scale, lins=91440, rins=91440, tins=45720, bins=45720):
    """(箱に収まるか, 段落ごとの行数)"""
    bp = txbody.find(qn("a:bodyPr"))
    wrap = True
    if bp is not None:
        lins = int(bp.get("lIns", lins))
        rins = int(bp.get("rIns", rins))
        tins = int(bp.get("tIns", tins))
        bins = int(bp.get("bIns", bins))
        wrap = bp.get("wrap") != "none"
        if bp.find(qn("a:normAutofit")) is not None or bp.get("vert") not in (None, "horz"):
            return False, []
    width = (w_emu - lins - rins) / EMU_PT
    height = (h_emu - tins - bins) / EMU_PT
    total, widest, counts = 0.0, 0.0, []
    for p in txbody.findall(qn("a:p")):
        wd, ht, n = para_height(p, width, scale, wrap)
        widest = max(widest, wd)
        total += ht
        counts.append(n)
    if not wrap and widest > width * 0.97:
        return False, counts
    return total <= height + 0.5, counts


def scaled_sizes(txbody):
    return [r for r in txbody.iter(qn("a:rPr"), qn("a:endParaRPr")) if r.get("sz")]


def best_k(txbody, w, h, **ins):
    """大きくできる一番大きい倍率（できなければ None）"""
    rprs = scaled_sizes(txbody)
    if not rprs or not "".join(t.text or "" for t in txbody.iter(qn("a:t"))).strip():
        return None
    if any(int(r.get("sz")) / 100 > MAX_PT for r in rprs if r.tag == qn("a:rPr")):
        return None                                          # 見出し・大きな数字の入った箱はそのまま
    ok, base = layout(txbody, w, h, 1.0, **ins)
    if not ok:
        return None                                          # 今の時点で収まっていない箱は触らない
    for k in STEPS:
        ok, counts = layout(txbody, w, h, k, **ins)
        if ok and all(c <= b for c, b in zip(counts, base)):     # 折り返しが増える（1行が2行になる）なら大きくしない
            return k
    return None


def apply(txbody, k):
    for r in scaled_sizes(txbody):
        old = int(r.get("sz")) / 100
        new = max(old, math.floor(old * k * 2 + 0.5) / 2)
        r.set("sz", str(int(round(new * 100))))


def walk(shapes):
    for sh in shapes:
        if sh.shape_type == 6:                               # グループ
            yield from walk(sh.shapes)
        else:
            yield sh


def sig(txbody):
    return tuple(sorted({r.get("sz") for r in scaled_sizes(txbody)}))


def enlarge(prs):
    """同じ形の箱（同じ大きさ・同じ文字の大きさ）や同じ表のセルは、同じ倍率にそろえる（一番小さい倍率に合わせる）"""
    stats = {}
    n_slides = len(prs.slides)
    for n, s in enumerate(prs.slides, 1):
        if n in (1, n_slides) or s._element.get("show") == "0":
            continue
        groups = {}
        for sh in walk(s.shapes):
            if sh.rotation:
                continue
            if sh.has_text_frame and sh.width and sh.height:
                tb = sh.text_frame._txBody
                key = ("sp", round(sh.width / 45720), round(sh.height / 45720), sig(tb))
                groups.setdefault(key, []).append((tb, best_k(tb, sh.width, sh.height)))
            elif getattr(sh, "has_table", False) and sh.has_table:
                tbl = sh.table
                for ri, row in enumerate(tbl.rows):
                    for ci, cell in enumerate(row.cells):
                        tc = cell._tc
                        if tc.get("hMerge") or tc.get("vMerge"):
                            continue
                        w = sum(tbl.columns[c].width for c in range(ci, ci + int(tc.get("gridSpan", 1))))
                        h = sum(tbl.rows[r].height for r in range(ri, ri + int(tc.get("rowSpan", 1))))
                        pr = tc.find(qn("a:tcPr"))
                        ins = {}
                        if pr is not None:
                            for a, k in (("marL", "lins"), ("marR", "rins"), ("marT", "tins"), ("marB", "bins")):
                                if pr.get(a):
                                    ins[k] = int(pr.get(a))
                        tb = tc.find(qn("a:txBody"))
                        if not "".join(x.text or "" for x in tb.iter(qn("a:t"))).strip():
                            continue
                        groups.setdefault(("tbl", id(tbl._tbl)), []).append((tb, best_k(tb, w, h, **ins)))
        done = 0
        for key, items in groups.items():
            ks = [k for _, k in items]
            if any(k is None for k in ks):
                continue                                     # 1つでも大きくできない箱があれば、そろえるために全部そのまま
            k = min(ks)
            for tb, _ in items:
                apply(tb, k)
                done += 1
        if done:
            stats[n] = done
    return stats


def main(base, out):
    prs = Presentation(base)
    stats = enlarge(prs)
    tmp = out + ".tmp.pptx"
    prs.save(tmp)
    jb.ORDER = list(range(len(prs.slides)))       # repack は並びの長さだけ見る
    jb.repack(base, tmp, out)
    Path(tmp).unlink()
    print("saved:", out, "／大きくした箱の数（ページ: 数）:", stats, "／合計", sum(stats.values()))


if __name__ == "__main__":
    main(*sys.argv[1:3])
