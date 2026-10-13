# -*- coding: utf-8 -*-
"""字体の検査：全ページの全文字が「メイリオ」か（2026-10-13 殿村さん指摘）。

    python -m _build.fontcheck <pptx>

文字（a:r）ごとに a:latin／a:ea を見て、メイリオ（"メイリオ" または "Meiryo"）以外・指定なしをページ別に数える。
指定なしの文字は、テーマ既定（Calibri／MSゴシック）の字体になる。納品前に必ず0件にする。
"""
import sys
import zipfile
from collections import Counter

from lxml import etree

from _build.arauchi_v4.build import slide_order

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
OK = {"メイリオ", "Meiryo"}


def scan(path):
    z = zipfile.ZipFile(path)
    bad = {}
    for no, name in enumerate(slide_order(z), 1):
        c = Counter()
        for r in etree.fromstring(z.read(name)).iter(A + "r"):
            if not "".join((t.text or "") for t in r.iter(A + "t")).strip():
                continue
            rpr = r.find(A + "rPr")
            lat = rpr.find(A + "latin") if rpr is not None else None
            ea = rpr.find(A + "ea") if rpr is not None else None
            lat = lat.get("typeface") if lat is not None else None
            ea = ea.get("typeface") if ea is not None else None
            if lat not in OK or ea not in OK:
                c[(lat, ea)] += 1
        if c:
            bad[no] = c
    return bad


if __name__ == "__main__":
    bad = scan(sys.argv[1])
    for no, c in bad.items():
        print(f"S{no}: " + ", ".join(f"{n}か所 latin={k[0]} ea={k[1]}" for k, n in c.items()))
    print("メイリオ以外・指定なしの文字：", sum(sum(c.values()) for c in bad.values()), "か所")
