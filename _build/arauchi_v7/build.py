# -*- coding: utf-8 -*-
"""あらうち提案資料 v7 をつくる（2026-10-13 殿村さん指示）。

    python -m _build.arauchi_v7.build <v6.pptx> <出力.pptx>

1. v6 に pages.py の直しを当てる（python-pptx。S1・S3・S8・S15 と、費用ページ内のページ参照）
2. v6 のzipに、直したスライドのXMLだけを差し込む（それ以外のパーツは v6 のまま）
3. 費用4枚（v6のS23〜S26）を、S13の直後へ並べ替える（presentation.xml の p:sldIdLst の並びだけ変える）
"""
import re
import sys
import zipfile
from pathlib import Path

from pptx import Presentation

from _build.arauchi_v4.build import c14n, normalize_decl, rel_set, slide_order
from _build.arauchi_v7.pages import apply

# 出来上がりの並び（v6の何枚目か、0始まり）：S1〜S13 → 費用4枚（v6のS23〜S26）→ v6のS14〜S22 → v6のS27〜S48
ORDER = list(range(0, 13)) + [22, 23, 24, 25] + list(range(13, 22)) + list(range(26, 48))


def repack(base_path, built_path, out_path):
    zb, zn = zipfile.ZipFile(base_path), zipfile.ZipFile(built_path)
    base_order, built_order = slide_order(zb), slide_order(zn)
    assert len(base_order) == len(built_order) == 48, (len(base_order), len(built_order))
    assert sorted(ORDER) == list(range(48))
    repl = {}
    for b_name, n_name in zip(base_order, built_order):
        if c14n(zb.read(b_name)) != c14n(zn.read(n_name)):
            repl[b_name] = normalize_decl(zn.read(n_name))
        b_rels = b_name.replace("slides/", "slides/_rels/") + ".rels"
        n_rels = n_name.replace("slides/", "slides/_rels/") + ".rels"
        if rel_set(zb.read(b_rels)) != rel_set(zn.read(n_rels)):
            raise SystemExit(f"想定外のrelsの違い: {b_rels}")
    pres = zb.read("ppt/presentation.xml").decode("utf-8")
    entries = re.findall(r"<p:sldId [^>]*/>", pres)
    assert len(entries) == 48
    new_entries = [entries[i] for i in ORDER]
    pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", lambda m: "<p:sldIdLst>" + "".join(new_entries) + "</p:sldIdLst>", pres, flags=re.S)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zb.infolist():
            name = info.filename
            if name == "ppt/presentation.xml":
                data = pres.encode("utf-8")
            elif name in repl:
                data = repl.pop(name)
            else:
                data = zb.read(name)
            zo.writestr(name, data)
    assert not repl, repl
    return sorted(n for n in base_order if c14n(zb.read(n)) != c14n(zn.read(built_order[base_order.index(n)])))


def main(base, out):
    tmp = out + ".tmp.pptx"
    prs = Presentation(base)
    apply(prs)
    prs.save(tmp)
    changed = repack(base, tmp, out)
    Path(tmp).unlink()
    print("saved:", out)
    print("差し替えたスライドのXML:", changed)


if __name__ == "__main__":
    main(*sys.argv[1:3])
