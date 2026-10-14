# -*- coding: utf-8 -*-
"""あらうち提案資料 v9（2026-10-14 殿村さん指示）：S15・S16のプラン③の月額費用の表記を整理する。

    python -m _build.arauchi_v9.build <殿村さんがPowerPointで保存したv8(2).pptx> <出力.pptx>

ベースのzipから、slide15.xml と slide16.xml だけを差し替える（ほかのパーツは1バイトも変えない）。

S15 料金プランの比較
  ・下から2行目の項目名「月額費用」→「月額ツール・LINE関連費」（35,000円／35,000円／プラン③＝利用ツールにより変動）
  ・最下段「月額費用」のプラン③＝「月額費用が発生／（別途お見積り）」
  ・プラン③の縦結合（2行ぶち抜き）を解いて、それぞれの行に入れる
S16 運用内容の比較
  ・プラン③の月額費用＝「月額費用が発生／（別途お見積り）」（2行になるので、行の高さと下の注記の位置を少し調整）
  ・プラン③の注記を差し替え（「DYM手数料」の表記をなくす）
"""
import copy
import sys
import zipfile

from lxml import etree

from _build.arauchi_v4.build import normalize_decl

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}


def q(tag):
    pre, name = tag.split(":")
    return "{%s}%s" % (NS[pre], name)


def cell_text(tc):
    return "\n".join("".join(t.text or "" for t in p.iter(q("a:t"))) for p in tc.iter(q("a:p")))


def set_run_text(p, text):
    """段落の最初の run の文字だけを差し替える（書式はそのまま）"""
    runs = p.findall(q("a:r"))
    assert runs, "run がない段落"
    runs[0].find(q("a:t")).text = text
    for r in runs[1:]:
        p.remove(r)


def set_size(p, sz):
    for rpr in list(p.iter(q("a:rPr"))) + list(p.iter(q("a:endParaRPr"))):
        rpr.set("sz", str(sz))


def table_rows(root):
    tbl = root.find(".//" + q("a:tbl"))
    return tbl, tbl.findall(q("a:tr"))


# ---------------------------------------------------------------- S15
def fix_s15(root):
    tbl, rows = table_rows(root)
    r3, r4 = rows[3], rows[4]
    c3, c4 = r3.findall(q("a:tc")), r4.findall(q("a:tc"))
    assert [cell_text(c) for c in c3[:3]] == ["月額費用", "35,000円", "35,000円"], [cell_text(c) for c in c3]
    assert [cell_text(c) for c in c4[:3]] == ["月額費用", "135,000円", "85,000円"], [cell_text(c) for c in c4]
    assert c3[3].get("rowSpan") == "2" and c4[3].get("vMerge") == "1"
    assert cell_text(c3[3]) == "月額費用が発生\n（別途お見積り）", cell_text(c3[3])

    # 1) 下から2行目の項目名（長くなるので、左右の余白を 0.08in → 0.03in にして1行に収める。文字は中央揃えなので見た目は変わらない）
    set_run_text(c3[0].find(".//" + q("a:p")), "月額ツール・LINE関連費")
    tcpr = c3[0].find(q("a:tcPr"))
    tcpr.set("marL", "27432")
    tcpr.set("marR", "27432")

    # 2) プラン③の縦結合を解く
    del c3[3].attrib["rowSpan"]
    del c4[3].attrib["vMerge"]

    # 3) 下から2行目のプラン③＝同じ行の他のセル（35,000円）と同じ書式で「利用ツールにより変動」
    new_tx = copy.deepcopy(c3[2].find(q("a:txBody")))
    set_run_text(new_tx.find(q("a:p")), "利用ツールにより変動")
    c3[3].replace(c3[3].find(q("a:txBody")), new_tx)
    c3[3].replace(c3[3].find(q("a:tcPr")), copy.deepcopy(c3[2].find(q("a:tcPr"))))

    # 4) 最下段のプラン③＝同じ行の他のセル（135,000円）と同じ書式（太字・紺）で2行
    tx = c4[3].find(q("a:txBody"))
    for p in tx.findall(q("a:p")):
        tx.remove(p)
    for line in ("月額費用が発生", "（別途お見積り）"):
        p = copy.deepcopy(c4[2].find(q("a:txBody")).find(q("a:p")))
        set_run_text(p, line)
        set_size(p, 1500)
        tx.append(p)


# ---------------------------------------------------------------- S16
def fix_s16(root):
    tbl, rows = table_rows(root)
    c = rows[1].findall(q("a:tc"))
    assert [cell_text(x) for x in c] == ["月額費用", "135,000円", "85,000円", "月額費用が発生"], [cell_text(x) for x in c]

    # プラン③の月額費用＝2行
    tx = c[3].find(q("a:txBody"))
    p1 = tx.find(q("a:p"))
    p2 = copy.deepcopy(p1)
    set_run_text(p2, "（別途お見積り）")
    p1.addnext(p2)

    # 2行ぶんの高さ（0.55in → 0.62in）。表の全体の高さと、下の注記の位置も同じだけずらす
    add = 566928 - int(rows[1].get("h"))
    assert add == 64008, add
    rows[1].set("h", "566928")
    gf = root.find(".//" + q("p:graphicFrame"))
    ext = gf.find(q("p:xfrm")).find(q("a:ext"))
    ext.set("cy", str(int(ext.get("cy")) + add))

    note = [sp for sp in root.iter(q("p:sp")) if sp.find(q("p:nvSpPr")).find(q("p:cNvPr")).get("name") == "注記"][0]
    xf = note.find(q("p:spPr")).find(q("a:xfrm"))
    off, ext_n = xf.find(q("a:off")), xf.find(q("a:ext"))
    off.set("y", str(int(off.get("y")) + add))
    ext_n.set("cy", str(int(ext_n.get("cy")) - add))             # 下端（フッターの手前）は動かさない

    # プラン③の注記
    paras = note.find(q("p:txBody")).findall(q("a:p"))
    assert len(paras) == 3 and "DYM手数料" in "".join(t.text for t in paras[2].iter(q("a:t")))
    set_run_text(paras[2], "※プラン③は継続利用するツールに応じて月額費用が発生します。金額は別途お見積りします。")


# ---------------------------------------------------------------- 差し替え
def main(base, out):
    zb = zipfile.ZipFile(base)
    targets = {"ppt/slides/slide15.xml": fix_s15, "ppt/slides/slide16.xml": fix_s16}
    new = {}
    for name, fn in targets.items():
        root = etree.fromstring(zb.read(name))
        fn(root)
        new[name] = normalize_decl(etree.tostring(root, encoding="UTF-8"))
    with zipfile.ZipFile(out, "w") as zo:
        for info in zb.infolist():
            data = new.get(info.filename, None)
            zo.writestr(info, zb.read(info.filename) if data is None else data, compress_type=info.compress_type)
    print("saved:", out)
    print("差し替えたパーツ:", sorted(new))


if __name__ == "__main__":
    main(*sys.argv[1:3])
