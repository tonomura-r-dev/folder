# -*- coding: utf-8 -*-
"""人材業界（派遣・バイト・単発）の業界資料：色使いを直す（2026-10-20 殿村さん「コレ修正してほしい 主に色使い」）。

    python -m _build.patch_haken_colors <資料.pptx> <出力.pptx>

  紺・黒・白・灰色だけで組む（緑はLINEの部分だけ）。黄色・オレンジ・赤・明るい青・紫・黄緑を外す。
  ・赤・オレンジの文字（リード2行目・★の行・※有料など）→ 紺
  ・「できないこと」「いま起きていること」「① 応募単価の高騰」「これまで」など、できない側・いまの側 → 灰色
  ・オレンジ・赤の札・枠（第6章の行、⑤の行、通知メッセージの3枚、別途費用の枠など）→ 紺の線＋白
  ・8フェーズの①〜④（薄いオレンジ）→ 薄い灰色（⑤〜⑧の薄い青はそのまま）
  ・黄緑（04FA0A）の帯「公式アカウントを友だち追加！」→ LINEの緑（06C755）。それ以外の黄緑・緑 → 紺
  ・前後検索の凡例の4色は画像（散布図）の色と対応しているので残す。画像の中の色は変えられない
  あわせて、別途費用の「サンクスLINE誘導 初期10万」→「初期15万」（CLAUDE.md の費用。月額3万円〜はそのまま）。
"""
import copy
import sys

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

NAVY = "1E2761"
GRAY = "7F7F7F"
DGRAY = "595959"
LGRAY = "F2F2F2"
LBLUE = "EEF2F9"
LBORDER = "C9D2E3"
LINEGRAY = "A6A6A6"
WHITE = "FFFFFF"
LINE_GREEN = "06C755"

# 置き換える色は "色" か (色, 付け直す明るさの指定)。明るさの指定を渡すと、元の lumMod などは捨てて付け直す
NOMODS = {}
# (色, 種類) → 置き換える色。種類＝text / fill / line / gs
DEFAULT = {
    ("FF0000", "text"): NAVY, ("C00000", "text"): NAVY, ("ED7D31", "text"): NAVY,
    ("00B050", "text"): NAVY, ("7A5200", "text"): NAVY, ("FFFF00", "text"): WHITE,
    ("FF0000", "fill"): NAVY, ("FF0000", "line"): NAVY,
    ("ED7D31", "fill"): NAVY, ("ED7D31", "line"): NAVY,
    ("FDF1E9", "fill"): WHITE, ("FDF2F2", "fill"): WHITE,
    ("4A90D9", "fill"): NAVY, ("4A90D9", "line"): NAVY,
    ("E8A838", "line"): NAVY, ("FFF8E8", "fill"): WHITE,
    ("FFC000", "line"): NAVY, ("2080FF", "line"): NAVY,
    ("04FA0A", "line"): NAVY, ("04FA0A", "fill"): NAVY, ("04FA0A", "gs"): NAVY,
    ("accent6", "fill"): GRAY, ("accent6", "gs"): ("8A95B5", NOMODS),
    ("8E4EC6", "fill"): NAVY, ("8E4EC6", "line"): NAVY, ("8E4EC6", "text"): NAVY,
    ("3467B2", "fill"): NAVY, ("3467B2", "line"): NAVY, ("3467B2", "text"): NAVY,
    ("00897B", "fill"): NAVY, ("00897B", "line"): NAVY, ("00897B", "text"): NAVY,
}
# できない側・いまの側を灰色にするページ
CONTRAST = {("FF0000", "fill"): GRAY, ("FF0000", "line"): GRAY, ("FDF2F2", "fill"): LGRAY, ("FDF1E9", "fill"): LBLUE}
KEEP_LEGEND = {(c, k) for c in ("8E4EC6", "3467B2", "00897B", "C00000") for k in ("fill", "text")}

# ページ番号 → (タイトルの頭, 色の上書き, 図形名ごとの上書き {名前: {(色, 種類): 色}})
SLIDES = {
    2: ("資料アジェンダ", {("ED7D31", "line"): LBORDER}, {}),
    3: ("［第1章］検討背景", CONTRAST, {}),
    4: ("［第1章］派遣・バイト求人のWebマーケ構造", CONTRAST, {"TextBox 7": {("FF0000", "text"): DGRAY}}),
    5: ("［第1章］Webマーケで「できること」", CONTRAST, {}),
    6: ("［第2章］全体設計｜カスタマージャーニー",
        {("FFC000", "line"): LINEGRAY, ("2080FF", "line"): LINEGRAY, ("accent1", "line"): NAVY}, {}),
    7: ("［第2章］全体設計｜LINEで対策できる領域", {}, {}),
    8: ("［第2章］全体設計｜施策の考え方", {}, {}),
    9: ("［第2章］ニーズ調査｜シーズナリティ", {}, {}),
    10: ("［第2章］ニーズ調査｜前後検索", None, {}),
    11: ("［第2章］求職者の心理", {}, {}),
    12: ("［第2章］他社分析", {}, {}),
    16: ("［第3章］構築｜施策展開図", {}, {}),
    17: ("［第3章］構築｜施策展開図（初期・月次）", {("accent1", "line"): NAVY}, {}),
    18: ("［第3章］構築｜友だち追加の動線", {}, {}),
    19: ("［第3章］構築｜あいさつメッセージ", {}, {}),
    20: ("［第3章］希望条件", {}, {}),
    21: ("［第3章］構築｜リッチメニュー", {}, {}),
    22: ("［第4章］配信設計｜追加後15日", {("04FA0A", "fill"): LINE_GREEN}, {}),
    23: ("［第4章］配信設計｜7日ステップ", {}, {}),
    24: ("［第4章］配信設計｜実文面 Day0", {}, {}),
    25: ("［第4章］配信設計｜実文面 Day7", {}, {}),
    26: ("［第4章］年間の企画投稿カレンダー", {("accent1", "fill"): (NAVY, NOMODS), ("accent1", "line"): (NAVY, NOMODS)}, {}),
    27: ("［第4章］配信設計｜通知メッセージ", {}, {}),
    29: ("［第4章］歩留まり｜手が止まるのは", CONTRAST, {}),
    31: ("［第5章］LINEOA実績・信頼性", {}, {}),
    32: ("［第5章］成果｜効果測定の設計", {("ED7D31", "text"): GRAY, **CONTRAST}, {"TextBox 26": {("FF0000", "text"): DGRAY}}),
    33: ("［第6章］費用プラン", {}, {}),
}


def kind_of(clr):
    par = clr.getparent()
    gp = par.getparent()
    ptag = etree.QName(par).localname
    gtag = etree.QName(gp).localname if gp is not None else ""
    if ptag == "gs":
        return "gs"
    if ptag != "solidFill":
        return None
    if gtag in ("rPr", "endParaRPr", "defRPr"):
        return "text"
    if gtag in ("spPr", "tcPr", "grpSpPr"):
        return "fill"
    if gtag == "ln":
        return "line"
    return None


def owner_name(clr):
    e = clr
    while e is not None and e.tag not in (qn("p:sp"), qn("p:cxnSp"), qn("p:pic")):
        e = e.getparent()
    if e is None:
        return None
    nv = e.find(".//" + qn("p:cNvPr"))
    return nv.get("name") if nv is not None else None


def recolor(clr, new_val):
    """new_val＝"色"（元の明るさ・透明度の指定はそのまま）か (色, 付け直す指定)"""
    val, mods = (new_val, None) if isinstance(new_val, str) else new_val
    new = etree.Element(qn("a:srgbClr"), val=val)
    if mods is None:
        for c in clr:
            new.append(copy.deepcopy(c))
    else:
        for k, v in mods.items():
            etree.SubElement(new, qn("a:" + k), val=v)
    clr.getparent().replace(clr, new)


def fix_slide(s, extra, by_name, legend):
    n = 0
    for clr in list(s.shapes._spTree.iter(qn("a:srgbClr"), qn("a:schemeClr"))):
        if clr.getparent() is None:
            continue
        k = kind_of(clr)
        if k is None:
            continue
        v = clr.get("val")
        key = (v.upper() if clr.tag == qn("a:srgbClr") else v, k)
        ov = by_name.get(owner_name(clr), {})
        if key in ov:
            recolor(clr, ov[key])
            n += 1
            continue
        if legend and key in KEEP_LEGEND:
            continue
        new = extra.get(key) or DEFAULT.get(key)
        if new:
            recolor(clr, new)
            n += 1
    return n


def fix_thanks_fee(s):
    """別途費用の「サンクスLINE誘導 初期10万＋月3万〜」→ 初期15万"""
    hit = 0
    for p in s.shapes._spTree.iter(qn("a:p")):
        text = "".join(t.text or "" for t in p.iter(qn("a:t")))
        if "サンクスLINE誘導" not in text or "初期10万" not in text:
            continue
        for t in p.iter(qn("a:t")):
            if "初期10万" in (t.text or ""):
                t.text = t.text.replace("初期10万", "初期15万")
                hit += 1
    assert hit == 1, hit


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    total = 0
    for no, (head, extra, by_name) in SLIDES.items():
        s = slides[no - 1]
        title = next(sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip())
        assert title.replace(" ", "").startswith(head.replace(" ", "")), (no, title)
        total += fix_slide(s, extra or {}, by_name, legend=extra is None)
        if no == 33:
            fix_thanks_fee(s)
    prs.save(out)
    print("saved:", out, "recolored:", total)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
