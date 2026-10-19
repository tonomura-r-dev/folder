# -*- coding: utf-8 -*-
"""塾ver3：明光義塾のステップ配信（公式事例の画面）を1ページ足す（2026-10-19 殿村さん「その二つは明らかにステップ配信の内容じゃない？入れてください」）。

    python -m _build.juku_fmt.meiko_step <塾ver3（54枚）> <出力.pptx>

  公式事例（https://www.lycbiz.com/jp/case-study/line-official-account/meikonet/）の図版2つ：
    図版5＝アンケート未回答の方への配信（クイズ「サボローを探せ！」。答えを見るボタンが「保護者の方／生徒の方／その他の方」）
    図版6＝塾に通っていない方への配信（「先ほどは友だち登録していただき…」から始まる＝友だち追加のあとに自動で届く。
           ボタンは迷い具合で3つ：授業を受けてみたい→無料体験／雰囲気を知りたい→教室見学／一度検討したい→資料請求）
  - 16枚目（新）＝明光義塾のステップ配信（公式事例の画面）。15枚目（実際の画面）の次。
  - 本文のページ参照（P16以降）は+1。型の分解（P20）の③とリード文、ステップ配信（P30）の注記、出典ページの行に、このページを足す。
  ほかのページは1バイトも変えない（直したスライドのXMLと、足した1枚・画像だけを差し込む）。
"""
import copy
import random
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn

import _build.juku_fmt.build as jb
from _build.arauchi_v4.build import slide_order
from _build.arauchi_v4.lib import add_box
from _build.arauchi_v8.pages import replace_in_paragraphs
from _build.gym_curves.build import add_pic, keep_only, set_one
from _build.gym_fmt.build import fix_fonts
from _build.juku_fmt.build import CT_SLIDE, RT_SLIDE, rels_of
from _build.juku_fmt.meiko import BORDER, GREY, INK, NAVY, PALE, W, X0, chip, title_of
from _build.juku_fmt.plan import set_lead

ROOT = Path(__file__).resolve().parent.parent.parent
FIGS = ROOT / "_data/参考資料/明光義塾LINE_20261019/公式事例の図版"
N_BASE = 54
TPL, SCREENS, PATTERN, SEG, STEP, SOURCES = 13, 15, 19, 28, 29, 53      # 54枚の番号
ORDER = list(range(1, SCREENS + 1)) + [("N", 1)] + list(range(SCREENS + 1, N_BASE + 1))
NEW_POS = {key: i for i, key in enumerate(ORDER, 1)}
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
# 図版の見出し（画像の上の文字）と左右の余白を切り落とす：(ファイル, (左, 上, 右, 下) px)
FIG_CROP = [("zuhan_5.png", (100, 52, 765, 450)), ("zuhan_6.png", (120, 48, 680, 450))]
FIG_H = 2.85                               # 画像の高さ（インチ）。幅は縦横比のまま、列の中央に置く


def fig_bytes(name, box):
    im = Image.open(FIGS / name).convert("RGB")
    im = im.crop(box)
    buf = BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue(), im.size


# ---------------------------------------------------------------- 1) 組み立て（zipのまま）
def assemble(base, out):
    z = zipfile.ZipFile(base)
    order = slide_order(z)
    assert len(order) == N_BASE
    for n, t in [(TPL, "現状分析｜各社LINEの今"), (SCREENS, "明光義塾の実際の画面"), (PATTERN, "明光義塾の型を分解すると、3つの仕組み"),
                 (STEP, "ステップ配信｜14日間の配信"), (SOURCES, "この資料の出典")]:
        assert title_of(z, order[n - 1]) == t, (n, title_of(z, order[n - 1]))
    names = set(z.namelist())
    ct = z.read("[Content_Types].xml").decode("utf-8")
    prels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    pres = z.read("ppt/presentation.xml").decode("utf-8")
    assert '<Default Extension="png"' in ct

    def rid_of_part(name):
        m = re.search(r'<Relationship [^>]*Target="slides/%s"[^>]*/>' % re.escape(name.split("/")[-1]), prels)
        return re.search(r'Id="(rId\d+)"', m.group(0)).group(1)

    rid_of = {n: rid_of_part(order[n - 1]) for n in range(1, N_BASE + 1)}
    next_rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', prels)) + 1
    next_slide = max(int(m.group(1)) for n in names for m in [re.match(r"ppt/slides/slide(\d+)\.xml$", n)] if m) + 1

    tpl = z.read(order[TPL - 1]).decode("utf-8")
    rels = etree.fromstring(z.read(rels_of(order[TPL - 1])))
    for r in list(rels):
        if r.get("Type").endswith("/notesSlide"):
            rels.remove(r)
    assert [r.get("Type").split("/")[-1] for r in rels] == ["slideLayout"]
    used = set(re.findall(r'creationId[^>]*val="(\d+)"', "".join(z.read(o).decode("utf-8") for o in order)))
    cid = str(random.Random(202610193).randrange(10 ** 9, 4 * 10 ** 9))
    assert cid not in used
    dst = f"ppt/slides/slide{next_slide}.xml"
    xml, c = re.subn(r'(<p14:creationId [^>]*val=")\d+(")', rf"\g<1>{cid}\2", tpl, count=1)
    assert c == 1
    add = {dst: xml.encode("utf-8")}
    for j, (fn, box) in enumerate(FIG_CROP, 1):
        media = f"ppt/media/meiko_step_{j}.png"
        assert media not in names
        add[media] = fig_bytes(fn, box)[0]
        etree.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                         Id=f"rId{100 + j}", Type=RT_IMAGE, Target=f"../media/meiko_step_{j}.png")
    add[rels_of(dst)] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    ct = ct.replace("</Types>", f'<Override PartName="/{dst}" ContentType="{CT_SLIDE}"/></Types>')
    rid = f"rId{next_rid}"
    prels = prels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{RT_SLIDE}" Target="slides/{dst.split("/")[-1]}"/></Relationships>')
    rid_of[("N", 1)] = rid

    m = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", pres, re.S)
    entries = {re.search(r'r:id="(rId\d+)"', e).group(1): e for e in re.findall(r"<p:sldId [^>]*/>", m.group(1))}
    next_id = max(int(x) for x in re.findall(r'<p:sldId id="(\d+)"', m.group(1))) + 1
    lst = [entries[rid_of[key]] if rid_of[key] in entries else f'<p:sldId id="{next_id}" r:id="{rid_of[key]}"/>' for key in ORDER]
    pres = pres.replace(m.group(0), "<p:sldIdLst>" + "".join(lst) + "</p:sldIdLst>")

    app = z.read("docProps/app.xml").decode("utf-8")
    n_new, ppt_t = len(ORDER), "<vt:lpstr>PowerPoint プレゼンテーション</vt:lpstr>"
    assert f"<Slides>{N_BASE}</Slides>" in app and app.count(ppt_t) == N_BASE
    app = app.replace(f"<Slides>{N_BASE}</Slides>", f"<Slides>{n_new}</Slides>")
    app, c = re.subn(r"(<vt:lpstr>スライド タイトル</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+(</vt:i4>)", rf"\g<1>{n_new}\2", app)
    assert c == 1
    mv = re.search(r'<vt:vector size="(\d+)" baseType="lpstr">', app)
    app = app.replace(mv.group(0), f'<vt:vector size="{int(mv.group(1)) + n_new - N_BASE}" baseType="lpstr">')
    i = app.rfind(ppt_t) + len(ppt_t)
    app = app[:i] + ppt_t * (n_new - N_BASE) + app[i:]
    replace = {"ppt/presentation.xml": pres, "ppt/_rels/presentation.xml.rels": prels,
               "[Content_Types].xml": ct, "docProps/app.xml": app}

    with zipfile.ZipFile(out, "w") as zo:
        for info in z.infolist():
            data = replace.get(info.filename)
            zo.writestr(info, data.encode("utf-8") if data is not None else z.read(info.filename),
                        compress_type=info.compress_type)
        for name, data in add.items():
            zo.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)


# ---------------------------------------------------------------- 2) 中身
def page_step(s):
    """16枚目：明光義塾のステップ配信（公式事例の画面）"""
    keep_only(s, {2, 3, 4, 67, 104})
    set_one(s, 2, "明光義塾のステップ配信（公式事例の画面）")
    set_lead(s, ["友だち追加のあと、相手に合わせた配信が自動で届きます（2023年の公式事例の画面）。",
                 "答えていない方にはクイズで回答を促し、塾に通っていない方には迷い具合に合わせて3つの行動を案内しています。"])
    cw, gap = 4.84, W - 2 * 4.84
    cols = [
        ("1　アンケートに答えていない方へ", 1,
         ["クイズ「サボローを探せ！」の答えを見るボタンが", "「保護者の方／生徒の方／その他の方」。",
          "楽しみながら答えてもらい、誰なのかが分かる。", "配信した方の55%が反応。"]),
        ("2　塾に通っていない方へ", 2,
         ["「先ほどは友だち登録していただき…」から始まる。", "ボタンは迷い具合で3つ：",
          "授業を受けてみたい→無料体験／雰囲気を知りたい→", "教室見学／一度検討したい→資料請求。"]),
    ]
    for i, (head, j, lines) in enumerate(cols):
        x = X0 + i * (cw + gap)
        size = fig_bytes(*FIG_CROP[j - 1])[1]
        iw = FIG_H * size[0] / size[1]
        assert iw <= cw
        chip(s, x, 1.62, cw, 0.34, head, size=11.5)
        add_pic(s, f"rId{100 + j}", f"明光義塾_ステップ配信_{j}", x + (cw - iw) / 2, 2.04, iw, FIG_H)
        add_box(s, x, 2.04 + FIG_H + 0.1, cw, 1.3, [[(t, {})] for t in lines], size=10.5, color=INK, fill=PALE,
                line=BORDER, anchor="ctr", ins=(0.12, 0.06, 0.1, 0.06), line_spacing=110)
    set_one(s, 104, "答えやすい形で聞き、迷い具合に合わせて次の一歩を選べるようにしている。")
    set_one(s, 67, "画面：LINEヤフー for Business 導入事例 明光義塾（2024年4月5日）掲載の図版（2023年の施策時のもの）。"
                   "55%は同事例の記載。画面の中のボタンの色は明光義塾のもの")
    return s


def remap_refs(s):
    """本文中の「P◯」（54枚の番号）を新しい番号に"""
    for t in s._element.iter(qn("a:t")):
        if t.text and re.search(r"P\d+", t.text):
            t.text = re.sub(r"P(\d+)", lambda m: f"P{NEW_POS[int(m.group(1))]}", t.text)


def add_refs(at):
    p16 = NEW_POS[("N", 1)]
    s = at[PATTERN]
    assert replace_in_paragraphs(s, "→ 配信した方の55%が反応", f"→ 配信した方の55%が反応（P{p16}）") == 1
    assert replace_in_paragraphs(s, f"・配信の出し分け（P{NEW_POS[SEG]}）に取り入れます。",
                                 f"・配信の出し分け（P{NEW_POS[SEG]}）・ステップ配信（P{NEW_POS[STEP]}）に取り入れます。") == 1
    s = at[STEP]
    assert replace_in_paragraphs(s, "配信の中身をこれに合わせる", f"配信の中身をこれに合わせる。明光義塾のステップ配信の例はP{p16}") == 1
    s = at[SOURCES]
    old = f"P14〜P15・P{NEW_POS[PATTERN]}"
    assert replace_in_paragraphs(s, old, f"P14〜P{p16}・P{NEW_POS[PATTERN]}") == 1
    assert replace_in_paragraphs(s, "明光義塾（2024年4月5日）／", "明光義塾（2024年4月5日・掲載の図版を含む）／") == 1


def edit(prs):
    sl = prs.slides
    at = {key: sl[i - 1] for key, i in NEW_POS.items()}
    for n in range(1, N_BASE + 1):
        remap_refs(at[n])
    add_refs(at)
    page_step(at[("N", 1)])
    return sum(fix_fonts(at[k]) for k in (("N", 1), PATTERN, STEP, SOURCES))


def main(base, out):
    mid, tmp = out + ".mid.pptx", out + ".tmp.pptx"
    assemble(base, mid)
    prs = Presentation(mid)
    fixed = edit(prs)
    prs.save(tmp)
    jb.ORDER = ORDER                      # repack は並びの長さだけ見る
    jb.repack(mid, tmp, out)
    Path(mid).unlink()
    Path(tmp).unlink()
    print("saved:", out, "／", len(ORDER), "枚／メイリオにそろえた文字:", fixed)


if __name__ == "__main__":
    main(*sys.argv[1:3])
