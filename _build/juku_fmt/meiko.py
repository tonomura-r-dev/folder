# -*- coding: utf-8 -*-
"""塾ver3：明光義塾を「運用の参考事例」として入れる（2026-10-19 殿村さん。ジムのカーブスと同じ扱い）。

    python -m _build.juku_fmt.recolor <塾ver3（51枚）> <色を直した版>
    python -m _build.juku_fmt.meiko <色を直した版（51枚）> <出力.pptx>

  殿村さん：「業界資料は、運用がうまくいっている企業の事例をメインで扱う」（ジムと同じ）。
           「明光義塾取ってきます」→ スクショ3枚（あいさつ＋相性チェック／LINEの中で体験予約／リッチメニュー）。
           「https://www.lycbiz.com/jp/case-study/line-official-account/meikonet/ が非常に役に立つ」
  - 14枚目（新）＝運用の参考事例｜明光義塾（公式事例の数字＋施策の流れ）
  - 15枚目（新）＝明光義塾の実際の画面（殿村さんのスクショ3枚。_data/参考資料/明光義塾LINE_20261019/）
  - 19枚目（新・区切り「施策の詳細」のあと）＝明光義塾の型を分解すると、3つの仕組み
  - 型を後ろのページに：入口（P23）・診断（P26）・リッチメニュー（P27）・配信の出し分け（P28）に参照を足す。
    診断の「無料体験を予約する」は「LINEの中で受付」に。
  - 13枚目（現状分析）：明光義塾はLINEの中で体験予約を受けている（10/19確認）。
    「診断→体験予約をLINEでつないでいる塾は確認できなかった／まだ空いている」→「確認した4社では明光義塾だけ」。
  - 出典ページに「P14〜P16」の行を足す。本文のページ参照（P◯）は新しい番号に。
  - ステップ配信は追加から日をおいて届くので、まだ見えない（殿村さん「ステップ配信が見えへん」）→ 届いたら足す。
  ほかのページは1バイトも変えない（直したスライドのXMLと、足した3枚・画像だけを差し込む）。
"""
import copy
import random
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

import _build.juku_fmt.build as jb
from _build.arauchi_v4.build import slide_order
from _build.arauchi_v4.lib import add_box
from _build.arauchi_v8.pages import replace_in_paragraphs, shape_by_id
from _build.gym_curves.build import add_pic, keep_only, place, round_box, set_one
from _build.gym_fmt.build import fix_fonts
from _build.juku_fmt.build import CT_SLIDE, RT_SLIDE, rels_of
from _build.juku_fmt.plan import set_lead

ROOT = Path(__file__).resolve().parent.parent.parent
SHOTS = ROOT / "_data/参考資料/明光義塾LINE_20261019"
N_BASE = 51
TPL, DIVIDER = 13, 16                     # 型にする「現状分析｜各社LINEの今」・区切り「施策の詳細」
ORDER = list(range(1, TPL + 1)) + [("N", 1), ("N", 2)] + list(range(TPL + 1, DIVIDER + 1)) + [("N", 3)] + \
        list(range(DIVIDER + 1, N_BASE + 1))
NEW_POS = {key: i for i, key in enumerate(ORDER, 1)}
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
GATE, DIAG, MENU, SEG, SOURCES = 20, 23, 24, 25, 50      # 入口・診断・リッチメニュー・配信の出し分け・出典（元の番号）

NAVY, INK, GREY, GREEN = "1F285A", "333333", "595959", "00897B"
PALE, BORDER = "F4F7FF", "D9D9D9"
X0, W = 0.47, 9.89

CASE_URL = "https://www.lycbiz.com/jp/case-study/line-official-account/meikonet/"
FRIENDS = ("30.8万人", "2026年9月の実測")          # page.line.me（2026-09-25 308,344人。13枚目と同じ）
SCREEN_SHOTS = [
    ("1　あいさつ＋相性チェック", "01_あいさつ_相性チェック.jpg", ["友だち追加の直後に", "30秒の相性チェック（全5問）"]),
    ("2　LINEの中で体験予約", "02_無料体験の予約_LINE内.jpg", ["「このままLINEで予約」（60秒）", "トークの中で入力（全7問）"]),
    ("3　リッチメニュー", "03_リッチメニュー.jpg", ["一番大きい枠に", "授業料シミュレーション"]),
]
SHOT_ASPECT = 277 / 600


def title_of(z, name):
    m = re.search(r"<a:t>([^<]*)</a:t>", z.read(name).decode("utf-8"))
    return m.group(1) if m else ""


# ---------------------------------------------------------------- 1) 組み立て（zipのまま）
def assemble(base, out):
    z = zipfile.ZipFile(base)
    order = slide_order(z)
    assert len(order) == N_BASE
    for n, t in [(TPL, "現状分析｜各社LINEの今"), (DIVIDER, "資料アジェンダ"), (GATE, "構築①｜友だち追加の入口"),
                 (DIAG, "構築③｜30秒ぴったり診断"), (MENU, "構築④｜リッチメニュー"), (SEG, "構築⑤｜配信の出し分け"),
                 (SOURCES, "この資料の出典")]:
        assert title_of(z, order[n - 1]) == t, (n, title_of(z, order[n - 1]))
    assert "施策の詳細" in z.read(order[DIVIDER - 1]).decode("utf-8")
    names = set(z.namelist())
    ct = z.read("[Content_Types].xml").decode("utf-8")
    prels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    pres = z.read("ppt/presentation.xml").decode("utf-8")
    assert '<Default Extension="jpeg"' in ct

    def rid_of_part(name):
        m = re.search(r'<Relationship [^>]*Target="slides/%s"[^>]*/>' % re.escape(name.split("/")[-1]), prels)
        return re.search(r'Id="(rId\d+)"', m.group(0)).group(1)

    rid_of = {n: rid_of_part(order[n - 1]) for n in range(1, N_BASE + 1)}
    next_rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', prels)) + 1
    next_slide = max(int(m.group(1)) for n in names for m in [re.match(r"ppt/slides/slide(\d+)\.xml$", n)] if m) + 1

    tpl = z.read(order[TPL - 1]).decode("utf-8")
    tpl_rels = etree.fromstring(z.read(rels_of(order[TPL - 1])))
    for r in list(tpl_rels):
        if r.get("Type").endswith("/notesSlide"):
            tpl_rels.remove(r)
    assert [r.get("Type").split("/")[-1] for r in tpl_rels] == ["slideLayout"]
    used = set(re.findall(r'creationId[^>]*val="(\d+)"', "".join(z.read(o).decode("utf-8") for o in order)))
    rnd = random.Random(202610192)
    add = {}
    for k in (1, 2, 3):
        dst = f"ppt/slides/slide{next_slide}.xml"
        next_slide += 1
        cid = str(rnd.randrange(10 ** 9, 4 * 10 ** 9))
        assert cid not in used
        xml, c = re.subn(r'(<p14:creationId [^>]*val=")\d+(")', rf"\g<1>{cid}\2", tpl, count=1)
        assert c == 1
        add[dst] = xml.encode("utf-8")
        rels = copy.deepcopy(tpl_rels)
        if k == 2:                                  # 実際の画面：スクショ3枚（パーツ名は英数字だけ）
            for j, (_, fn, _) in enumerate(SCREEN_SHOTS, 1):
                media = f"ppt/media/meiko_{j}.jpeg"
                assert media not in names
                add[media] = (SHOTS / fn).read_bytes()
                etree.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                                 Id=f"rId{100 + j}", Type=RT_IMAGE, Target=f"../media/meiko_{j}.jpeg")
        add[rels_of(dst)] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        ct = ct.replace("</Types>", f'<Override PartName="/{dst}" ContentType="{CT_SLIDE}"/></Types>')
        rid = f"rId{next_rid}"
        next_rid += 1
        prels = prels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{RT_SLIDE}" Target="slides/{dst.split("/")[-1]}"/></Relationships>')
        rid_of[("N", k)] = rid

    m = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", pres, re.S)
    entries = {re.search(r'r:id="(rId\d+)"', e).group(1): e for e in re.findall(r"<p:sldId [^>]*/>", m.group(1))}
    next_id = max(int(x) for x in re.findall(r'<p:sldId id="(\d+)"', m.group(1))) + 1
    lst = []
    for key in ORDER:
        if rid_of[key] in entries:
            lst.append(entries[rid_of[key]])
        else:
            lst.append(f'<p:sldId id="{next_id}" r:id="{rid_of[key]}"/>')
            next_id += 1
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
            zo.writestr(name, data, compress_type=zipfile.ZIP_STORED if name.endswith(".jpeg") else zipfile.ZIP_DEFLATED)


# ---------------------------------------------------------------- 2) 中身
def chip(s, x, y, w, h, text, size=11):
    return round_box(s, x, y, w, h, [[(text, {})]], adj=20000, size=size, bold=True, color="FFFFFF", fill=NAVY,
                     align="ctr", ins=(0.04, 0, 0.04, 0))


def page_case(s):
    """14枚目：運用の参考事例｜明光義塾"""
    keep_only(s, {2, 3, 4, 67, 104})
    set_one(s, 2, "運用の参考事例｜明光義塾")
    set_lead(s, ["LINEで資料請求を伸ばした塾の事例です（LINEヤフー公式の導入事例）。",
                 "「塾はすぐには申し込まれない」前提で、サイトから友だちになってもらい、比べている間をLINEで支えています。"])
    kw, gap = 3.15, (W - 3 * 3.15) / 2
    for i, (lab, big, sub) in enumerate([
            ("友だち数", FRIENDS[0], FRIENDS[1]),
            ("スタンプ施策の約5か月", "約20万人増", "うち約15万人は友だちからの広がり"),
            ("資料請求（2023年6〜7月）", "前年比 約120%", "施策を行った時期")]):
        round_box(s, X0 + i * (kw + gap), 1.65, kw, 1.3,
                  [[(lab, {"size": 11, "bold": True, "color": NAVY})],
                   [(big, {"size": 26, "bold": True, "color": NAVY})],
                   [(sub, {"size": 9.5, "color": GREY})]], adj=6000, fill=PALE, line=BORDER, align="ctr",
                  line_spacing=100)
    add_box(s, X0, 3.1, W, 0.78,
            [[("「学習塾は他塾と念入りに比較検討されるお客さまが多く、Webサイトを訪れてすぐに資料請求や体験授業の"
               "申し込みなどにつながるケースはそこまで多くありません」", {})],
             [("― 明光ネットワークジャパン ご担当者（公式事例より）", {"size": 9, "color": GREY})]],
            size=10.5, color=INK, fill="FFFFFF", line=BORDER, ins=(0.14, 0.05, 0.14, 0.05), line_spacing=105)
    add_box(s, X0, 3.98, W, 0.3, [[("公式事例の施策の流れ", {})]], size=12, bold=True, color=NAVY, ins=(0, 0, 0, 0))
    flow = [("サイトで", ["初回訪問時・離脱時の", "ポップアップから", "友だち追加"]),
            ("追加のあと", ["チャットと配信で", "良さ・他塾との違いを", "じっくり伝える"]),
            ("スタンプで", ["タップで答えるアンケートに", "答えるとスタンプを配布", "→ 友だちから広がる"]),
            ("配信で", ["回答をもとに、", "まだ塾に通っていない", "お子さま・保護者に絞る"])]
    cw = (W - 3 * 0.12) / 4
    for i, (head, lines) in enumerate(flow):
        x = X0 + i * (cw + 0.12)
        add_box(s, x, 4.62, cw, 1.1, [[(t, {})] for t in lines], size=10.5, color=INK, fill="FFFFFF", line=BORDER,
                align="ctr", anchor="ctr", ins=(0.04, 0.04, 0.04, 0.04), line_spacing=110)
        chip(s, x, 4.3, cw, 0.32, f"{i + 1}　{head}", size=10.5)
    add_box(s, X0, 5.8, W, 0.5,
            [[("※ スタンプで来た方には、後日クイズ形式の配信で回答を促し、配信した方の55%が反応。友だちが増えても、ブロック率は施策前とほぼ変わらず。", {})],
             [("※ 明光義塾は、チャットの接客にチャットコマース「ジールス」（株式会社Zeals）を使用。", {})]],
            size=9, color=GREY, ins=(0, 0, 0, 0), line_spacing=100)
    set_one(s, 104, "大手でも、やっていることは「友だちになってもらい、比べている間をLINEで支える」こと。")
    set_one(s, 67, "出典｜LINEヤフー for Business 導入事例 明光義塾（2024年4月5日）。友だち数：page.line.me（2026-09-25実測）")
    return s


def page_screens(s):
    """15枚目：明光義塾の実際の画面（スクショ3枚）"""
    keep_only(s, {2, 3, 4, 67})
    set_one(s, 2, "明光義塾の実際の画面")
    set_lead(s, ["友だち追加の直後に、30秒の相性チェックと「このままLINEで無料体験の予約」を案内しています。",
                 "リッチメニューの一番大きい枠は「授業料シミュレーション」。費用の不安に先回りしています。"])
    colw = 3.0
    gap = (W - 3 * colw) / 2
    ih = 4.1
    iw = ih * SHOT_ASPECT
    for j, (head, _, cap) in enumerate(SCREEN_SHOTS):
        x = X0 + j * (colw + gap)
        chip(s, x, 1.62, colw, 0.34, head, size=11.5)
        add_pic(s, f"rId{101 + j}", f"明光義塾_{head[2:]}", x + (colw - iw) / 2, 2.06, iw, ih)
        add_box(s, x, 6.22, colw, 0.5, [[(t, {})] for t in cap], size=10.5, color=INK, align="ctr", anchor="t",
                ins=(0, 0, 0, 0), line_spacing=105)
    set_one(s, 67, "画面：明光義塾 LINE公式アカウント（2026-10-19 弊社で友だち追加して確認）")
    return s


def page_pattern(s):
    """19枚目：明光義塾の型を分解すると、3つの仕組み"""
    keep_only(s, {2, 3, 4, 67, 104})
    set_one(s, 2, "明光義塾の型を分解すると、3つの仕組み")
    set_lead(s, ["やっていることは3つ。どれも「すぐには申し込まれない」前提でつくられています。",
                 f"この型を、このあとの入口（P{NEW_POS[GATE]}）・診断（P{NEW_POS[DIAG]}）・"
                 f"配信の出し分け（P{NEW_POS[SEG]}）に取り入れます。"])
    kw, gap = 3.15, (W - 3 * 3.15) / 2
    cols = [
        ("① サイトから、LINEの友だちに",
         ["サイトの初回訪問時・離脱時に、ポップアップで友だち追加を案内する。",
          "スタンプの無料配布で、友だちから友だちへ広げる。"],
         ["約5か月で友だち約20万人増", "うち約15万人は友だちからの広がり"]),
        ("② 追加の直後に、診断から予約まで",
         ["友だち追加の直後に、30秒の相性チェック（ボタンで答える5問）。",
          "そのまま「LINEで無料体験の予約」へ。予約はトークの中で入力できる。"],
         ["相性チェック 全5問", "体験予約 60秒・全7問"]),
        ("③ 答えをもとに、配信の相手を絞る",
         ["アンケートで、保護者か生徒か、塾に通っているかを聞く。",
          "「まだ塾に通っていない」お子さま・保護者に絞って配信する。"],
         ["資料請求 前年比 約120%", "（2023年6〜7月）"]),
    ]
    for i, (head, body, ex) in enumerate(cols):
        x = X0 + i * (kw + gap)
        chip(s, x, 1.65, kw, 0.42, head, size=12)
        round_box(s, x, 2.17, kw, 2.15, [[(t, {})] for t in body], adj=5000, size=12.5, color=INK, fill=PALE,
                  line=BORDER, anchor="t", ins=(0.12, 0.12, 0.1, 0.06), after=8, line_spacing=110)
        round_box(s, x, 4.45, kw, 1.35,
                  [[("実際の例", {"size": 10.5, "bold": True, "color": NAVY})]] + [[(t, {})] for t in ex],
                  adj=5000, size=11.5, color=INK, fill="FFFFFF", line=BORDER, anchor="t", ins=(0.1, 0.08, 0.08, 0.04),
                  line_spacing=105)
    set_one(s, 104, "友だちになってもらい、答えてもらい、合う方に届ける。この順番が成果につながっている。")
    set_one(s, 67, f"※ 出典：LINEヤフー for Business 導入事例（2024年4月5日）・明光義塾 LINE公式アカウント（2026-10-19 弊社確認）。"
                   f"実際の画面はP{NEW_POS[('N', 2)]}。")
    return s


def remap_refs(s):
    """本文中の「P◯」（51枚の番号）を新しい番号に"""
    for t in s._element.iter(qn("a:t")):
        if t.text and re.search(r"P\d+", t.text):
            t.text = re.sub(r"P(\d+)", lambda m: f"P{NEW_POS[int(m.group(1))]}", t.text)


def set_runs(sh, texts):
    """段落ごとに、最初のrunの文字だけ差し替える（書式はそのまま）"""
    ps = sh.text_frame.paragraphs
    assert len(ps) == len(texts), (sh.shape_id, len(ps))
    for p, t in zip(ps, texts):
        if t is None:
            continue
        runs = p.runs
        runs[0].text = t
        for r in runs[1:]:
            r._r.getparent().remove(r._r)


def fix_landscape(s):
    """13枚目：明光義塾はLINEの中で体験予約を受けている（10/19確認）→「空いている」の言い方を直す"""
    cell = shape_by_id(s, 79)
    set_runs(cell, ["あり", "（体験予約はLINEの中で）"])
    for c in cell.text_frame.paragraphs[0].runs[0]._r.iter(qn("a:srgbClr")):
        c.set("val", GREEN)
    set_runs(shape_by_id(s, 102), ["明光義塾はLINEで成果を出している（公式事例・P14）", None, None, None])
    set_runs(shape_by_id(s, 103), [None, "・保護者向けに「診断 → 体験予約」をLINEの中で",
                                   "　つないでいるのは、確認した4社では明光義塾だけ", None])
    set_one(s, 104, "保護者向けに、診断から体験予約までをLINEでつないでいる塾は、まだ少ない。")


def add_refs(at):
    """型を後ろのページに：入口・診断・リッチメニュー・配信の出し分け"""
    p14, p15 = NEW_POS[("N", 1)], NEW_POS[("N", 2)]
    s = at[GATE]
    assert replace_in_paragraphs(s, "Sitelead等（別途費用）",
                                 f"Sitelead等（別途費用）。明光義塾も、サイトの初回訪問時・離脱時のポップアップで友だちを集めている（P{p14}）") == 1
    s = at[DIAG]
    sh = shape_by_id(s, 68)
    assert sh.text_frame.paragraphs[1].text == "CV② 体験予約"
    sh.text_frame.paragraphs[1].runs[0].text = "CV② 体験予約（LINEの中で受付）"
    note = shape_by_id(s, 31)
    assert note.text_frame.text == "※競合の診断（トライの性格別学習法診断など）はP13の調査結果を参照"
    set_runs(note, [f"※明光義塾は、相性チェック（5問）のあと、LINEの中で体験予約まで受け付けている（P{p15}）。"
                    "競合の診断（トライの性格別学習法診断など）はP13を参照"])
    s = at[MENU]
    assert replace_in_paragraphs(s, "日程変更など）を出し分ける",
                                 f"日程変更など）を出し分ける。明光義塾は、一番大きい枠に「授業料シミュレーション」を置いている（P{p15}）") == 1
    s = at[SEG]                                   # 注記の枠が無いので、リッチメニューの注記を写して足す
    src = shape_by_id(at[MENU], 45)
    el = copy.deepcopy(src._element)
    el.find(qn("p:nvSpPr")).find(qn("p:cNvPr")).set("id", str(max(x.shape_id for x in s.shapes) + 1))
    s.shapes._spTree.append(el)
    new = s.shapes[-1]
    set_runs(new, [f"※明光義塾は、アンケートの回答をもとに「まだ塾に通っていない」お子さま・保護者に絞って配信している（P{p14}）"])


def add_source_row(s):
    """出典ページ：P13の下に明光義塾の3枚の行を足し、下の2行を1行ぶん下げる"""
    by_id = {sh.shape_id: sh for sh in s.shapes}
    assert by_id[25].text_frame.text == "P13" and by_id[27].text_frame.text == f"P{NEW_POS[21]}・P{NEW_POS[31]}"
    row_y, step = by_id[27].top, by_id[27].top - by_id[25].top
    pages = f"P{NEW_POS[('N', 1)]}・P{NEW_POS[('N', 2)]}・P{NEW_POS[('N', 3)]}"
    for src_id, text in ((27, pages),
                         (28, "LINEヤフー for Business 導入事例 明光義塾（2024年4月5日）／明光義塾 LINE公式アカウント（2026-10-19 弊社確認）")):
        el = copy.deepcopy(by_id[src_id]._element)
        el.find(qn("p:nvSpPr")).find(qn("p:cNvPr")).set("id", str(max(x.shape_id for x in s.shapes) + 1))
        by_id[30]._element.addnext(el)
        new = [x for x in s.shapes if x._element is el][0]
        set_runs(new, [text])
        new.top = row_y
    for sid in (27, 28, 29, 30):
        by_id[sid].top = by_id[sid].top + step
    fill = {sid: copy.deepcopy(by_id[sid]._element.spPr.find(qn("a:solidFill"))) for sid in (25, 27)}
    for sid, src in ((27, 25), (28, 25), (29, 27), (30, 27)):          # 行の色の交互（薄い紺・白）をそろえる
        old = by_id[sid]._element.spPr.find(qn("a:solidFill"))
        old.addprevious(copy.deepcopy(fill[src]))
        old.getparent().remove(old)
    assert by_id[29].top + by_id[29].height < by_id[31].top
    assert replace_in_paragraphs(s, "競合の実名掲載（P13）",
                                 f"競合の実名掲載（P13・{pages}）") == 1


def edit(prs):
    sl = prs.slides
    at = {key: sl[i - 1] for key, i in NEW_POS.items()}
    for n in range(1, N_BASE + 1):
        remap_refs(at[n])
    fix_landscape(at[TPL])
    add_refs(at)
    add_source_row(at[SOURCES])
    page_case(at[("N", 1)])
    page_screens(at[("N", 2)])
    page_pattern(at[("N", 3)])
    return sum(fix_fonts(at[k]) for k in (("N", 1), ("N", 2), ("N", 3), TPL, GATE, DIAG, MENU, SEG, SOURCES))


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
