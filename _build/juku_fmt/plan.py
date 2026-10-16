# -*- coding: utf-8 -*-
"""塾ver3の「具体プラン提示」を、殿村さんの提案の組み立てに作り替える（2026-10-16 殿村さん指示）。

    python -m _build.juku_fmt.plan <塾ver3（48枚）> <出力.pptx>

  殿村さん：
    ○改善したいフローの8フェーズのうちでの最大の課題
    ○予算の確認（いくら出せるか、CPAいくらならGO出るか）
    ○松竹梅の提示（初訪）
        松＝初期＋4-5投稿（週1）＋10ステップ配信／API連携等を含めてここまでできます
        竹＝初期＋投稿1-2＋新規獲得に向けたツール・機能の導入
        梅＝ツール導入支援＋Fee（手数料）／サンクスLINE誘導のみ＋初期＋Fee（手数料）
    ○松竹梅の提示（クロージング）※イメージ
        松 MAX＝開発含むプラン（Hachidoriなどと連携）※最適運用プランで提示
        竹 MID＝初期＋4-5投稿（週1）＋10ステップ配信
        梅 MIN＝サンクスLINE誘導のみ＋初期＋Fee（手数料）
  → 区切り「具体プラン提示」（37枚目）のあとに4枚：最大の課題／ご予算の確認／松竹梅（初訪）／松竹梅（クロージング・非表示）。
    旧「ご提供プラン（運用構成）」（月額20万円〜（3投稿））は松竹梅に置き換えて外す。金額は未確定なので「◯」。

  1) 組み立て：48枚のzipに、全体設計想定（14枚目）のXMLを型にした4枚を足し、旧ご提供プランを外して並べ替える
  2) 中身：python-pptx で4枚を組み、直したスライドのXMLだけを1)のzipに差し込む（python-pptxで保存し直さない）
"""
import random
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

from _build.arauchi_v4.build import slide_order
from _build.arauchi_v4.lib import NAVY, RED, TEXT, add_box
from _build.arauchi_v8.pages import replace_in_paragraphs
from _build.gym_fmt.build import fix_fonts, set_title
from _build.juku_fmt.build import CT_SLIDE, RT_SLIDE, rels_of, repack

N_BASE = 48
TEMPLATE = 14            # 全体設計想定（見出し・説明・区切り線・8フェーズの帯がある）
OLD_PLAN = 37            # 旧「ご提供プラン（運用構成）」
AFTER = 36               # 区切り「資料アジェンダ：具体プラン提示」
N_NEW = 4
HIDDEN = [4]             # 新しい4枚のうち、クロージング用は非表示
ORDER = list(range(1, AFTER + 1)) + [("N", k) for k in range(1, N_NEW + 1)] + \
        [n for n in range(AFTER + 1, N_BASE + 1) if n != OLD_PLAN]
NEW_POS = {key: i for i, key in enumerate(ORDER, 1)}
KEEP_IDS = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}   # 見出し・説明・区切り線・上の8フェーズの帯

PHASES = ["①接触", "②離脱", "③育成", "④リード獲得", "⑤リード有効化", "⑥再育成", "⑦入会（契約）", "⑧紹介"]
TITLE = "002060"
TINT = "C9D3E6"          # 帯（全部）・薄い紺
BG = "F4F7FF"
LINE = "D9D9D9"
MID = "3E5B9A"           # 竹
OFF_FILL, OFF_TXT = "EEEEEE", "A6A6A6"
X0, COL_W, GAP = 0.47, 3.15, 0.22
COLS = [X0 + i * (COL_W + GAP) for i in range(3)]


# ---------------------------------------------------------------- 1) 組み立て
def assemble(base, out):
    z = zipfile.ZipFile(base)
    order = slide_order(z)
    assert len(order) == N_BASE
    names = set(z.namelist())
    ct = z.read("[Content_Types].xml").decode("utf-8")
    prels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    pres = z.read("ppt/presentation.xml").decode("utf-8")

    def rid_for(name):
        m = re.search(r'<Relationship [^>]*Target="slides/%s"[^>]*/>' % re.escape(name.split("/")[-1]), prels)
        return m, re.search(r'Id="(rId\d+)"', m.group(0)).group(1)

    rid_of = {n: rid_for(order[n - 1])[1] for n in range(1, N_BASE + 1)}
    next_rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', prels)) + 1
    next_slide = max(int(m.group(1)) for n in names for m in [re.match(r"ppt/slides/slide(\d+)\.xml$", n)] if m) + 1

    # 型：14枚目のXML（ノートは付けない・creationIdは振り直す）
    tpl = z.read(order[TEMPLATE - 1]).decode("utf-8")
    tpl_rels = etree.fromstring(z.read(rels_of(order[TEMPLATE - 1])))
    for r in list(tpl_rels):
        if r.get("Type").endswith("/notesSlide"):
            tpl_rels.remove(r)
    assert [r.get("Type").split("/")[-1] for r in tpl_rels] == ["slideLayout"]
    used = set(re.findall(r'creationId[^>]*val="(\d+)"', "".join(z.read(o).decode("utf-8") for o in order)))
    rnd = random.Random(20261016)
    add = {}
    for k in range(1, N_NEW + 1):
        dst = f"ppt/slides/slide{next_slide}.xml"
        next_slide += 1
        cid = str(rnd.randrange(10 ** 9, 4 * 10 ** 9))
        assert cid not in used
        xml = re.sub(r'(<p14:creationId [^>]*val=")\d+(")', rf"\g<1>{cid}\2", tpl, count=1)
        if k in HIDDEN:
            xml = xml.replace("<p:sld ", '<p:sld show="0" ', 1)
        add[dst] = xml.encode("utf-8")
        add[rels_of(dst)] = etree.tostring(tpl_rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        ct = ct.replace("</Types>", f'<Override PartName="/{dst}" ContentType="{CT_SLIDE}"/></Types>')
        rid = f"rId{next_rid}"
        next_rid += 1
        prels = prels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{RT_SLIDE}" Target="slides/{dst.split("/")[-1]}"/></Relationships>')
        rid_of[("N", k)] = rid

    # 旧ご提供プランを外す（ほかから参照されていないこと）
    old = order[OLD_PLAN - 1]
    for r in etree.fromstring(z.read(rels_of(old))):
        assert r.get("Type").endswith("/slideLayout")
    for n in names:
        if n.endswith(".rels") and n not in (rels_of(old), "ppt/_rels/presentation.xml.rels"):
            assert old.split("/")[-1] not in z.read(n).decode("utf-8"), n
    gone = {old, rels_of(old)}
    m, _ = rid_for(old)
    prels = prels.replace(m.group(0), "")
    ct, k = re.subn(r'<Override [^>]*PartName="/%s"[^>]*/>' % re.escape(old), "", ct)
    assert k == 1

    # p:sldIdLst を新しい並びに
    m = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", pres, re.S)
    entries = {re.search(r'r:id="(rId\d+)"', e).group(1): e for e in re.findall(r"<p:sldId [^>]*/>", m.group(1))}
    next_id = max(int(x) for x in re.findall(r'<p:sldId id="(\d+)"', m.group(1))) + 1
    lst = []
    for key in ORDER:
        rid = rid_of[key]
        if rid in entries:
            lst.append(entries[rid])
        else:
            lst.append(f'<p:sldId id="{next_id}" r:id="{rid}"/>')
            next_id += 1
    pres = pres.replace(m.group(0), "<p:sldIdLst>" + "".join(lst) + "</p:sldIdLst>")

    # docProps/app.xml（スライド数・タイトル一覧）
    app = z.read("docProps/app.xml").decode("utf-8")
    n_new, ppt_t = len(ORDER), "<vt:lpstr>PowerPoint プレゼンテーション</vt:lpstr>"
    assert f"<Slides>{N_BASE}</Slides>" in app and app.count(ppt_t) == N_BASE
    app = app.replace(f"<Slides>{N_BASE}</Slides>", f"<Slides>{n_new}</Slides>")
    app = re.sub(r"(<vt:lpstr>スライド タイトル</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+(</vt:i4>)", rf"\g<1>{n_new}\2", app)
    mv = re.search(r'<vt:vector size="(\d+)" baseType="lpstr">', app)
    app = app.replace(mv.group(0), f'<vt:vector size="{int(mv.group(1)) + n_new - N_BASE}" baseType="lpstr">')
    i = app.rfind(ppt_t) + len(ppt_t)
    app = app[:i] + ppt_t * (n_new - N_BASE) + app[i:]

    with zipfile.ZipFile(out, "w") as zo:
        for info in z.infolist():
            if info.filename in gone:
                continue
            data = {"ppt/presentation.xml": pres.encode("utf-8"),
                    "ppt/_rels/presentation.xml.rels": prels.encode("utf-8"),
                    "[Content_Types].xml": ct.encode("utf-8"),
                    "docProps/app.xml": app.encode("utf-8")}.get(info.filename)
            zo.writestr(info, data if data is not None else z.read(info.filename), compress_type=info.compress_type)
        for name, data in add.items():
            zo.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)


# ---------------------------------------------------------------- 2) 中身
def clear_body(s):
    for sh in list(s.shapes):
        if sh.shape_id not in KEEP_IDS:
            sh._element.getparent().remove(sh._element)


def set_lead(s, lines):
    """説明（TextBox 2）を1〜2行に。1行目の書式のまま"""
    sh = [x for x in s.shapes if x.shape_id == 3][0]
    tb = sh.text_frame._txBody
    ps = tb.findall(qn("a:p"))
    for p in ps[1:]:
        tb.remove(p)
    first = ps[0]
    for k, text in enumerate(lines):
        p = first if k == 0 else etree.fromstring(etree.tostring(first))
        runs = p.findall(qn("a:r"))
        runs[0].find(qn("a:t")).text = text
        for r in runs[1:]:
            p.remove(r)
        if k:
            tb.append(p)


def set_header(s, on=None):
    """上の8フェーズの帯。on＝強調するフェーズの番号（Noneなら全部を薄い紺＝元の型のまま）"""
    shapes = sorted([x for x in s.shapes if 5 <= x.shape_id <= 12], key=lambda x: x.shape_id)
    assert [x.text_frame.text for x in shapes] == PHASES
    if on is None:
        return
    for i, sh in enumerate(shapes, 1):
        fill, txt = ("1F285A", "FFFFFF") if i == on else (OFF_FILL, OFF_TXT)
        sh._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr")).set("val", fill)
        for c in sh._element.txBody.iter(qn("a:srgbClr")):
            c.set("val", txt)


def phase_shape(s, x, y, w, h, text, fill, color, first):
    shp = add_box(s, x, y, w, h, [[(text, {})]], size=10, bold=True, color=color, fill=fill, align="ctr",
                  ins=(0.1, 0.02, 0.02, 0.02))
    geom = shp._element.spPr.find(qn("a:prstGeom"))
    geom.set("prst", "homePlate" if first else "chevron")
    return shp


def corner(shp, adj):
    """角の丸み（短い辺に対する割合×100000）"""
    av = shp._element.spPr.find(qn("a:prstGeom")).find(qn("a:avLst"))
    for g in list(av):
        av.remove(g)
    etree.SubElement(av, qn("a:gd"), name="adj", fmla=f"val {adj}")
    return shp


def hang(shp, inch=0.16):
    """箇条書きの2行目以降を「・」のあとにそろえる"""
    for p in shp.text_frame._txBody.findall(qn("a:p")):
        ppr = p.find(qn("a:pPr"))
        ppr.set("marL", str(int(inch * 914400)))
        ppr.set("indent", str(-int(inch * 914400)))
    return shp


def card(s, x, y, w, h, line=LINE, line_w=0.75, fill="FFFFFF"):
    return corner(add_box(s, x, y, w, h, [[("", {})]], fill=fill, line=line, line_w=line_w, shape="round"), 5000)


def page_issue(s):
    """最大の課題（8フェーズのうち）"""
    set_title(s.shapes[0] if s.shapes[0].shape_id == 2 else [x for x in s.shapes if x.shape_id == 2][0],
              "改善したいフロー｜8フェーズのうち最大の課題")
    set_lead(s, ["入会までの8フェーズのうち、保護者が一番離れているのはどこか。",
                 "塾では、資料請求から体験予約までの間（⑤リード有効化）が最大の課題になりやすいと考えています（仮説）。"])
    set_header(s, on=5)
    step, w = 9.89 / 8, 1.21
    for i, name in enumerate(PHASES):
        if i == 4:
            fill, color = NAVY, "FFFFFF"
        elif i in (1, 5):
            fill, color = TINT, NAVY
        else:
            fill, color = OFF_FILL, OFF_TXT
        phase_shape(s, X0 + i * step, 1.75, w, 0.42, name, fill, color, i == 0)
    x5 = X0 + 4 * step
    add_box(s, x5 + w / 2 - 0.8, 2.25, 1.6, 0.3, [[("最大の課題（仮説）", {})]], size=10, bold=True, color="FFFFFF",
            fill=RED, align="ctr", shape="round", ins=(0.02, 0, 0.02, 0))

    cards = [
        ("② 離脱", "サイトに来た保護者の多くが、名前を残さずに帰っている。", "離脱防止ポップアップで、帰る前にLINEの友だちになってもらう。", False),
        ("⑤ リード有効化", "資料請求のあと、体験予約までの間に「読んで・比べて・迷う」期間があり、ここで他の塾に流れる。",
         "迷っている間も配信を届け続け、体験予約につなげる。", True),
        ("⑥ 再育成", "体験に来なかった・入会しなかった保護者との接点が、そこで切れる。", "時期に合わせた企画配信で、もう一度ご案内する。", False),
    ]
    y0, h = 2.75, 2.85
    for x, (head, now, line_, main) in zip(COLS, cards):
        card(s, x, y0, COL_W, h, line=NAVY if main else LINE, line_w=2.25 if main else 0.75)
        corner(add_box(s, x, y0, COL_W, 0.45, [[(head, {})]], size=14, bold=True, color="FFFFFF" if main else NAVY,
                       fill=NAVY if main else TINT, align="ctr", shape="round"), 12000)
        add_box(s, x + 0.15, y0 + 0.55, COL_W - 0.3, 1.2,
                [[("起きていること", {"size": 10, "bold": True, "color": "7F7F7F"})], [(now, {})]],
                size=12.5, color=TEXT, anchor="t", ins=(0, 0, 0, 0), line_spacing=115)
        add_box(s, x + 0.15, y0 + 1.8, COL_W - 0.3, 0.95,
                [[("LINEでの打ち手", {"size": 10, "bold": True, "color": "7F7F7F"})], [(line_, {"bold": True, "color": NAVY})]],
                size=12.5, color=TEXT, anchor="t", ins=(0, 0, 0, 0), line_spacing=115)
    add_box(s, X0, 5.9, 9.89, 0.5, [[("本日、貴塾で一番改善したいフェーズを確認し、そこに合わせてプランをご提案します。", {})]],
            size=14, bold=True, color="FFFFFF", fill=NAVY, align="ctr", shape="round")


def page_budget(s):
    """ご予算の確認（月々のご予算・獲得単価の目安）"""
    set_title([x for x in s.shapes if x.shape_id == 2][0], "ご予算の確認")
    set_lead(s, ["プランをお選びいただく前に、2つ確認させてください。"])
    set_header(s)
    items = [
        ("1", "月々のご予算", "LINEの運用に、月々いくらまで\nかけられますか。", "月額　＿＿＿＿　万円まで"),
        ("2", "獲得単価（CPA）の目安", "資料請求・体験予約の1件あたり、\nいくらまでなら導入をご判断いただけますか。", "1件あたり　＿＿＿＿　円まで"),
    ]
    w, y0, h = 4.83, 1.85, 3.5
    for k, (no, head, q, fill_in) in enumerate(items):
        x = X0 + k * (w + 0.23)
        card(s, x, y0, w, h, fill=BG)
        add_box(s, x + 0.25, y0 + 0.3, 0.5, 0.5, [[(no, {})]], size=18, bold=True, color="FFFFFF", fill=NAVY,
                align="ctr", shape="round", ins=(0, 0, 0, 0))
        add_box(s, x + 0.9, y0 + 0.3, w - 1.1, 0.5, [[(head, {})]], size=18, bold=True, color=NAVY, ins=(0, 0, 0, 0))
        add_box(s, x + 0.25, y0 + 1.0, w - 0.5, 1.0, [[(q, {})]], size=13, color=TEXT, anchor="t", ins=(0, 0, 0, 0),
                line_spacing=120)
        corner(add_box(s, x + 0.25, y0 + 2.35, w - 0.5, 0.8, [[(fill_in, {})]], size=18, bold=True, color=NAVY,
                       fill="FFFFFF", line=NAVY, dash="dash", align="ctr", shape="round"), 10000)
    add_box(s, X0, 5.65, 9.89, 0.5, [[("この2つをもとに、次のページの松・竹・梅から、貴塾に合うプランをご提案します。", {})]],
            size=14, bold=True, color="FFFFFF", fill=NAVY, align="ctr", shape="round")


PLAN_LEAD = ["ご予算と最大の課題に合わせて、3つのプランからお選びいただけます。",
             "6か月〜・税抜。金額は貴塾の教室数・配信の規模に合わせて確定します。"]
HEAD_STYLE = [(NAVY, "FFFFFF"), (MID, "FFFFFF"), (TINT, NAVY)]       # 松・竹・梅


def plan_columns(s, plans):
    """plans = [(名前, 目安の言葉, ひとこと, [項目], 費用)]×3"""
    y0 = 1.75
    for x, (fill, txt), (name, tag, tagline, items, fee) in zip(COLS, HEAD_STYLE, plans):
        runs = [(name, {"size": 26})] + ([("　" + tag, {"size": 12})] if tag else [])
        corner(add_box(s, x, y0, COL_W, 0.95, [runs, [(tagline, {"size": 11.5, "bold": False})]], size=26, bold=True,
                       color=txt, fill=fill, align="ctr", shape="round", ins=(0.06, 0.02, 0.06, 0.04), line_spacing=100),
               10000)
        card(s, x, y0 + 1.05, COL_W, 2.75)
        hang(add_box(s, x + 0.18, y0 + 1.25, COL_W - 0.32, 2.45,
                     [[("・" + t, {"size": 11, "color": "7F7F7F"} if t.startswith("例：") else {})] for t in items],
                     size=13, color=TEXT, anchor="t", ins=(0, 0, 0, 0), after=6, line_spacing=115))
        corner(add_box(s, x, y0 + 3.9, COL_W, 0.6, [[(fee, {})]], size=12.5, bold=True, color=NAVY, fill=BG, line=LINE,
                       align="ctr", shape="round", ins=(0.04, 0.02, 0.04, 0.02)), 12000)


def page_plan_first(s):
    """松竹梅（初訪）"""
    set_title([x for x in s.shapes if x.shape_id == 2][0], "ご提供プラン｜松・竹・梅")
    set_lead(s, PLAN_LEAD)
    set_header(s)
    plan_columns(s, [
        ("松", "", "構築から配信の運用まで",
         ["初期構築", "企画投稿 月4〜5本（週1回）", "ステップ配信 10通", "API連携なども含めて、ここまで対応できます"],
         "初期費用 ◯万円／月額 ◯万円"),
        ("竹", "", "構築＋新しい友だちを増やす",
         ["初期構築", "企画投稿 月1〜2本", "新しい友だちを増やすツール・機能の導入", "例：離脱防止ポップアップ"],
         "初期費用 ◯万円／月額 ◯万円"),
        ("梅", "", "ツールの導入を支援",
         ["ツールの導入支援", "例：サンクスLINE誘導のみ（申し込み完了画面からLINEへ）", "初期構築", "月額は手数料のみ"],
         "初期費用 ◯万円／月額 手数料 ◯万円"),
    ])


def page_plan_closing(s):
    """松竹梅（クロージング）＝非表示。松を開発込みの最適運用プランに上げ、初訪の松を竹に"""
    set_title([x for x in s.shapes if x.shape_id == 2][0], "ご提供プラン｜松・竹・梅")
    set_lead(s, PLAN_LEAD)
    set_header(s)
    plan_columns(s, [
        ("松", "MAX", "最適運用プラン（開発を含む）",
         ["開発を含むプラン", "Hachidoriなどのツールと連携", "貴塾に合わせた最適な運用"],
         "初期費用 ◯万円／月額 ◯万円"),
        ("竹", "MID", "構築から配信の運用まで",
         ["初期構築", "企画投稿 月4〜5本（週1回）", "ステップ配信 10通"],
         "初期費用 ◯万円／月額 ◯万円"),
        ("梅", "MIN", "サンクスLINE誘導のみ",
         ["サンクスLINE誘導のみ", "例：申し込み完了画面からLINEへ", "初期構築", "月額は手数料のみ"],
         "初期費用 ◯万円／月額 手数料 ◯万円"),
    ])


def edit(prs):
    sl = prs.slides
    at = {key: sl[i - 1] for key, i in NEW_POS.items()}
    for k, fn in zip(range(1, N_NEW + 1), (page_issue, page_budget, page_plan_first, page_plan_closing)):
        s = at[("N", k)]
        clear_body(s)
        fn(s)
    # 出典ページの「P46」（企画のご参考）を新しい番号に
    src = at[N_BASE - 1]
    assert replace_in_paragraphs(src, "P46", f"P{NEW_POS[46]}") == 1
    return sum(fix_fonts(at[("N", k)]) for k in range(1, N_NEW + 1))


def main(base, out):
    mid, tmp = out + ".mid.pptx", out + ".tmp.pptx"
    assemble(base, mid)
    prs = Presentation(mid)
    fixed = edit(prs)
    prs.save(tmp)
    import _build.juku_fmt.build as jb
    jb.ORDER = ORDER                      # repack は並びの長さだけ見る
    repack(mid, tmp, out)
    Path(mid).unlink()
    Path(tmp).unlink()
    print("saved:", out, "／", len(ORDER), "枚／メイリオにそろえた文字:", fixed)


if __name__ == "__main__":
    main(*sys.argv[1:3])
