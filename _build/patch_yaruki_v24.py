# -*- coding: utf-8 -*-
"""やる気スイッチ ver2_3（殿村さん最新版）→ ver2_4（2026-09-30 上司の指摘を反映）

  S5   本提案の考え方に「CPAの改善・CPOの改善 → 許容CPAを引き上げられる」を組み込む（資料の大前提）
  S7   リードの誤り（前後検索ではなく調査データのページ）を直す
  S14  前提整理を図にする
  S20  6か月後の欄にCPA・CPOを追加。直後にCPA・CPOの月別推移（1〜6か月目）を新設
  S22  アジェンダを「施策の詳細（既存の効率化）」に。直後に全体の導線（S17の友だち獲得部分をぼかした版）
  S24  入会12件 → 11件（体験予約44件×25%）
  S27  QA自動化・個別チャットの中身を習い事向けに直す（別業種の文面が残っていた）
  S28  直後に全体の導線（S17の配信部分をぼかした版）
  S35  サンクスLINE誘導の費用 初期15万円・月額5万円。直後に【オプション】QA自動化（チャットボット）を新設
  S37  成果報酬型の記載。直後に「CPO改善のご提案｜サンクスLINE誘導」を新設
  S38  サンクスLINE誘導ツール 初期¥150,000・月額¥50,000
  S8・S9 前後検索「子ども 習い事」「習い事」：タイトルを 前後検索「KW名」 に、画像を色分け版に、リードを分析の結論に、凡例と出典を追加
  S10  前後検索のまとめを「項目別」（色分けと同じ3項目 × 検索前・当日・後 ＋ LINEでの打ち手）に作り直す
       色分け画像は _build/make_yaruki_zengo_color.py で作る（分類は _data/zengo/やる気スイッチ_クエリ分類_*.csv）

数字の出どころ：株式会社やる気スイッチグループ_LINEOA施策提案【SIM】_ver3.00.xlsx（成果報酬運営）。
件数・金額は整数で出す（小数は書かない。殿村さん指示）。

  python3 _build/patch_yaruki_v24.py <ver2_3.pptx>
"""
import sys
from copy import deepcopy
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_src = (ROOT / "_build/build_chengrowth_v11.py").read_text(encoding="utf-8")
_helpers = _src[_src.index("# ================= helpers"):_src.index("# ============================================================\n# 流用するページの素材")]
_head = _src[_src.index("import shutil"):_src.index("shutil.copyfile(SRC, OUT)")]
exec(_head)
exec(_helpers)
SRC = Path(sys.argv[1])
OUT = ROOT / "202609_株式会社やる気スイッチグループ御中_LINE公式アカウント運用のご提案_ver2_4.pptx"
prs = Presentation(str(SRC))
S = list(prs.slides)
DGREEN, GRAY = "0B7A3B", "7F7F7F"
PBLUE = "DCE6F5"


# ============================================================
# 部品
# ============================================================
def by(s, name, nth=0):
    return [x for x in s.shapes if x.name == name][nth]


def by_text(s, text):
    return next(x for x in s.shapes if x.has_text_frame and text in x.text_frame.text)


def drop(*shapes):
    for sh in shapes:
        sh._element.getparent().remove(sh._element)


def set_lines(sh, lines, sz=None, color=None, bold=None, algn=None):
    """段落ごと書き換える（最初の段落の書式・最初のrunの書式を引き継ぐ）"""
    tx = sh.text_frame._txBody
    ps = tx.findall(qn("a:p"))
    ppr = ps[0].find(qn("a:pPr"))
    r0 = next(tx.iter(qn("a:r")))
    rpr = r0.find(qn("a:rPr"))
    for p in ps:
        tx.remove(p)
    for line in lines:
        p = tx.makeelement(qn("a:p"), {})
        if ppr is not None:
            p.append(deepcopy(ppr))
            if algn:
                p[0].set("algn", algn)
        elif algn:
            p.append(p.makeelement(qn("a:pPr"), {"algn": algn}))
        r = p.makeelement(qn("a:r"), {})
        rp = deepcopy(rpr) if rpr is not None else r.makeelement(qn("a:rPr"), {"lang": "ja-JP"})
        if sz:
            rp.set("sz", str(int(sz * 100)))
        if bold is not None:
            rp.set("b", "1" if bold else "0")
        if color:
            for f in rp.findall(qn("a:solidFill")):
                rp.remove(f)
            sf = rp.makeelement(qn("a:solidFill"), {})
            sf.append(sf.makeelement(qn("a:srgbClr"), {"val": color}))
            rp.insert(0, sf)
        r.append(rp)
        t = r.makeelement(qn("a:t"), {})
        t.text = line
        r.append(t)
        p.append(r)
        tx.append(p)


def replace_in_para(sh, old, new):
    """runが分かれていても段落単位で置き換える"""
    hit = False
    el = sh._element if hasattr(sh, "_element") else sh._tc
    for p in el.iter(qn("a:p")):
        ts = list(p.iter(qn("a:t")))
        full = "".join(t.text or "" for t in ts)
        if old in full:
            ts[0].text = full.replace(old, new)
            for t in ts[1:]:
                t.text = ""
            hit = True
    assert hit, old


def move_by(shapes, dx=0, dy=0):
    for sh in shapes:
        sh.left = sh.left + Cm(dx)
        sh.top = sh.top + Cm(dy)


def clone_slide(src, names=None):
    """同じレイアウトで新しいページを作り、src の図形をコピーする（names 指定時はその名前だけ）"""
    new = prs.slides.add_slide(src.slide_layout)
    for ph in list(new.shapes):
        ph._element.getparent().remove(ph._element)
    tree = new.shapes._spTree
    for el in src.shapes._spTree:
        if el.tag.split("}")[-1] not in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
            continue
        nv = el.find(".//" + qn("p:cNvPr"))
        if names is not None and (nv is None or nv.get("name") not in names):
            continue
        cp = deepcopy(el)
        for blip in cp.iter(qn("a:blip")):          # 画像があれば付け替える
            rid = blip.get(qn("r:embed"))
            part = src.part.related_part(rid)
            _, new_rid = new.part.get_or_add_image_part(io.BytesIO(part.blob))
            blip.set(qn("r:embed"), new_rid)
        tree.append(cp)
    return new


ORDER = list(S)            # 並べ替え用：最後にこの順で並べる


def insert_after(anchor, new):
    ORDER.insert(ORDER.index(anchor) + 1, new)


FRAME = ["正方形/長方形 2", "正方形/長方形 3", "Google Shape;156;p7", "直線コネクタ 1", "テキスト ボックス 39"]


def new_page(title, lead, after):
    """S20 と同じ枠（タイトル・リードの帯・区切り線）で新しいページを作る"""
    s = clone_slide(S[19], FRAME)
    set_lines(by(s, "Google Shape;156;p7"), [title])
    set_lines(by(s, "テキスト ボックス 39"), lead)
    insert_after(after, s)
    return s


def fade(slide, x, y, w, h, label=None):
    """ぼかし：半透明の白で覆い、ふちをぼかす"""
    sp = box(slide, x, y, w, h, fill=WHITE, shape=MSO_SHAPE.RECTANGLE)
    clr = sp.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    clr.append(clr.makeelement(qn("a:alpha"), {"val": "85000"}))
    spPr = sp._element.spPr
    eff = spPr.makeelement(qn("a:effectLst"), {})
    eff.append(eff.makeelement(qn("a:softEdge"), {"rad": "63500"}))
    spPr.append(eff)
    if label:
        put_text(sp.text_frame, [one(label, 11, True, GRAY, align="c")], anchor="m")
    return sp


def center_body(slide, top=DIV_Y + 0.4, bottom=FOOT_Y - 0.3, skip=()):
    body = [sh for sh in slide.shapes if Cm(DIV_Y + 0.1) <= sh.top < Cm(FOOT_Y - 0.1) and sh not in skip]
    t = min(sh.top for sh in body)
    b = max(sh.top + sh.height for sh in body)
    dy = int((Cm(top) + Cm(bottom)) / 2 - (t + b) / 2)
    for sh in body:
        sh.top = sh.top + dy


yen = lambda v: f"{int(round(v)):,}円"

# ============================================================
# SIM（成果報酬運営）の数字
# ============================================================
FIX = [20000, 20000, 20000, 40000, 40000, 40000]      # 固定費（初動3か月は2万円）
UNIT, CONSUL = 2000, 50000                             # 成果単価（体験予約1件）／コンサル費（入会案内）
SIM = {"忍者ナイン": {"cv": [2, 4, 6, 8, 11, 13], "rate": 0.40},
       "チャイルドアイズ": {"cv": [2, 4, 7, 9, 10, 12], "rate": 0.25}}


def split_int(vals, total):
    """月ごとの入会を整数にし、合計を total に合わせる（端数の大きい月から繰り上げ。同じなら早い月）"""
    base = [int(v) for v in vals]
    rest = sorted(range(len(vals)), key=lambda i: (-(vals[i] - base[i]), i))
    for i in rest[:total - sum(base)]:
        base[i] += 1
    return base


for k, d in SIM.items():
    d["join"] = split_int([c * d["rate"] for c in d["cv"]], int(sum(d["cv"]) * d["rate"]))
    d["cost_a"] = [f + UNIT * c for f, c in zip(FIX, d["cv"])]
    d["cpa"] = [a / c for a, c in zip(d["cost_a"], d["cv"])]
    d["cpo"] = [(a + CONSUL) / j if j else None for a, j in zip(d["cost_a"], d["join"])]
    print(k, "入会", d["join"], sum(d["join"]), "CPA", [round(v) for v in d["cpa"]], "CPO", [round(v) for v in d["cpo"]])
assert sum(SIM["忍者ナイン"]["join"]) == 17 and sum(SIM["チャイルドアイズ"]["join"]) == 11

# ============================================================
# S5 本提案の考え方：CPAの改善・CPOの改善 → 許容CPAの引き上げ
# ============================================================
s = S[4]
by(s, "TextBox 2").top = Cm(2.0)
by(s, "TextBox 3").top = Cm(3.45)
by(s, "TextBox 4").top = Cm(5.3)
set_lines(by(s, "TextBox 4"), ["一度つながれば、接点を持ち続けられる。情報を渡し続けられる。",
                               "体験予約で終わらず、来場 → 入会 → 継続まで並走できる——これがLINE最大の魅力です。"], sz=14)
row1 = [by(s, n) for n in ("TextBox 5", "Rounded Rectangle 6", "TextBox 7", "Rounded Rectangle 8", "TextBox 9")]
row2 = [by(s, n) for n in ("TextBox 10", "Rounded Rectangle 11", "TextBox 12", "Rounded Rectangle 13", "TextBox 14",
                           "Rounded Rectangle 15", "TextBox 16", "Rounded Rectangle 17", "TextBox 18", "Rounded Rectangle 19")]
move_by(row1, dy=7.35 - 11.56)
move_by(row2, dy=9.35 - 14.35)
lab = by(s, "TextBox 5")                       # 上司の指摘の言葉に合わせて「広告・アフィ」
lab.left, lab.width = Cm(0.9), Cm(3.35)
set_lines(lab, ["広告・アフィ"], sz=12)
# ① CPAの改善／② CPOの改善
y1 = 11.25
for x, w, hd, subs in [(4.32, 9.45, "① CPAの改善", ["LPで離脱した方をLINEで回収し、", "同じ広告費で体験予約を増やす"]),
                       (14.35, 11.76, "② CPOの改善", ["体験予約から来場・入会までLINEで後押しし、", "入会率を上げる"])]:
    sp = box(s, x, y1, w, 2.35, fill=PALE, line=BORDER, radius=0.08)
    put_text(sp.text_frame, [one(hd, 14, True, NAVY, align="c", sa=3)] + [one(b, 11.5, None, INK, align="c", ls=1.2) for b in subs],
             anchor="m", ml=0.2, mr=0.2)
down(s, 4.32 + 9.45 / 2 - 0.6, 13.75, 1.2, 0.6)
down(s, 14.35 + 11.76 / 2 - 0.6, 13.75, 1.2, 0.6)
band(s, 14.5, "①＋②で、許容CPAを引き上げられる ＝ 広告費を増やしても採算が合う", sz=15, h=1.35, x=1.52, w=24.59)
T(s, 1.52, 16.0, 24.59, 0.9,
  [one("※許容CPA＝入会1件にかけられる費用（許容CPO）×入会率。入会率が上がるほど、体験予約1件にかけられる広告費が増えます。",
       10, None, MUT, align="c")], anchor="m", ml=0, mr=0)

# ============================================================
# S7 リードの誤り（調査データのページに前後検索と書いてあった）
# ============================================================
set_lines(by(S[6], "Text 7"), ["習い事は「3〜5歳」で始める家庭が最も多く、", "未就学児の保護者の87.5%が「させたい」と考えています。"])

# ============================================================
# S14 前提整理を図に
# ============================================================
s = S[13]
drop(*[sh for sh in s.shapes if sh.top >= Cm(4.5)])
chip(s, CX0, 4.6, 7.2, 0.85, "1　対象範囲（4つの領域）", fill="06C755", sz=12)
areas = ["① 友だち追加", "② 配信の効率改善", "③ ユーザー満足度改善", "④ 管理側の業務効率化"]
bw, gap = (CW - 0.4 * 3) / 4, 0.4
for i, a in enumerate(areas):
    sp = box(s, CX0 + i * (bw + gap), 5.7, bw, 1.9, fill=NAVY, radius=0.10)
    put_text(sp.text_frame, [one(a, 14, True, WHITE, align="c")], anchor="m")
T(s, CX0, 7.75, CW, 0.8, [one("LINE公式アカウントの4つの領域を対象に、施策を設計します。", 12, None, INK, align="c")],
  anchor="m", ml=0, mr=0)
cards = [("2　想定UU規模", ["友だち数は、導入できる動線の本数と整備の進み具合で変わります", "数値はレンジの中間を基準にした試算例です"]),
         ("3　数値の位置づけ", ["開封率・クリック率・CVR・獲得件数は、業界一般値・弊社実績にもとづく数値です", "実装できる範囲に合わせ、すり合わせのうえ精緻化します"]),
         ("4　費用について", ["記載の費用（初期費用・月額費用）は前提の参考値です", "施策の範囲が決まったのち、別途お見積りをご案内します"])]
cw3 = (CW - 0.4 * 2) / 3
for i, (hd, body) in enumerate(cards):
    x = CX0 + i * (cw3 + 0.4)
    chip(s, x, 9.0, cw3, 0.85, hd, fill="06C755", sz=12)
    sp = box(s, x, 9.95, cw3, 4.2, fill=PALE, line=BORDER, radius=0.06)
    put_text(sp.text_frame, [one("・" + b, 11.5, None, INK, ls=1.3, sa=6) for b in body], anchor="m", ml=0.35, mr=0.3)
center_body(s)

# ============================================================
# S20 6か月後の欄に CPA・CPO（6か月目の1か月分）
# ============================================================
s = S[19]
lab_cv, val_cv = by(s, "正方形/長方形 85"), by(s, "テキスト ボックス 26")
n, c = SIM["忍者ナイン"], SIM["チャイルドアイズ"]
for i, (name, val) in enumerate([("CPA", f"{min(n['cpa'][5], c['cpa'][5]):,.0f}〜{max(n['cpa'][5], c['cpa'][5]):,.0f}円"),
                                 ("CPO", f"{min(n['cpo'][5], c['cpo'][5]):,.0f}〜{max(n['cpo'][5], c['cpo'][5]):,.0f}円")]):
    y = 12.25 + i * 1.18
    le = deepcopy(lab_cv._element); s.shapes._spTree.append(le)
    ve = deepcopy(val_cv._element); s.shapes._spTree.append(ve)
    L, V = s.shapes[-2], s.shapes[-1]
    L.top = Cm(y)
    set_lines(L, [name])
    V.left, V.top, V.width = Cm(22.45), Cm(y + 0.1), Cm(4.75)
    set_lines(V, [val], color=NAVY, sz=11)
T(s, 19.0, 14.7, 8.0, 0.6, [one("※6か月目の1か月分。月ごとの推移は次ページ", 9, None, MUT)], ml=0, mr=0)

# 次ページ：CPA・CPOの推移（1〜6か月目）
s = new_page("想定の費用対効果（CPA・CPOの推移）",
             ["運用が育つにつれて体験予約・入会が増え、",
              "体験予約1件あたり（CPA）・入会1件あたり（CPO）の費用は下がっていきます。"], after=S[19])
hdr = ["", "1か月目", "2か月目", "3か月目", "4か月目", "5か月目", "6か月目"]
colw = [5.62] + [3.25] * 6
y = 4.35
for k, d in SIM.items():
    chip(s, CX0, y, 5.0, 0.8, k, fill=NAVY, sz=12)
    rows = [["体験予約"] + [f"{v}件" for v in d["cv"]],
            ["CPA（体験予約1件）"] + [yen(v) for v in d["cpa"]],
            ["入会"] + [f"{v}件" for v in d["join"]],
            ["CPO（入会1件）"] + [yen(v) if v else "－" for v in d["cpo"]]]
    table(s, CX0, y + 0.95, CW, 4.0, hdr, rows, col_w=colw, hsz=11, bsz=11,
          align=["l"] + ["r"] * 6, bold_rows=(2, 4))
    y += 5.75
T(s, CX0, 15.9, CW, 1.4, [one("※CPA＝（固定費＋成果報酬 体験予約1件2,000円）÷体験予約。CPO＝（固定費＋成果報酬＋入会案内のコンサル費 50,000円）÷入会。"
        "初期費用は含みません。固定費は1〜3か月目20,000円、4か月目から40,000円。"
        "入会＝体験予約×入会引き上げ率（忍者ナイン40%・チャイルドアイズ25%）。各月の入会は整数にし、6か月の合計（17件・11件）に合わせています。", 9, None, MUT, ls=1.2)], ml=0, mr=0)

# ============================================================
# S22 アジェンダ → 施策の詳細（既存の効率化）＋全体の導線（友だち獲得部分をぼかす）
# ============================================================
set_lines(by(S[21], "正方形/長方形 1"), ["施策の詳細（既存の効率化）"])


def flow_page(after, title, lead, fades):
    s = clone_slide(S[16])
    set_lines(by(s, "Google Shape;156;p7"), [title])
    set_lines(by(s, "テキスト ボックス 69"), lead)
    for f in fades:
        fade(s, *f)
    insert_after(after, s)
    return s


flow_page(S[21], "想定動線（既存の効率化）",
          ["このパートでは、友だち追加後の配信と、ユーザー対応・管理側の効率化をご説明します。",
           "（薄く表示している部分は、新規友だち獲得のパートでご説明します）"],
          [(1.3, 5.25, 8.75, 5.95), (1.55, 16.15, 7.3, 1.65)])

# ============================================================
# S24 入会12件 → 11件（44件×25%）
# ============================================================
replace_in_para(by(S[23], "TextBox 46"), "入会12件", "入会11件")

# ============================================================
# S27 QA自動化・個別チャットを習い事向けに（別業種の文面が残っていた）
# ============================================================
s = S[26]
set_lines(by_text(s, "料金・使用量"), ["・よくある質問（月謝・送迎・振替など）にLINEで自動回答",
                                     "・営業時間外の疑問も解消し、体験予約前の離脱を防ぐ"])
set_lines(by_text(s, "契約内容等"), ["・入会手続きなど個別のご相談は有人チャットで対応",
                                  "・タグ管理で、入会検討中の方のお問い合わせを優先"])

# ============================================================
# S28 アジェンダ（新規友だち獲得）の直後に全体の導線（配信部分をぼかす）
# ============================================================
flow_page(S[27], "想定動線（新規友だち獲得）",
          ["このパートでは、友だち追加までの動線（離脱防止バナー・サンクスLINE誘導など）をご説明します。",
           "（薄く表示している部分は、既存の効率化のパートでご説明します）"],
          [(14.55, 4.25, 12.3, 7.0), (10.35, 7.55, 4.15, 3.7), (9.95, 12.35, 16.9, 5.6), (1.55, 12.9, 4.6, 3.2)])

# ============================================================
# S35 サンクスLINE誘導の費用 初期15万円・月額5万円
# ============================================================
s = S[34]
replace_in_para(by_text(s, "初期：10万円"), "初期：10万円", "初期：15万円")
replace_in_para(by_text(s, "月額3万円"), "月額3万円～", "月額5万円")

# S35 の直後：【オプション】QA自動化（チャットボット）
s = new_page("【オプション】QA自動化（チャットボット）",
             ["体験予約の前に保護者が気にする疑問（月謝・送迎・振替など）に、",
              "LINEで24時間すぐにお答えし、熱量が下がる前に体験予約へつなげます。"], after=S[34])
cols = [("目的", ["機会損失を防ぐ：夜間・休日も疑問をその場で解消し、他の教室への流出・離脱を防ぐ",
                  "問い合わせ対応の工数を減らす：電話・メールで多い定型の質問に自動で回答",
                  "大事なご相談に集中：スタッフは個別相談・体験後のフォローに時間を使える"]),
        ("組み込み方", ["リッチメニューの押しやすい位置に「よくある質問」「個別相談」を常設",
                        "一次対応：定型の質問はタップで選ぶ形の自動回答ですぐに解決",
                        "二次対応：複雑なご相談だけ1対1のトーク（有人）に引き継ぐ"]),
        ("設計のポイント", ["貴社の電話・メールでの問い合わせ実績と競合の情報から、質問を洗い出す",
                            "体験予約の前に必ず気になる点（月謝・入会金・送迎・振替・持ち物）を優先してセット",
                            "質問を増やしすぎず、タップで迷わず答えにたどり着ける階層にする"])]
cw3 = (CW - 0.4 * 2) / 3
for i, (hd, body) in enumerate(cols):
    x = CX0 + i * (cw3 + 0.4)
    chip(s, x, 4.5, cw3, 0.95, hd, fill=NAVY, sz=13)
    sp = box(s, x, 5.6, cw3, 8.3, fill=PALE, line=BORDER, radius=0.05)
    put_text(sp.text_frame, [one("・" + b, 11.5, None, INK, ls=1.35, sa=8) for b in body], anchor="t",
             ml=0.35, mr=0.3, mt=0.35)
chip(s, CX0 + (CW - 6.5) / 2, 14.35, 6.5, 1.1, "初期 5万円〜", fill=NAVY, sz=16)
T(s, CX0, 15.55, CW, 0.7, [one("※質問の数・階層により変動します", 10, None, MUT, align="c")], anchor="m", ml=0, mr=0)
center_body(s)

# ============================================================
# S37 成果報酬型の記載
# ============================================================
s = S[36]
set_lines(by_text(s, "施策成果件数に伴う"), ["成果報酬型：成果件数に応じた費用で、リスクを抑えて運用を始められます。"], sz=14)
chipDYM = by_text(s, "DYMプラン")
chipDYM.width = Cm(7.3)
set_lines(chipDYM, ["DYMプラン（成果報酬型）"])

# S37 の直後：CPO改善のご提案｜サンクスLINE誘導
s = new_page("CPO改善のご提案｜サンクスLINE誘導",
             ["体験予約の完了画面からLINEへつなげ、来場・入会までを後押しします。",
              "入会率が上がるほど、入会1件あたりの費用（CPO）が下がります。"], after=S[36])


def node(s, x, y, w, h, lines, fill, col=INK, sz=12, line=None, lw=1.0):
    sp = box(s, x, y, w, h, fill=fill, line=line, lw=lw, radius=0.10)
    put_text(sp.text_frame, [one(t, sz if i == 0 else sz - 2, i == 0, col, align="c", ls=1.2)
                             for i, t in enumerate(lines)], anchor="m", ml=0.15, mr=0.15, mt=0.05, mb=0.05)
    return sp


y, h = 4.6, 2.5
node(s, CX0, y, 4.3, h, ["体験予約", "フォーム"], PALE, NAVY, sz=13, line=BORDER)
arrow(s, 5.7, y + 0.8, 0.8, 0.9)
node(s, 6.7, y, 4.3, h, ["予約完了", "の画面"], PALE, NAVY, sz=13, line=BORDER)
arrow(s, 11.2, y + 0.8, 0.8, 0.9, fill=GREEN)
node(s, 12.2, y, 4.0, h, ["DYMの", "ツールで", "LINEへ"], GREEN, WHITE, sz=13)
arrow(s, 16.4, y + 0.8, 0.8, 0.9, fill=GREEN)
node(s, 17.4, y, 4.4, h, ["体験前の案内", "体験後のフォロー"], PGREEN, DGREEN, sz=12, line=GREEN)
arrow(s, 22.0, y + 0.8, 0.8, 0.9, fill=GREEN)
node(s, 23.0, y, 3.32, h, ["入会"], GREEN, WHITE, sz=15)
pts = [("予約の直後に案内", "関心が一番高いタイミングなので、友だちになっていただきやすい"),
       ("予約内容をトークに反映", "体験内容に合わせて、案内を出し分けます"),
       ("体験後のフォロー", "来場・入会までLINEで後押しします")]
cw3 = (CW - 0.26 * 2) / 3
for i, (hd, sub) in enumerate(pts):
    card(s, CX0 + i * (cw3 + 0.26), 7.6, cw3, 2.5, hd, sub, hsz=12.5, bsz=10.5, anchor="m")
band(s, 10.5, "弊社実績：美容クリニックで、予約後の来院率が40〜50%改善", sz=14, h=1.3)
bw2 = 6.5
x1 = CX0 + (CW - (bw2 * 2 + 0.3)) / 2
chip(s, x1, 12.25, bw2, 1.3, "初期 15万円", fill=NAVY, sz=16)
chip(s, x1 + bw2 + 0.3, 12.25, bw2, 1.3, "月額 5万円", fill=NAVY, sz=16)
T(s, CX0, 13.7, CW, 0.7, [one("※APIツール（Lステップ等）と併用可／完了画面を増やす場合は追加費用", 10, None, MUT, align="c")],
  anchor="m", ml=0, mr=0)
center_body(s)

# ============================================================
# S38 サンクスLINE誘導ツール 初期¥150,000・月額¥50,000
# ============================================================
tbl = next(sh for sh in S[37].shapes if sh.has_table).table
for row in tbl.rows:
    cells = row.cells
    if "サンクスLINE誘導ツール" in cells[4].text:
        replace_in_para(cells[5], "¥50,000", "¥150,000")
        tx6 = cells[6].text_frame._txBody
        for p in tx6.findall(qn("a:p")):
            tx6.remove(p)
        for p in cells[5].text_frame._txBody.findall(qn("a:p")):
            tx6.append(deepcopy(p))
        replace_in_para(cells[6], "¥150,000", "¥50,000")

# ============================================================
# S8・S9 前後検索（色分け・結論・凡例・出典）
# ============================================================
CAT = [("習い事の比較・検討", "3467B2", "E7EEF8"), ("子育て・学び", "D9661F", "FBEBDD"), ("家族のお出かけ・楽しみ", "8E4EC6", "F1E9F8")]


def swap_pic(slide, pic, path):
    _, rid = slide.part.get_or_add_image_part(str(path))
    pic._element.blipFill.find(qn("a:blip")).set(qn("r:embed"), rid)


for s, kw, img, lead, nocolor in [
        (S[7], "子ども 習い事", "yaruki_zengo_kodomo_color.png",
         ["習い事の比較は検索当日に集中し、その前後は「家族のお出かけ・楽しみ」と「子育て」への関心が中心です。",
          "当日に決めきれなかった方をLINEでつなぎとめ、日々の関心に合わせた配信で接点を保ちます。"], "色なし：その他"),
        (S[8], "習い事", "yaruki_zengo_naraigoto_color.png",
         ["習い事の比較は検索の前から後まで続き、検索後は「忍者ナイン 評判」「くもん 月謝」など評判・費用の確認に進みます。",
          "比較が続くあいだ、LINEで評判・月謝などの疑問にお答えし、体験予約へつなげます。"], "色なし：大人の習い事・趣味、その他")]:
    set_lines(by(s, "Google Shape;156;p7"), [f"前後検索「{kw}」"])
    ld = by(s, "Text 7")
    ld.top, ld.height = Cm(1.8), Cm(1.85)
    set_lines(ld, lead, sz=13)
    swap_pic(s, next(sh for sh in s.shapes if sh.shape_type == 13), IMG / img)
    x = 3.7
    for name, col, _ in CAT:
        w = 0.42 * len(name) + 0.8
        chip(s, x, 3.95, w, 0.58, name, fill=col, sz=10)
        x += w + 0.2
    T(s, x + 0.1, 3.95, 7.0, 0.58, [one(nocolor, 9, None, MUT)], anchor="m", ml=0, mr=0)
    T(s, 1.07, 17.45, 17.0, 0.5, [one(f"出典：LINEヤフー社提供の前後検索データ（検索起点：「{kw}」）", 8, None, MUT)],
      anchor="m", ml=0, mr=0)

# ============================================================
# S10 前後検索のまとめ（項目別）
# ============================================================
s = S[9]
set_lines(by(s, "Google Shape;156;p7"), ["前後検索のまとめ（項目別）"])
ld = by(s, "Text 7")
set_lines(ld, ["習い事の比較は検索の当日から後まで続き、その前後には子育て・家族のお出かけへの関心があります。",
               "項目ごとに、LINEでの打ち手を用意します。"], sz=13)
drop(*[sh for sh in s.shapes if sh.name not in ("Google Shape;156;p7", "Text 7")])
LW, PW, RW, G = 4.3, 5.3, 4.72, 0.1
xs = [CX0, CX0 + LW + G]
for i in range(3):
    xs.append(xs[-1] + PW + G)
y0 = 4.25
for i, (h, fill) in enumerate([("検索前（-15〜-1日）", NAVY), ("検索当日", NAVY), ("検索後（+1〜+15日）", NAVY), ("LINEでの打ち手", "06C755")]):
    chip(s, xs[i + 1], y0, RW if i == 3 else PW, 0.85, h, fill=fill, sz=11.5)
ROWS = [
    (["「公文式教室」", "「ピアノ 習い事」", "「そろばん 効果」", "「英語教室 おすすめ 子供」"],
     ["「習い事 ランキング」", "「ヤマハ音楽教室」", "「そろばん」", "「体操教室」"],
     ["「忍者ナイン 評判」", "「くもん 月謝」", "「小学生 習い事 いくつ」", "「幼児教室」"],
     "比較の最中に友だち追加していただき、評判・月謝などの疑問にお答えして体験予約へ"),
    (["「こどもちゃれんじ」", "「スマイルゼミ」", "「集団行動が苦手な子供」", "「9歳の壁」"],
     ["「rsウイルス」"],
     ["「小学校受験」", "「小学校一年生」", "「ポピー 教材」", "「児童手当」"],
     "子どもの成長・悩みに寄り添う情報をお届けし、接点を保つ"),
    (["「キッザニア」", "「スタジオアリス」", "「夏休み 子供 過ごし方」"],
     ["－"],
     ["「ポケモンセンター」", "「七五三」", "「ハーモニーランド」"],
     "季節の行事（七五三・夏休みなど）に合わせた企画配信"),
]
RH = 3.35
for r, ((name, col, pale), (pre, day, post, line)) in enumerate(zip(CAT, ROWS)):
    y = y0 + 0.95 + r * (RH + 0.12)
    sp = box(s, xs[0], y, LW, RH, fill=col, radius=0.06)
    put_text(sp.text_frame, [one(name, 12.5, True, WHITE, align="c")], anchor="m", ml=0.2, mr=0.2)
    for i, qs in enumerate((pre, day, post)):
        sp = box(s, xs[i + 1], y, PW, RH, fill=pale, radius=0.04)
        put_text(sp.text_frame, [one(q, 10.5, None, INK, align="c", sa=1) for q in qs], anchor="m", ml=0.15, mr=0.15)
    sp = box(s, xs[4], y, RW, RH, fill=PGREEN, line="06C755", radius=0.06)
    put_text(sp.text_frame, [one(line, 10.5, True, DGREEN, align="l", ls=1.25)], anchor="m", ml=0.3, mr=0.25)
yt = y0 + 0.95 + 3 * (RH + 0.12) + 0.1
T(s, xs[0], yt, LW, 0.8, [one("LINEへの動線", 11, True, NAVY, align="c")], anchor="m", ml=0, mr=0)
for i, (lab, fill) in enumerate([("離脱防止バナーの掲載", NAVY), ("LINE友だち追加", "06C755"), ("LINE施策（投稿）", "06C755")]):
    chip(s, xs[i + 1], yt, PW, 0.8, lab, fill=fill, sz=11)
T(s, CX0, yt + 1.0, CW, 0.5, [one("出典：LINEヤフー社提供の前後検索データ（検索起点：「子ども 習い事」「習い事」）", 8, None, MUT)],
  anchor="m", ml=0, mr=0)

# ============================================================
# 並べ替え
# ============================================================
lst = prs.slides._sldIdLst
sid = [sl.slide_id for sl in ORDER]
ids = {int(el.get("id")): el for el in list(lst)}
for el in list(lst):
    lst.remove(el)
for i in sid:
    lst.append(ids[i])
prs.save(str(OUT))
print("saved:", OUT.name, len(Presentation(str(OUT)).slides), "枚")
