# -*- coding: utf-8 -*-
"""サンクスLINEの簡易資料（5ページ・表紙／裏表紙／アジェンダなし）。

ねらい：LINE公式アカウント単体の提案は通りにくいので、ADの予算型提案に組み込む入口の資料。
話が通れば具体提案（動線・配信・シミュレーション）に進む。業界は問わない汎用版。
図形はコンパクトに、情報量は最小限（殿村さん指示）。

書式はチェングロウス本資料（DYMの書式）をコピーし、不要ページを使い回して中身を作り直す。
部品は build_chengrowth_v11.py と同じ。

  python3 _build/build_thanks_line.py
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_src = (ROOT / "_build/build_chengrowth_v11.py").read_text(encoding="utf-8")
_helpers = _src[_src.index("# ================= helpers"):_src.index("# ============================================================\n# 流用するページの素材")]
_head = _src[_src.index("import shutil"):_src.index("shutil.copyfile(SRC, OUT)")]
_head = _head.replace('SRC = ROOT / "20260922_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver1.2.pptx"',
                      'SRC = ROOT / "20260929_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver2.5.pptx"')
_head = _head.replace('OUT = ROOT / "20260929_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver2.0.pptx"',
                      'OUT = ROOT / "20260930_サンクスLINEのご提案（簡易版）ver1.0.pptx"')
exec(_head)
shutil.copyfile(SRC, OUT)
prs = Presentation(str(OUT))
S = list(prs.slides)
exec(_helpers)
GRAY, PGRAY, DGREEN = "7F7F7F", "EDEDED", "0B7A3B"


def node(s, x, y, w, h, lines, fill, col=INK, sz=12, line=None, dash=None, lw=1.0):
    sp = box(s, x, y, w, h, fill=fill, line=line, lw=lw, dash=dash, radius=0.10)
    put_text(sp.text_frame, [one(t, sz if i == 0 else sz - 2, i == 0, col, align="c", ls=1.2)
                             for i, t in enumerate(lines)], anchor="m", ml=0.15, mr=0.15, mt=0.05, mb=0.05)
    return sp


def price(s, x, y):
    chip(s, x, y, 4.6, 0.9, "初期 10万円", fill=NAVY, sz=13)
    chip(s, x + 4.8, y, 4.6, 0.9, "月額 3万円〜", fill=NAVY, sz=13)


ORDER = []
spare = iter([S[i] for i in (1, 2, 3, 4, 5)])


def new(title, lead):
    s = next(spare)
    frame(s, title, lead)
    ORDER.append(s)
    return s


# ============================================================
# P1 LINEの成果は、申込みの「後」で決まる
# ============================================================
s = new("LINEの成果は、申込みの「後」で決まる",
        ["広告で申込みまではとれても、来店・成約・リピートは、その後の追い方次第です。"])
y, h = 5.6, 3.0
node(s, CX0, y, 5.2, h, ["広告"], PALE, NAVY, sz=15, line=BORDER)
arrow(s, 6.7, y + 1.05, 0.9, 0.9)
node(s, 7.9, y, 5.2, h, ["申込み", "資料請求・予約・購入"], PALE, NAVY, sz=15, line=BORDER)
arrow(s, 13.4, y + 1.05, 0.9, 0.9, fill=GRAY)
node(s, 14.6, y, 5.2, h, ["メール・電話で", "フォロー"], WHITE, GRAY, sz=14, line=GRAY, dash=MSO_LINE_DASH_STYLE.DASH)
arrow(s, 20.1, y + 1.05, 0.9, 0.9, fill=GRAY)
node(s, 21.3, y, 5.02, h, ["来店・成約", "リピート"], PGRAY, INK, sz=15)
# 取りこぼしの印
sp = box(s, 14.6, y + h + 0.4, 11.72, 1.5, fill=PORANGE, radius=0.10)
put_text(sp.text_frame, [one("読まれない・つながらない＝取りこぼし", 13, True, ORANGE, align="c")], anchor="m")
T(s, CX0, y + h + 0.4, 12.9, 1.5, [one("広告の成果（申込み数）", 12, True, NAVY, align="c")], anchor="m")
band(s, 12.6, "申込みの「後」の追い方を変えれば、同じ広告費のまま成果が増えます", sz=14, h=1.4)

# ============================================================
# P2 サンクスLINE
# ============================================================
s = new("サンクスLINE｜申込みの直後に、LINEでつながる",
        ["申込みの完了画面から、そのままLINEへ。関心が一番高いタイミングなので、友だちになっていただきやすい仕組みです。"])
y, h = 4.6, 2.6
node(s, CX0, y, 4.4, h, ["フォーム入力"], PALE, NAVY, sz=13, line=BORDER)
arrow(s, 5.8, y + 0.85, 0.8, 0.9)
node(s, 6.8, y, 4.9, h, ["申込み完了", "購入・予約・資料請求"], PALE, NAVY, sz=13, line=BORDER)
arrow(s, 11.9, y + 0.85, 0.8, 0.9, fill=GREEN)
node(s, 12.9, y, 3.6, h, ["DYMの", "ツールで", "LINEへ"], GREEN, WHITE, sz=13)
arrow(s, 16.7, y + 0.85, 0.8, 0.9, fill=GREEN)
node(s, 17.7, y, 8.62, 1.2, ["はじめての方 → 友だち追加の画面へ"], PGREEN, DGREEN, sz=12, line=GREEN)
node(s, 17.7, y + 1.4, 8.62, 1.2, ["友だちの方 → トーク画面へ"], PGREEN, DGREEN, sz=12, line=GREEN)
# 強み3つ
yc, hc, wc = 8.0, 2.6, 8.2
pts = [("関心が一番高い直後に案内", "申込みの直後なので、友だちになっていただきやすい"),
       ("友だち集めの広告費ゼロ", "今の広告で集めた申込みを、そのままLINEへ"),
       ("入力内容をそのまま活用", "トークに反映し、配信の出し分けにも使える")]
for i, (hd, sub) in enumerate(pts):
    card(s, CX0 + i * (wc + 0.26), yc, wc, hc, hd, sub, hsz=12.5, bsz=10.5, anchor="m")
price(s, CX0, 11.2)
T(s, CX0 + 9.8, 11.2, 15.3, 0.9, [one("既存のAPIツールとの併用も可能です", 11, None, INK)], anchor="m", ml=0)

# ============================================================
# P3 LINEで、成果まで育てる
# ============================================================
s = new("LINEで、成果まで育てる",
        ["申込んだ方に合わせて、LINEで次の一歩を後押しします。"])
cols = [("資料請求した方", "比較のポイント・事例", "来店・商談へ"),
        ("予約した方", "前日・当日のリマインド", "キャンセルを減らす"),
        ("購入した方", "使い方・次回のご案内", "リピートへ")]
wc = 8.2
for i, (who, what, goal) in enumerate(cols):
    x = CX0 + i * (wc + 0.26)
    node(s, x, 4.6, wc, 1.4, [who], NAVY, WHITE, sz=14)
    down(s, x + wc / 2 - 0.6, 6.15, 1.2, 0.7)
    node(s, x, 7.0, wc, 1.8, [what], PALE, NAVY, sz=13, line=BORDER)
    down(s, x + wc / 2 - 0.6, 8.95, 1.2, 0.7, fill=GREEN)
    node(s, x, 9.8, wc, 1.4, [goal], GREEN, WHITE, sz=14)
band(s, 12.2, "弊社実績：美容クリニックで、予約後の来院率が40〜50%改善", sz=14, h=1.4)

# ============================================================
# P4 広告は「申込みの数」、LINEは「売上」をつくる
# ============================================================
s = new("広告は「申込みの数」、LINEは「売上」をつくる",
        ["同じ広告費のまま、申込みの「後」を伸ばして成約を増やします。"])
chip(s, 4.4, 4.4, 8.9, 0.75, "広告の役割：申込みを集める", fill=NAVY, sz=11)
chip(s, 13.6, 4.4, 12.72, 0.75, "LINEの役割：申込んだ方を売上へ", fill=GREEN, sz=11)
for (y, lab, lf, lc, mid, mf, mc, ml, mdash, res, rf, rc, af) in [
        (5.5, "広告だけ", PGRAY, GRAY, "メール・電話", WHITE, GRAY, GRAY, MSO_LINE_DASH_STYLE.DASH,
         ["成約 20件", "1件あたり 5.0万円"], PGRAY, INK, GRAY),
        (8.4, "広告\n＋LINE", GREEN, WHITE, "サンクスLINE", PGREEN, DGREEN, GREEN, None,
         ["成約 24件", "1件あたり 約4.3万円"], GREEN, WHITE, GREEN)]:
    h = 2.4
    sp = node(s, CX0, y, 2.9, h, lab.split("\n")[:1] + lab.split("\n")[1:], lf, lc, sz=12)
    node(s, 4.4, y, 8.9, h, ["申込み 100件", "広告費 月100万円"], PALE, NAVY, sz=14, line=BORDER)
    arrow(s, 13.5, y + 0.75, 0.8, 0.9, fill=af)
    node(s, 14.5, y, 5.4, h, [mid], mf, mc, sz=13, line=ml, dash=mdash, lw=1.5)
    arrow(s, 20.1, y + 0.75, 0.8, 0.9, fill=af)
    node(s, 21.1, y, 5.22, h, res, rf, rc, sz=15)
band(s, 11.6, "LINEの友だちは毎月たまり、広告を止めても案内を続けられます", sz=14, h=1.4)
foot(s, "※試算例：成約率 20%（広告だけ）→ 24%（広告＋LINE）。1件あたり＝（広告費＋LINE費用 月3万円）÷成約件数")

# ============================================================
# P5 進め方
# ============================================================
s = new("進め方｜広告のご提案と、あわせて始められます",
        ["大きな開発は不要です。広告の予算の一部から、小さく始められます。"])
steps = [("1", "具体的なご提案", "LINEへの動線・配信の内容・成果のシミュレーション"),
         ("2", "広告と同時に開始", "完了画面にLINEへの案内を設置し、配信をスタート")]
for i, (no, hd, sub) in enumerate(steps):
    y = 4.8 + i * 3.0
    node(s, CX0, y, 2.4, 2.4, [no], NAVY, WHITE, sz=24)
    card(s, CX0 + 2.7, y, 22.42, 2.4, hd, sub, hsz=15, bsz=12, anchor="m")
price(s, CX0, 11.1)
T(s, CX0 + 9.8, 11.1, 15.3, 0.9, [one("※設定する地点（完了画面）を増やす場合は追加費用", 10, None, MUT)], anchor="m", ml=0)
band(s, 12.8, "まずは、広告のご提案とあわせて、LINEの設計をご提案します", sz=14, h=1.4)

# ============================================================
# 並べ替え（5枚だけ残す）
# ============================================================
keep = {id(x) for x in ORDER}
lst = prs.slides._sldIdLst
items = list(lst)
by = {id(sl): el for sl, el in zip(S, items)}
for sl, el in zip(S, items):
    if id(sl) not in keep:
        prs.part.drop_rel(el.rId); lst.remove(el)
for el in list(lst):
    lst.remove(el)
for sl in ORDER:
    lst.append(by[id(sl)])
prs.save(str(OUT))
print("saved:", OUT.name, len(Presentation(str(OUT)).slides), "枚")
