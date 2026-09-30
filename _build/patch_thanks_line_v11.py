# -*- coding: utf-8 -*-
"""サンクスLINE簡易資料 ver1.0（殿村さんPC修正版）→ ver1.1

S1：LINEがどこに関わるかを見えるように（メール・電話の箱をLINEに置き換え、「ここをLINEに」の目印）
S2：下の費用とAPIの話を、費用の枠にまとめ直す
S5：上の文章（リード）を強くする

  python3 _build/patch_thanks_line_v11.py <ver1.0.pptx>
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_src = (ROOT / "_build/build_chengrowth_v11.py").read_text(encoding="utf-8")
_helpers = _src[_src.index("# ================= helpers"):_src.index("# ============================================================\n# 流用するページの素材")]
_head = _src[_src.index("import shutil"):_src.index("shutil.copyfile(SRC, OUT)")]
exec(_head)
exec(_helpers)
SRC = Path(sys.argv[1])
OUT = ROOT / "20260930_サンクスLINEのご提案（簡易版）ver1.1.pptx"
prs = Presentation(str(SRC))
S = list(prs.slides)
DGREEN, GRAY = "0B7A3B", "7F7F7F"
by = lambda s, n: next(x for x in s.shapes if x.name == n)


def drop(s, *names):
    for n in names:
        el = by(s, n)._element
        el.getparent().remove(el)


def node(s, x, y, w, h, lines, fill, col=INK, sz=12, line=None, lw=1.0):
    sp = box(s, x, y, w, h, fill=fill, line=line, lw=lw, radius=0.10)
    put_text(sp.text_frame, [one(t, sz if i == 0 else sz - 2, i == 0, col, align="c", ls=1.2)
                             for i, t in enumerate(lines)], anchor="m", ml=0.15, mr=0.15, mt=0.05, mb=0.05)
    return sp


def fill(sp, color):
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(color)


# ============================================================
# S1：LINEの位置を見せる
# ============================================================
s = S[0]
# タイトルの誤字（LINEのは → LINEは）
for para in by(s, "TextBox 1")._element.iter(qn("a:p")):
    ts = list(para.iter(qn("a:t")))
    full = "".join(t.text or "" for t in ts)
    if "LINEのは" in full:          # runが分かれているので段落ごと書き換える
        ts[0].text = full.replace("LINEのは", "LINEは")
        for t in ts[1:]:
            t.text = ""
b8 = by(s, "Rounded Rectangle 8")
x, y, w, h = b8.left / 360000, b8.top / 360000, b8.width / 360000, b8.height / 360000
drop(s, "Rounded Rectangle 8", "Rounded Rectangle 10", "Rounded Rectangle 11")
fill(by(s, "Right Arrow 7"), GREEN); fill(by(s, "Right Arrow 9"), GREEN)
node(s, x, y, w, h, ["LINE", "リマインド・配信", "1対1のトーク"], GREEN, WHITE, sz=16)
node(s, 21.3, y, 5.02, h, ["来店・成約", "リピート"], PGREEN, DGREEN, sz=15, line=GREEN, lw=1.5)
# 広告で追えるところ／LINEで追えるところ（流れの真下に範囲の帯）
drop(s, "TextBox 12")
yb = y + h + 0.35
chip(s, CX0, yb, 13.1 - CX0, 0.85, "広告で追えるところ：申込みまで", fill=NAVY, sz=12)
chip(s, x, yb, 26.32 - x, 0.85, "LINEで追えるところ：来店・成約・リピートまで", fill=GREEN, sz=12)
sp = box(s, x, yb + 1.15, 26.32 - x, 1.1, fill="EDEDED", radius=0.10)
put_text(sp.text_frame, [one("今はメール・電話だけ → 読まれない・つながらない", 11, True, GRAY, align="c")], anchor="m")

# ============================================================
# S2：費用とAPIの話を1つの枠に
# ============================================================
s = S[1]
drop(s, "Rounded Rectangle 15", "Rounded Rectangle 16", "TextBox 17")
y0 = 13.6
box(s, CX0, y0, CW, 1.2, fill=WHITE, line=BORDER, radius=0.06)
chip(s, CX0 + 0.3, y0 + 0.25, 1.8, 0.7, "費用", fill=NAVY, sz=11)
T(s, CX0 + 2.4, y0, 8.4, 1.2,
  [multi([("初期 ", 10.5, True, INK), ("10万円", 14, True, NAVY), ("　月額 ", 10.5, True, INK), ("3万円〜", 14, True, NAVY)])],
  anchor="m", ml=0)
ln = s.shapes.add_connector(1, Cm(CX0 + 10.9), Cm(y0 + 0.25), Cm(CX0 + 10.9), Cm(y0 + 0.95))
ln.line.color.rgb = RGBColor.from_string(BORDER); ln.line.width = Pt(1.0)
T(s, CX0 + 11.2, y0, 13.7, 1.2,
  [one("✓ 今お使いのAPIツール（Lステップ等）と併用できます", 9.5, None, INK, sa=1),
   one("✓ 案内を置く完了画面を増やす場合は、追加費用", 9.5, None, INK)], anchor="m", ml=0)

# ============================================================
# S5：リードを強く
# ============================================================
s = S[4]
put_text(by(s, "TextBox 2").text_frame,
         [one("広告の予算に、月3万円〜を足すだけ。", 14, True, NAVY, ls=1.3, sa=2),
          one("完了画面にLINEへの案内を置けば、広告で集めた申込みを売上までつなげる仕組みが、広告と同時に始められます。",
              12.5, None, INK, ls=1.3)], anchor="m", ml=0, mr=0)

prs.save(str(OUT))
print("saved:", OUT.name)
