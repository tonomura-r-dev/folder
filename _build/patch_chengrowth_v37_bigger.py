# -*- coding: utf-8 -*-
"""チェングロウス ver3.6 → ver3.7：内容ページの文字を大きくする（2026-10-02 殿村さん指示）
本文の文字サイズを倍率Fで拡大（タイトルは据え置き、リード文は18ptに）。折り返しは個別に調整（ADJ）。
  python3 _build/patch_chengrowth_v37_bigger.py <ver3.6.pptx> [F]
"""
import sys
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20260929_チェングロウス_LINE公式アカウント運用のご提案ver3.7.pptx"
F0 = float(sys.argv[2]) if len(sys.argv) > 2 else 1.12
prs = Presentation(sys.argv[1])
SKIP_SLIDES = set()   # 表紙・扉・弊社実績・裏表紙は対象外（下で判定）
F_BY_SLIDE = {4: 1.1, 5: 1.1, 6: 1.0, 7: 1.1, 8: 1.1, 9: 1.1, 11: 1.05, 12: 1.05, 13: 1.0, 15: 1.0, 16: 1.1,
              17: 1.0, 18: 1.05, 19: 1.05, 21: 1.05, 22: 1.05, 23: 1.1, 25: 1.1, 26: 1.05, 28: 1.1, 29: 1.1,
              30: 1.0, 31: 1.0}      # ページごとの倍率（折り返しが崩れないよう個別に決めた）
NO_STRETCH = {13, 17, 30, 31}        # 縦に広げないページ（表・画像が多い）


def is_target(i, s):
    texts = [sh.text_frame.text for sh in s.shapes if sh.has_text_frame]
    if i in (1, len(prs.slides)):
        return False
    if any(t.strip() == "資料アジェンダ" for t in texts):
        return False
    if any("弊社実績" == t.strip() for t in texts):
        return False
    return True


def scale(el, f):
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for e in el.iter(qn(tag)):
            sz = e.get("sz")
            if sz:
                e.set("sz", str(int(round(int(sz) * f / 50.0)) * 50))


EMU = 360000
TARGET_BOTTOM = 16.6 * EMU      # 本文の下端（注記は16.8cm付近）
LEAD_SZ = {8: "1700", 11: "1500", 13: "1600", 16: "1600", 17: "1600", 25: "1450", 31: "1600"}


def is_fixed(sh):
    """大きさを変えず、位置だけ動かす図形（画像・表・矢印）"""
    return sh.shape_type == 13 or getattr(sh, "has_table", False) and sh.has_table or "Arrow" in sh.name


for i, s in enumerate(prs.slides, 1):
    if not is_target(i, s):
        continue
    f = F_BY_SLIDE.get(i, F0)
    body = []
    for sh in s.shapes:
        if sh.name in ("TextBox 1", "Connector 3"):
            continue
        if sh.name == "TextBox 2":
            for e in sh._element.iter(qn("a:rPr")):
                e.set("sz", LEAD_SZ.get(i, "1800"))
            sh.width = int(24.8 * EMU)
            continue
        # 幅の狭い小さな見出し枠は、折り返さないよう倍率を1.0以下にする
        small = sh.height < 1.4 * EMU and sh.width < 12 * EMU
        scale(sh._element, min(f, 1.0) if small else f)
        is_note = sh.top >= 16.5 * EMU and sh.height < 1.2 * EMU
        if not is_note:
            body.append(sh)
    if not body:
        continue
    top0 = min(sh.top for sh in body)
    bottom0 = max(sh.top + sh.height for sh in body)
    sy = 1.0 if i in NO_STRETCH else max(1.0, min(1.25, (TARGET_BOTTOM - top0) / (bottom0 - top0)))
    for sh in body:
        T, H = sh.top, sh.height
        new_top = top0 + (T - top0) * sy
        if is_fixed(sh):
            sh.top = int(new_top + (H * sy - H) / 2)
        else:
            sh.top = int(new_top)
            sh.height = int(H * sy)


def para_sub(n, old, new):
    """段落に old を含む箇所を、段落ごと new に置き換える（最初のrunの書式を残す）"""
    for sh in prs.slides[n - 1].shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            if old in p.text:
                runs = p.runs
                runs[0].text = new
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
                return
    raise SystemExit(f"not found: {n} {old}")


para_sub(11, "この3つは、", "この3つは、サイトに「条件別の検索結果ページ」と「LINEとの会員連携」があれば再現可能")
for sh in prs.slides[5].shapes:           # 6枚目のSTEP見出しは折り返さないよう10pt
    if sh.has_text_frame and sh.text_frame.text.startswith("STEP"):
        for e in sh._element.iter(qn("a:rPr")):
            e.set("sz", "1000")
# 21枚目：カード見出しが折り返さないよう短くする
para_sub(21, "データを初日から", "① データを初日から蓄積")
para_sub(21, "開発のやり直しを", "② 開発の作り直しを避ける")
para_sub(21, "3月の転職のピークに", "③ 3月のピークに間に合う")
para_sub(21, "リニューアルの要件に入れておけば", "リニューアルの要件に入れれば、")
para_sub(21, "サイトの改修が二度手間", "サイトの二度手間を避けられる。")
# 13枚目：表の下の箱が表に重ならないよう、少し下げる
for sh in prs.slides[12].shapes:
    if sh.has_text_frame and (sh.text_frame.text.startswith("掲載中の求人を") or sh.text_frame.text.startswith("「手軽な登録」")):
        sh.top += int(0.8 * EMU)
# 16枚目：自社サイトの流れの箱は13ptにそろえる（「応募フォーム」が折り返さないように）
for sh in prs.slides[15].shapes:
    if sh.has_text_frame and sh.text_frame.text.replace("\n", "") in ("求人ページ", "応募フォーム", "完了画面", "サンクスLINE（オプション）", "LINE追加") and sh.top < 8.0 * EMU:
        for e in sh._element.iter(qn("a:rPr")):
            e.set("sz", "1300")
para_sub(4, "LINE経由の応募の中で", "LINE経由の中で、配信より多い")
para_sub(4, "単発・日払い・在宅など", "単発・日払い・在宅など、")
para_sub(4, "条件をボタンに", "よく探される条件をボタンに")
para_sub(12, "登録と同じ画面で", "・登録と同じ画面で、友だち追加もご案内")
para_sub(12, "応募しなかった方とも、登録者として", "・応募しなかった方ともつながり続ける")
para_sub(15, "月約17人＝", "※月約17人＝サイトからの応募（月約47件）×完了画面の表示80%×友だち追加45%の想定／求人ボックス・Indeed経由は対象外／画面はイメージです")
prs.save(str(OUT))
print("saved:", OUT.name, "F =", F0)
