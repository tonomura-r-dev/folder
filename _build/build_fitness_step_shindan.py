# -*- coding: utf-8 -*-
"""ジム・フィットネス業界：ステップ配信（追加から14日）に診断を足して厚くした版を、1枚の試作で作る（2026-10-29 殿村さん）。

    python -m _build.build_fitness_step_shindan <ジム・フィットネス業界_LINEOA施策提案ver3.pptx> <出力.pptx>

  ver3の25ページ「ステップ配信｜追加から14日で体験予約を取り切る」だけを残し、表に診断の配信を3行足す
  （Day3 性格タイプ診断／Day5 ダイエットタイプ診断／Day10 続けられるか診断）。Day8 の会員の声は診断のタイプ別に変える。
  「MBTI」の名前は使わない（登録商標）。足した行は薄い紺で塗る。表の行を写したら a16:rowId を振り直す（CLAUDE.md 2026-09-24）。
  実際の資料に入れるかは、この試作を見て殿村さんが決める。
"""
import copy
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from _build.pptx_parts import drop_slide

NEW_FILL = "DCE3F2"
ROWS = [  # (元の行（写す書式）, 追加か, Day, 目的, 配信コンテンツ, CTA, 通知プレビュー冒頭)。\n は改行
    (1, False, "Day0", "体験の案内\n＋属性取得", "あいさつ＋無料体験（期限つき）のバナー＋\n30秒 体づくりタイプ診断", "体験予約／\n診断回答", "【〇月〇日まで】無料体験のご案内"),
    (2, False, "Day2", "初心者の不安", "タイプ別「最初の1ヶ月はこれだけ」\n＋マシンの使い方動画", "動画視聴", "【初めての方へ】まずはマシン3種類"),
    (2, True, "Day3", "続け方を見つける", "性格タイプ診断（3問・4タイプ）→\n計画派・気分派・仲間派・記録派ごとの通い方", "性格タイプ診断", "【3問】あなたに合うジムの通い方は？"),
    (3, False, "Day4", "費用の納得", "1分料金シミュレーター（週2回なら1回いくら？）", "料金診断", "【1回いくら？】週2回なら…"),
    (2, True, "Day5", "目的別の深掘り", "ダイエットタイプ診断（4問：食事・間食・睡眠・運動）\n→ タイプ別「ジムでやること＋家での1つ」", "ダイエット\nタイプ診断", "【4問】あなたのダイエットタイプは？"),
    (4, False, "Day6", "2回目のオファー", "無料体験のご案内＋\n「見学だけでもOK・入会を迫りません」", "体験予約 CV①", "【見学だけOK】今週の空き枠"),
    (5, False, "Day8", "社会的証明", "同じタイプの会員の声（診断のタイプ別）\n／続いている人の共通点", "事例閲覧", "【コツコツ型の方へ】3ヶ月後の変化"),
    (2, True, "Day10", "不安の見える化", "続けられるか診断（3問）→ 結果に合わせた\n体験の回り方と、最初の2週間の予定", "続けられるか診断", "【3問】あなたは続けられる？"),
    (6, False, "Day11", "疑問の除去", "不安Q&A（勧誘・解約・シャワー・混雑・持ち物）\n＋チャット相談", "チャット相談", "【解約もOK】よくあるご質問"),
    (7, False, "Day14", "最終オファー", "入会金無料・事務手数料0円＋体験枠の残数", "体験予約（再）", "【残3枠】今月の体験ご予約"),
]
LEAD = ["友だち追加から15日でアクティブ率が急落する。その14日間に、答えるだけの診断を4回挟み、",
        "反応を切らさずに体験予約へつなげる。"]
CARDS = [
    ("① 診断は4回、どれも1タップ", ["Day0・3・5・10に答えるだけの配信"]),
    ("② 答えはそのまま出し分けに", ["診断のタイプを配信のタグにする"]),
    ("③ 診断の結果から体験予約へ", ["結果の最後に「無料体験を予約」"]),
]
NOTE = ["※ 薄い紺の行が、今回足した診断の配信。Day5は目的に「ダイエット」を選んだ方だけ（他の目的の方は送らない）。",
        "※ 性格タイプ診断・ダイエットタイプ診断を入れる場合は、リッチメニュー「② 診断・シミュレーター」に項目を足す。"]


def set_cell_text(tc, text, sz="950"):
    p = tc.find(qn("a:txBody")).find(qn("a:p"))
    rs = p.findall(qn("a:r"))
    for r in rs[1:]:
        p.remove(r)
    r0 = rs[0]
    r0.find(qn("a:rPr")).set("sz", sz)
    for k, part in enumerate(text.split("\n")):
        if k == 0:
            r0.find(qn("a:t")).text = part
            continue
        br = p.makeelement(qn("a:br"), {})
        br.append(copy.deepcopy(r0.find(qn("a:rPr"))))
        r = copy.deepcopy(r0)
        r.find(qn("a:t")).text = part
        p.append(br)
        p.append(r)


def set_fill(tc, hexv):
    pr = tc.find(qn("a:tcPr"))
    sf = pr.find(qn("a:solidFill"))
    sf.find(qn("a:srgbClr")).set("val", hexv)


def set_paras(sh, lines):
    txb = sh._element.find(qn("p:txBody"))
    ps = txb.findall(qn("a:p"))
    tmpl = copy.deepcopy(ps[0])
    for r in tmpl.findall(qn("a:r"))[1:] + tmpl.findall(qn("a:br")):
        tmpl.remove(r)
    for p in ps:
        txb.remove(p)
    for line in lines:
        p = copy.deepcopy(tmpl)
        p.find(qn("a:r")).find(qn("a:t")).text = line
        txb.append(p)


def set_card(sh, head, body):
    txb = sh._element.find(qn("p:txBody"))
    ps = txb.findall(qn("a:p"))
    h, b = copy.deepcopy(ps[0]), copy.deepcopy(ps[1])
    for p in (h, b):
        for r in p.findall(qn("a:r"))[1:] + p.findall(qn("a:br")):
            p.remove(r)
    for p in ps:
        txb.remove(p)
    h.find(qn("a:r")).find(qn("a:t")).text = head
    txb.append(h)
    for line in body:
        p = copy.deepcopy(b)
        p.find(qn("a:r")).find(qn("a:t")).text = line
        txb.append(p)


def main(src, out):
    prs = Presentation(src)
    keep = next(s for s in prs.slides
                if any(sh.has_text_frame and sh.text_frame.text.startswith("ステップ配信｜追加から14日") for sh in sh_list(s)))
    for s in [s for s in prs.slides if s.slide_id != keep.slide_id]:
        drop_slide(prs, s)
    s = prs.slides[0]
    by = {sh.shape_id: sh for sh in s.shapes}

    tbl = by[4]._element.find(".//" + qn("a:tbl"))
    trs = tbl.findall(qn("a:tr"))
    src_rows = [copy.deepcopy(tr) for tr in trs]
    for tr in trs[1:]:
        tbl.remove(tr)
    row_h = Inches(0.42)
    trs[0].set("h", str(int(Inches(0.4))))
    for k, (base, new, *texts) in enumerate(ROWS):
        tr = copy.deepcopy(src_rows[base])
        tr.set("h", str(int(row_h)))
        tcs = tr.findall(qn("a:tc"))
        for tc, t in zip(tcs, texts):
            set_cell_text(tc, t)
        if new:
            for tc in tcs:
                set_fill(tc, NEW_FILL)
        tbl.append(tr)
    for n, e in enumerate(tbl.iter("{http://schemas.microsoft.com/office/drawing/2014/main}rowId")):
        e.set("val", str(10000 + n))
    by[4].height = int(Inches(0.4) + row_h * len(ROWS))

    set_paras(by[3], LEAD)
    top = by[4].top + by[4].height + Inches(0.1)
    for sid, (head, body) in zip((5, 6, 7), CARDS):
        set_card(by[sid], head, body)
        by[sid].top, by[sid].height = top, Inches(0.52)
    set_paras(by[8], NOTE)
    by[8].top, by[8].height = top + Inches(0.6), Inches(0.36)
    prs.save(out)
    print("saved:", out)


def sh_list(s):
    return list(s.shapes)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
