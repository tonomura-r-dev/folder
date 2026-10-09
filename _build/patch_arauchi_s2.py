# -*- coding: utf-8 -*-
"""あらうち提案資料のS2（現状把握）を、あらうちの数字で作り直す（2026-10-09 殿村さん指示）。

S2の「Lstep運用中（リヴトラスト名義）」・友だち18,200・UU2,700・広告費100万は、他社（リヴトラスト）の資料の残りだった。
あらうちは Lステップなし（LINE公式アカウントの標準機能）・友だち1,241・UU1,737・広告なし。
SIM ver1.3 由来の数字（サイトのフォームの問い合わせ30件・ブロック率35%）は「（仮）」と書く。
S2以外のページ・図形の位置・書式（フォント・サイズ・太字・色）は触らない。

使い方:
    python _build/patch_arauchi_s2.py <元のpptx> <出力.pptx>
"""
import copy
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]


def set_paras(shape, paras, plain=False):
    """paras＝段落ごとの [(文字, 太字 or None), ...]。各段落の最初の run の書式を土台にする。
    plain=True なら色の指定（灰色など）を外して、ほかの値と同じ色にする。"""
    tf = shape.text_frame
    ps = tf.paragraphs
    assert len(ps) >= len(paras), (shape.shape_id, len(ps), len(paras))
    for p_el, runs in zip(ps, paras):
        base = p_el.runs[0]._r
        for r in p_el.runs[1:]:
            p_el._p.remove(r._r)
        for i, (text, bold) in enumerate(runs):
            r = base if i == 0 else copy.deepcopy(base)
            if i > 0:
                p_el._p.insert(p_el._p.index(prev) + 1, r)
            r.find(qn("a:t")).text = text
            rpr = r.find(qn("a:rPr"))
            if bold is not None:
                rpr.set("b", "1" if bold else "0")
            if plain:
                for f in rpr.findall(qn("a:solidFill")):
                    rpr.remove(f)
            prev = r
    for p_el in ps[len(paras):]:
        tf._txBody.remove(p_el._p)
    if plain:
        for end in tf._txBody.iter(qn("a:endParaRPr")):
            for f in end.findall(qn("a:solidFill")):
                end.remove(f)


prs = Presentation(SRC)
s2 = prs.slides[1]
by_id = {sh.shape_id: sh for sh in s2.shapes}
assert by_id[171].text_frame.text == "18,200" and by_id[179].text_frame.text == "Lstep", "S2の中身が想定と違う"

# 上の要約
set_paras(by_id[26], [
    [("LINE公式アカウントの標準機能で運用中（Lステップ等のツールなし）。", None)],
    [("友だち追加の入口・追加後の動線ともに、改善の余地があると想定。", None)],
])
# WEBサイトステータス
set_paras(by_id[106], [[("1,737", None)]])
set_paras(by_id[108], [[("0（広告なし）", None)]])
set_paras(by_id[110], [[("なし", True)]])
set_paras(by_id[118], [[("30件（仮）", None)]], plain=True)
set_paras(by_id[120], [[("問い合わせ", None)]])
set_paras(by_id[124], [[("0％", None)]])
set_paras(by_id[114], [[("YouTube", None)]], plain=True)
set_paras(by_id[112], [[("不明（広告なし）", None)]])
# LINEステータス
set_paras(by_id[171], [[("1,241", None)]])
set_paras(by_id[173], [[("約807（仮）", None)]])
set_paras(by_id[175], [[("35%（仮）", None)]])
set_paras(by_id[179], [[("なし（標準機能）", None)]])
set_paras(by_id[92], [[("リッチメニュー", None)]])
for r in by_id[92].text_frame._txBody.iter(qn("a:rPr")):
    r.set("sz", "900")              # 10ptだと折り返す
set_paras(by_id[177], [[("なし", True)]])
set_paras(by_id[169], [[("あいさつ", None)]])
set_paras(by_id[191], [[("ボタンなし", True)]])
# 注記・下の一言
set_paras(by_id[67], [
    [("※", None), ("いずれも貴社ヒヤリング、外部計測", None)],
    [("　をもとにした想定。（仮）は仮置き。", None)],
])
set_paras(by_id[43], [
    [("全ページに「LINEでお問い合わせ」ボタンを設置。", None)],
    [("100円チェッカー・内見依頼", True), ("など、問い合わせ入口を整備。", False)],
])
set_paras(by_id[53], [[("リッチメニュー未設定・あいさつにボタンなし。改善の余地あり。", None)]])

prs.save(OUT)
print("saved:", OUT)
