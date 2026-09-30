# -*- coding: utf-8 -*-
"""やる気スイッチ ver3.1（殿村さんPC修正版）→ ver3.2：資料の文末を言い切りに揃える（です・ます を外す）＋字体をメイリオに統一。
S8の「をLINEで見られます」はLINEの訴求文（ユーザーに見せる文言）なので残す。
  python3 _build/patch_yaruki_v32_iikiri.py <ver3.1.pptx>
"""
import sys
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "202609_株式会社やる気スイッチグループ御中_LINE公式アカウント運用のご提案_ver3.2.pptx"
REP = [
    ("をもとにした想定数値となります。", "をもとにした想定数値"),
    ("体験予約1件にかけられる広告費が増えます。", "体験予約1件にかけられる広告費が増える"),
    ("複数機能の導入による調整もご相談ください。", "複数機能の導入による調整もご相談可能"),
    ("別途お見積りをご案内します", "別途お見積りをご案内"),
    ("ご説明します", "ご説明"),
    ("入会までLINEで追いきります", "入会までLINEで追いきる"),
    ("友だち追加動線について、★を記載しております。", "友だち追加動線に★を記載"),
    ("同バナーから実施が可能です。", "同バナーから実施可能"),
    ("ステータス通知を行うことが出来る機能です。", "ステータス通知を行うことができる機能"),
    ("APIを使用した機能です。", "APIを使用した機能"),
    ("料金プランにより通数単価、利用可能機能が異なります。", "料金プランにより通数単価・利用可能な機能が異なる"),
    ("CPOの改善に最適なオプションです。", "CPOの改善に最適なオプション"),
    ("リスクを抑えて運用を始められます。", "リスクを抑えて運用を始められる"),
    ("月次対応内に含んでおります。", "月次対応内に含む"),
    ("ご依頼いただくことも可能です。", "ご依頼いただくことも可能"),
    ("都度発注でのご依頼も各単価で受けさせていただきます。", "都度発注でのご依頼も各単価で対応可能"),
    ("来場・入会までを後押しします。", "来場・入会までを後押し"),
    ("入会1件あたりの費用（CPO）が下がります。", "入会1件あたりの費用（CPO）が下がる"),
    ("案内を出し分けます", "案内を出し分ける"),
    ("体験後のフォローで入会を後押しします", "体験後のフォローで入会を後押し"),
    ("追加のご提案です", "追加のご提案"),
    ("初期費用は含みません。", "初期費用は含まない。"),
    ("に合わせています。", "に合わせている。"),
    ("の費用は含みません。", "の費用は含まない。"),
    # 補足資料など（S44以降）
    ("前後する可能性がございます。", "前後する可能性あり"),
    ("メールマガジンの6倍を誇ると言われています。", "メールマガジンの6倍とされる"),
    ("ポイントの導入支援が可能です。", "ポイントの導入支援が可能"),
    ("連携の実施も可能です。", "連携の実施も可能"),
    ("動線設計が可能です。", "動線設計が可能"),
]
prs = Presentation(sys.argv[1])
hits = {}


def fix_p(p, where):
    runs = list(p.iter(qn("a:r")))
    for old, new in REP:
        done = False
        for r in runs:                      # まず1つのrunの中で置き換える（書式を保つ）
            t = r.find(qn("a:t"))
            if t is not None and t.text and old in t.text:
                t.text = t.text.replace(old, new); done = True
        if not done:
            ts = [r.find(qn("a:t")) for r in runs if r.find(qn("a:t")) is not None]
            full = "".join(t.text or "" for t in ts)
            if old in full:                 # runをまたぐときは段落ごと
                ts[0].text = full.replace(old, new)
                for t in ts[1:]:
                    t.text = ""
                done = True
        if done:
            hits.setdefault(old, []).append(where)


for i, s in enumerate(prs.slides, 1):
    for el in s.shapes._spTree.iter(qn("a:p")):
        fix_p(el, i)
# 字体をメイリオに統一（"Meiryo" 表記・指定なしも含めて、英数字・日本語・記号すべて）
nfont = 0
for s in prs.slides:
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for rpr in s.shapes._spTree.iter(qn(tag)):
            for ft in ("a:latin", "a:ea", "a:cs"):
                e = rpr.find(qn(ft))
                if e is None:
                    e = rpr.makeelement(qn(ft), {})
                    # latin/ea/cs は solidFill などの後ろに置く決まり
                    anchor = [c for c in rpr if c.tag in (qn("a:hlinkClick"), qn("a:hlinkMouseOver"), qn("a:rtl"), qn("a:extLst"))]
                    if anchor:
                        anchor[0].addprevious(e)
                    else:
                        rpr.append(e)
                if e.get("typeface") != "メイリオ":
                    e.set("typeface", "メイリオ"); nfont += 1
    for r in s.shapes._spTree.iter(qn("a:r")):          # rPr の無い run にも付ける
        if r.find(qn("a:rPr")) is None:
            rpr = r.makeelement(qn("a:rPr"), {"lang": "ja-JP"})
            for ft in ("a:latin", "a:ea", "a:cs"):
                rpr.append(rpr.makeelement(qn(ft), {"typeface": "メイリオ"}))
            r.insert(0, rpr); nfont += 1
# latin → ea → cs の順に並べ直す（順番が崩れるとPowerPointで開けないことがある）
AFTER = {qn(x) for x in ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")}
for s in prs.slides:
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for rpr in s.shapes._spTree.iter(qn(tag)):
            fs = [rpr.find(qn(ft)) for ft in ("a:latin", "a:ea", "a:cs")]
            for e in fs:
                rpr.remove(e)
            nxt = next((c for c in rpr if c.tag in AFTER), None)
            for e in fs:
                if nxt is not None:
                    nxt.addprevious(e)
                else:
                    rpr.append(e)
print("字体を直した箇所:", nfont)
prs.save(str(OUT))
for old, _ in REP:
    print(f"{old[:24]:26} → p{hits.get(old)}")
print("saved:", OUT.name)
