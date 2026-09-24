# -*- coding: utf-8 -*-
"""チェングロウスver1.1（PC保存版 2026-09-24）にレビュー指摘を反映。
- P1表紙：やる気スイッチのロゴ（忍者ナイン・チャイルドアイズ）→ 自動車求人Naviロゴ
- 他社案件の残り／てにをは 10箇所（run単位で書式保持）
※ 1回きりのパッチ。再実行するとPC調整が消える可能性があるので、以後はPC保存版を正とする。
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Cm

ROOT = Path(__file__).resolve().parent.parent
OUT = str(ROOT / "20260922_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver1.1.pptx")
LOGO = str(ROOT / "_images/chengrowth_logo.png")

prs = Presentation(OUT)
SW = prs.slide_width

# --- P1 表紙ロゴ差し替え ---
s = prs.slides[0]
old = [sh for sh in s.shapes if sh.shape_type == 13 and sh.top > Cm(9)]
assert len(old) == 2, len(old)
top = min(sh.top for sh in old)
for sh in old:
    sh._element.getparent().remove(sh._element)
w = Cm(8.0)
s.shapes.add_picture(LOGO, int((SW - w) / 2), top + Cm(0.3), width=w)

# --- 文言 ---
REPL = {
    11: [("展開をさせて", "展開させて")],
    12: [("。詳細なセグメント設定", "や、詳細なセグメント設定")],
    15: [("再学習", "再育成（再アプローチ）")],
    22: [("料金・使用量等のよくある質問", "求人内容や選考プロセス等のよくある質問")],
    24: [("マイサーラ会員登録案内", "マイページ登録・求職者会員登録案内"),
         ("地域別WEB申込キャンペーン", "地域別・希望条件別 求人相談キャンペーン")],
    25: [("で尚且つ", "で、尚且つ"),
         ("利用までナーチャリング", "面談・就業までナーチャリングが可能")],
    31: [("ご依頼依頼も。", "ご依頼いただくことも可能です。")],
    38: [("ハードル抑えつつ運営が可能。", "ハードルを抑えつつ運用することが可能。")],
}

def paras(sh):
    if sh.shape_type == 6:
        for c in sh.shapes: yield from paras(c)
        return
    if getattr(sh, "has_table", False) and sh.has_table:
        for row in sh.table.rows:
            for cell in row.cells: yield from cell.text_frame.paragraphs
    if sh.has_text_frame: yield from sh.text_frame.paragraphs

n = 0
for si, pairs in REPL.items():
    for sh in prs.slides[si - 1].shapes:
        for p in paras(sh):
            for r in p.runs:
                for a, b in pairs:
                    if a in r.text:
                        r.text = r.text.replace(a, b); n += 1
                        print(f"s{si:02d}: {a} → {b}")

# P10 「実装可能領域」「で」「によって」の余分な「で」
for sh in prs.slides[9].shapes:
    for p in paras(sh):
        rl = list(p.runs)
        for i, r in enumerate(rl):
            if r.text == "で" and i and rl[i - 1].text.endswith("実装可能領域"):
                r.text = ""; n += 1; print("s10: 領域でによって → 領域によって")

prs.save(OUT)
print("done", n)

# 追加（同日・リスト外の他社残り。P24 は上で保存済みのため別実行で適用済み）
# 料金プラン診断コンテンツ → 30秒 適職診断コンテンツ
# 料金確定通知もLINEで受け取れる → 新着求人の通知もLINEで受け取れる
# 「料金プラン診断」 → 「30秒 適職診断」
