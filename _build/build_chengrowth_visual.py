# -*- coding: utf-8 -*-
"""チェングロウス：施策のビジュアル（動線別・4ブロックの型）を、本資料とは別のPPTXで作る。

型（CLAUDE.md 2026-09-25）：動線ごとに ①友だち追加の動線 ②効率改善（ステップ・企画配信）
③満足度改善（ユーザー） ④効率改善（管理側）の4ブロック。
数字は SIM ver2.6（殿村さん修正版）の9月時点。画像は殿村さんが画像生成したもの（_images/chengrowth_v11_*.png）。

  python3 _build/build_chengrowth_visual.py
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_src = (ROOT / "_build/build_chengrowth_v11.py").read_text(encoding="utf-8")
# 本資料のビルドと同じ部品（書式・図形・画像の置き方）をそのまま使う
_helpers = _src[_src.index("# ================= helpers"):_src.index("# ============================================================\n# 流用するページの素材")]
_head = _src[_src.index("import shutil"):_src.index("shutil.copyfile(SRC, OUT)")]
_head = _head.replace('SRC = ROOT / "20260922_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver1.2.pptx"',
                      'SRC = ROOT / "20260929_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver2.1.pptx"')
_head = _head.replace('OUT = ROOT / "20260929_株式会社チェングロウス御中_LINE公式アカウント運用のご提案ver2.0.pptx"',
                      'OUT = ROOT / "20260929_株式会社チェングロウス御中_LINE施策の設計（動線別）ver1.0.pptx"')
exec(_head)
shutil.copyfile(SRC, OUT)
prs = Presentation(str(OUT))
S = list(prs.slides)
exec(_helpers)

BLUE_T = "0B7A3B"
LX, LW = CX0, 17.9          # 左：4ブロック
RX, RW = 19.5, 6.82         # 右：画面イメージ
HW = 3.4                    # ブロック見出しの幅


def block(slide, y, h, head, lines, fill=PALE, hcol=NAVY):
    chip(slide, LX, y, HW, h, head, fill=hcol, sz=10.5)
    sp = box(slide, LX + HW + 0.15, y, LW - HW - 0.15, h, fill=fill, line=BORDER)
    paras = []
    for ln in lines:
        if isinstance(ln, tuple):
            tag, body = ln
            paras.append(multi([(tag + "　", 9.5, True, hcol), (body, 9.5, None, INK)], ls=1.2, sa=3))
        else:
            paras.append(one(ln, 9.5, None, INK, ls=1.2, sa=3))
    put_text(sp.text_frame, paras, anchor="m", ml=0.25, mr=0.2, mt=0.1, mb=0.1)


def page(slide, title, lead, b1, b2, b3, b4, pics, note):
    frame(slide, title, lead)
    block(slide, 4.2, 2.1, "①友だち追加の動線", b1, fill=PORANGE, hcol=ORANGE)
    block(slide, 6.45, 5.4, "②効率改善\n（配信）", b2)
    block(slide, 12.2, 2.3, "③満足度改善\n（ユーザー）", b3, fill=PGREEN, hcol=BLUE_T)
    block(slide, 14.85, 2.0, "④効率改善\n（管理側）", b4, fill="F2F2F2", hcol=MUT)
    y = 4.2
    hs = [h for _, h in pics]
    for name, h in pics:
        fitpic(slide, name, RX, y, RW, h)
        y += h + 0.25
    foot(slide, note)


NOTE = ("※件数はシミュレーション（SIM ver2.6）の9月の応募数15件を、動線ごとの友だち数とクリック数の比率で振り分けたもの"
        "（反応率の補正・複数回のクリックを含むため、開封率×クリック率×応募率の単純な掛け算とは一致しません）／画面はイメージです")

# ---------------- 表紙 ----------------
cover = S[0]
replace_on(cover, {"―サイトリニューアルに伴うご提案―": "―施策の設計（動線別）―"})

# ---------------- 動線00 離脱防止 ----------------
s = S[10]
page(s, "動線00｜離脱防止ポップアップ",
     ["求人ページから離れようとした方に、LINEで新着求人を受け取れることをご案内します。",
      "応募しなかった方ともつながり続け、配信とリッチメニューで会員登録・応募へつなげます。"],
     ["求人ページから離れようとした方 → 「条件に合う新着求人をLINEでお届けします」 → LINE追加（月約41人）",
      ("", "※月3万円（初期1.5万円）")],
     [("[ステップ]", "1〜14日後）離脱防止から追加した方（約27人）→ 希望の職種を聞く → 条件に合う求人 → 会員登録・応募のご案内（全10通）→ 開封72.5%×クリック12%×応募2.0% → 約1件／月"),
      ("[企画配信]", "離脱防止から追加した方（約110人）／3月の転職ピーク前（1〜2月）→ 新着求人・資格取得支援・働き方の特集（月4本）→ 開封78%×クリック10%×応募2.0% → 約2件／月"),
      ("[リッチメニュー]", "有効な友だち（約137人）→ 職種・地域・資格で求人を探す（月30%が利用×2回）→ 応募2.0% → 約1件／月"),
      ("", "※配信・リッチメニューの作成はコンサル費（月20万円）に含む")],
     [("[QA自動化]", "資格・受験料・勤務地などのよくある質問に、LINEで24時間すぐ回答"),
      ("[個別チャット]", "求人の条件や働き方の相談に、LINEで対応（電話は不要）")],
     [("[自動応答]", "よくある質問への回答を自動化し、電話・メールでの問い合わせ対応を減らす")],
     [("chengrowth_v11_popup.png", 4.3)],
     NOTE)

# ---------------- 動線01 LINEで登録 ----------------
s = S[11]
page(s, "動線01｜「LINEで登録」ボタン（会員登録＝CV①）",
     ["求人詳細・会員登録ページに「LINEで登録」を置き、入力ほぼなしで会員登録と友だち追加を同時に完了させます。",
      "登録した方を、リッチメニューと配信で応募（CV②）へつなげます。"],
     ["求人詳細・会員登録ページ → 「LINEで登録（入力はほぼ不要）」 → 会員登録＋LINE追加（月約98人）",
      ("", "※サイト側で実装（リニューアルの要件）。LINE Profile+は申請・審査のうえ別途")],
     [("[ステップ]", "1〜14日後）LINEで登録した方（約66人）→ 希望の職種を聞く → 条件に合う求人 → 応募のご案内（全10通）→ 開封72.5%×クリック12%×応募2.0% → 約3件／月"),
      ("[企画配信]", "LINEで登録した方（約262人）／3月の転職ピーク前（1〜2月）→ 新着求人・資格取得支援・働き方の特集（月4本）→ 開封78%×クリック10%×応募2.0% → 約4件／月"),
      ("[リッチメニュー]", "有効な友だち（約328人）→ 職種・地域・資格で求人を探す（月30%が利用×2回）→ 応募2.0% → 約4件／月"),
      ("", "※配信・リッチメニューの作成はコンサル費（月20万円）に含む")],
     [("[QA自動化]", "資格・受験料・勤務地などのよくある質問に、LINEで24時間すぐ回答"),
      ("[個別チャット]", "求人の条件や面談の相談に、LINEで対応（電話は不要）"),
      ("[その他]", "希望条件（職種・地域）の変更を、リッチメニューからいつでも")],
     [("[自動応答]", "よくある質問への回答を自動化し、問い合わせ対応を減らす"),
      ("[その他]", "会員の情報とLINEをひもづけ、条件別の配信を自動化する")],
     [("chengrowth_v11_register.png", 3.5), ("chengrowth_v11_richmenu.png", 3.9), ("chengrowth_v11_talk_step.png", 4.6)],
     NOTE)

# ---------------- 動線02 サンクスLINE ----------------
s = S[12]
page(s, "動線02｜応募の完了画面（サンクスLINE・オプション）",
     ["応募の完了画面でLINEの友だち追加をご案内し、応募した方とLINEでつながります。",
      "応募の数ではなく、応募から面談・採用までの取りこぼしを減らすための動線です（シミュレーションには含めていません）。"],
     ["応募の完了画面 → 「今後のご連絡はこのLINEでお送りします」 → LINE追加（月約17人の想定）",
      ("", "※オプション：月5万円（初期10万円）")],
     [("[ステップ]", "応募当日）応募した方 → 応募受付のお礼・面談までの流れ → 面談日程の確定"),
      ("[ステップ]", "面談前日・当日）面談予定の方 → 日時・場所・持ち物のご案内 → 無断キャンセルを防ぐ"),
      ("[企画配信]", "応募済みの方には、求人の案内などの配信は行わない（連絡・リマインドのみ）")],
     [("[QA自動化]", "面談の場所・持ち物・服装などの質問に、LINEで24時間すぐ回答"),
      ("[個別チャット]", "日程の変更や相談を、電話なしでLINEで完結")],
     [("[自動応答]", "面談日程の確認・変更の受付を自動化し、電話でのやりとりを減らす")],
     [("chengrowth_v11_talk_thanks.png", 12.4)],
     "※月約17人＝サイトからの応募（月約47件）×完了画面の表示80%×友だち追加45%の想定。シミュレーションの応募数・費用には含めていません／画面はイメージです")

ORDER = [S[0], S[10], S[11], S[12], S[-1]]
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
