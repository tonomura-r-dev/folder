# -*- coding: utf-8 -*-
"""すなお ver7.2 → ver7.3（案）：「配信の内容」の章の4ページを、提案FMT（MAX版）のページを写して作る（2026-10-27 殿村さん）。

    python -m _build.patch_sunao_v73_fmt <ver7.2.pptx> <提案FMT_MAX版_20261024d.pptx> <ver7.3.pptx>

  ・前後検索「投資セミナー」＝FMTのS15（施策の考え方：短期・前後検索の図）。前後検索の画像を色分け済みの「投資セミナー」に差し替え、
    右の札を3分類（比べる／確かめる／始める）に、緑の帯（友だち追加の想定）を画像の0〜1日に合わせる
  ・診断＝FMTのS13の「LINEの画面」と「タグ・セグメント」の図を写し、右に2問とタグの使い方の表
  ・ステップ配信＝FMTのS16の形（Dayの帯 → メッセージの見出し → 中身 → ［検索データ］ → ボタン）。8通を4枚×2段（殿村さん「B」）
  ・16日以降＝FMTのS17（施策の考え方：長期・企画配信）。下の月のバー（トレンド）は外す（殿村さん「B」）
  色は紺・白・黒・灰（LINEの部分だけ緑）。FMTの「ご提案時に詳細記載」の札・黄色や金色の札は使わない。
  並び：P32（重要な考え方）→ 前後検索 → 診断 → ステップ配信（P34と差し替え）→ 16日以降 → P33 → P35。ほかのページは変えない。
"""
import copy
import sys

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches

from _build.patch_sunao_v73_haishin import note, table
from _build.patch_tanaka_v5 import copy_slide
from _build.pptx_parts import drop_slide, rect, rich

NAVY, BLACK, INK, SUB, LINE, LIGHT, PALE, WHITE = "1F2A5C", "1A1A1A", "333333", "777777", "D0D5E0", "EEF1F8", "F5F7FA", "FFFFFF"
ZENGO = [("3467B2", "比べる：ほかのセミナー"), ("8E4EC6", "確かめる：不安・安全性"), ("00897B", "始める：始め方")]
IMG = "_images/sunao_zengo_toushi_seminar_color.png"
IN = 2.54


def i(v):
    return v * IN


def sid(s, n):
    return next(sh for sh in s.shapes if sh.shape_id == n)


def remove(s, keep):
    for sh in list(s.shapes):
        if sh.shape_id not in keep:
            sh._element.getparent().remove(sh._element)


def set_lines(sh, lines):
    """最初の段落・最初の文字の書式のまま、行ごとに段落を作り直す"""
    txb = sh.text_frame._txBody
    ps = txb.findall(qn("a:p"))
    base = ps[0]
    for p in ps:
        txb.remove(p)
    for line in lines:
        p = copy.deepcopy(base)
        rs = p.findall(qn("a:r"))
        for r in rs[1:]:
            p.remove(r)
        for br in p.findall(qn("a:br")):
            p.remove(br)
        rs[0].find(qn("a:t")).text = line
        txb.append(p)


def recolor(sh, old, new):
    for e in sh._element.iter(qn("a:srgbClr")):
        if e.get("val").upper() == old:
            e.set("val", new)


# ── 1. 前後検索「投資セミナー」（FMT S15） ─────────────────────────
def page_zengo(prs, f15):
    s = copy_slide(prs, f15)
    remove(s, {4, 8, 10, 263, 264, 265, 266, 267, 268, 269, 849, 850, 851, 852, 3, 6})
    set_lines(sid(s, 4), ["前後検索「投資セミナー」"])
    set_lines(sid(s, 10), ["セミナーを探した方は、その日にほかのセミナーと比べて不安を確かめ、",
                           "1〜4日後に始め方を調べる。関心が続く15日間をステップ配信で追う。"])
    set_lines(sid(s, 852), ["投資セミナー"])
    x, y, w = 0.95, 2.36, 7.7
    pic = s.shapes.add_picture(IMG, Inches(x), Inches(y), Inches(w), Inches(w * 1167 / 2000))
    band = sid(s, 267)                                      # 緑の帯（友だち追加の想定）を、画像の0日〜1日に合わせて画像の後ろへ
    band.left, band.width = Inches(x + w * 1030 / 2000), Inches(w * (1264 - 1030) / 2000)
    band.top, band.height = Inches(1.86), Inches(y + w * 1167 / 2000 - 1.86)
    band._element.addprevious(pic._element)                 # 画像を帯の前（下）に置き、帯を画像の上に重ねる
    sid(s, 268).left = band.left - Inches(0.15)
    sid(s, 269).left = band.left + Inches(0.21)
    sid(s, 263).left, sid(s, 263).width = Inches(x), band.left - Inches(x) - Inches(0.05)
    sid(s, 264).left, sid(s, 264).width = band.left, Inches(x + w) - band.left
    for k, (col, text) in enumerate(ZENGO):                 # 右の札＝色分けの3分類
        rect(s, i(8.82), i(2.02 + 0.36 * k), i(1.85), i(0.3), col, [(text, 10.5, True, WHITE, 0)], align=PP_ALIGN.CENTER,
             margins=(0.05, 0.0, 0.05, 0.0))
    set_lines(sid(s, 6), ["16日以降へ", "", "（企画配信で", "再アプローチ）"])
    return s


# ── 2. 診断（FMT S13 の画面とタグの図） ─────────────────────────
PHONE = {39, 40, 41, 42, 43, 44, 45, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 59, 60, 91, 93}
TAGCARD = {32, 33, 34, 35, 36, 37, 38}


def page_shindan(prs, f13):
    s = copy_slide(prs, f13)
    remove(s, {3, 5, 7, 9, 58} | PHONE | TAGCARD)
    set_lines(sid(s, 9), [" 施策の考え方"])
    title = "診断：最初の2問で、どんな方かを分かるようにする"
    set_lines(sid(s, 5), [title])
    sid(s, 5).width = Inches(8.5)
    sid(s, 58).width = Inches(min(0.25 * len(title) + 0.2, 8.0))
    set_lines(sid(s, 7), ["答えでタグを付け、その後の配信をその方に合わせる"])
    texts = {41: "友だち追加ありがとうございます！参加特典は…", 43: "あてはまる方は？", 44: "一般", 45: "保険営業",
             49: "【8日目】受講生の声", 51: "【今月】次回セミナーのご案内", 52: "日程を教えて", 54: "次回の日程はこちらです→",
             60: "講師", 33: "タグ・セグメント（例）", 36: "保険営業", 37: "紹介を増やしたい",
             }
    for n, t in texts.items():
        set_lines(sid(s, n), [t])
    set_lines(sid(s, 38), ["タグで対象を絞る", "→ 紹介に悩んでいた受講生の声を配信"])
    for n in (32, 33, 36, 37, 47):                           # 赤い見出し・枠、金色のタグは紺に
        recolor(sid(s, n), "C8102E", NAVY)
        recolor(sid(s, n), "D9A441", NAVY)
    dx, dy = Inches(0.55 - 3.50), Inches(1.98 - 2.91)        # 画面を左へ
    for n in PHONE:
        sh = sid(s, n)
        sh.left, sh.top = sh.left + dx, sh.top + dy
    tx, ty = Inches(3.6 - 0.42), Inches(4.72 - 5.36)         # タグの図を表の下へ
    for n in TAGCARD:
        sh = sid(s, n)
        sh.left, sh.top = sh.left + tx, sh.top + ty
    end = table(s, 3.6, 1.98, [1.55, 2.75, 2.44], ["質問", "選択肢（付けるタグ）", "配信での使い方"],
                [["1問目\nあてはまる方", "保険営業の方／一般の方", "一般の方には、一般の方向け\nのセミナーを案内"],
                 ["2問目\nいちばんの関心", "保険営業の方：成約／紹介／\n働く時間\n一般の方：資産形成／見直し",
                  "受講生の声を、同じ悩み\nだった方の声にする"]], rh=0.98)
    rect(s, i(6.5), i(4.72), i(3.9), i(1.6), PALE,
         [("示唆", 13, True, NAVY, 2), ("一度の診断で、その方に合う受講生の声とセミナーを届けられる。聞くのは友だち追加直後の1回だけにする", 12, False, INK, 0)],
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, margins=(0.25, 0.12, 0.2, 0.1))
    note(s, 6.5, ["※答えはボタンで選ぶ形。押したボタンで、その後の配信を分ける"])
    assert end < 4.7, end
    return s


# ── 3. ステップ配信（FMT S16 の形・8通を4枚×2段） ───────────────────
STEPS = [  # (Day, 役割, メッセージの見出し, 中身2行, 検索データ, ボタン)
    ("0日目", "追加直後：診断", "ご登録ありがとう\nございます！", ["参加特典のご案内", "2問の診断で関心を確認"],
     "—（診断の答えで出し分け）", "2問で診断する"),
    ("2日目", "比べる", "ほかのセミナーと、\nここが違います", ["両想い営業・ライフプラン", "ブロックを動画で紹介"],
     "検索した日に「投資セミナー 東京」「資産運用セミナー」と比べている", "見どころ動画を見る"),
    ("4日目", "確かめる", "安心してご参加ください", ["講師の経歴と実績", "参加は無料"],
     "検索した日に「投資詐欺」「金融庁」で不安を確かめている", "講師紹介を見る"),
    ("6日目", "始める", "セミナーで持ち帰れること", ["週2日で成果を出す", "面談メソッド"],
     "1〜4日後に「口座開設」「nisa つみたて投資枠」など始め方を検索", "内容を詳しく見る"),
    ("8日目", "受講生の声", "同じ悩みから、\n成果を出した方の声", ["診断の答えに合わせる", "（成約／紹介／働く時間）"],
     "—（診断の答えで出し分け）", "受講生の声を見る"),
    ("10日目", "よくある質問", "参加前のよくある質問", ["地方からも参加できる", "初心者でも分かる"],
     "7〜15日後も「インデックスファンド」など学びの検索が続く", "よくある質問を見る"),
    ("12日目", "予約", "次回セミナーのご案内", ["開催日程", "予約ページ"],
     "—（募集期間に合わせて配信）", "日程を見て予約する"),
    ("14日目", "締切", "予約の締切が\n近づいています", ["締切日のお知らせ", "予約ページ"],
     "—（募集期間に合わせて配信）", "予約する"),
]


def page_step(prs, f16):
    s = copy_slide(prs, f16)
    remove(s, {4, 8, 10})
    set_lines(sid(s, 10), ["セミナーを探した直後から、比べる → 確かめる → 始めるの順に不安を解消し、",
                           "15日以内にセミナーの予約へつなげる配信を構築。"])
    w, gap, h = 2.4, 0.12, 2.4
    for k, (day, role, head, body, data, btn) in enumerate(STEPS):
        x = 0.42 + (w + gap) * (k % 4)
        y = 1.62 + (h + 0.1) * (k // 4)
        rect(s, i(x), i(y), i(w), i(h), WHITE, None, line=LINE, lw=0.75)
        rect(s, i(x), i(y), i(w), i(0.5), BLACK if k == len(STEPS) - 1 else NAVY,
             [(day, 14, True, WHITE, 0), (role, 10.5, False, WHITE, 0)], align=PP_ALIGN.CENTER, margins=(0.05, 0.02, 0.05, 0.02))
        rect(s, i(x + 0.08), i(y + 0.54), i(w - 0.16), i(0.48), WHITE, [(t, 12, True, NAVY, 0) for t in head.split("\n")], align=PP_ALIGN.CENTER,
             margins=(0.05, 0.0, 0.05, 0.0))
        rect(s, i(x + 0.08), i(y + 1.04), i(w - 0.16), i(0.38), WHITE, [(t, 11, False, INK, 0) for t in body],
             align=PP_ALIGN.CENTER, margins=(0.05, 0.0, 0.05, 0.0))
        rect(s, i(x + 0.1), i(y + 1.46), i(w - 0.2), i(0.48), "F2F2F2",
             [("［検索データ］", 8, True, SUB, 0), (data, 8.5, False, INK, 0)], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             margins=(0.08, 0.03, 0.06, 0.02))
        rect(s, i(x + 0.25), i(y + 2.02), i(w - 0.5), i(0.28), WHITE, [(btn, 10, True, NAVY, 0)], align=PP_ALIGN.CENTER,
             line=NAVY, lw=1.0, margins=(0.03, 0.0, 0.03, 0.0))
    note(s, 6.58, ["※募集期間外は、次回の案内開始時に届ける。入口（離脱防止ポップアップ／通常の友だち追加）で、0日目のあいさつの言葉を変える",
                   "※一般の方には一般の方向けの案内を中心に配信。配信内容は、貴社LPの「生産性4倍セミナーの特徴」「受講生の声」「よくある質問」「講師紹介」をもとにした例"])
    return s


# ── 4. 16日以降（FMT S17・月のバーは外す） ────────────────────────
def page_after(prs, f17):
    s = copy_slide(prs, f17)
    remove(s, {4, 8, 10})
    set_lines(sid(s, 10), ["15日を過ぎた友だちには回数を絞ってつながりを保ち、",
                           "開催月は企画配信で、予約・参加まで後押しする。"])
    lanes = [
        ("開催のない月", "月1〜2回", 2.2, [("講師のコラム", "営業のヒントを届ける"), ("受講生の声", "成果が出た方の話"),
                                       ("次回の予告", "次のセミナーの時期")]),
        ("開催月", "月4〜5回", 3.9, [("開催のご案内", "日程と予約ページ"), ("見どころ", "動画・受講生の声"),
                                   ("予約締切", "締切をお知らせ"), ("前日・当日", "予約した方へ\nリマインド"),
                                   ("参加後", "サロンのご案内\n（次のページ）")]),
    ]
    for name, freq, y, boxes in lanes:
        hh = 1.4
        rect(s, i(0.46), i(y), i(1.75), i(hh), NAVY, [(name, 13, True, WHITE, 2), (freq, 12, False, WHITE, 0)], align=PP_ALIGN.CENTER)
        n = len(boxes)
        aw = 0.24
        bw = (10.36 - 2.36 - aw * (n - 1)) / n
        for k, (head, body) in enumerate(boxes):
            x = 2.36 + (bw + aw) * k
            rect(s, i(x), i(y), i(bw), i(hh), WHITE, [(head, 13, True, NAVY, 3)] + [(t, 12, False, INK, 0) for t in body.split("\n")],
                 align=PP_ALIGN.CENTER, line=LINE, lw=0.75, margins=(0.08, 0.05, 0.08, 0.05))
            if k < n - 1:
                rich(s, i(x + bw), i(y), i(aw), i(hh), [[("▶", 11, False, NAVY)]], align=PP_ALIGN.CENTER)
    note(s, 5.5, ["※前日・当日のリマインドと参加後のフォローは、予約・参加した方を特定できる範囲で実施（予約システム等との連携は、可否を確認のうえ設計）"])
    return s


def main(src, fmt_path, out):
    prs = Presentation(src)
    fmt = Presentation(fmt_path)
    fs = list(fmt.slides)
    slides = list(prs.slides)
    assert len(slides) == 61 and len(fs) == 43
    p32, p33, p34 = slides[31], slides[32], slides[33]
    assert p34.shapes[0].text_frame.text.startswith("友だち追加直後のステップ配信")
    new = [page_zengo(prs, fs[14]), page_shindan(prs, fs[12]), page_step(prs, fs[15]), page_after(prs, fs[16])]
    lst = prs.slides._sldIdLst
    el = {int(e.get("id")): e for e in lst}
    order = [el[x.slide_id] for x in [p32] + new + [p33]]
    for e in order[1:]:
        lst.remove(e)
    anchor = order[0]
    for e in order[1:]:
        anchor.addnext(e)
        anchor = e
    drop_slide(prs, p34)
    assert len(prs.slides) == 64
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
