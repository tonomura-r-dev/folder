# -*- coding: utf-8 -*-
"""株式会社ワクトリ：提案資料の骨組み（v1）を、提案FMT（MAX版）から作る（2026-10-26 殿村さん「ざっくりでいいから骨組みだけ」）。

    python -m _build.build_wakutori_v1 <提案FMT_MAX版_20261024d.pptx> <タナカホーム_v3.pptx> <出力.pptx>

  並びは `_drafts/ワクトリ_与件整理（2026-10-25）.md` の「提案概要（殿村さん確定版）」のとおり。
  01 施策背景：FMTの3枚の後ろに、競合5社の比較／各社1枚×4／まとめ／ワクトリの現状／ヒアリング を足す
  02 LINE施策全体図：FMTのまま。短期・長期の3枚は、転職の例を外して差し込み枠（前後検索・トレンドはこれから）
  03 施策詳細：FMTの「具体施策のイメージ」を写して8施策（札＝集客／即時分類／分岐育成／成果回収）。中身は枠のまま
  04 料金形態：タナカホームの3ページ（概算お見積り／料金プラン／運用内容）を写して、ワクトリのプランに。SIMの数字は「--」
  05 サポート体制：FMTの提供施策の後ろに、効果が悪いときの改善例／報告の粒度／体制・スケジュール を足す
  06 次回ご提案に向けて：FMTのまま
"""
import copy
import os
import sys

from pptx import Presentation
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from _build.patch_tanaka_v5 import copy_slide
from _build.pptx_parts import drop_slide, rect, rich, shape

NAVY, INK, SUB, LINE, LIGHT, WHITE, GRAY = "1F2A5C", "333333", "777777", "D0D5E0", "EEF1F8", "FFFFFF", "F3F5FA"
IN = 2.54
IMG = "_data/wakutori/line/"


def i(v):
    return v * IN


# ── 文字の差し替え ────────────────────────────────────────────
def set_p(p, text):
    rs = p.findall(qn("a:r"))
    for r in rs[1:]:
        p.remove(r)
    rs[0].find(qn("a:t")).text = text


def replace_text(s, old, new, must=True):
    hit = 0
    for p in s._element.iter(qn("a:p")):
        rs = p.findall(qn("a:r"))
        if rs and "".join(r.find(qn("a:t")).text or "" for r in rs) == old:
            set_p(p, new)
            hit += 1
    assert hit or not must, old
    return hit


def set_cell(cell, text):
    """表のセル：最初の段落の書式で、改行ごとに段落を作り直す"""
    txb = cell._tc.find(qn("a:txBody"))
    ps = txb.findall(qn("a:p"))
    base = ps[0]
    for p in ps:
        txb.remove(p)
    for line in text.split("\n"):
        p = copy.deepcopy(base)
        rs = p.findall(qn("a:r"))
        if not rs:
            continue
        for r in rs[1:]:
            p.remove(r)
        for br in p.findall(qn("a:br")):
            p.remove(br)
        rs[0].find(qn("a:t")).text = line
        txb.append(p)


# ── ページの土台 ──────────────────────────────────────────────
def page_a(prs, base, header, title, lead):
    """S12 の形（上に「前提調査」・タイトル＋下線・リード・区切り線）"""
    s = copy_slide(prs, base)
    keep = {"正方形/長方形 57", "正方形/長方形 2", "テキスト ボックス 4", "テキスト ボックス 6", "テキスト ボックス 8"}
    for sh in list(s.shapes):
        if sh.name not in keep:
            sh._element.getparent().remove(sh._element)
    by = {sh.name: sh for sh in s.shapes}
    set_p(by["テキスト ボックス 8"].text_frame.paragraphs[0]._p, " " + header)
    set_p(by["テキスト ボックス 4"].text_frame.paragraphs[0]._p, title)
    set_p(by["テキスト ボックス 6"].text_frame.paragraphs[0]._p, lead)
    by["正方形/長方形 57"].width = Inches(min(0.25 * len(title) + 0.2, 6.0))
    return s


def page_b(prs, base, title, lead):
    """S39 の形（上のバーにタイトル・リード・区切り線）"""
    s = copy_slide(prs, base)
    keep = {"テキスト ボックス 47", "Rectangle 2", "正方形/長方形 52"}
    for sh in list(s.shapes):
        if sh.name not in keep:
            sh._element.getparent().remove(sh._element)
    by = {sh.name: sh for sh in s.shapes}
    set_p(by["テキスト ボックス 47"].text_frame.paragraphs[0]._p, title)
    set_p(by["Rectangle 2"].text_frame.paragraphs[0]._p, lead)
    return s


def placeholder(s, x, y, w, h, label):
    sh = shape(s, MSO_SHAPE.RECTANGLE, i(x), i(y), i(w), i(h), [(label, 12, False, SUB, 0)], fill=None, line="A6A6A6", lw=1.0)
    sh.line.dash_style = MSO_LINE.DASH
    return sh


def note(s, text, y=6.95):
    rich(s, i(0.42), i(y), i(10.0), i(0.25), [[(text, 9, False, SUB)]], align=PP_ALIGN.LEFT)


def table(s, x, y, widths, head, rows, rh=0.62, hh=0.4, size=11.5, first_bold=True):
    xs = [x]
    for w in widths[:-1]:
        xs.append(xs[-1] + w + 0.03)
    for (cx, w), t in zip(zip(xs, widths), head):
        rect(s, i(cx), i(y), i(w), i(hh), NAVY, [(t, 12, True, WHITE, 0)], align=PP_ALIGN.CENTER)
    for r, row in enumerate(rows):
        ry = y + hh + 0.04 + (rh + 0.04) * r
        for k, ((cx, w), t) in enumerate(zip(zip(xs, widths), row)):
            first = k == 0
            rect(s, i(cx), i(ry), i(w), i(rh), LIGHT if first else WHITE,
                 [(line, size, first and first_bold, NAVY if first else INK, 0) for line in t.split("\n")],
                 align=PP_ALIGN.LEFT, line=LINE, lw=0.75, margins=(0.2, 0.05, 0.15, 0.05))
    return y + hh + 0.04 + (rh + 0.04) * len(rows)


def phones(s, files, x, y, h, gap=0.12, captions=None):
    w = h * 277 / 600
    for k, f in enumerate(files):
        px = x + (w + gap) * k
        src = IMG + "伏せ字/" + f if os.path.exists(IMG + "伏せ字/" + f) else IMG + f   # 宛名を伏せた画像があればそちら
        s.shapes.add_picture(src, Inches(px), Inches(y), Inches(w), Inches(h))
        if captions:
            rich(s, i(px - 0.05), i(y + h + 0.03), i(w + 0.1), i(0.3), [[(captions[k], 10.5, True, NAVY)]], align=PP_ALIGN.CENTER)
    return x + (w + gap) * len(files) - gap


def facts(s, x, y, w, items, bh=1.35):
    for k, (head, body) in enumerate(items):
        by_ = y + (bh + 0.12) * k
        rect(s, i(x), i(by_), i(w), i(0.34), NAVY, [(head, 12, True, WHITE, 0)], align=PP_ALIGN.LEFT, margins=(0.25, 0.0, 0.1, 0.0))
        rect(s, i(x), i(by_ + 0.34), i(w), i(bh - 0.34), WHITE, [(t, 11.5, False, INK, 2) for t in body.split("\n")],
             align=PP_ALIGN.LEFT, line=LINE, lw=0.75, anchor=MSO_ANCHOR.TOP, margins=(0.25, 0.1, 0.15, 0.05))


# ── 01 施策背景に足すページ ────────────────────────────────────
COMPANIES = [  # 会社, 友だち数, 追加直後, 配信の頻度, 相談の入口, リッチメニュー
    ("ワクトリ", "79,904", "アンケート3問\n（エントリー・開始時期・エリア）", "翌日にENTRY\nその後は週1回", "専属エージェント\n（ENTRY）", "お仕事探し"),
    ("リゾートバイトダイブ", "256,680", "マイページ登録のURL\n＋LINEでできること", "毎日1通", "初回カウンセリングの予約\n（電話15分）", "6マス\n開始月別の求人あり"),
    ("リゾバ.com", "156,010", "リゾートバイト経験の\n有無を質問", "1日2通", "まずは相談する\n担当者がチャットで返信", "タブ2つ\n開始月別の求人あり"),
    ("リゾートバイト.com", "32,225", "アンケート\n（性別から）", "—", "お問い合わせ\n（メニュー）", "タブ2つ\n開始月別の求人あり"),
    ("アルファリゾート", "26,846", "会員登録と\nカウンセリング予約のURL", "—", "初回カウンセリング必須\n担当者がチャットで返信", "6マス\n開始月別の求人あり"),
]
DETAIL = {
    "リゾートバイトダイブ": (
        "追加の直後から毎日1通。登録 → 流れ → カウンセリング → 稼げる額の順に、不安を消して登録・相談へつなぐ。",
        ["ダイブ_01.jpg", "ダイブ_03.jpg", "ダイブ_04.jpg"],
        [("あいさつ", "マイページ登録のURLを最初に置き、LINEでできること（求人を探す・探してもらう・相談）を案内"),
         ("追加後の配信（毎日1通）", "開始までの5ステップ → 初回カウンセリング（電話15分・履歴書不要）→ 稼げる額（平均月収 約25万円）→ 口コミ・SNS"),
         ("リッチメニュー", "求人を探す／会員登録／ログイン／11月・12月開始のお仕事／よくある質問")],
        "出典｜リゾートバイトダイブ LINE公式アカウント（2026年10月20〜24日の配信）"),
    "リゾバ.com": (
        "追加の直後に経験の有無を聞き、1日2通、相談・タイプ別・条件別・エリア別の特集で求人へつなぐ。",
        ["リゾバ_01.jpg", "リゾバ_02.jpg", "リゾバ_05.jpg"],
        [("あいさつ", "採用コーディネーターがサポートすると伝え、会員登録（まずは相談する）へ。続けて「リゾートバイト経験はありますか？」"),
         ("追加後の配信（1日2通）", "「話を聞いてみませんか？」／「あなたはどのタイプ？」（タイプ別の求人）／条件特集・エリア特集・レア求人"),
         ("リッチメニュー", "探したい・働きたい（新着・条件検索・応募・11〜1月開始）／知りたい（はじめての方へ・体験談・特典）")],
        "出典｜リゾバ.com LINE公式アカウント（2026年10月20〜22日の配信）"),
    "リゾートバイト.com": (
        "追加の直後にアンケートから始め、リッチメニューで開始月別の求人と登録へつなぐ。",
        ["リゾートバイトcom_01.jpg", "リゾートバイトcom_02.jpg"],
        [("あいさつ", "アンケート「Q1.性別を教えてください」から始める"),
         ("リッチメニュー（おしごと探し）", "リゾートバイト検索／10月・11月勤務開始／スキー場バイト／かんたんWEB登録（約50秒）"),
         ("リッチメニュー（サポートナビ）", "はじめての方へ／質問はこちら／体験談／お問い合わせ／マイページ／配信設定")],
        "出典｜リゾートバイト.com LINE公式アカウント（2026年10月20日の画面）"),
    "アルファリゾート": (
        "あいさつで会員登録と初回カウンセリングの予約を案内し、担当者がチャットで相談に乗る。",
        ["アルファ_01.jpg", "アルファ_02.jpg"],
        [("あいさつ", "まずは無料で会員登録（URL）／初回カウンセリングの予約（URL・必須）／気になることは担当者が返信"),
         ("リッチメニュー", "マイページ ログイン／求人検索／お役立ち情報／マイページ作成／よくある質問／10月・11月スタートのお仕事"),
         ("ねらい", "LINEを、登録とカウンセリングの受付の窓口にしている")],
        "出典｜リゾートバイトならアルファリゾート LINE公式アカウント（2026年10月20日の画面）"),
}


def add_chapter1(prs, s12, s39):
    out = []
    s = page_a(prs, s12, "前提調査", "リゾートバイト業界のLINE活用",
               "主要5社のLINEを比べると、友だち追加の直後から、登録・相談へつなぐ配信が当たり前になっている。")
    table(s, 0.42, 1.98, [1.9, 1.05, 2.0, 1.45, 1.95, 1.5],
          ["会社", "友だち数", "追加の直後", "配信の頻度", "相談の入口", "リッチメニュー"],
          [list(c) for c in COMPANIES], rh=0.78, size=11)
    note(s, "出典｜各社のLINE公式アカウント（2026年10月に確認）。友だち数はプロフィールの表示。「—」は確認した画面では配信が見えなかったもの")
    out.append(s)
    for name, (lead, files, items, src) in DETAIL.items():
        s = page_a(prs, s12, "前提調査", "各社の取り組み｜" + name, lead)
        end = phones(s, files, 0.45, 1.98, 4.6)
        facts(s, end + 0.35, 1.98, 10.43 - end - 0.35, items, bh=1.45)
        note(s, src)
        out.append(s)
    s = page_a(prs, s12, "前提調査", "他社のLINE活用のまとめ",
               "他社は、友だち追加の直後から毎日つないでいる。リゾートバイトの集客で、LINEは必須の接点になっている。")
    table(s, 0.42, 2.0, [3.6, 6.37], ["他社がやっていること", "どこがやっているか"],
          [["1. 追加の直後に、登録・相談へ誘導する", "リゾートバイトダイブ／リゾバ.com／アルファリゾート"],
           ["2. 追加の後も、毎日配信する", "リゾートバイトダイブ（毎日1通）／リゾバ.com（1日2通）"],
           ["3. 経験・タイプを聞き、合う求人を出し分ける", "リゾバ.com（経験の有無・タイプ別）／リゾートバイト.com（アンケート）"],
           ["4. リッチメニューに、開始月別の求人を置く", "4社とも"]], rh=0.7, size=12.5)
    out.append(s)
    s = page_a(prs, s12, "前提調査", "ワクトリの現状のLINE",
               "今の流れは、あいさつ → アンケート3問 → 翌日のENTRY → 週1回の企画配信。絵と文面は見やすく整っている。")
    phones(s, ["ワクトリ_01.jpg", "ワクトリ_02.jpg", "ワクトリ_04.jpg", "ワクトリ_07.jpg"], 0.75, 1.95, 3.3, gap=0.75,
           captions=["あいさつ", "アンケート3問", "翌日のENTRY", "企画配信（毎週火曜）"])
    table(s, 0.42, 5.68, [2.6, 7.37], ["企画投稿の2種", "中身"],
          [["直接応募", "悩み3つ → 条件の特集10選 → 求人カード → 今すぐチェック"],
           ["会員登録の引き上げ", "専属エージェントがサポート → 重視したいポイント（立地・稼げる など）→ ENTRY"]],
          rh=0.4, hh=0.32, size=11)
    out.append(s)
    s = page_b(prs, s39, "これまでの運用について", "これまでの運用の課題感と、これからやりたい施策を伺う。")
    for k, head in enumerate(["これまでの運用の課題感", "これからやりたい施策"]):
        x = 0.42 + 5.04 * k
        rect(s, i(x), i(2.0), i(4.96), i(0.4), NAVY, [(head, 13, True, WHITE, 0)], align=PP_ALIGN.CENTER)
        rect(s, i(x), i(2.44), i(4.96), i(4.0), WHITE, None, line=LINE, lw=0.75)
    out.append(s)
    return out


# ── 03 施策詳細 ───────────────────────────────────────────────
MEASURES = [  # (写すFMTのページ, 札, 施策名, 補足)
    (20, "集客", "友だちの入口", "サイトのLINEボタン・広告の受け皿"),
    (20, "集客", "ポップアップ", "サイトを離れる方をLINEへ"),
    (21, "即時分類", "適正バイト診断", "今のアンケートと「重視したいポイント」を発展"),
    (22, "分岐育成", "ステップ配信", "追加後15日で、ENTRY・カウンセリングへ"),
    (23, "分岐育成", "開始時期・エリア別の配信", "診断の答えで、配信を出し分ける"),
    (19, "分岐育成", "リッチメニュー", "開始月別の求人・相談・ENTRY"),
    (24, "成果回収", "サンクスLINE", "マイページ登録の完了画面からLINEへ、カウンセリング・応募まで"),
    (25, "成果回収", "契約満了前の次のお仕事案内", "稼働中・満了前に、次のお仕事をご案内"),
]


def add_chapter3(prs, fs):
    out = []
    navy_fill = next(sh for sh in fs[19].shapes if sh.name == "正方形/長方形 4")._element.spPr.find(qn("a:solidFill"))
    for n, tag, name, sub in MEASURES:
        s = copy_slide(prs, fs[n - 1])
        t = next(sh for sh in s.shapes if sh.name == "正方形/長方形 4")
        set_p(t.text_frame.paragraphs[0]._p, tag)
        old = t._element.spPr.find(qn("a:solidFill"))                  # 札の色は紺にそろえる（FMTの1枚だけ緑）
        old.addprevious(copy.deepcopy(navy_fill))
        t._element.spPr.remove(old)
        rich(s, i(0.45), i(2.15), i(6.1), i(0.75), [[(name, 16, True, NAVY)], [(sub, 12, False, INK)]], anchor=MSO_ANCHOR.TOP)
        out.append(s)
    return out


# ── 04 料金形態（タナカホームの3ページ） ─────────────────────────
def add_chapter4(prs, tanaka):
    ts = list(tanaka.slides)
    est, plan, scope = (copy_slide(prs, ts[k]) for k in (12, 13, 14))
    for old, new in [
        ("概算お見積り：月26万円で、6か月目に資料請求 月11件・来場予約 月8件", "概算お見積り：月--万円で、6か月目に友だち --人・ENTRY 月--件"),
        ("プラン②｜総合運用で、友だちの入口を増やし、検討の順番に沿った配信で資料請求・来場予約を増やす",
         "プラン①（月3投稿）で、友だちの入口を増やし、検討の順番に沿った配信でENTRY・応募を増やす"),
        ("¥265,000〜", "¥--"), ("¥260,000〜", "¥--"), ("資料請求", "ENTRY"), ("来場予約", "応募"),
        ("1,493人", "--人"), ("11件/月", "--件/月"), ("8件/月", "--件/月"),
        ("離脱防止", "ポップアップ"), ("来場前の", "適正バイト"), ("お知らせ", "診断"), ("型をつくる", "設計・構築"),
        ("月4回程度", "月3回"), ("来場前のお知らせ", "開始時期・エリア別"), ("型の改善", "配信"),
    ]:
        replace_text(est, old, new, must=False)
    sx = next(sh for sh in est.shapes if sh.name == "正方形/長方形 17")   # サンクスLINEはオプションなので、初期はリッチメニューに
    sx.text_frame.paragraphs[0].runs[0].text = "リッチメニュー"
    for para in sx.text_frame.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    set_p(next(sh for sh in est.shapes if sh.name == "正方形/長方形 21").text_frame.paragraphs[0]._p, "設計・構築")
    head = ["項目", "プラン①（月3投稿）\nコンサル基本", "プラン②（月5投稿）\nコンサル基本", "プラン③\n成果報酬型"]
    tb = next(sh for sh in plan.shapes if sh.has_table).table
    body = [head,
            ["初期費用", "115,000円〜", "115,000円〜", "--"],
            ["月額コンサルティング費", "200,000円〜", "300,000円〜", "成果単価 × 件数"],
            ["月額ツール・LINE関連費", "30,000円〜", "30,000円〜", "--"],
            ["月額費用", "230,000円〜", "330,000円〜", "--"]]
    for r, row in enumerate(body):
        for c, t in enumerate(row):
            set_cell(tb.cell(r, c), t)
    for old, new in [
        ("※初期費用＝運用設計・構築 100,000円〜＋離脱防止 15,000円＋サンクスLINE誘導 150,000円", "※初期費用＝運用設計・構築 100,000円〜＋ポップアップ 15,000円"),
        ("※月額ツール・LINE関連費＝離脱防止 30,000円＋サンクスLINE誘導 30,000円〜。LINE公式アカウント利用料は別（貴社のご契約のまま）",
         "※月額ツール・LINE関連費＝ポップアップ 30,000円。サンクスLINE誘導はオプション（初期150,000円・月額30,000円〜）。LINE公式アカウント利用料は別"),
        ("※契約期間は6か月〜。概算お見積りのページはプラン②の場合", "※契約期間は6か月〜。概算お見積りのページはプラン①の場合"),
    ]:
        replace_text(plan, old, new, must=False)
    tb = next(sh for sh in scope.shapes if sh.has_table).table
    body = [head,
            ["ポップアップ", "導入・運用", "導入・運用", "--"],
            ["適正バイト診断・\nステップ配信", "設計・構築・改善", "設計・構築・改善", "--"],
            ["企画配信", "月3回程度", "月5回程度", "--"],
            ["開始時期・エリア別の配信", "対応", "対応", "--"],
            ["結果の分析・改善\n（月次レポート・定例会）", "対応", "対応", "--"]]
    for r, row in enumerate(body):
        for c, t in enumerate(row):
            set_cell(tb.cell(r, c), t)
    note(scope, "※サンクスLINE誘導はオプション（初期150,000円・月額30,000円〜）", y=5.8)
    return [est, plan, scope]


# ── 05 サポート体制に足すページ ─────────────────────────────────
def add_chapter5(prs, s39):
    out = []
    s = page_b(prs, s39, "効果が悪いときの改善例", "数字が伸びないときは、どこで止まっているかを見て、配信の中身・順番・対象を直す。")
    table(s, 0.42, 2.0, [2.6, 2.6, 4.77], ["止まっているところ", "見る数字", "直し方"],
          [["友だちが増えない", "友だち追加数・ポップアップの表示回数", "ポップアップの出し方と訴求、広告の着地ページの導線を見直す"],
           ["ブロックが多い", "ブロック率（配信ごと）", "配信の頻度と対象を絞る。診断の答えに合わない配信を止める"],
           ["開封・クリックが低い", "開封率・クリック率", "見出し・画像・送る時間を変えて比べる"],
           ["ENTRYにつながらない", "ENTRY数・ENTRY率", "診断の結果に合う求人と、ステップ配信の順番を見直す"],
           ["登録後に応募が進まない", "カウンセリング・応募の数", "サンクスLINEの後の案内と、カウンセリングの予約の導線を見直す"]], rh=0.66, size=12)
    out.append(s)
    s = page_b(prs, s39, "報告の粒度", "毎月のレポートと定例会で、数字と次の打ち手を共有する。")
    table(s, 0.42, 2.0, [1.6, 3.2, 5.17], ["頻度", "やること", "中身"],
          [["毎月", "月次レポート", "友だち数・ブロック率、配信ごとの開封・クリック、ENTRY・応募"],
           ["毎月", "定例会", "先月の振り返りと、翌月の企画配信の内容"],
           ["四半期", "設計の見直し", "診断の項目、ステップ配信の順番、リッチメニューの並び"]], rh=0.62, size=12)
    placeholder(s, 0.42, 4.55, 9.97, 2.15, "差し込み：月次レポートの見本")
    out.append(s)
    s = page_b(prs, s39, "体制・スケジュール", "設計 → 構築 → 運用開始 → 改善の順に進める。")
    table(s, 0.42, 2.0, [2.0, 7.97], ["段階", "やること"],
          [["設計", "友だちの入口、適正バイト診断の項目、ステップ配信の内容と順番"],
           ["構築", "ポップアップ・サンクスLINE誘導、診断、ステップ配信、リッチメニュー"],
           ["運用開始", "企画配信と、開始時期・エリア別の配信"],
           ["改善", "月次レポート・定例会で、配信の中身・順番・対象を見直す"]], rh=0.6, size=12)
    placeholder(s, 0.42, 4.95, 9.97, 1.75, "差し込み：体制図（弊社と貴社の役割）")
    out.append(s)
    return out


# ── 02 の短期・長期（転職の例を外して差し込み枠に） ─────────────────
def clear_body(s, label):
    for sh in list(s.shapes):
        if sh.top >= Inches(1.5):
            sh._element.getparent().remove(sh._element)
    placeholder(s, 0.42, 1.75, 9.97, 5.0, label)


def main(fmt_path, tanaka_path, out):
    prs = Presentation(fmt_path)
    fs = list(prs.slides)
    assert len(fs) == 43
    replace_text(fs[0], "MAX版", "株式会社ワクトリ 御中")
    replace_text(fs[0], "最上位パートナーの知見を活用し、獲得に強い動線を構築する", "追客・ファン化を再構築する、LINE公式アカウントの設計見直し")
    for n in (39, 40, 41, 42):                                          # 確認したいことの表の黄色の欄は白に
        for e in fs[n - 1]._element.iter(qn("a:srgbClr")):
            if e.get("val").upper() == "FFF8E1":
                e.set("val", "FFFFFF")
    for s in fs:                                                        # 紺の札の上の黄色い文字は白に（赤はFMTのまま）
        for e in s._element.iter(qn("a:srgbClr")):
            if e.get("val").upper() == "FFFF00":
                e.set("val", "FFFFFF")
    sub = next(sh for sh in fs[0].shapes if sh.has_text_frame and sh.text_frame.text.startswith("追客"))
    sub.left, sub.width = Inches(1.46), Inches(7.91)                    # 表紙の小見出しを下線の幅まで広げて1行に
    for n in (5, 9, 18, 27, 29, 38):                                    # 章の区切りの章名が2行に折れないように
        max((sh for sh in fs[n - 1].shapes if sh.has_text_frame and sh.top > Inches(1)), key=lambda sh: sh.width).width = Inches(8.0)
    clear_body(fs[14], "差し込み：前後検索「リゾートバイト」（LINEヤフーのデータ）")
    clear_body(fs[15], "差し込み：Day0〜15のステップ配信（前後検索をもとに作成）")
    clear_body(fs[16], "差し込み：Googleトレンド「リゾートバイト」と、時期ごとの企画配信")

    ch1 = add_chapter1(prs, fs[11], fs[38])
    ch3 = add_chapter3(prs, fs)
    ch4 = add_chapter4(prs, Presentation(tanaka_path))
    ch5 = add_chapter5(prs, fs[38])

    def el(s):
        return next(e for e in prs.slides._sldIdLst if int(e.get("id")) == s.slide_id)

    order = [fs[k] for k in range(0, 8)] + ch1 + [fs[k] for k in range(8, 18)] + ch3 + [fs[26]] + ch4 \
        + [fs[k] for k in range(28, 37)] + ch5 + [fs[k] for k in range(37, 43)]
    els = [el(s) for s in order]
    for s in [fs[k] for k in range(18, 26)] + [fs[27]]:                 # 具体施策のイメージ8枚（写し済み）と空の料金プラン
        drop_slide(prs, s)
    lst = prs.slides._sldIdLst
    for e in els:
        lst.remove(e)
        lst.append(e)
    assert len(prs.slides) == len(order), (len(prs.slides), len(order))
    prs.save(out)
    print("saved:", out, len(prs.slides), "枚")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
