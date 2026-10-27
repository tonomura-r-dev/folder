# -*- coding: utf-8 -*-
"""LINEログイン・Profile+のご提案 ver1.0（汎用・簡易版の資料化。2026-10-27 殿村さん）。

    python -m _build.build_line_login_profileplus_v1 <サンクスLINE誘導のご提案_Appendix_ver1.1.pptx> <出力.pptx>

  骨組みはサンクスLINEのAppendix（4:3・16枚）を横展開する。
  ・1〜9ページ（表紙・会社概要・弊社の強み・受賞・ポジショニング・アジェンダ）はそのまま使い、表紙とアジェンダの資料名だけ差し替える
  ・10〜16ページ（サンクスLINEの本編）は外し、本編8ページを新しく組む（13ページ「ターゲット業界別」の見出しの形を使う）
    10 課題／11 全体の導線／12 LINEログイン＆友だち追加／13 ID連携／14 Profile+／15 期待効果／16 費用／17 導入の流れ
  ・デモ動画は後から（AIのデモ化）。12・14ページに動画の差し込み枠（QRコードの位置）だけ置く
  ・費用は汎用資料の表のまま：ID連携＝DYM共通資料の成果報酬型プランの表（初期20万円〜／月額5万円〜）、LINEログイン＝0円。
    Profile+は汎用資料に金額が無いので「別途お見積もり」（佐村さんに確認中）
  色は紺・白・黒・灰色。緑はLINEの機能の部分だけ。字体はメイリオ。
"""
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm, Pt

from _build.pptx_parts import (GRAY, GREEN, INK, LGRAY, NAVY, WHITE, arrow_line, drop_slide, keep_header_only, label,
                               rect, retitle, rgb, shape, style_run)
from _build.patch_tanaka_v5 import copy_slide
from _build.patch_tanaka_v4_thanks_option import set_lines, sid

DOC = "LINEログイン・Profile+のご提案"
PALE = "F2F2F2"          # 示唆・注記の枠（薄い灰色の塗りだけ。左の縦線は付けない）
X0, W0 = 1.8, 23.9       # 本文の左端と幅（cm）


def box(s, x, y, w, h, text, size=12, fill=WHITE, line=NAVY, color=INK, bold=True, kind=MSO_SHAPE.ROUNDED_RECTANGLE):
    lines = text.split("\n")
    return shape(s, kind, x, y, w, h, [(t, size, bold, color, 0) for t in lines], fill=fill, line=line, lw=1.0,
                 adj=0.12 if kind == MSO_SHAPE.ROUNDED_RECTANGLE else None, margins=(0.15, 0.05, 0.15, 0.05))


def note(s, y, text, size=9):
    return label(s, X0, y, W0, 0.6, [(text, size, False, GRAY, 0)])


def takeaway(s, y, text, h=1.1):
    return shape(s, MSO_SHAPE.RECTANGLE, X0, y, W0, h, [(text, 13, True, NAVY, 0)], fill=PALE)


def table(s, x, y, w, col_w, rows, head_size=11, body_size=11, row_h=0.95):
    """rows[0]＝見出し。列幅は col_w（cm）"""
    gs = s.shapes.add_table(len(rows), len(col_w), Cm(x), Cm(y), Cm(w), Cm(row_h * len(rows)))
    tb = gs.table
    tbl = tb._tbl
    tbl.tblPr.set("firstRow", "1")
    tbl.tblPr.set("bandRow", "0")
    for c, cw in enumerate(col_w):
        tb.columns[c].width = Cm(cw)
    for r, row in enumerate(rows):
        tb.rows[r].height = Cm(row_h)
        for c, text in enumerate(row):
            cell = tb.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(NAVY if r == 0 else (PALE if c == 0 else WHITE))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Cm(0.2)
            cell.margin_top = cell.margin_bottom = Cm(0.08)
            tf = cell.text_frame
            tf.word_wrap = True
            for k, line in enumerate(text.split("\n")):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.CENTER if (r == 0 or c == 0) else PP_ALIGN.LEFT
                run = p.add_run()
                run.text = line
                style_run(run, head_size if r == 0 else body_size, r == 0 or c == 0, WHITE if r == 0 else (NAVY if c == 0 else INK))
    return gs


def video_frame(s, x, y, w, h, name):
    """デモ動画の差し込み枠（QRコードと「動画を見る」の位置だけ押さえる）"""
    fr = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, None, fill=WHITE, line=GRAY, lw=1.0, adj=0.06)
    fr.line.dash_style = 4
    q = min(h - 1.0, 2.6)
    shape(s, MSO_SHAPE.RECTANGLE, x + 0.4, y + (h - q) / 2, q, q, [("QRコード", 9, False, GRAY, 0)], fill=PALE, line=LGRAY, lw=0.75)
    label(s, x + q + 0.7, y + h / 2 - 1.1, w - q - 1.0, 2.2,
          [(f"デモ動画（{name}）", 11, True, NAVY, 2), ("QRコードを読み取って", 10, False, INK, 0), ("動画を見る", 10, False, INK, 0)],
          anchor=MSO_ANCHOR.MIDDLE)


def flow(s, y, steps, h=1.9, gap=0.75):
    """steps = [(text, kind)]  kind: 'line'＝LINEの機能（緑）／'cv'＝紺の塗り／''＝白＋紺枠"""
    n = len(steps)
    w = (W0 - gap * (n - 1)) / n
    xs = []
    for i, (text, kind) in enumerate(steps):
        x = X0 + i * (w + gap)
        xs.append(x)
        if kind == "line":
            box(s, x, y, w, h, text, fill=GREEN, line=None, color=WHITE)
        elif kind == "cv":
            box(s, x, y, w, h, text, fill=NAVY, line=None, color=WHITE)
        else:
            box(s, x, y, w, h, text)
        if i:
            arrow_line(s, x - gap + 0.1, y + h / 2, x - 0.1, y + h / 2, NAVY, 1.75)
    return xs, w


# ---------------- 本編 ----------------
def p_kadai(s):
    retitle(s, "LINEログイン・Profile+｜解決したい課題", "会員登録・フォーム入力・登録後の連絡。CVまでの3か所で、ユーザーは離脱しやすい")
    xs, w = flow(s, 4.8, [("サイト訪問", ""), ("会員登録", ""), ("予約・購入\nフォーム", ""), ("CV", "cv"), ("その後の連絡", "")])
    issues = {1: ("会員登録の面倒さ", "ID・パスワード・メールアドレスの入力が多い"),
              2: ("フォーム入力時の離脱", "氏名・電話番号・住所を一から入力"),
              4: ("連絡が届きにくい", "メルマガ等の開封率が下がり、追客が続かない")}
    for i, (head, body) in issues.items():
        x = xs[i]
        arrow_line(s, x + w / 2, 6.8, x + w / 2, 7.6, GRAY, 1.25)
        label(s, x - 0.3, 7.7, w + 0.6, 2.6, [("離脱", 10, True, GRAY, 1), (head, 12, True, NAVY, 2), (body, 11, False, INK, 0)],
              align=PP_ALIGN.CENTER)
    takeaway(s, 11.6, "3か所の手間を、LINEでまとめて解消する仕組みが必要")


def p_zentai(s):
    retitle(s, "LINEログイン・Profile+｜全体の導線", "会員登録・友だち追加・フォーム入力の3か所の手間を、LINEでワンストップで解消")
    steps = [("サイト訪問", ""), ("LINEで\nログイン", "line"), ("友だち追加", "line"), ("ID連携", "line"),
             ("Profile+で\n自動入力", "line"), ("CV", "cv"), ("LINEで\n追客", "line")]
    xs, w = flow(s, 4.8, steps, h=2.0, gap=0.55)
    subs = {1: "ID・PWなしで\n会員登録", 5: "予約・購入", 2: "同意画面で\n同時に案内", 3: "会員情報と\nLINEを紐づけ",
            4: "フォームの\n入力の手間を削減", 6: "リマインド・\n会員ごとの配信"}
    for i, t in subs.items():
        label(s, xs[i] - 0.2, 7.0, w + 0.4, 1.6, [(x, 10, False, INK, 0) for x in t.split("\n")], align=PP_ALIGN.CENTER)
    table(s, X0, 9.3, W0, [5.0, 9.45, 9.45], [
        ["手間がかかる所", "これまで", "LINEログイン・Profile+の導入後"],
        ["会員登録", "ID・パスワード・メールアドレスを入力", "「LINEでログイン」ボタンから登録"],
        ["友だち追加", "サイトとは別に、LINEで友だち追加", "ログインの同意画面で、同時に案内"],
        ["フォーム入力", "氏名・電話番号・住所を一から入力", "LINEの登録情報で自動入力"],
    ], row_h=0.95)
    note(s, 13.4, "※緑はLINEの機能。友だち追加は、同意画面のチェックを外さずにログインした方が対象")


def p_login(s):
    retitle(s, "LINEログイン｜ワンタップで登録＋友だち追加", "「LINEでログイン」から、ID・パスワードなしで登録＋友だち追加まで案内")
    ys = [4.5, 7.3, 10.1]
    steps = [("1", "サイトの「LINEでログイン」ボタンをタップ", "会員登録・ログインのページにボタンを設置"),
             ("2", "LINEの同意画面", "友だち追加のチェックボックスを表示（初期状態はチェックあり）"),
             ("3", "登録完了＋公式アカウントの友だちに", "チェックを外さずにログインした方が、そのまま友だちに")]
    for (num, head, body), y in zip(steps, ys):
        shape(s, MSO_SHAPE.OVAL, X0, y + 0.35, 1.2, 1.2, [(num, 14, True, WHITE, 0)], fill=GREEN)
        label(s, X0 + 1.6, y, 12.6, 2.0, [(head, 13, True, NAVY, 3), (body, 11, False, INK, 0)], anchor=MSO_ANCHOR.MIDDLE)
        if num != "3":
            arrow_line(s, X0 + 0.6, y + 1.65, X0 + 0.6, y + 2.75, GREEN, 1.5)
    video_frame(s, 16.6, 4.6, 9.1, 7.3, "LINEログイン・友だち追加")
    shape(s, MSO_SHAPE.RECTANGLE, X0, 12.9, W0, 1.6,
          [("同意画面を介した、自然な友だち化", 12, True, NAVY, 2),
           ("LINEログインの「友だち追加オプション」を有効にして表示。チェックを外した方は友だちにならない", 10, False, INK, 0)], fill=PALE)


def p_id(s):
    retitle(s, "ID連携｜会員情報とLINEを紐づけ、1to1の配信へ", "自社の会員IDとLINEアカウントを裏側で紐づけ、会員ごとに合わせたメッセージを届ける")
    box(s, X0, 4.6, 7.0, 2.4, "自社の会員データ\n会員ID・購入履歴・予約", size=12)
    box(s, X0 + 8.45, 4.95, 7.0, 1.7, "ID連携", size=14, fill=GREEN, line=None, color=WHITE)
    box(s, X0 + 16.9, 4.6, 7.0, 2.4, "LINEアカウント\n（公式アカウントの友だち）", size=12)
    arrow_line(s, X0 + 7.1, 5.8, X0 + 8.35, 5.8, NAVY, 1.75)
    arrow_line(s, X0 + 15.55, 5.8, X0 + 16.8, 5.8, NAVY, 1.75)
    label(s, X0, 7.6, W0, 0.7, [("紐づけると、できること（例）", 12, True, NAVY, 0)])
    table(s, X0, 8.4, W0, [8.4, 15.5], [
        ["配信の例", "使う会員データ"],
        ["発送通知・次回購入のリマインド", "購入履歴"],
        ["かご落ちのお知らせ", "カートに入れたまま購入していない商品"],
        ["会員ランクに応じた限定クーポンの出し分け", "会員ランク"],
        ["予約のリマインド", "予約日時・来場予約の内容"],
    ], row_h=0.92)
    note(s, 13.3, "※配信には、ID連携した方が公式アカウントの友だちであることが必要")


def p_profile(s):
    retitle(s, "Profile+｜予約・購入フォームを自動入力", "LINEの登録情報で予約・購入フォームを自動入力し、入力の手間による離脱を防ぐ")
    label(s, X0, 4.3, 14.0, 0.7, [("例：不動産の来場予約・EC購入", 12, True, NAVY, 0)])
    box(s, X0, 5.2, 6.4, 1.0, "通常のフォーム", size=12, fill=PALE, line=None, color=NAVY)
    box(s, X0 + 7.7, 5.2, 6.4, 1.0, "Profile+", size=12, fill=GREEN, line=None, color=WHITE)
    for k, (a, b) in enumerate([("氏名を入力", "LINEで同意"), ("電話番号を入力", "氏名・電話番号・住所を\n自動入力"), ("住所を入力", "内容を確認"), ("送信", "送信")]):
        y = 6.6 + k * 1.45
        box(s, X0, y, 6.4, 1.05, a, size=11, bold=False)
        box(s, X0 + 7.7, y, 6.4, 1.05, b, size=11, bold=False, line=GREEN)
    arrow_line(s, X0 + 6.55, 8.6, X0 + 7.55, 8.6, NAVY, 1.75)
    video_frame(s, 16.6, 5.2, 9.1, 6.7, "Profile+")
    shape(s, MSO_SHAPE.RECTANGLE, X0, 12.7, W0, 1.0,
          [("※Profile+のご利用には、LINEヤフー株式会社による事前審査・承認が必要", 11, True, NAVY, 0)], fill=PALE)


def p_kouka(s):
    retitle(s, "期待効果｜CVR改善とLTV向上の両立", "入力の手間を減らしてCVを増やし、LINEでつながった方にリピート・追客を続ける")
    table(s, X0, 4.6, W0, [5.0, 12.4, 6.5], [
        ["期待効果", "何が変わるか", "関わる機能"],
        ["CVR改善", "会員登録とフォーム入力のステップが減り、\n予約・購入までの離脱を抑える", "LINEログイン\nProfile+"],
        ["友だちの増加", "ログインと同時に友だち追加を案内し、\nCV前の方ともつながる", "LINEログイン\n（友だち追加オプション）"],
        ["LTV向上", "CV後もLINEで連絡が届き、\nリマインド・会員ごとの配信でリピート・追客を続ける", "ID連携\nLINE公式アカウント"],
    ], body_size=12, row_h=1.7)
    takeaway(s, 12.1, "入口（登録・入力）の手間を減らし、CV後もLINEでつながり続ける")


def p_hiyou(s):
    retitle(s, "費用｜LINEヤフー社の利用料と、弊社の開発・保守費用", "LINEログインの機能は無料。ID連携・Profile+の開発・保守は弊社で対応")
    table(s, X0, 4.6, W0, [5.2, 6.5, 5.6, 6.6], [
        ["区分", "項目", "初期", "月額"],
        ["LINEヤフー社の\n利用料", "LINEログイン", "0円", "0円"],
        ["LINEヤフー社の\n利用料", "LINE公式アカウント", "―", "月額プラン料金＋\nメッセージ通数に応じた費用"],
        ["弊社の\n開発・保守費用", "ID連携", "200,000円〜", "50,000円〜"],
        ["弊社の\n開発・保守費用", "Profile+", "別途お見積もり", "別途お見積もり"],
    ], body_size=12, row_h=1.45)
    tb = next(sh for sh in s.shapes if sh.has_table).table
    for top, low in ((1, 2), (3, 4)):                  # 区分のセルを縦に結合（下のセルの文字は消してから）
        for para in list(tb.cell(low, 0).text_frame.paragraphs)[1:]:
            para._p.getparent().remove(para._p)
        tb.cell(low, 0).text_frame.paragraphs[0].runs[0].text = ""
        tb.cell(top, 0).merge(tb.cell(low, 0))
    note(s, 12.2, "※LINE公式アカウントのメッセージ配信料等は、別途発生")
    note(s, 12.8, "※施策内容・連携範囲等により、費用は個別にお見積もりいたします。")


def p_nagare(s):
    retitle(s, "導入の流れ｜お申込みから公開まで", "要件定義の段階でProfile+の審査申請を進め、開発・検証を経て公開")
    steps = [("お申込み", "ご要望・現在の会員登録・\nフォームの確認"),
             ("要件定義・設計", "導線と連携範囲を設計\nProfile+の審査申請"),
             ("開発", "LINEログイン・ID連携・\nProfile+の実装"),
             ("検証・公開", "動作を確認して公開\n公開後は配信・改善へ")]
    n, gap = len(steps), 0.3
    w = (W0 - gap * (n - 1)) / n
    for i, (head, body) in enumerate(steps):
        x = X0 + i * (w + gap)
        shape(s, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON, x, 4.8, w, 1.6, [(head, 13, True, WHITE, 0)],
              fill=NAVY, adj=0.25, margins=(0.6, 0.05, 0.4, 0.05))
        label(s, x + 0.2, 6.8, w - 0.4, 2.0, [(t, 11, False, INK, 0) for t in body.split("\n")], align=PP_ALIGN.CENTER)


def center_body(s, bottom=17.6):
    body = [sh for sh in s.shapes if sh.name not in ("TextBox 1", "TextBox 2", "Connector 3")]
    low = max((sh.top + sh.height) for sh in body) / 360000
    shift = (bottom - low) / 3
    if shift > 0:
        for sh in body:
            sh.top = sh.top + Cm(shift)


PAGES = [p_kadai, p_zentai, p_login, p_id, p_profile, p_kouka, p_hiyou, p_nagare]


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    assert len(slides) == 16
    assert slides[0].shapes[0].text_frame.text == "サンクスLINE誘導のご提案"
    set_lines(sid(slides[0], 2).text_frame, [DOC])
    set_lines(sid(slides[8], 2).text_frame, ["LINEログイン・Profile+導入のご提案"])
    base = slides[12]                                   # ターゲット業界別（見出し＝TextBox 1・2＋区切り線）
    assert base.shapes[0].text_frame.text.startswith("ターゲット業界別")
    for build in PAGES:
        s = copy_slide(prs, base)
        keep_header_only(s)
        build(s)
        center_body(s)
    for s in slides[9:]:
        drop_slide(prs, s)
    assert len(prs.slides) == 17
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
