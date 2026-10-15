# -*- coding: utf-8 -*-
"""タナカホーム株式会社「青森店 受注最大化に向けたマーケティング戦略」v2（2026-10-15 殿村さんの骨子・6枚＋裏表紙）。

    python -m _build.build_tanaka_v2 <あらうちv10.pptx（器）> <出力.pptx>

  骨子（殿村さん）：Step1 現状・集客（広告改善・SEOで質の高い流入）→ Step2 課題（即決しない層の離脱・競合流出）
  → Step3 解決策（LINEでファン化・営業×事務）→ Step4 拡張性（青森店で型を作り全店へ）→ だから今LINEを入れる。
  数字は2026-10-15の商談議事録のもの（来場予約 目標50〜60／現状35〜40・CPA約3万・資料請求 月一桁・青森店キャンセル2〜3割）。
  v1（15枚）とは別ファイル。費用・試算のページは入れていない（骨子に無いため）。
"""
import sys

from pptx import Presentation

from _build.arauchi_v4.lib import NAVY, RED, TEXT, add_box, add_table, move, set_paras
from _build.arauchi_v10.build import TINT, LINE_C, by_id, clone, meiryo, runs
from _build.build_tanaka_v1 import GRAY, HDR, COVER, BACK, HDR_SRC, new_page, head_cell, cell, label_cell, note, band

W = 9.73
X0 = 0.55


def heading(s, x, y, w, text):
    add_box(s, x, y, w, 0.38, [runs((text, dict(size=12, bold=True, color="FFFFFF")))], fill=NAVY, align="ctr",
            anchor="ctr", name=f"見出し {text}")


def flow(s, items, y, h, x0=X0, w_total=W, arrow_w=0.32, size=11):
    """横一列の流れ図：items=[(文字, 塗り or None, 文字色)]"""
    n = len(items)
    bw = (w_total - arrow_w * (n - 1)) / n
    for k, (text, fill, color) in enumerate(items):
        x = x0 + k * (bw + arrow_w)
        add_box(s, x, y, bw, h, [runs((t, dict(size=size, bold=True, color=color))) for t in text.split("\n")],
                fill=fill, line=None if fill else NAVY, align="ctr", anchor="ctr", line_spacing=115,
                ins=(0.06, 0.04, 0.06, 0.04), name=f"流れ{k + 1}")
        if k < n - 1:
            add_box(s, x + bw, y, arrow_w, h, [runs(("▶", dict(size=14, color=NAVY)))], align="ctr", anchor="ctr",
                    ins=(0, 0, 0, 0), name="矢印")


# ---------------------------------------------------------------- 1 表紙
def p_cover(s):
    set_paras(by_id(s, 31), ["タナカホーム株式会社 御中"])
    t = by_id(s, 26)
    set_paras(t, ["青森店 受注最大化に向けたマーケティング戦略"], size=26)
    move(t, x=0.6, w=9.63)
    set_paras(by_id(s, 27), ["集客強化から「ファン化・成約」へ繋げる", "LINE活用のご提案"])
    for sid in (31, 26, 27):
        meiryo(by_id(s, sid))


# ---------------------------------------------------------------- 2 現状と打ち手
def p_genjo(prs, src):
    s = new_page(prs, src, "現状と打ち手：広告改善とSEOで、質の高い流入を確保する",
                 "青森店の来場は目標に届いていない。集客フェーズは、広告の改善と地域特化のSEOで立て直す")
    lw, rx, rw = 4.75, 5.55, 4.73
    heading(s, X0, 1.85, lw, "青森店の現在地")
    rows = [[label_cell("来場予約\n（4展示場）", 11), cell("目標 月50〜60件に対し、現状 月35〜40件", 11)],
            [label_cell("来場予約のCPA", 11), cell("Web広告で約30,000円\n（ポータル・特典費を含めると注文住宅 約50,000円）", 10.5)],
            [label_cell("青森店", 11), cell("予約のキャンセルが2〜3割\n八戸・おいらせより認知が浅い", 11)],
            [label_cell("自社サイトの\n資料請求", 11), cell("月一桁", 11)],
            [label_cell("検索広告のCPA\n（9月）", 11), cell("14,796円 → 47,500円（前年同月比）", 11)]]
    add_table(s, X0, 2.3, [1.55, lw - 1.55], [0.62] * 5, rows, name="現在地")

    heading(s, rx, 1.85, rw, "集客フェーズの打ち手")
    add_box(s, rx, 2.3, rw, 1.05, [
        runs(("広告の改善", dict(size=12.5, bold=True, color=NAVY))),
        runs(("青森店の媒体配分とクリエイティブを見直し、来場予約のCPAを下げる（別紙：集客改善シミュレーション）",
              dict(size=10.5, color=TEXT)))], line=NAVY, anchor="ctr", line_spacing=120,
        ins=(0.15, 0.06, 0.12, 0.06), name="打ち手1")
    add_box(s, rx, 3.45, rw, 1.05, [
        runs(("地域特化のSEO", dict(size=12.5, bold=True, color=NAVY))),
        runs(("青森・八戸・弘前の地域キーワードと、AI検索への対応で、家づくりを具体的に検討している方の流入を増やす",
              dict(size=10.5, color=TEXT)))], line=NAVY, anchor="ctr", line_spacing=120,
        ins=(0.15, 0.06, 0.12, 0.06), name="打ち手2")
    add_box(s, rx + rw / 2 - 0.2, 4.52, 0.4, 0.28, [runs(("▼", dict(size=12, color=NAVY)))], align="ctr", anchor="ctr",
            ins=(0, 0, 0, 0), name="矢印下")
    add_box(s, rx, 4.82, rw, 0.62, [runs(("質が高く、熱量の高い見込み客の流入が増える", dict(size=12, bold=True, color="FFFFFF")))],
            fill=NAVY, align="ctr", anchor="ctr", name="結果")
    note(s, "※数字は2026年10月15日の打ち合わせで共有いただいた内容。広告・SEOの詳細は別紙", y=6.4)


# ---------------------------------------------------------------- 3 課題
def p_kadai(prs, src):
    s = new_page(prs, src, "課題：集めた見込み客の「即決しない層」が離脱し、競合に流れる",
                 "流入から成約までの各段階で接点が切れる。検討中の方を受け止める仕組みが、今は無い")
    flow(s, [("流入\n広告・SEO・ポータル", NAVY, "FFFFFF"), ("CV\n資料請求・来場予約 月35〜40件", NAVY, "FFFFFF"),
             ("来場", NAVY, "FFFFFF"), ("商談・成約", NAVY, "FFFFFF")], 1.85, 0.78, size=11)
    walls = ["来訪者の大半はCVせずに離れる。自社サイトの資料請求は月一桁で、検討中の方の受け皿がない",
             "予約の後は電話・メールだけ。青森店は予約のキャンセルが2〜3割",
             "検討が浅い段階での来場があり、その場では決まらない",
             "検討が長い方を追い続ける仕組みがなく、競合（一条工務店・大進建設）に流れる"]
    n, aw = 4, 0.32
    bw = (W - aw * (n - 1)) / n
    for k, t in enumerate(walls):
        x = X0 + k * (bw + aw)
        add_box(s, x, 2.75, bw, 2.0, [runs((f"壁{k + 1}", dict(size=10.5, bold=True, color=RED))),
                                      runs((t, dict(size=10.5, color=TEXT)))],
                line=LINE_C, anchor="t", line_spacing=120, ins=(0.1, 0.08, 0.1, 0.06), name=f"壁{k + 1}")
    band(s, [("最大の損失は、", dict(size=13, bold=True, color=NAVY)),
             ("熱量の高い見込み客を集めても「即決しない層」との接点が切れること", dict(size=13, bold=True, color=RED))],
         y=5.55, h=0.62)
    note(s, "※現状の追客は予約後の電話・メール。既存のLINE公式アカウントはイベント告知が中心で、追客には使えていない", y=6.4)


# ---------------------------------------------------------------- 4 解決策
def p_fan(prs, src):
    s = new_page(prs, src, "解決策：LINEの配信で「この家が好き・タナカホームが好き」を育てる",
                 "売り込みではなく「住みたいイメージ」を配信で育て、来場・商談のときの成約率を底上げする")
    rows = [[head_cell("配信コンテンツ"), head_cell("内容"), head_cell("感じてもらうこと")],
            [label_cell("ルームツアー動画"), cell("完成した家の中を歩く動画。間取り・収納・暮らしの動線", 10.5),
             cell("この家に住みたい", 11.5, True, NAVY, align="ctr")],
            [label_cell("暮らしの風景"), cell("入居後の暮らし・季節ごとの様子・施主の声", 10.5),
             cell("この暮らしがいい", 11.5, True, NAVY, align="ctr")],
            [label_cell("街並み・エリアの魅力"), cell("八戸・おいらせ・青森・弘前の街並み、土地の情報", 10.5),
             cell("この街に住みたい", 11.5, True, NAVY, align="ctr")],
            [label_cell("家づくりの知識"), cell("青森の冬を乗り切る性能、費用の考え方、補助金", 10.5),
             cell("この会社は信頼できる", 11.5, True, NAVY, align="ctr")]]
    add_table(s, X0, 1.85, [2.1, 4.9, 2.73], [0.4] + [0.6] * 4, rows, name="配信")
    flow(s, [("友だち追加\nサイト・来場予約・展示場", None, NAVY), ("配信でファン化", NAVY, "FFFFFF"),
             ("「取れるとき」に来場", NAVY, "FFFFFF"), ("商談の成約率を底上げ", RED, "FFFFFF")], 4.75, 0.72, size=11)
    note(s, "※動画・写真はInstagram・Meta広告の素材を流用し、制作の負担を抑える／既存のLINE公式アカウントを活用", y=6.4)


# ---------------------------------------------------------------- 5 現場体制
def p_taisei(prs, src):
    s = new_page(prs, src, "現場体制：営業の追客の負担を減らし、事務との連携で数字を追う",
                 "営業は関係構築と商談に集中し、事務が配信と顧客リストの状況を管理する")
    rows = [[head_cell("担当"), head_cell("役割"), head_cell("LINEでやること")],
            [label_cell("営業"), cell("関係構築と商談に集中", 11, True, NAVY),
             cell("・検討度が上がった方（配信への反応・質問）に個別チャットで対応\n・来場・商談へつなげる", 10.5)],
            [label_cell("事務\n（数字の管理）"), cell("配信の管理と\n顧客リストの可視化", 11, True, NAVY),
             cell("・配信の予約と反応（開封・クリック）の確認\n・チャットタグで検討状況を管理（検討中・予約済・来場済）\n"
                  "・来場予約の3日前・前日のリマインド配信（キャンセル対策）", 10.5)],
            [label_cell("弊社"), cell("配信の設計・制作・分析", 11, True, NAVY),
             cell("・配信の企画と、文面・画像の制作\n・月次で友だち数・反応・来場予約を分析し、配信を改善", 10.5)]]
    add_table(s, X0, 1.85, [1.5, 2.4, 5.83], [0.4, 0.95, 1.2, 0.95], rows, name="体制")
    band(s, [("属人化しない：", dict(size=13, bold=True, color=RED)),
             ("担当が替わっても、同じ配信・同じ追客が続く仕組み", dict(size=13, bold=True, color=NAVY))], y=5.55, h=0.62)
    note(s, "※来場予約のリマインドは、予約の受け方に合わせて設定（連携の方法は構築時に決定）", y=6.4)


# ---------------------------------------------------------------- 6 今後の展開
def p_tenkai(prs, src):
    s = new_page(prs, src, "今後の展開：青森店で型を作り、全店・全エリアへ横展開する",
                 "配信の仕組みとコンテンツは、他の店舗・別のターゲットにそのまま転用できる")
    lw, rx, rw = 3.7, 4.85, 5.43
    heading(s, X0, 1.85, lw, "青森店で確立する型")
    for k, (h1, h2) in enumerate([("① 入口", "サイト・来場予約・展示場からLINEへ"),
                                  ("② 配信", "ルームツアー・暮らし・街並みの型"),
                                  ("③ 体制", "営業 × 事務 × 弊社の運用")]):
        add_box(s, X0, 2.3 + k * 0.86, lw, 0.74, [runs((h1, dict(size=12, bold=True, color=NAVY))),
                                                  runs((h2, dict(size=10.5, color=TEXT)))],
                line=NAVY, anchor="ctr", line_spacing=115, ins=(0.15, 0.04, 0.1, 0.04), name=f"型{k + 1}")
    add_box(s, X0 + lw + 0.05, 2.3, rx - X0 - lw - 0.1, 2.58, [runs(("▶", dict(size=18, color=NAVY)))], align="ctr",
            anchor="ctr", ins=(0, 0, 0, 0), name="矢印")
    heading(s, rx, 1.85, rw, "横展開")
    rows = [[label_cell("店舗", 11), cell("八戸・おいらせ・弘前", 11), cell("店舗ごとのQR・URLで、配信を分ける", 10.5)],
            [label_cell("商品", 11), cell("注文住宅・分譲・土地", 11), cell("希望に合わせた配信", 10.5)],
            [label_cell("層", 11), cell("検討中・来場済・OB施主", 11), cell("段階に合わせた配信（OB施主には紹介・点検の案内）", 10.5)]]
    add_table(s, rx, 2.3, [0.95, 1.95, rw - 2.9], [0.86] * 3, rows, name="横展開")
    band(s, [("集客投資の効果を最大化するために、", dict(size=13, bold=True, color=NAVY)),
             ("今、LINE公式アカウントの基盤を構築する", dict(size=13, bold=True, color=RED))], y=5.55, h=0.62)
    note(s, "※1つのアカウントで運用し、店舗・商品・層は友だち追加の経路とタグで分ける／既存のLINE公式アカウントを活用", y=6.4)


ORDER = [p_genjo, p_kadai, p_fan, p_taisei, p_tenkai]


def main(base, out):
    prs = Presentation(base)
    slides = list(prs.slides)
    src, cover, back = slides[HDR_SRC - 1], slides[COVER - 1], slides[BACK - 1]
    assert by_id(src, HDR[0]).text_frame.text == "サマリー"
    new = []
    for f in ORDER:
        f(prs, src)
        new.append(prs.slides[-1])
    p_cover(cover)
    keep = [cover] + new + [back]
    lst = prs.slides._sldIdLst
    by_part = {prs.part.related_part(e.rId): e for e in lst}
    for e in list(lst):
        lst.remove(e)
    keep_parts = {s.part for s in keep}
    for part, e in by_part.items():
        if part not in keep_parts:
            prs.part.drop_rel(e.rId)
    for s in keep:
        lst.append(by_part[s.part])
    prs.save(out)
    print("saved:", out, len(prs.slides), "枚")


if __name__ == "__main__":
    main(*sys.argv[1:3])
