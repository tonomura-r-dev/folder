# -*- coding: utf-8 -*-
"""S15：「LINEによる問い合わせ対応の効率化」に作り直す（2026-10-12 殿村さん指示）。

上段：①自動応答 → ②個別チャット → ③チャットタグ の3枚（LINEの機能で効率化できることが流れで分かるように矢印でつなぐ）。
下段：役割分担（当社＝仕組みづくり・改善支援／あらうち様＝お客様への個別対応）。
当社が個別チャットの返信を請け負うと誤解されないよう、要旨と下段の両方で対応主体を明記する。
元のS15の情報（自動回答・応答の用意・時間外の自動返信・100円チェッカー・空室確認と内見の手配・チャットタグの管理・
確認漏れ防止・配信の重複防止・物件の確認に集中）はすべて上段のどこかに残す。"""
from _build.arauchi_v3.common import (Inches, clone_into, ids, move, qn, set_anchor, set_insets, set_line_spacing,
                                      set_paras, set_size_all, set_spacing, set_color)

NAVY = "1F2A5C"


def _title(shape, text):
    set_paras(shape, [[(text, None)]])


def apply(ctx):
    s = ctx.orig[15]
    b = ids(s)
    assert b[45].text_frame.text.startswith("その他")
    _title(b[45], "LINEによる問い合わせ対応の効率化")

    # 要旨（2行・14pt）
    set_paras(b[2], [
        [("LINEの機能で、問い合わせへの一次対応と対応状況の管理を効率化します。", None)],
        [("お客様への個別対応はあらうち様、仕組みづくりと改善支援は当社が担当します。", True)],
    ], size=14)
    move(b[2], x=0.47, y=0.98, w=9.90, h=0.62)
    set_anchor(b[2], "ctr")

    # 上段：3枚のカード（幅3.0インチ・間0.35インチ、間に矢印）
    xs = (0.47, 3.82, 7.17)
    CW = 3.05
    heads = ((5, "① 自動応答"), (7, "② 個別チャット"), (11, "③ チャットタグ"))
    for (sid, text), x in zip(heads, xs):
        set_paras(b[sid], [[(text, None)]], size=16)
        move(b[sid], x=x, y=1.62, w=CW, h=0.45)
    bodies = (
        (6, [
            [("設定：当社　対応：自動", True)],
            "・よくある質問（初期費用・お申込みの流れなど）にキーワード応答で自動回答",
            "・リッチメニューのボタンと同じ言葉で応答を用意",
            "・受付時間外は、時間を指定して受付の旨を自動で返信",
            "・リッチメニューから100円チェッカーへ案内",
        ]),
        (9, [
            [("対応：あらうち様", True)],
            "・ご担当者様が、物件に関する相談・問い合わせに対応",
            "・送られた物件（URL・物件名・画像）の空室確認・内見の手配",
            "・一次対応は自動応答に任せ、物件の確認に集中",
        ]),
        (13, [
            [("設計・設定：当社", True)],
            "・問い合わせ状況・対応ステータス（受付・確認中・返信済）を管理",
            "・タグの付け替えで、確認漏れ・返信漏れを防止",
            "・問い合わせ済みの方を除いた再案内で、重複配信を防止",
        ]),
    )
    head_colors = {6: NAVY, 9: "4A6FA5", 13: "5B6478"}
    for (sid, paras), x in zip(bodies, xs):
        set_paras(b[sid], paras, size=14)
        move(b[sid], x=x, y=2.12, w=CW, h=3.02)
        set_insets(b[sid], l=0.10, t=0.10, r=0.08, b=0.06)
        set_line_spacing(b[sid], 100)
        set_spacing(b[sid], before=3, after=0)
        set_color(b[sid], 0, head_colors[sid])
        for i in range(1, len(paras)):
            set_color(b[sid], i, "333333")
    # カードの間の矢印（元のS17の右向き矢印を複製）
    arrow_src = ids(ctx.orig[17])[18]
    for x in (3.52 + 0.035, 6.87 + 0.035):
        a = clone_into(s, arrow_src)
        move(a, x=x, y=3.52, w=0.23, h=0.22)

    # 下段：役割分担
    # 元の見出し「その他（コンサルに含む）」の情報（＝当社の作業はコンサル費用に含む）を残す
    set_paras(b[15], [[("役割分担", None), ("　※当社の対応は運用コンサル費用に含みます", None)]])
    r2 = b[15].text_frame._txBody.findall(qn("a:p"))[0].findall(qn("a:r"))[1].find(qn("a:rPr"))
    r2.set("sz", "1200")
    r2.set("b", "0")
    move(b[15], x=0.47, y=5.17, w=9.75, h=0.37)
    set_paras(b[17], [[("当社（DYM）｜仕組みづくり・改善支援", None)]], size=14)
    set_paras(b[21], [[("あらうち様｜お客様への個別対応", None)]], size=14)
    RW = 4.725
    move(b[17], x=0.47, y=5.53, w=RW, h=0.40)
    move(b[21], x=10.22 - RW, y=5.53, w=RW, h=0.40)
    # あらうち様の見出しは、②個別チャット（あらうち様が対応）と同じ青にして、対応関係を色で示す
    sppr = b[21]._element.spPr
    for fill in [sppr.find(qn("a:solidFill"))] + list(sppr.find(qn("a:ln")).findall(qn("a:solidFill"))):
        fill.find(qn("a:srgbClr")).set("val", "4A6FA5")     # 塗りと枠線だけ（影の色は変えない）
    set_paras(b[19], [
        "・LINE公式アカウントの初期設計",
        "・自動応答の設定",
        "・チャットタグの設計・設定",
        "・配信運用・改善支援",
    ], size=14)
    set_paras(b[23], [
        "・個別チャットでの返信",
        "・物件の空室確認",
        "・内見の手配",
        "・お申込み・ご契約の対応",
    ], size=14)
    for sid, x in ((19, 0.47), (23, 10.22 - RW)):
        move(b[sid], x=x, y=5.97, w=RW, h=1.10)
        set_insets(b[sid], l=0.12, t=0.04, r=0.08, b=0.04)
        set_line_spacing(b[sid], 100)
        set_spacing(b[sid], before=0, after=0)
        set_anchor(b[sid], "ctr")
        for i in range(4):
            set_color(b[sid], i, "333333")
