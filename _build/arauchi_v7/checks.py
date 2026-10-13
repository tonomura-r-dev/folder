# -*- coding: utf-8 -*-
"""v7 の検品：v6から変わったページ・並び・ページ参照・料金/数値が動いていないこと・指示の中身。

    python -m _build.arauchi_v7.checks <v6.pptx> <v7.pptx>
"""
import re
import sys
import zipfile

from lxml import etree
from pptx import Presentation

from _build.arauchi_v4.build import c14n, slide_order
from _build.arauchi_v7.build import ORDER

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def page_text(slide):
    return "\n".join("".join((t.text or "") for t in p.iter(A + "t")) for p in slide._element.iter(A + "p"))


def shape_fill(sp):
    sppr = sp.find(P + "spPr")
    if sppr is None:
        return None
    f = sppr.find(A + "solidFill")
    if f is None:
        return "none" if sppr.find(A + "noFill") is not None else None
    c = f[0]
    return (c.get("val") or "").upper()


def shapes_with(slide, text):
    return [sp for sp in slide._element.iter(P + "sp") if text in "".join((t.text or "") for t in sp.iter(A + "t"))]


def main(v6, v7):
    ok = True

    def check(cond, msg):
        nonlocal ok
        print(("OK  " if cond else "NG  ") + msg)
        ok = ok and cond

    z6, z7 = zipfile.ZipFile(v6), zipfile.ZipFile(v7)
    o6, o7 = slide_order(z6), slide_order(z7)
    check(len(o6) == 48 and len(o7) == 48, f"枚数 v6={len(o6)} → v7={len(o7)}（48枚のまま）")
    check(o7 == [o6[k] for k in ORDER], "並び：S1〜S13 → 費用4枚（v6のS23〜S26）→ v6のS14〜S22 → v6のS27〜S48")

    # ---- v6から中身が変わったページ（パーツ単位）
    diff = [i + 1 for i, n in enumerate(o6) if c14n(z6.read(n)) != c14n(z7.read(n))]
    check(diff == [1, 3, 8, 15, 24, 25], f"v6と中身が違うページ（v6の番号）＝{diff}（期待＝[1,3,8,15]＋ページ参照を直した[24,25]。ほかの{48 - len(diff)}枚は同じ）")
    # rels（画像・リンク）が変わっていない
    same_rels = all(z6.read(n.replace('slides/', 'slides/_rels/') + '.rels') == z7.read(n.replace('slides/', 'slides/_rels/') + '.rels') for n in o6)
    check(same_rels, "スライドのrels（画像・レイアウト・ノート）は全ページ同じ")

    p6, p7 = Presentation(v6), Presentation(v7)
    t6 = {i + 1: page_text(s) for i, s in enumerate(p6.slides)}
    t7 = {i + 1: page_text(s) for i, s in enumerate(p7.slides)}

    # ---- 費用4枚の位置
    exp = {13: "想定の費用対効果", 14: "費用感", 15: "料金プランの比較", 16: "運用内容の比較", 17: "プラン④｜成果報酬型", 18: "実施施策",
           19: "LINEによる問い合わせ対応の効率化", 20: "施策詳細｜新規友だちの獲得"}
    for n, k in exp.items():
        check(k in t7[n], f"S{n}＝『{k}』")
    # 費用4枚の中身：ページ参照（26→17）以外は v6 と同じ
    for new_n, old_n in ((14, 23), (15, 24), (16, 25), (17, 26)):
        a = t7[new_n].replace("17ページ", "26ページ")
        check(a == t6[old_n], f"S{new_n}（v6のS{old_n}）＝ページ参照以外は同じ文言（料金・条件・数値は不変）")
    # 指示対象外のページ：文言が v6 と同じ（v7のS18〜S26＝v6のS14〜S22。S15＝v6のS15は直したページなので外す）
    moved_same = all(t7[n] == t6[o] for n, o in [(i, i) for i in range(1, 14) if i not in (1, 3, 8)] + [(i + 4, i) for i in range(14, 23) if i != 15] + [(i, i) for i in range(27, 49)])
    check(moved_same, "指示対象外のページ（S1・S3・S8・S15・費用4枚以外の41枚）の文言は v6 と同じ")

    # ---- ページ参照
    alltext = "\n".join(t7[i] for i in range(1, 49))
    check("26ページ" not in alltext, "『26ページ』が残っていない")
    refs = {}
    for n in range(1, 49):
        for m in re.finditer(r"(\d+)ページ", t7[n]):
            refs.setdefault((n, int(m.group(1))), 0)
            refs[(n, int(m.group(1)))] += 1
    want = {17: "プラン④｜成果報酬型", 13: "想定の費用対効果", 6: "市場分析"}
    for (n, tgt), c in sorted(refs.items()):
        check(tgt in want and want[tgt] in t7[tgt], f"S{n}の『{tgt}ページ』×{c} → S{tgt}は『{want.get(tgt)}』")

    # ---- S1 表紙（元のDYM資料の表紙）
    s1 = p7.slides[0]
    t = t7[1]
    check("LINE公式アカウント運用のご提案" in t, "S1：タイトル『LINE公式アカウント運用のご提案』")
    check("LINEへの流入と友だち追加後の導線を整備し、\n物件問い合わせの増加を実現する" in t, "S1：サブタイトルの文言（変更なし）")
    check("株式会社あらうち 御中" in t, "S1：宛名『株式会社あらうち 御中』")
    sh = {x.shape_id: x for x in s1.shapes}
    check(sh[26].top == 2533105 and sh[28].top == 3199966, "S1：タイトル・線の位置＝元のDYM資料の表紙と同じ（y=2.770in／3.500in）")
    sub = sh[27]._element
    check(shape_fill(sub) == "none" or shape_fill(sub) is None, "S1：サブタイトルの帯（塗り）が無い")
    xml_sub = etree.tostring(sub).decode().lower()
    body = xml_sub[xml_sub.find("<p:txbody"):]
    check('sz="1600"' in body and 'b="1"' not in body, "S1：サブタイトル＝16pt・太字なし（元のDYM表紙と同じ強弱）")
    check("c00000" not in xml_sub, "S1：サブタイトルに赤字が無い")

    # ---- S3
    s3 = p7.slides[2]
    check("物件情報の送信（CV）" in t7[3] and "あらうち様による個別対応（参考" in t7[3], "S3：『物件情報の送信（CV）』『あらうち様による個別対応（参考）』")
    check("当社の計測・評価はここまで" in t7[3], "S3：ゴール＝CV。当社の計測・評価はここまで")
    red = [sp for sp in s3._element.iter(P + "sp") if shape_fill(sp) in ("C00000", "FF0000")]
    red_txt = ["".join((x.text or "") for x in sp.iter(A + "t")) for sp in red]
    check(not any("ご契約" in x or "お申込み" in x for x in red_txt), f"S3：契約・申込の箱に赤の強調が無い（赤の箱＝{red_txt}）")
    check(any("④" in x and "CV" in x for x in red_txt) and any("⑤" in x and "CV" in x for x in red_txt), "S3：赤＝⑤リード有効化（CV）と④物件情報の送信（CV）")
    # ---- S8
    s8 = p7.slides[7]
    check("物件情報の送信\n（CV）" in t7[8] and "運用フローの終点" in t7[8], "S8：運用フローの終点＝物件情報の送信（CV）")
    check("参考：CV以降の個別対応（あらうち様）／当社の計測対象外" in t7[8] and "あらうち様による\n個別対応（参考）" in t7[8], "S8：CV以降は参考表示（当社の計測対象外）")
    navy_after = [sp for sp in s8._element.iter(P + "sp") if any(k in "".join((x.text or "") for x in sp.iter(A + "t")) for k in ("内見・お申込みの案内", "別の物件URLの受付", "ご成約", "入居までの案内", "チャットタグで対応状況を管理")) and shape_fill(sp) is not None]
    check(all(shape_fill(sp) == "7F7F7F" for sp in navy_after) and len(navy_after) == 5, "S8：CV以降の箱は灰色（紺の箱が無い）")
    # ---- S15（v7ではS19）
    check("物件確認・内見等の個別対応" in t7[19] and "ご契約" not in t7[19], "S15（S19）：あらうち様＝『物件確認・内見等の個別対応』。『ご契約』と書いていない")
    for k in ("LINE公式アカウントの初期設計", "自動応答の設定", "チャットタグの設計・設定", "配信運用・改善支援", "当社（DYM）｜仕組みづくり・改善支援"):
        check(k in t7[19], f"S15（S19）：DYMの役割『{k}』を維持")
    check("※物件情報の送信（CV）以降は、当社の計測対象外です" in t7[19], "S15（S19）：CV以降は計測対象外の注記")

    # ---- 禁止語・文字サイズ
    mine = t7[1] + t7[3] + t7[8] + t7[19]
    for bad in ("松竹梅", "予算", "ヒアリング", "CPA", "スコープ", "∞"):
        check(bad not in mine, f"変更ページに『{bad}』が無い")
    mn = 99
    for n in (3, 8, 19):
        for rpr in list(p7.slides[n - 1]._element.iter(A + "rPr")):
            sz = rpr.get("sz")
            if sz and any((t.text or "").strip() for t in rpr.getparent().iter(A + "t")):
                mn = min(mn, int(sz) / 100)
    check(mn >= 11, f"S3・S8・S15の文字の最小サイズ＝{mn}pt（11pt以上）")

    print("\n総合:", "全部OK" if ok else "NGあり")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:3]))
