# -*- coding: utf-8 -*-
"""あらうち資料 v9（CV一本化・フォーム）の自動チェック。

    python -m _build.arauchi_v9.checks_form <直す前のv9（506784f）.pptx> <直した後のv9.pptx> [<前回納品したv9（66a63c8）.pptx>]

変えたページ・消すべき言葉・新しい数字（SIMの合計CV・累計・S2のアクティブ数とブロック率）・棒グラフの寸法・1行に収まるはずの文・字体・図形ID・パッケージ（ほかのパーツが1バイトも違わない）を見る。
3つ目のファイルを渡すと、前回納品した版から変わったスライドが S9・S12・S13・S16 だけであることも見る。
（見た目のはみ出しは、LibreOfficeの画像と、行間を1.29倍にした確認用コピーの画像で別に見る）
"""
import difflib
import re
import sys
import zipfile

from pptx import Presentation
from pptx.util import Emu

from _build.arauchi_v4.build import c14n, slide_order
from _build.arauchi_v9.form import ACTIVE, BLOCK, CHANGED, CV, FRIENDS6, TOTAL

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
FAILS, OKS = [], []


def check(cond, msg):
    (OKS if cond else FAILS).append(msg)
    if not cond:
        print("NG:", msg)


def paras(slide):
    out = []
    for p in slide._element.iter("{%s}p" % A):
        t = "".join(x.text or "" for x in p.iter("{%s}t" % A))
        if t:
            out.append(t)
    return out


def est_width_in(text, pt):
    em = 0.0
    for ch in text:
        if ord(ch) < 0x80:
            em += 0.3 if ch == " " else (0.85 if ch == "%" else 0.58)
        else:
            em += 1.0
    return em * pt / 72


def walk(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from walk(sh.shapes)


def main(base, out, prev=None):
    pb, po = Presentation(base), Presentation(out)
    sb, so = list(pb.slides), list(po.slides)
    check(len(sb) == len(so) == 47, "スライド数 47 のまま")

    # ---- どのスライドが変わったか／ほかのパーツは1バイトも変わらない
    zb, zo = zipfile.ZipFile(base), zipfile.ZipFile(out)
    ob, oo = slide_order(zb), slide_order(zo)
    check(ob == oo, "スライドの並び（パーツ名）が同じ")
    changed = [i + 1 for i, n in enumerate(oo) if c14n(zb.read(ob[i])) != c14n(zo.read(n))]
    check(changed == CHANGED, f"変えたスライド＝{CHANGED} / 実際 {changed}")
    diff_parts = [n for n in zb.namelist() if zb.read(n) != zo.read(n)]
    check(sorted(diff_parts) == sorted(ob[i - 1] for i in CHANGED), f"変わったパーツは直したスライドのXMLだけ {diff_parts}")
    check(zb.namelist() == zo.namelist(), "zipの中のパーツの一覧・順序が同じ")

    # ---- 文字の違い（目で確認する用）
    print("\n--- 文字の違い（v9の番号）")
    for no in CHANGED:
        a, b = paras(sb[no - 1]), paras(so[no - 1])
        sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag != "equal":
                print(f"S{no} {tag}: {a[i1:i2]} → {b[j1:j2]}")
    print()

    texts = ["\n".join(paras(s)) for s in so]
    allt = "\n".join(texts)

    # ---- 消すべき言葉
    banned = ["33件", "126件", "問い合わせ①", "問い合わせ②", "20件/月", "約15件", "約18件", "①＋②", "13＋20", "月10件",
              "改善の余地", "一方で、あいさつに送る", "DYM手数料", "成果報酬", "スコープ", "専用フォーム", "松竹梅", "おすすめ",
              "計測できない", "月4本", "有効CV", "上積み", "既存の友だち経由は", "（試算には含みません）",
              # 前の版（新規33＋既存54＝87件）の言葉・数字
              "87件", "21件", "約8件", "約13件", "新規友だち経由", "既存友だち経由", "新規＋既存", "グラフ_新規",
              # S2の前の値
              "約807", "35%（仮）",
              # 15時版の数字
              "55件", "1,377"]
    for w in banned:
        hits = [i + 1 for i, t in enumerate(texts) if w in t]
        check(not hits, f"「{w}」がどのページにも無い {hits if hits else ''}")

    # 「運用開始から／運用一か月目」の断り書きは入れない（殿村さん指示）。S31（スケジュール）の元からある「運用開始」は対象外＝数字を直したページだけ見る
    for w in ("運用開始", "運用一か月目", "運用1か月目"):
        hits = [n for n in (9, 12, 13, 16) if w in texts[n - 1]]
        check(not hits, f"「{w}」が数字を直したページ（S9・S12・S13・S16）に無い {hits if hits else ''}")

    # ---- 新しい数字（殿村さんのSIM・16時保存：CV 13,14,14,15,16,17／累計89・友だち約1,450・ブロック率30.5%・アクティブ数約862）
    n_last, n_first = CV[-1], CV[0]
    t = {i + 1: x for i, x in enumerate(texts)}
    check(CV == [13, 14, 14, 15, 16, 17] and TOTAL == 89 and ACTIVE == 862 and BLOCK == "30.5%" and FRIENDS6 == "1,450", "数字の正本（13,14,14,15,16,17／累計89／友だち約1,450／アクティブ数862／ブロック率30.5%）")
    check(f"約{ACTIVE}（仮）" in t[2] and f"{BLOCK}（仮）" in t[2] and "1,241" in t[2], "S2 のアクティブ数・ブロック率（SIMの現状値）")
    check(f"1か月目 月{n_first}件 → 6か月目 月{n_last}件（試算）" in t[12], "S12 の配信目標の文")
    check(f"6か月目にLINE経由の問い合わせ月{n_last}件を見込みます（試算）。" in t[13], "S13 の見出し")
    check(f"{n_last}件/月" in t[13] and "6か月累計" in t[13] and f"{TOTAL}件" in t[13] and f"約{FRIENDS6}人" in t[13], "S13 の問い合わせ・累計・友だち数")
    check("問い合わせ①" not in t[13] and "合計" not in t[13] and "問い合わせ" in t[13], "S13 に①②・合計の箱が無い")
    check("¥215,000" in t[13] and "¥135,000" in t[13] and "21.5万円" in t[13] and "13.5万円" in t[13], "S13 の費用はそのまま")
    box13 = [p for p in paras(so[12]) if p.startswith("＜試算の前提＞") or p.startswith("プラン①｜総合運用の基本施策") or p.startswith("API連携") or p.startswith("成果を保証")]
    check(len(box13) == 4 and "紹介による追加成果は含みません" in box13[2] and not any("既存" in p for p in box13), f"S13 の試算の前提は元の4行のまま（紹介を含まない） {box13}")
    check("問い合わせ＝フォームでの物件情報の送信（CV）の件数" in t[13], "S13 の注記")
    check(f"約{n_last}件" in t[9] and "新規・既存の合計" in t[9] and t[9].count("（フォーム送信）") == 1, "S9 右の列＝合計1つ（約12件・新規・既存の合計）")
    check("今の友だちは見込み顧客として活用し、将来的な紹介にもつなげます（紹介による成果は試算に含みません）。" in t[9], "S9 見出し2行目（元のまま）")
    check("※問い合わせ＝フォームでの物件情報の送信（CV）の件数。新しい友だち・今の友だちの合計の試算" in t[9], "S9 の注記")
    check(f"6か月目に月{n_last}件" in t[16], "S16 の注記の件数")
    check("フォームで計測をします" in t[4] and "厚くし" in t[4], "S4 フォームで計測をする・厚くする")
    check("標準機能のまま厚くできると想定" in t[2] and "標準機能で厚くできる余地あり" in t[2], "S2 厚くできる")
    check("フォームで物件を送信（CV）" in t[3], "S3 フォームで物件を送信（CV）")
    check("フォーム送信（CV）につなげます" in t[8] and "フォームで送信" in t[8], "S8 フォーム送信（CV）")
    check("「フォームで物件を送る」まで案内します" in t[11], "S11 フォームで物件を送る")
    check("フォームでの物件URLの送信を後押しします" in t[17], "S17 リード文")
    for n in (2, 3, 4, 8, 9, 11, 12, 13, 17):
        check("フォーム" in t[n] or n == 2, f"S{n} に「フォーム」がある")
    # フォームの費用は書かない・初期設計項目に入れない
    s8_items = [p for p in paras(so[7]) if p.startswith("・")]
    check(not any("フォーム" in p for p in s8_items), f"S8 の初期設計項目にフォームを入れていない {s8_items}")
    check(not re.search(r"フォーム.{0,12}(費用|料金|円|無料|込み)", allt), "フォームの費用・料金に触れていない")

    # ---- S12 の棒グラフ（1系列）
    sh12 = {sh.name: sh for sh in walk(so[11].shapes)}
    grp = [sh for sh in so[11].shapes if sh.shape_type == 6]
    check(len(grp) == 1 and grp[0].name == "グラフ（問い合わせ件数・月間の試算）", "S12 のグラフが1つのグループ")
    scale = Emu(914400 * 0.55).emu / max(CV)
    base_line = sh12["グラフ_基準線"]
    for i, v in enumerate(CV, 1):
        lab = sh12[f"グラフ_件数_{i}か月目"].text_frame.text
        check(lab in (str(v), f"{v}件"), f"S12 {i}か月目のラベル {lab} == {v}")
        bar = sh12[f"グラフ_問い合わせ_{i}か月目"]
        check(abs(bar.height - v * scale) <= 2, f"S12 {i}か月目の棒の高さ＝{v}件分")
        check(abs((bar.top + bar.height) - base_line.top) <= 2, f"S12 {i}か月目の棒の底が基準線に載っている")
    check(f"{TOTAL}件（試算）" in sh12["グラフ_累計"].text_frame.text and "6か月累計" in sh12["グラフ_累計"].text_frame.text, "S12 の累計")
    check(sh12["グラフ_凡例文字"].text_frame.text.replace("\n", "").replace("\x0b", "") in ("問い合わせ（フォーム送信）",), "S12 の凡例")
    check(not [n for n in sh12 if "①" in n or "②" in n or "内訳" in n or "新規" in n or "既存" in n], "S12 のグラフに①②・内訳・新規・既存の図形が無い")
    check(sum(CV) == TOTAL == 89, "累計 89 ＝ 13+14+14+15+16+17")

    # ---- 1行に収まるはずの文：元の行（殿村さんのPowerPointで1行に収まっていた）より長くないこと
    def find(paras_, head):
        return [x for x in paras_ if x.startswith(head)][0]

    def not_longer(new_head, old_head, no, pt, label_):
        n_, o_ = find(paras(so[no - 1]), new_head), find(paras(sb[no - 1]), old_head)
        check(est_width_in(n_, pt) <= est_width_in(o_, pt) + 0.05,
              f"{label_} が元の行より長くない（新 {est_width_in(n_, pt):.2f}in／元 {est_width_in(o_, pt):.2f}in）")

    not_longer("ポータルで物件検索", "ポータルサイトで物件検索", 3, 14, "S3 1行目")
    p4 = find(paras(so[3]), "送るメリット・安心材料")
    check(est_width_in(p4, 14) * 1.10 <= 9.92 - 0.2 - 0.4, f"S4 2行目が1行に収まる（{est_width_in(p4, 14):.2f}in／箱9.92in）")
    not_longer("離脱防止ポップアップとサイト", "離脱防止ポップアップとサイト", 8, 14, "S8 1行目")
    not_longer("今の友だちは見込み顧客", "今の友だちは見込み顧客", 9, 13, "S9 見出し2行目")
    not_longer("※問い合わせ＝", "※問い合わせ①", 9, 11, "S9 見出し3行目")
    not_longer("1か月目 月", "1か月目 月", 12, 18, "S12 配信目標の文")
    check(est_width_in("新規・既存の合計", 11) <= 1.95 - 0.1 - 0.2 and est_width_in("（フォーム送信）", 11) <= 1.95 - 0.1 - 0.2, "S9 右の列の補足が1行に収まる")
    note13 = find(paras(so[12]), "※6か月後は6か月目の月間の試算。")
    check(est_width_in(note13, 10) * 1.10 <= 10.3 - 0.2, f"S13 注記が1行に収まる（{est_width_in(note13, 10):.2f}in／箱10.3in）")

    # ---- 字体・サイズ
    from _build.fontcheck import scan
    bad = scan(out)
    check(not bad or sum(sum(c.values()) for c in bad.values()) == 0, "メイリオ以外・指定なしの文字が0か所")
    small = []
    for no in CHANGED:
        for rpr in so[no - 1]._element.iter("{%s}rPr" % A):
            sz = rpr.get("sz")
            if sz and int(sz) < 1000:
                small.append((no, sz))
    check(not [x for x in small if x[0] != 13], f"変えたページに10pt未満の文字が無い（S13の注記10ptは元のまま） {small[:3]}")

    # ---- 図形ID・パッケージ
    ids_ok = True
    for s in so:
        ids = [e.get("id") for e in s._element.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")]
        ids_ok &= len(ids) == len(set(ids))
    check(ids_ok, "各スライドの図形IDが重複していない")
    ct = zo.read("[Content_Types].xml").decode("utf-8")
    names = set(zo.namelist())
    check(set(re.findall(r'PartName="/([^"]+)"', ct)) <= names, "[Content_Types] の Override が全部実在する")

    # ---- 前回納品した版（66a63c8）から変わったのは S9・S12・S13・S16 だけ。ほかのパーツは1バイトも変わらない
    if prev:
        zp = zipfile.ZipFile(prev)
        op = slide_order(zp)
        check(op == oo, "前回版とスライドの並び（パーツ名）が同じ")
        ch = [i + 1 for i, n in enumerate(oo) if c14n(zp.read(op[i])) != c14n(zo.read(n))]
        check(ch == [9, 12, 13, 16], f"前回版から変わったスライド＝[9, 12, 13, 16] / 実際 {ch}")
        dp = [n for n in zp.namelist() if zp.read(n) != zo.read(n)]
        check(sorted(dp) == sorted(oo[i - 1] for i in (9, 12, 13, 16)), f"前回版から変わったパーツはその4枚のXMLだけ {dp}")

    print(f"\nOK {len(OKS)} / NG {len(FAILS)}")
    return not FAILS


if __name__ == "__main__":
    sys.exit(0 if main(*sys.argv[1:4]) else 1)
