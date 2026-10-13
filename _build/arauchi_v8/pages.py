# -*- coding: utf-8 -*-
"""あらうち提案資料 v8（2026-10-13 殿村さん指示）の直し。元＝殿村さんがPowerPointで保存したv7（48枚）。

  S9  想定セグメント×成果 … 新規友だち／既存友だちそれぞれの施策・導線・6か月目の成果（試算）を明確に。転換率の表示は外す
  S12 要件定義           … 配信目標を新SIM（1か月目10件→6か月目33件）に。棒グラフ（問い合わせ①＋②の積み上げ）を足す
  S13 想定の費用対効果   … 問い合わせ②20件／合計33件／6か月累計126件。試算の前提に「紹介による追加成果は含まない」を足す
  S15 料金プランの比較   … 4プラン → 3プラン（成果報酬型を削除）
  S16 運用内容の比較     … 4プラン → 3プラン（プラン①〜③の費用・実施内容・配信頻度を S15 と一致）
  S17 プラン④の詳細      … 削除（build.py で zip から外す）
  そのほか、残っていた矛盾する表現を直す：
    S8・S18 … 企画配信の頻度を「月4回程度」にそろえる／S18 … ステップ・企画配信のねらい「お申込みの後押し」→「物件の送信の後押し」
    S21 … CPF広告の説明「成果報酬で」→「友だち1人あたりの課金で」（削除したプラン④と紛らわしいため）

数字の正本（殿村さんの改訂SIM）：問い合わせ① 5,7,8,10,12,13／② 5,7,10,13,16,20／合計 10,14,18,23,28,33（6か月累計126件）。
CV①②の定義は元のSIMのまま（①＝友だち追加直後の物件の問い合わせ／②＝ステップ配信・企画配信などからの問い合わせ）。
新規15件・既存18件は、6か月目の②をステップ配信と企画配信のクリック数の比（10.4%：89.6%）で振り分けて再現できた
（新規＝①13＋②約2／既存＝企画配信経由の②約18）ので「試算上の内訳」として載せる。
"""
import copy

from lxml import etree

from _build.arauchi_v4.lib import NAVY, TEXT, add_box, add_table, delete, emu, make_para, move, qn, set_paras
from _build.arauchi_v6.pages import GRID, NAVY_TINT, body_cell, hdr_cell, merged_away, plan_head
from _build.arauchi_v7.pages import set_body

GREY_TXT = "444444"
CV1 = [5, 7, 8, 10, 12, 13]                    # 問い合わせ①（改訂SIM）
CV2 = [5, 7, 10, 13, 16, 20]                   # 問い合わせ②（改訂SIM）
TOTAL = [a + b for a, b in zip(CV1, CV2)]
assert TOTAL == [10, 14, 18, 23, 28, 33] and sum(TOTAL) == 126

C1_BAR = "1F497D"        # 問い合わせ①（S12の見出しの紺）
C2_BAR = "7FA7D9"        # 問い合わせ②（同じ青の明るい側）
RED_TXT = "FF3300"       # S12の目標の赤


# ---------------------------------------------------------------- 小さな部品
def shape_by_id(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid:
            return sh
    raise KeyError(sid)


def replace_in_paragraphs(slide, old, new):
    """段落の中の文字を差し替える（runをまたいでもよい）。差し替えた数を返す"""
    n = 0
    for p in slide._element.iter(qn("a:p")):
        runs = p.findall(qn("a:r"))
        if not runs:
            continue

        def texts():
            return [(r.find(qn("a:t")).text or "") for r in runs]

        start = "".join(texts()).find(old)
        while start != -1:
            end, pos, first = start + len(old), 0, True
            for r, t in zip(runs, texts()):
                r0, r1 = pos, pos + len(t)
                pos = r1
                if r1 <= start or r0 >= end:
                    continue
                a, b = max(start, r0) - r0, min(end, r1) - r0
                tt = r.find(qn("a:t"))
                if first:
                    tt.text = t[:a] + new + t[b:]
                    first = False
                else:
                    tt.text = t[:a] + t[b:]
            n += 1
            start = "".join(texts()).find(old, start + len(new))
    return n


def set_br_text(shape, text):
    """1つの段落のまま、"\n" の位置で改行する（最初のrunの書式を使う）"""
    lines = text.split("\n")
    set_paras(shape, ["".join(lines)])
    p = shape.text_frame._txBody.findall(qn("a:p"))[0]
    r = p.find(qn("a:r"))
    rpr = r.find(qn("a:rPr"))
    r.find(qn("a:t")).text = lines[0]
    prev = r
    for ln in lines[1:]:
        br = etree.Element(qn("a:br"))
        br.append(copy.deepcopy(rpr))
        prev.addnext(br)
        r2 = copy.deepcopy(r)
        r2.find(qn("a:t")).text = ln
        br.addnext(r2)
        prev = r2


def set_run_color(shape, hex_):
    for rpr in list(shape.text_frame._txBody.iter(qn("a:rPr"))) + list(shape.text_frame._txBody.iter(qn("a:endParaRPr"))):
        for f in rpr.findall(qn("a:solidFill")):
            rpr.remove(f)
        fill = etree.Element(qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr"), val=hex_)
        ln = rpr.find(qn("a:ln"))
        (ln.addnext(fill) if ln is not None else rpr.insert(0, fill))


def label(s, x, y, w, h, text, size, bold=False, color=TEXT, align="ctr", anchor="ctr", name=None):
    """枠・塗りなしの文字だけの箱（グラフの目盛りや凡例に使う）"""
    paras = [[(t, {})] for t in text] if isinstance(text, list) else [[(text, {})]]
    return add_box(s, x, y, w, h, paras, size=size, bold=bold, color=color, fill=None, line=None, align=align, anchor=anchor,
                   ins=(0, 0, 0, 0), name=name)


# ---------------------------------------------------------------- S8・S18・S21 の表記
def p08(s):
    n = replace_in_paragraphs(s, "月4本", "月4回程度")
    assert n == 2, n


def p18(s):
    head = shape_by_id(s, 386)
    assert "月4本" in head.text_frame.text
    set_br_text(head, "友だち追加から14日間のステップ配信（全10通）と月4回程度の企画配信で、\n物件URLの送信を後押しします")
    m = replace_in_paragraphs(s, "お申込みの後押し", "物件の送信の後押し")
    assert m == 2, m


def p21(s):
    sh = shape_by_id(s, 505)
    assert "成果報酬" in sh.text_frame.text
    set_paras(sh, ["CPF約100-200円台。", "かなり安価で、友だち1人あたりの", "課金で友だちを増やす広告。"])


# ---------------------------------------------------------------- S9 想定セグメント×成果
def p09(s):
    set_paras(shape_by_id(s, 236), ["対象"])                      # 「流入元」→「対象」（下の4行は状況別の対象）
    new_box = shape_by_id(s, 242)                                # 新規友だち（サイトから）＝離脱防止ポップアップとサイトのボタンの2行をまとめる
    delete(shape_by_id(s, 243))
    move(new_box, y=2.19, h=1.28)
    set_paras(new_box, [[("新規友だち", True)], [("（100円賃貸の", False)], [("サイトから）", False)]])
    set_paras(shape_by_id(s, 248), [[("既存友だち（今の友だち 1,241人）", True)], [("＋追加後の新しい友だちにも配信", False)]])
    n = replace_in_paragraphs(s, "月4本", "月4回程度")
    assert n == 1, n
    set_paras(shape_by_id(s, 282), ["お申込み・ご契約の案内", "（あらうち様が対応）"])   # CV以降は当社の対象外
    set_paras(shape_by_id(s, 279), ["リッチメニュー", "→ 物件の送信（CV）"])          # 既存友だちの導線の終点もCV（物件情報の送信）に

    # 右の列：転換率 → 6か月目の問い合わせ件数（試算）
    set_paras(shape_by_id(s, 271), ["6か月目の問い合わせ", "※試算の想定（月間）"])
    set_paras(shape_by_id(s, 272), ["約15件", "新規友だち経由", "（①13件＋②約2件）"])
    set_paras(shape_by_id(s, 273), ["約18件", "既存友だち経由", "（企画配信経由の②）"])
    d = shape_by_id(s, 274)                                       # 問い合わせた方の行（CV以降）
    set_paras(d, ["計測対象外"], size=12)
    set_run_color(d, "7F7F7F")
    move(d, y=4.34, h=0.61)
    set_body(d, anchor="ctr", autofit="noAutofit")
    # 見出しの文：既存友だち＝見込み顧客として活用し、将来的な紹介にもつなげる
    set_paras(shape_by_id(s, 293), [
        [("新しい友だちには物件の送り方を、今の友だちには配信とリッチメニューで問い合わせのきっかけを届けます。", None)],
        [("今の友だちは見込み顧客として活用し、将来的な紹介にもつなげます（紹介による成果は試算に含みません）。", None)],
        [("※問い合わせ①＝友だち追加直後の物件の送信、②＝配信経由。新規・既存の内訳は②をクリック数の比で振り分けた試算", None)],
    ])


# ---------------------------------------------------------------- S12 要件定義
def p12(s):
    # 上段：プロモーション目的（左）と配信目標（右）を横に並べ、右にグラフの場所を作る
    move(shape_by_id(s, 321), x=0.55, y=0.85, w=3.05, h=0.40)
    body = shape_by_id(s, 322)
    move(body, x=0.55, y=1.25, w=3.05, h=1.72)
    set_br_text(body, "LINE経由の\n物件問い合わせを\n増やす")
    move(shape_by_id(s, 324), x=3.70, y=0.85, w=6.58, h=0.40)
    goal = shape_by_id(s, 325)
    move(goal, x=3.70, y=1.25, w=6.58, h=1.72)
    set_paras(goal, ["1か月目 月10件 → 6か月目 月33件（試算）"], size=18)
    set_body(goal, anchor="t", ins=(0.1, 0.03, 0.1, 0.04))

    # 棒グラフ（問い合わせ①＋②の積み上げ。合計を棒の上に）
    n = replace_in_paragraphs(s, "月4本", "月4回程度")
    assert n == 1, n
    x0, pitch, bar_w, base_y, scale = 4.18, 0.70, 0.44, 2.52, 0.55 / 33
    before = {sh.shape_id for sh in s.shapes}
    add_box(s, x0, base_y, pitch * 6, 0.012, [[("", {})]], fill="A6A6A6", line=None, name="グラフ_基準線")
    for i, (c1, c2, t) in enumerate(zip(CV1, CV2, TOTAL)):
        cx = x0 + pitch * i + pitch / 2
        h1, h2 = c1 * scale, c2 * scale
        add_box(s, cx - bar_w / 2, base_y - h1, bar_w, h1, [[("", {})]], fill=C1_BAR, line=None, name=f"グラフ_問い合わせ①_{i + 1}か月目")
        add_box(s, cx - bar_w / 2, base_y - h1 - h2, bar_w, h2, [[("", {})]], fill=C2_BAR, line=None, name=f"グラフ_問い合わせ②_{i + 1}か月目")
        edge = i in (0, 5)
        label(s, cx - pitch / 2, base_y - h1 - h2 - 0.32, pitch, 0.30, f"{t}件" if edge else str(t), 14 if edge else 12,
              bold=True, color=RED_TXT if edge else TEXT, name=f"グラフ_合計_{i + 1}か月目")
        label(s, cx - pitch / 2, base_y + 0.03, pitch, 0.21, f"{i + 1}か月目", 10, color="595959", name=f"グラフ_月_{i + 1}")
        label(s, cx - pitch / 2, base_y + 0.23, pitch, 0.21, f"{c1}＋{c2}", 10, color="595959", name=f"グラフ_内訳_{i + 1}か月目")
    tag = label(s, 3.76, base_y + 0.23, 0.42, 0.21, "①＋②", 10, color="595959", align="l", name="グラフ_内訳の見出し")
    tag.text_frame._txBody.find(qn("a:bodyPr")).set("wrap", "none")        # 狭い箱で折り返さない
    # 凡例と累計
    for y, color, lines, h, nm in ((1.70, C1_BAR, ["①友だち追加直後"], 0.26, "①"),
                                   (1.96, C2_BAR, ["②ステップ配信・", "企画配信など"], 0.48, "②")):
        add_box(s, 8.58, y + 0.06, 0.14, 0.14, [[("", {})]], fill=color, line=None, name=f"グラフ_凡例_{nm}")
        label(s, 8.78, y, 1.45, h, lines, 11, color=TEXT, align="l", anchor="t" if len(lines) > 1 else "ctr", name=f"グラフ_凡例文字_{nm}")
    label(s, 8.58, 2.46, 1.65, 0.47, ["6か月累計", "126件（試算）"], 11, bold=True, color=TEXT, align="l", anchor="t", name="グラフ_累計")
    grp = s.shapes.add_group_shape(shapes=[sh for sh in s.shapes if sh.shape_id not in before])
    grp.name = "グラフ（問い合わせ件数・月間の試算）"


# ---------------------------------------------------------------- S13 想定の費用対効果
def p13(s):
    lead = shape_by_id(s, 381)
    set_paras(lead, [[("プラン①｜総合運用：初期21.5万円・月額13.5万円で運用します。", None)],
                     [("友だち追加の増加と導線改善で、6か月目にLINE経由の問い合わせ月33件を見込みます（試算）。", None)]])
    set_paras(shape_by_id(s, 377), ["20件/月"])
    set_paras(shape_by_id(s, 378), ["33件/月"])
    total = shape_by_id(s, 379)
    set_paras(total, ["（6か月累計126件）"])
    move(total, x=8.40, w=2.40)                                    # 90件→126件で字数が増えるので広げる（右端は元のまま）
    n = replace_in_paragraphs(s, "月4本", "月4回程度")
    assert n == 1, n
    for sid in (351, 352):                                         # 「リッチメニュー修正変更」「月次運用範囲内で適宜調整」＝2行で箱（0.53in）ぎりぎり
        cell = shape_by_id(s, sid)
        set_paras(cell, [t for t in [p.text for p in cell.text_frame.paragraphs]], size=12)
        set_body(cell, ins=(0.04, 0, 0.04, 0))
    box = shape_by_id(s, 382)
    move(box, y=4.78, h=1.75)
    set_paras(box, [[("＜試算の前提＞", None)],
                    [("プラン①｜総合運用の基本施策を対象とした想定値です。", None)],
                    [("API連携・追加開発、紹介による追加成果は含みません。", None)],
                    [("成果を保証するものではありません。", None)]])
    # 見出し以外は11ptにする（4段落で箱に収めるため）
    for i, p in enumerate(box.text_frame._txBody.findall(qn("a:p"))):
        if i:
            for rpr in list(p.iter(qn("a:rPr"))) + list(p.iter(qn("a:endParaRPr"))):
                rpr.set("sz", "1100")


# ---------------------------------------------------------------- S15・S16（3プラン）
COLS3 = [2.5, 2.5, 2.5, 2.5]
NOTE_COLOR = GREY_TXT
TOOL3 = "ツール利用料＋DYM手数料"


def _drop(s, names=(), table=True):
    for sh in list(s.shapes):
        if (table and getattr(sh, "has_table", False) and sh.has_table) or sh.name in names:
            delete(sh)


def p15(s):
    for sh in list(s.shapes):
        if sh.has_text_frame and "4つのプラン" in sh.text_frame.text:
            set_paras(sh, ["ご希望の運用範囲に合わせて、3つのプランをご用意しました。"])
    _drop(s, names=("コンサル型の説明", "成果報酬型の説明", "注記"))
    rows = [[hdr_cell("項目", NAVY, size=14),
             plan_head("プラン①", "総合運用", NAVY), plan_head("プラン②", "獲得重視", NAVY), plan_head("プラン③", "初期導入", NAVY)]]

    def row(label_, vals, size=15, bold=False):
        return [body_cell(label_, fill=NAVY_TINT, size=14, bold=True, color=NAVY)] + [body_cell(v, size=size, bold=bold) for v in vals]

    rows.append(row("初期費用", ["215,000円"] * 3))
    rows.append(row("月額コンサルティング費", ["100,000円", "50,000円", "なし"]))
    r3 = row("月額ツール・\nLINE関連費", ["35,000円", "35,000円"])
    r3.append(body_cell([[("導入後も継続して利用する", {})], [("ツールについて、", {})], [("ツール利用料＋DYM手数料", {"bold": True})], [("が発生します", {})]],
                        fill=NAVY_TINT, size=13, color=NAVY, rowSpan=2))
    rows.append(r3)
    r4 = [body_cell("月額費用", fill=NAVY, size=14, bold=True, color="FFFFFF"),
          body_cell("135,000円", fill=NAVY_TINT, size=20, bold=True, color=NAVY),
          body_cell("85,000円", fill=NAVY_TINT, size=20, bold=True, color=NAVY),
          merged_away(NAVY_TINT, vMerge=True)]
    for c in r4:
        c["borders"] = dict(c["borders"], t=(1.5, NAVY))
    rows.append(r4)
    heights = [0.85, 0.7, 0.7, 0.85, 1.0]
    add_table(s, 0.42, 1.48, COLS3, heights, rows, name="料金プラン比較表")
    y = 1.48 + sum(heights) + 0.2
    notes = [
        "※価格はあらうち様向けの個別提案価格です（DYMの標準価格ではありません）。初期費用215,000円は3プラン共通です。",
        "※プラン①の月額ツール・LINE関連費35,000円の内訳は、離脱防止ポップアップ30,000円、LINE公式アカウント5,000円です。",
        "※プラン③は、月額のコンサルティング費用はありません。導入後も継続して利用するツールについては、ツール利用料およびDYM手数料が発生します。",
    ]
    add_box(s, 0.42, y, 10.0, 7.1 - y, [[(t, {})] for t in notes], size=11, color=NOTE_COLOR, align="l", anchor="t",
            ins=(0.02, 0.02, 0.02, 0.02), after=3, name="注記")


def p16(s):
    for sh in list(s.shapes):
        if sh.has_text_frame and "コンサル型の運用内容" in sh.text_frame.text:
            set_paras(sh, ["3つのプランの運用内容を比較します。"])
    _drop(s, names=("注記",))
    rows = [[hdr_cell("項目", NAVY, size=14),
             plan_head("プラン①", "総合運用", NAVY), plan_head("プラン②", "獲得重視", NAVY), plan_head("プラン③", "初期導入", NAVY)]]
    rows.append([body_cell("月額費用", fill=NAVY, size=13, bold=True, color="FFFFFF"),
                 body_cell("135,000円", fill=NAVY_TINT, size=16, bold=True, color=NAVY),
                 body_cell("85,000円", fill=NAVY_TINT, size=16, bold=True, color=NAVY),
                 body_cell(TOOL3, fill=NAVY_TINT, size=13, bold=True, color=NAVY)])
    ops = [("離脱防止ポップアップ", ["導入・運用", "導入・運用", "初期導入支援"]),
           ("あいさつメッセージ改善", ["対応", "対応", "初期構築"]),
           ("リッチメニュー", ["設置", "設置", "初期構築"]),
           ("ステップ配信", ["14日間・10通", "なし", "なし"]),
           ("企画配信（定期配信）", ["月4回程度", "月1〜2回程度", "なし"]),
           ("運用改善支援", ["対応", "限定的", "なし"]),
           ("API連携・開発", ["必要に応じて\n別途見積もり", "なし", "なし"])]
    for label_, vals in ops:
        rows.append([body_cell(label_, fill=NAVY_TINT, size=13, bold=True, color=NAVY)] + [body_cell(v, size=13) for v in vals])
    heights = [0.62, 0.55] + [0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.62]
    add_table(s, 0.42, 1.48, COLS3, heights, rows, name="運用内容比較表")
    y = 1.48 + sum(heights) + 0.14
    notes = [
        "※API連携・開発は基本料金に含まれません。必要に応じて、プラン①で別途お見積りします。",
        "※13ページの試算（6か月目に月33件）は、プラン①｜総合運用の基本施策を対象としたものです。プラン②・③の成果は含みません。",
        "※プラン③は、導入後も継続して利用するツールについて、ツール利用料およびDYM手数料が発生します。",
    ]
    add_box(s, 0.42, y, 10.0, 7.1 - y, [[(t, {})] for t in notes], size=11, color=NOTE_COLOR, align="l", anchor="t",
            ins=(0.02, 0.02, 0.02, 0.02), after=2, name="注記")


def apply(prs):
    sl = list(prs.slides)
    p08(sl[7])
    p09(sl[8])
    p12(sl[11])
    p13(sl[12])
    p15(sl[14])
    p16(sl[15])
    p18(sl[17])
    p21(sl[20])
