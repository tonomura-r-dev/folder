# -*- coding: utf-8 -*-
"""タナカホーム株式会社 LINE公式アカウント活用のご提案 v1（2026-10-15 殿村さん指示「この要件で資料に。確認は後で」）。

    python -m _build.build_tanaka_v1 <あらうちv10.pptx（器）> <出力.pptx>

  器：あらうち v10（4:3・DYMのヘッダー／表紙／裏表紙）。表紙と裏表紙だけ残し、本文は全部作り直す。
  構成はJ.P.Returnsのお手本（タイトル＝論点：結論／リード1行／表 or KEY TAKEAWAY＋補足3つ）。
  要件は殿村さんがチャットに貼った他AIの方針書（ヒアリングQ1〜Q7の回答入り）。社内のルールに合わない所は直した：
   ・SIMはFMT ver3.00の式で作り直した（CVR＝配信クリック×0.4〜0.8%）。他AIの「6か月目に来場48件」は使わない
   ・出典のない数字（UA値0.25・平屋1,500万円台・月々5万円台・QUOカード5,000円分 など）は載せない
   ・費用：コンサル 初期30万・月25万（Q7「月20万〜30万強」）／離脱防止 初期1.5万・月3万／LINE広告は広告費から仮に月20万
   ・あいさつは0日目（ステップ配信は最短1日後）
  確認待ちの点は `_drafts/タナカホーム_与件整理（2026-10-15）.md` にまとめた。
"""
import sys

from pptx import Presentation
from pptx.util import Inches

from _build.arauchi_v4.lib import NAVY, RED, TEXT, add_box, add_table, move, set_paras
from _build.arauchi_v10.build import MID, TAKE, TINT, LINE_C, by_id, clone, key_takeaway, meiryo, runs, three_points

GRAY = "7F7F7F"
HDR = (633, 634, 635, 636, 637)          # あらうちv10 2ページ（サマリー）のタイトル・リード帯・区切り線・リード
COVER, BACK, HDR_SRC = 1, 23, 2

# ---------------------------------------------------------------- 試算（FMT ver3.00の式。build_tanaka_sim() で計算）
UU = 13228
CPF_BUDGET, CPF_UNIT = 200000, 200
CVR = [0.004, 0.005, 0.006, 0.007, 0.008, 0.008]
MONTH_COST = 250000 + 30000 + CPF_BUDGET       # コンサル＋離脱防止＋LINE広告（LINE公式アカウントの利用料は含めない）


def build_tanaka_sim():
    p36 = 1.1 ** 4                               # FMTの「想定の係数（反応率）」（特典・クリエイティブ・知名度・動線＝すべてTrue）
    rows, cum = [], 0
    for i in range(6):
        blk = 0.316 * 1.025 ** i                 # FMTのブロック率（初月31.6%・月1.025倍）
        cpf = CPF_BUDGET / CPF_UNIT * 1.01 ** i  # FMT：CPF増分＝広告費÷単価、2か月目以降は前月の1.01倍
        exitp = UU * 0.55 * 0.12 * 0.12 * 1.01 ** max(0, i - 1)   # 離脱防止＝UU×表示55%×クリック12%×追加12%
        base = UU * 0.005                        # FMT：AD以外のLINE誘導（サイトのLINEボタンなど）
        add = cpf + exitp + base
        cum += add
        tgt = cum * (1 - blk)
        k = 0.8 if i >= 3 else 1                 # FMT：4〜6か月後係数
        plan = tgt * 4 * 0.8 * 0.78 * 0.1 * p36 * k      # 企画配信 月4本
        step = tgt * 5 * 0.725 * 0.12 * 0.2 * p36 * k    # ステップ配信 5本（1・3・7・10・15日後）
        clk = (plan + step) * 2.05
        cv = round(clk * CVR[i])
        rows.append(dict(add=round(add), cum=round(cum), blk=blk, tgt=round(tgt), clk=round(clk), cv=cv,
                         unit=round(MONTH_COST / cv)))
    return rows


SIM = build_tanaka_sim()


# ---------------------------------------------------------------- 部品
def yen(n):
    return f"{n:,}円"


def new_page(prs, src, title, lead, lead_size=None):
    s = prs.slides.add_slide(src.slide_layout)
    for sh in list(s.shapes):
        sh._element.getparent().remove(sh._element)
    t = clone(s, by_id(src, HDR[0]))
    set_paras(t, [title])
    meiryo(t)
    move(t, w=8.4)
    for sid in HDR[1:4]:
        clone(s, by_id(src, sid))
    ld = clone(s, by_id(src, HDR[4]))
    set_paras(ld, lead if isinstance(lead, list) else [lead], size=lead_size)
    meiryo(ld)
    return s


B = {k: (0.75, LINE_C) for k in "lrtb"}


def head_cell(text, size=11.5):
    return dict(text=text, fill=NAVY, color="FFFFFF", bold=True, size=size, borders=B)


def cell(text, size=11, bold=False, color=TEXT, fill="FFFFFF", align="l"):
    return dict(text=text, size=size, bold=bold, color=color, fill=fill, align=align, borders=B,
                ins=(0.1, 0.04, 0.08, 0.04))


def label_cell(text, size=11.5):
    return cell(text, size=size, bold=True, color=NAVY, fill=TINT, align="ctr")


def note(s, text, y=6.72, h=0.3, size=9):
    add_box(s, 0.55, y, 9.73, h, [runs((text, dict(size=size, color=GRAY)))], anchor="t", ins=(0, 0, 0, 0),
            line_spacing=115, name="注記")


def band(s, text_runs, y=6.0, h=0.62):
    add_box(s, 0.55, y, 9.73, h, [runs(*text_runs)], fill=TINT, align="ctr", anchor="ctr", name="示唆")


# ---------------------------------------------------------------- ページ
def p_cover(s):
    set_paras(by_id(s, 31), ["タナカホーム株式会社 御中"])
    set_paras(by_id(s, 26), ["LINE公式アカウント活用のご提案"])
    set_paras(by_id(s, 27), ["資料請求の後もLINEでつながり、", "モデルハウスへの来場につなげる"])
    for sid in (31, 26, 27):
        meiryo(by_id(s, sid))


def p_summary(prs, src):
    s = new_page(prs, src, "サマリー", "来場数の目標（月50〜60件）に向けて、LINEで資料請求の後を追う")
    key_takeaway(s, [runs(("資料請求で終わらせず、LINEで来場までつなげる", dict(size=20, bold=True, color=NAVY))),
                     runs(("離脱防止・LINE広告で友だちを集め、ステップ配信と予約のリマインドで来場を増やす",
                           dict(size=13, color=NAVY)))])
    three_points(s, "補足ポイント", [
        ("01 現状・課題", "資料請求は目標のCPAで取れている。資料請求から来場予約への引き上げは月3〜4件にとどまる。", NAVY),
        ("02 提案", "サイトの離脱防止・資料請求の完了ページ・LINE広告から友だちを集め、前後検索に合わせたステップ配信で来場予約まで案内する。", MID),
        ("03 試算", f"6か月目に、LINE経由の来場予約 月{SIM[5]['cv']}件。初期31.5万円・月48万円（LINE公式アカウントの利用料を除く）。", RED),
    ])


def p_genjo(prs, src):
    s = new_page(prs, src, "現状：資料請求は目標どおり。来場への引き上げが足りない",
                 "資料請求のCPAは目標の3,000円以内。一方で、来場は目標に届いていない")
    rows = [[head_cell("項目"), head_cell("数字"), head_cell("補足")],
            [label_cell("資料請求のCPA"), cell("2,969円", 13, True, NAVY, align="ctr"), cell("目標 3,000円以内")],
            [label_cell("来場数"), cell("月 約30件", 13, True, NAVY, align="ctr"), cell("目標 月50〜60件")],
            [label_cell("資料請求 → 来場予約"), cell("月3〜4件", 13, True, RED, align="ctr"), cell("今のフォロー：メール・電話")],
            [label_cell("検索広告のCPA"), cell("14,796円 → 47,500円", 13, True, NAVY, align="ctr"),
             cell("前年同月との比較（CV 1件）")],
            [label_cell("青森店"), cell("予約のキャンセルが多い", 12, True, NAVY, align="ctr"), cell("来場の質にも課題")]]
    add_table(s, 0.55, 1.85, [2.6, 3.4, 3.73], [0.42] + [0.62] * 5, rows, name="現状")
    note(s, "※数字は貴社よりご提供いただいた資料より", y=5.5)


def p_kadai(prs, src):
    s = new_page(prs, src, "課題：資料請求の後・予約の後・検索広告の3か所",
                 "3つの課題は、どれも「資料請求の後に追う手段がない」ことにつながる")
    rows = [[head_cell("課題"), head_cell("起きていること"), head_cell("LINEでの打ち手")],
            [label_cell("01 資料請求の後"), cell("資料請求から来場予約に進む方が少ない"),
             cell("ステップ配信で、来場予約まで段階的に案内", bold=True, color=NAVY)],
            [label_cell("02 予約の後"), cell("青森店で、予約のキャンセルが多い"),
             cell("来場3日前・前日のリマインドで、当日まで接点を保つ", bold=True, color=NAVY)],
            [label_cell("03 検索広告"), cell("検索広告のCPAが上がり、12月は検討が止まりやすい"),
             cell("友だちに残し、検討が再開する時期に配信で呼び戻す", bold=True, color=NAVY)]]
    add_table(s, 0.55, 1.85, [2.2, 3.65, 3.88], [0.42] + [0.95] * 3, rows, name="課題")
    band(s, [("ポイントは、", dict(size=13, bold=True, color=RED)),
             ("資料請求をした方とサイトを離れた方を、LINEに残すこと", dict(size=13, bold=True, color=NAVY))], y=5.4)


def p_senryaku(prs, src):
    s = new_page(prs, src, "戦略：資料請求者とサイトを離れた方をLINEに集め、来場まで追う",
                 "集める・追う・来てもらうの3段階で、来場を増やす")
    key_takeaway(s, [runs(("LINEを、資料請求と来場の間をつなぐ場所にする", dict(size=20, bold=True, color=NAVY)))],
                 h=1.1)
    three_points(s, "3つの段階", [
        ("01 集める", "サイトを離れる方には離脱防止、資料請求をした方には完了ページ、まだ接点のない方にはLINE広告で、友だちを集める。", NAVY),
        ("02 追う", "3問のアンケートで希望を聞き、前後検索の関心の順番に合わせたステップ配信で、来場予約まで案内する。", MID),
        ("03 来てもらう", "来場予約の後は3日前・前日のリマインドで当日まで接点を保ち、来場後はお礼と次の案内につなげる。", RED),
    ], y=3.7)


def p_zengo(prs, src):
    s = new_page(prs, src, "前後検索：検索前は予算、当日は住まいの種類、検索後は会社と性能",
                 "住まいの検索の前後15日間で、関心はこの順番に移る")
    rows = [[head_cell("時期"), head_cell("主な検索語"), head_cell("関心"), head_cell("LINEで届けること")],
            [label_cell("検索前\n（15日前〜前日）"), cell("家を買う タイミング／平屋 間取り 20坪／住宅ローン 年収／結婚補助金", 10.5),
             cell("予算・タイミング", 11.5, True, NAVY, align="ctr"), cell("月々の返済の目安・補助金", 11)],
            [label_cell("当日"), cell("新築一戸建て／注文住宅 相場／土地探し／建売 注文住宅 どっち", 10.5),
             cell("住まいの種類・相場", 11.5, True, NAVY, align="ctr"), cell("間取り・施工事例", 11)],
            [label_cell("検索後\n（翌日〜15日後）"),
             cell("一条工務店／坪単価 ハウスメーカー／住宅ローン シミュレーション／断熱等級／フラット35／ゼッチ", 10.5),
             cell("会社・性能・ローンの比較", 11.5, True, NAVY, align="ctr"), cell("性能の比べ方・モデルハウスの見学", 11)]]
    add_table(s, 0.55, 1.85, [1.75, 4.0, 1.85, 2.13], [0.42, 1.0, 0.9, 1.15], rows, name="前後検索")
    band(s, [("ステップ配信は、", dict(size=13, bold=True, color=NAVY)),
             ("予算 → 住まいの種類 → 性能・見学", dict(size=13, bold=True, color=RED)),
             (" の順番で届ける", dict(size=13, bold=True, color=NAVY))], y=5.55)
    note(s, "出典：LINEヤフー社提供の前後検索データ", y=6.3)


def p_hurdle(prs, src):
    s = new_page(prs, src, "3つの不安：お金・営業・性能を、来場の前にLINEで解く",
                 "来場をためらう理由を、配信の中身で先に解消する")
    rows = [[head_cell("不安"), head_cell("具体的には"), head_cell("LINEでの解き方")],
            [label_cell("お金"), cell("注文住宅は高そう／ローンが組めるか不安"),
             cell("月々の返済の目安・補助金の案内・ローンの相談をLINEで受け付け", bold=True, color=NAVY)],
            [label_cell("営業"), cell("展示場に行くと売り込まれそう／まだ決めていない"),
             cell("見学の流れを先に伝え、質問はLINEで気軽に受け付け", bold=True, color=NAVY)],
            [label_cell("性能・比較"), cell("大手と何が違うのか／青森の冬は寒くないか"),
             cell("断熱・暖かさの比べ方と、施主の声（冬の光熱費など）を配信", bold=True, color=NAVY)]]
    add_table(s, 0.55, 1.85, [1.6, 3.7, 4.43], [0.42] + [0.95] * 3, rows, name="不安")


def p_zentai(prs, src):
    s = new_page(prs, src, "全体像：3つの入口で集め、LINEの中で来場予約まで案内する",
                 "サイト・資料請求・LINE広告から友だちを集め、来場の後まで追う")
    add_box(s, 0.55, 1.75, 2.35, 0.3, [runs(("入口", dict(size=11, bold=True, color=GRAY)))], ins=(0, 0, 0, 0))
    entries = [("サイトを離れる方", "離脱防止ポップアップ"), ("資料請求をした方", "完了ページのLINEボタン"),
               ("まだ接点のない方", "LINE広告（友だち追加）")]
    for k, (who, how) in enumerate(entries):
        add_box(s, 0.55, 2.1 + k * 0.92, 2.35, 0.8, [runs((who, dict(size=10.5, color=TEXT))),
                                                     runs((how, dict(size=12, bold=True, color=NAVY)))],
                fill="FFFFFF", line=NAVY, align="ctr", anchor="ctr", line_spacing=115, name=f"入口{k + 1}")
    arrow = lambda x, y: add_box(s, x, y, 0.35, 0.5, [runs(("▶", dict(size=16, color=NAVY)))], align="ctr",
                                 anchor="ctr", ins=(0, 0, 0, 0), name="矢印")
    arrow(2.95, 3.0)
    add_box(s, 3.35, 1.75, 2.75, 0.3, [runs(("LINE公式アカウント", dict(size=11, bold=True, color=GRAY)))], ins=(0, 0, 0, 0))
    steps = ["友だち追加：あいさつ＋3問アンケート", "ステップ配信（1〜15日後）", "来場予約：3日前・前日のリマインド"]
    for k, t in enumerate(steps):
        add_box(s, 3.35, 2.1 + k * 0.92, 2.75, 0.8, [runs((t, dict(size=11.5, bold=True, color="FFFFFF")))],
                fill=NAVY, align="ctr", anchor="ctr", line_spacing=115, name=f"LINE{k + 1}")
        if k < 2:
            add_box(s, 4.55, 2.9 + k * 0.92, 0.35, 0.12, [runs(("▼", dict(size=8, color=NAVY)))], align="ctr",
                    anchor="ctr", ins=(0, 0, 0, 0), name="矢印下")
    arrow(6.15, 3.0)
    add_box(s, 6.55, 1.75, 1.7, 0.3, [runs(("来場", dict(size=11, bold=True, color=GRAY)))], ins=(0, 0, 0, 0))
    add_box(s, 6.55, 2.1, 1.7, 2.64, [runs(("モデルハウス\nへの来場", dict(size=13, bold=True, color=RED))),
                                      runs(("八戸・おいらせ\n青森・弘前", dict(size=10.5, color=TEXT)))],
            fill="FFFFFF", line=RED, line_w=1.5, align="ctr", anchor="ctr", line_spacing=120, name="来場")
    arrow(8.3, 3.0)
    add_box(s, 8.7, 1.75, 1.58, 0.3, [runs(("来場の後", dict(size=11, bold=True, color=GRAY)))], ins=(0, 0, 0, 0))
    add_box(s, 8.7, 2.1, 1.58, 2.64, [runs(("お礼・アンケート", dict(size=11, bold=True, color=NAVY))),
                                      runs(("ローンの事前審査の案内", dict(size=11, bold=True, color=NAVY))),
                                      runs(("→ 成約へ", dict(size=12, bold=True, color=RED)))],
            fill=TINT, line=NAVY, align="ctr", anchor="ctr", line_spacing=140, name="来場後")
    rows = [[head_cell("毎月見る数字", 11), cell("友だち数", 11.5, True, NAVY, align="ctr"),
             cell("ブロック率", 11.5, True, NAVY, align="ctr"), cell("配信のクリック数", 11.5, True, NAVY, align="ctr"),
             cell("LINE経由の来場予約数", 11.5, True, RED, align="ctr")]]
    add_table(s, 0.55, 5.2, [1.85, 1.75, 1.75, 2.1, 2.28], [0.55], rows, name="KPI")


def p_atsumeru(prs, src):
    s = new_page(prs, src, "施策①集める：入口ごとに、LINEで受け取れるものを案内する",
                 "サイトを離れる方・資料請求をした方・まだ接点のない方を、それぞれの入口でLINEへ")
    rows = [[head_cell("入口"), head_cell("対象"), head_cell("案内の内容"), head_cell("費用")],
            [label_cell("離脱防止\nポップアップ"), cell("サイトを離れようとした方"),
             cell("資料請求をしなくても、LINEで間取り・施工事例を受け取れる", bold=True, color=NAVY),
             cell("初期1.5万円\n月3万円", align="ctr")],
            [label_cell("資料請求の\n完了ページ"), cell("資料請求をした方"),
             cell("来場予約・ご質問はLINEで受け付け", bold=True, color=NAVY), cell("初期構築に含む", align="ctr")],
            [label_cell("LINE広告\n（友だち追加）"), cell("八戸・おいらせ・青森・弘前で住まいを検討している方"),
             cell("間取り・施工事例をLINEで受け取れる", bold=True, color=NAVY),
             cell("広告費から\n仮に月20万円", align="ctr")]]
    add_table(s, 0.55, 1.85, [1.75, 2.75, 3.45, 1.78], [0.42] + [0.95] * 3, rows, name="集める")
    note(s, "※資料請求の完了後にLINEへ自動で案内するツール（サンクスLINE誘導）はオプション：初期15万円・月3万円〜", y=5.2)


def p_ou(prs, src):
    s = new_page(prs, src, "施策②追う：希望を聞いて配信を分け、来場当日まで接点を保つ",
                 "友だち追加の直後から来場の後まで、段階ごとに案内を変える")
    rows = [[head_cell("段階"), head_cell("やること"), head_cell("ねらい")],
            [label_cell("友だち追加の直後"), cell("あいさつ＋3問アンケート（エリア・住まいの種類・検討の時期）"),
             cell("回答で配信を分ける", bold=True, color=NAVY)],
            [label_cell("1〜15日後"), cell("ステップ配信（次のページ）"), cell("来場予約まで案内", bold=True, color=NAVY)],
            [label_cell("来場予約の後"), cell("3日前・前日に、駐車場・当日の流れ・担当者からのあいさつ"),
             cell("青森店のキャンセルを減らす", bold=True, color=RED)],
            [label_cell("来場の後"), cell("お礼・アンケート・住宅ローンの事前審査の案内"), cell("成約へつなげる", bold=True, color=NAVY)],
            [label_cell("検討が長い方"), cell("施主の声（冬の光熱費など）・施工事例を定期配信"),
             cell("検討が再開したときに思い出してもらう", bold=True, color=NAVY)]]
    add_table(s, 0.55, 1.85, [2.1, 4.65, 2.98], [0.42] + [0.72] * 5, rows, name="追う")
    note(s, "※予約の後のリマインドは、貴社の予約管理と連携して配信（連携の方法は構築時に決定）", y=5.9)


def p_step(prs, src):
    s = new_page(prs, src, "ステップ配信：前後検索の関心の順番に合わせて、15日間で来場予約へ",
                 "予算 → 住まいの種類 → 性能・見学の順に届け、最後に来場予約を案内する")
    data = [("あいさつ\n（追加直後）", "ご挨拶・3問アンケート", "回答に合わせて、間取り・施工事例を案内", "資料・間取りを見たい"),
            ("1日後", "月々いくらで建てられるか", "家賃と比べた月々の返済の目安・補助金", "予算・ローン"),
            ("3日後", "青森の冬でも暖かい家", "断熱・窓の性能と、会社の比べ方", "性能・会社の比較"),
            ("7日後", "平屋と2階建て・土地探し", "平屋と2階建ての違い・エリアの土地探し", "住まいの種類"),
            ("10日後", "モデルハウス見学の案内", "見学の流れ・当日できること", "見学"),
            ("15日後", "来場予約のご案内", "今月の来場特典・LINEから来場予約", "予約を迷っている")]
    rows = [[head_cell("配信"), head_cell("テーマ"), head_cell("内容"), head_cell("前後検索の関心")]]
    for k, (when, theme, body, need) in enumerate(data):
        last = k == len(data) - 1
        rows.append([label_cell(when, 11), cell(theme, 11.5, True, RED if last else NAVY), cell(body, 10.5),
                     cell(need, 10.5, align="ctr")])
    add_table(s, 0.55, 1.85, [1.45, 2.85, 3.5, 1.93], [0.42] + [0.62] * 6, rows, name="ステップ")
    note(s, "※配信の時刻は20時前後を基本に、反応を見て調整。来場特典の内容は貴社と決定", y=6.05)


def p_sim(prs, src):
    s = new_page(prs, src, f"試算：6か月目に、LINE経由の来場予約 月{SIM[5]['cv']}件",
                 "友だちを増やしながら、配信のクリックから来場予約を積み上げる")
    hdr = [head_cell("項目", 11)] + [head_cell(f"{m}か月目", 11) for m in range(1, 7)]
    def row(label, vals, bold=False, color=TEXT):
        return [label_cell(label, 10.5)] + [cell(v, 11, bold, color, align="ctr") for v in vals]
    rows = [hdr,
            row("友だちの増加", [f"{r['add']:,}人" for r in SIM]),
            row("累計の友だち", [f"{r['cum']:,}人" for r in SIM]),
            row("ブロック率", [f"{round(r['blk'] * 100)}%" for r in SIM]),
            row("届く友だち", [f"{r['tgt']:,}人" for r in SIM]),
            row("配信のクリック", [f"{r['clk']:,}回" for r in SIM]),
            row("来場予約（LINE経由）", [f"{r['cv']}件" for r in SIM], True, RED),
            row("来場予約1件あたりの費用", [yen(r["unit"]) for r in SIM], True, NAVY)]
    add_table(s, 0.55, 1.85, [2.35] + [1.23] * 6, [0.42] + [0.45] * 7, rows, name="試算")
    note(s, "※前提（仮の数字を含む）：サイトの月間UU 13,228／離脱防止＝UU×表示55%×クリック12%×追加12%／"
            "LINE広告＝仮に月20万円・友だち1人200円／ブロック率＝31.6%から月1.025倍／来場予約＝配信のクリック×0.4〜0.8%"
            "（運用が育つ想定で月ごとに上げる）／今の友だち数は含めない\n"
            f"※1件あたりの費用＝月{MONTH_COST // 10000}万円（コンサル25万円・離脱防止3万円・LINE広告20万円）÷来場予約。"
            "初期費用・LINE公式アカウントの利用料は含めない", y=5.85, h=0.9)


def p_hiyo(prs, src):
    s = new_page(prs, src, "費用：初期31.5万円・月48万円（LINE広告は広告費から配分）",
                 "LINE運用コンサルを軸に、離脱防止とLINE広告で友だちを集める")
    rows = [[head_cell("項目"), head_cell("初期"), head_cell("月額"), head_cell("内容")],
            [label_cell("LINE運用コンサル"), cell("30万円", 12, True, NAVY, align="ctr"), cell("25万円", 12, True, NAVY, align="ctr"),
             cell("初期：アカウント設計・リッチメニュー・3問アンケート・ステップ配信・予約管理との連携の構築\n"
                  "月額：月2〜4回の配信・数字の分析と改善・月1回の定例", 10)],
            [label_cell("離脱防止ポップアップ"), cell("1.5万円", 12, True, NAVY, align="ctr"),
             cell("3万円", 12, True, NAVY, align="ctr"), cell("サイトを離れる方に、LINEを案内", 10)],
            [label_cell("LINE広告（友だち追加）"), cell("—", 12, align="ctr"), cell("仮に20万円", 12, True, NAVY, align="ctr"),
             cell("広告費からの配分（広告費は実費）", 10)],
            [label_cell("LINE公式アカウント利用料"), cell("—", 12, align="ctr"), cell("実費", 12, align="ctr"),
             cell("配信の数によって変わる", 10)],
            [cell("合計", 12, True, "FFFFFF", NAVY, "ctr"), cell("31.5万円", 13, True, RED, TAKE, "ctr"),
             cell("48万円＋利用料", 13, True, RED, TAKE, "ctr"), cell("", 10, fill=TAKE)]]
    add_table(s, 0.55, 1.85, [2.3, 1.35, 1.6, 4.48], [0.42, 0.95, 0.6, 0.6, 0.6, 0.55], rows, name="費用")
    note(s, "※オプション：サンクスLINE誘導ツール（資料請求の完了後、LINEへ自動で案内）初期15万円・月3万円〜\n"
            "※施策内容・配信本数・連携範囲等により、費用は個別にお見積もりいたします。", y=5.75, h=0.55)


def p_schedule(prs, src):
    s = new_page(prs, src, "スケジュール：4週間を目安に構築し、運用を始める",
                 "設計・構築・配信の準備・確認の4段階で進める")
    rows = [[head_cell("時期"), head_cell("段階"), head_cell("やること")],
            [label_cell("1週目"), cell("設計", 12, True, NAVY, align="ctr"),
             cell("ヒアリング・エリアごとの配信の分け方・3問アンケート・リッチメニューの構成")],
            [label_cell("2週目"), cell("構築", 12, True, NAVY, align="ctr"),
             cell("LINE公式アカウントの設定・アンケート・予約管理との連携")],
            [label_cell("3週目"), cell("配信の準備", 12, True, NAVY, align="ctr"),
             cell("ステップ配信の文面・画像の作成・離脱防止ポップアップの設置")],
            [label_cell("4週目"), cell("確認・公開", 12, True, RED, align="ctr"),
             cell("実機での動作確認（予約・リマインド）・公開・LINE広告の配信開始")]]
    add_table(s, 0.55, 1.85, [1.6, 1.9, 6.23], [0.42] + [0.8] * 4, rows, name="スケジュール")


ORDER = [p_summary, p_genjo, p_kadai, p_senryaku, p_zengo, p_hurdle, p_zentai, p_atsumeru, p_ou, p_step, p_sim,
         p_hiyo, p_schedule]


def main(base, out):
    prs = Presentation(base)
    slides = list(prs.slides)
    src, cover, back = slides[HDR_SRC - 1], slides[COVER - 1], slides[BACK - 1]
    assert by_id(src, HDR[0]).text_frame.text == "サマリー"
    assert "貴社の明るい未来" in "".join(sh.text_frame.text for sh in back.shapes if sh.has_text_frame)
    new = []
    for f in ORDER:
        f(prs, src)
        new.append(prs.slides[-1])                    # add_slide は末尾に足す
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
    for r in SIM:
        print(r)


if __name__ == "__main__":
    main(*sys.argv[1:3])
