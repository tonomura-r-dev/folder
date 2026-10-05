# -*- coding: utf-8 -*-
"""教育（塾）業界：施策の設計図①〜④（16:9・4枚）を見やすく作り直す（2026-09-25）
殿村さんが抜き出した4枚（16:9）をベースに、各スライドの中身を消して組み直す。
方針：1ページ1メッセージ／1つの箱は1行／表は列を減らし数字を大きく／本文14pt以上・数字20pt以上
  python3 _build/build_juku_blueprint_169.py <入力pptx(4枚)> <出力pptx>
  本体（4:3）に差し替えるとき：python3 _build/build_juku_blueprint_169.py <ver2.pptx> <出力> 26
  （4枚目の開始ページを渡す。4:3のときは座標と文字を自動で縮める）
"""
import sys
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
START = int(sys.argv[3]) if len(sys.argv) > 3 else 1
prs = Presentation(SRC)
SW = Emu(prs.slide_width).cm
X0, W = 1.2, 31.47          # 設計は16:9の座標で書き、4:3のときは下の係数で縮める
WIDE = SW > 30
KX = 1.0 if WIDE else 25.12 / 31.47
KY = 1.0 if WIDE else (17.25 - 3.85) / (18.1 - 3.85)
KF = 1.0 if WIDE else 0.88


def mx(x):
    return X0 + (x - X0) * KX


def map_y(y):
    return y if y < 3.85 else 3.85 + (y - 3.85) * KY
NAVY, TNAVY, INK, GRAY, LGRAY = "1F285A", "002060", "333333", "7F7F7F", "D9D9D9"
PALE, RED, GREEN, LINE_GREEN, BEIGE, ORANGE = "F4F7FF", "C00000", "00897B", "06C755", "FFF2CC", "ED7D31"
PHASES = ["①接触", "②離脱", "③育成", "④リード獲得", "⑤リード有効化", "⑥再育成", "⑦入会（契約）", "⑧紹介"]


def meiryo(run):
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", "メイリオ")


def box(s, x, y, w, h, fill, lines, color=INK, size=14, bold=False, anchor=MSO_ANCHOR.MIDDLE,
        align=PP_ALIGN.CENTER, line=None, lw=1.25, shape=MSO_SHAPE.ROUNDED_RECTANGLE, ml=0.25, adj=0.08):
    h = h * KY if y >= 3.85 else h
    sp = s.shapes.add_shape(shape, Cm(mx(x)), Cm(map_y(y)), Cm(w * KX), Cm(h))
    if fill:
        sp.fill.solid()
        sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sp.fill.background()
    if line:
        sp.line.color.rgb = RGBColor.from_string(line)
        sp.line.width = Pt(lw)
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sp.adjustments[0] = adj
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(ml)
    tf.margin_top = tf.margin_bottom = Cm(0.05)
    for i, ln in enumerate(lines):
        t, sz, b, c = (ln, size, bold, color) if isinstance(ln, str) else (list(ln) + [None] * 4)[:4]
        sz, b, c = sz or size, bold if b is None else b, c or color
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = t
        r.font.size = Pt(round(sz * KF * 2) / 2)
        r.font.bold = b
        r.font.color.rgb = RGBColor.from_string(c)
        meiryo(r)
    return sp


def text(s, x, y, w, h, lines, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    return box(s, x, y, w, h, None, lines, color=color, size=size, bold=bold, align=align, anchor=anchor,
               shape=MSO_SHAPE.RECTANGLE, ml=0.05)


def arrow(s, x1, y1, x2, y2, color=NAVY, w=2.5):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(mx(x1)), Cm(map_y(y1)), Cm(mx(x2)), Cm(map_y(y2)))
    c.line.color.rgb = RGBColor.from_string(color)
    c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle"}))


def header(s, title, lead, phases, foot=None):
    tree = s.shapes._spTree
    for sh in list(s.shapes):
        keep = sh.is_placeholder or (Emu(sh.top).cm < 1.6 and not (sh.has_text_frame and sh.text_frame.text.strip()))
        if not keep:
            tree.remove(sh._element)
    text(s, 2.0, 0.3, 28, 1.0, [(title, 18, True, TNAVY)], anchor=MSO_ANCHOR.MIDDLE)
    g = 0.06
    w = (W - g * 7) / 8
    for i, t in enumerate(PHASES):
        n = i + 1
        fill, col = (NAVY, "FFFFFF") if n in phases else ("EEEEEE", "A6A6A6")
        box(s, X0 + i * (w + g), 1.7, w, 0.55, fill, [(t, 9, True, col)],
            shape=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON, ml=0.1)
    text(s, X0, 2.45, W, 1.2, [(lead, 15, False, INK)], anchor=MSO_ANCHOR.MIDDLE)
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, Cm(3.85), Cm(SW), Cm(3.85))
    c.line.color.rgb = RGBColor.from_string(LGRAY)
    if foot:
        text(s, X0, 18.15, W, 0.6, [(foot, 9, False, GRAY)])


def phone(s, x, y, w, h, name="◯◯塾 △△教室"):
    box(s, x, y, w, h, "2B2B2B", [], adj=0.08)
    sx, sy, sw, sh_ = x + 0.3, y + 0.6, w - 0.6, h - 1.1
    box(s, sx, sy, sw, sh_, "C9D6E8", [], shape=MSO_SHAPE.RECTANGLE)
    box(s, sx, sy, sw, 0.8, "FFFFFF", [("‹  " + name, 11, True, INK)], shape=MSO_SHAPE.RECTANGLE, align=PP_ALIGN.LEFT)
    return sx, sy + 0.9, sw, sh_ - 0.9


def bubble(s, x, y, w, h, lines, size=11, fill="FFFFFF"):
    return box(s, x, y, w, h, fill, [(t, size, False, INK) for t in lines], align=PP_ALIGN.LEFT,
               anchor=MSO_ANCHOR.MIDDLE, adj=0.12, ml=0.25)


S = list(prs.slides)[START - 1:START + 3]

# ================= ① 友だち追加の動線 =================
s = S[0]
header(s, "施策の設計図①　友だち追加の動線",
       "どこで、何を訴求して、月に何人がLINEの友だちになるか（大手1社・サイトUU 約28万／月のモデル値）",
       [1, 2, 4], "※サイトUUは大手の有名塾の実数（SimilarWeb）を参考にしたモデル値。係数はDYM SIMの型。チラシ・教室のQRは計測できないため数えていない。費用は税抜")
COLS = [("箇所", 8.0), ("訴求", 12.6), ("LINE追加", 5.2), ("費用", 5.67)]
x = X0
for t, w_ in COLS:
    box(s, x, 4.3, w_ - 0.2, 0.8, "EEEEEE", [(t, 13, True, INK)], shape=MSO_SHAPE.RECTANGLE)
    x += w_
ROUTES = [("00", "離脱防止\nポップアップ", "「まだ決めなくて大丈夫です」", "2,200", "月3万円\n初期1.5万円", "28.4万UU×表示55%×クリック12%×追加12%"),
          ("01", "完了画面からの誘導", "「日程の確認・変更はLINEで」", "300", "月5万円\n初期10万円", "資料請求・体験予約の完了 月2,000件×15%")]
for i, (no, where, msg, n, fee, calc) in enumerate(ROUTES):
    y = 5.4 + i * 3.9
    h = 3.5
    box(s, X0, y, 7.8, h, "FFFFFF", [], line="8EA9DB", shape=MSO_SHAPE.RECTANGLE)
    box(s, X0 + 0.3, y + h / 2 - 0.62, 1.4, 1.25, NAVY, [(no, 16, True, "FFFFFF")], shape=MSO_SHAPE.OVAL, ml=0.02)
    text(s, X0 + 1.9, y, 5.85, h, [(t, 15, True, NAVY) for t in where.split("\n")], anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, X0 + 7.85, y + h / 2, X0 + 8.45, y + h / 2)
    box(s, X0 + 8.0 + 0.5, y, 12.1, h, "FFFFFF", [(msg, 17, True, INK)], line="8EA9DB", shape=MSO_SHAPE.RECTANGLE, align=PP_ALIGN.LEFT, ml=0.5)
    arrow(s, X0 + 20.65, y + h / 2, X0 + 21.2, y + h / 2)
    box(s, X0 + 20.6 + 0.65, y, 4.35, h, LINE_GREEN, [(n, 24, True, "FFFFFF"), ("人／月", 13, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE)
    text(s, X0 + 8.6, y + h - 0.75, 11.9, 0.6, [(calc, 10.5, False, GRAY)], align=PP_ALIGN.LEFT)
    box(s, X0 + 25.8, y, 5.47, h, "FFFFFF", [(ln, 14, False, INK) for ln in fee.split("\n")], line="C9D3E6", shape=MSO_SHAPE.RECTANGLE)
box(s, X0, 13.4, W, 1.7, LINE_GREEN, [("新しい友だち　合計 月2,500人", 24, True, "FFFFFF")], adj=0.15)
text(s, X0, 15.6, W, 1.6, [("一番大きいのは、サイトから帰ろうとする保護者を拾う「離脱防止ポップアップ」。", 15, True, NAVY),
                            ("広告費を足さずに、すでにサイトに来ている保護者とつながれる。", 14, False, INK)])

# ================= ② ステップ配信・企画配信 =================
s = S[1]
header(s, "施策の設計図②　効率改善（ステップ配信・企画配信）",
       "14日間のステップ配信と、時期に合わせた企画配信で、資料請求（CV①）と体験予約（CV②）を取る",
       [3, 5, 6], "※大手1社（サイトUU約28万／月）のモデル値。開封率・クリック率＝DYM SIMの係数、CVR＝0.3〜1.0%、件数は四捨五入。通知メッセージのみ別途費用")
HW = (W - 0.8) / 2


def table(x, title, rows, sub):
    box(s, x, 4.3, HW, 0.9, NAVY, [(title, 15, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE, align=PP_ALIGN.LEFT, ml=0.4)
    text(s, x, 5.25, HW, 0.6, [(sub, 11, False, GRAY)])
    cw = [2.7, HW - 2.7 - 4.2, 4.2]
    for i, (when, msg, n, cv, calc, hit) in enumerate(rows):
        y = 5.95 + i * 1.95
        h = 1.75
        box(s, x, y, cw[0] - 0.12, h, PALE, [(when, 15, True, NAVY)], shape=MSO_SHAPE.RECTANGLE, line="C9D3E6", ml=0.05)
        box(s, x + cw[0], y, cw[1] - 0.12, h, "FFFFFF", [(msg, 14, False, INK), (calc, 9, False, GRAY)], shape=MSO_SHAPE.RECTANGLE, line="C9D3E6", align=PP_ALIGN.LEFT, ml=0.3)
        f, c = (GREEN, "FFFFFF") if hit else ("F2F2F2", GRAY)
        lines = [(n, 18 if hit else 13, True, c), (cv, 10.5, True, c)]
        box(s, x + cw[0] + cw[1], y, cw[2], h, f, lines, shape=MSO_SHAPE.RECTANGLE, ml=0.05)


table(X0, "[ステップ配信]　友だち追加から14日間・自動",
      [("0日後", "あいさつ＋30秒診断 → 資料を受け取る", "3件", "CV① 資料請求", "2,500人×100%×12%×1.0%", True),
       ("3日後", "家でできる勉強のコツ", "信頼づくり", "CVは狙わない", "2,500人×72.5%×12%", False),
       ("5日後", "月謝＋講習費の年間の目安", "2件", "CV① 資料請求", "2,500人×72.5%×12%×1.0%", True),
       ("7日後", "同じタイプの子の事例＋体験のご案内", "2件", "CV② 体験予約", "2,500人×72.5%×12%×0.8%", True),
       ("14日後", "次のテストから逆算＋体験のご案内", "2件", "CV② 体験予約", "2,500人×72.5%×12%×1.0%", True)],
      "対象：新しい友だち 月2,500人")
table(X0 + HW + 0.8, "[企画配信]　時期に合わせて月1〜2本",
      [("1〜2月", "新学年で変わること（学年別）＋体験のご案内", "5件", "CV② 体験予約", "6,800人×78%×10%×1.0%", True),
       ("5月", "中間テスト後：苦手単元チェック（中学生の保護者）", "2件", "CV② 体験予約", "2,700人×78%×10%×0.8%", True),
       ("6月", "夏期講習の早期申込特典＋無料体験", "5件", "CV② 体験予約", "6,800人×78%×10%×1.0%", True),
       ("10月", "2学期の中間テスト後＋冬期講習の早期案内", "3件", "CV② 体験予約", "6,800人×78%×10%×0.6%", True),
       ("11月", "冬期講習の前：体験・入会に至らなかった人へ再案内", "実績で計算", "⑥再育成", "通知メッセージ・別途費用", False)],
      "対象：友だち 累計約10,000人（ブロックを除いて約6,800人）")
box(s, X0, 15.95, W, 1.9, BEIGE, [("ステップ配信で毎月　資料請求 約5件・体験予約 約4件", 18, True, INK),
                                   ("企画配信は1回あたり　体験予約 2〜5件（年間の全体は別ページ）", 15, False, INK)], adj=0.12)

# ================= ③ 満足度改善 =================
s = S[2]
header(s, "施策の設計図③　満足度改善（保護者）",
       "保護者の疑問に、待たせずに答える。すぐ答えられることは自動で、相談は人が答える",
       [3, 5], "※有人対応の時間帯（平日14〜21時）は例。教室の運用に合わせて決める。いずれも月額費用内")
ITEMS = [("QA自動化", "料金・コース・対象学年などは、\n自動応答ですぐ答える"),
         ("個別チャット", "志望校の相談・日程の調整は、\n平日14〜21時にスタッフが1対1で"),
         ("その他", "リッチメニューの右下に\n「無料体験・見学の予約」")]
for i, (t, d) in enumerate(ITEMS):
    y = 4.4 + i * 3.25
    box(s, X0, y, 4.8, 2.9, NAVY, [(t, 17, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE, ml=0.1)
    box(s, X0 + 4.9, y, 15.3, 2.9, "FFFFFF", [(ln, 16, k == 0, INK) for k, ln in enumerate(d.split("\n"))],
        shape=MSO_SHAPE.RECTANGLE, line="C9D3E6", align=PP_ALIGN.LEFT, ml=0.5)
box(s, X0, 14.4, 20.2, 2.0, BEIGE, [("すぐ答える仕組みが、", 17, True, INK), ("資料請求・体験予約の取りこぼしを減らす", 17, True, INK)], adj=0.12)
px, py, pw, ph = phone(s, 22.6, 4.3, 10.0, 13.6)
box(s, px + pw - 4.6, py + 0.3, 4.3, 1.0, "A6E3A1", [("料金はいくら？", 12, False, INK)], adj=0.3)
bubble(s, px + 0.3, py + 1.5, pw - 1.6, 2.6, ["（自動応答）", "コース別の月謝の目安です", "・小学生　◯◯円〜", "・中学生　◯◯円〜"], size=11)
box(s, px + pw - 5.6, py + 4.4, 5.3, 1.0, "A6E3A1", [("志望校の相談をしたい", 12, False, INK)], adj=0.3)
bubble(s, px + 0.3, py + 5.6, pw - 1.6, 1.9, ["（教室長の佐藤です）", "お子さまの学年を教えて", "いただけますか？"], size=11)
my = py + ph - 3.6
cw_, ch_ = pw / 3, 1.8
for k, t in enumerate(["30秒診断", "コース・料金", "教室・講師", "資料を\n受け取る", "よくある\n質問", "無料体験・\n見学の予約"]):
    last = k == 5
    box(s, px + (k % 3) * cw_, my + (k // 3) * ch_, cw_, ch_, RED if last else "FFFFFF",
        [(ln, 10.5, True, "FFFFFF" if last else NAVY) for ln in t.split("\n")], line="BFBFBF", shape=MSO_SHAPE.RECTANGLE, ml=0.03)

# ================= ④ 効率改善（教室・運用側） =================
s = S[3]
header(s, "施策の設計図④　効率改善（教室・運用側）",
       "予約の受付・日程変更・持ち物の案内はLINEで自動に。LINE経由の成果も毎月数字で見える",
       [5, 7], "※GA4のパラメータ（utm）でLINE経由の流入と資料請求・体験予約を計測。いずれも月額費用内")
box(s, X0, 4.35, W, 0.9, NAVY, [("[自動応答]　体験予約の流れ（教室は当日迎えるだけ）", 15, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE, align=PP_ALIGN.LEFT, ml=0.4)
FL = [("LINEで", "体験予約"), ("日時を選ぶ", "空き枠から"), ("予約完了を", "自動で送る"), ("前日に", "持ち物・地図"), ("日程の変更も", "LINEで"), ("教室は", "迎えるだけ")]
n = len(FL)
g = 0.75
w = (W - g * (n - 1)) / n
for i, (a, b) in enumerate(FL):
    x = X0 + i * (w + g)
    last = i == n - 1
    box(s, x, 5.6, w, 3.0, GREEN if last else "FFFFFF", [(a, 15, False, "FFFFFF" if last else INK), (b, 17, True, "FFFFFF" if last else NAVY)],
        line=None if last else "8EA9DB", lw=1.75, ml=0.08)
    if not last:
        arrow(s, x + w + 0.05, 7.1, x + w + g - 0.05, 7.1)
box(s, X0, 9.4, W, 0.9, NAVY, [("[その他]　LINE経由の成果を数える", 15, True, "FFFFFF")], shape=MSO_SHAPE.RECTANGLE, align=PP_ALIGN.LEFT, ml=0.4)
FL2 = [("LINEの配信", "リンクに目印"), ("サイト", "GA4で見分ける"), ("資料請求", "CV①を計測"), ("体験予約", "CV②を計測"), ("月次レポート", "定例会で共有")]
n = len(FL2)
w = (W - g * (n - 1)) / n
for i, (a, b) in enumerate(FL2):
    x = X0 + i * (w + g)
    col = GREEN if i in (2, 3) else NAVY
    box(s, x, 10.65, w, 3.0, "FFFFFF", [(a, 17, True, col), (b, 14, False, GRAY)], line="8EA9DB", lw=1.75)
    if i < n - 1:
        arrow(s, x + w + 0.05, 12.15, x + w + g - 0.05, 12.15)
box(s, X0, 14.6, W, 1.9, BEIGE, [("教室の手間を増やさずに、LINE経由の成果が毎月数字で見える", 19, True, INK)], adj=0.12)

# ページ番号（4:3の本体で、番号の枠が無いスライドだけ足す）
if not WIDE:
    for k, sl in enumerate(S):
        if sl.slide_layout.name.startswith("2_"):   # 2_レイアウトはページ番号が出ないので足す
            tb = sl.shapes.add_textbox(Cm(25.6), Cm(18.28), Cm(1.6), Cm(0.7))
            p = tb.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = str(START + k)
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = RGBColor.from_string("FFFFFF")
            meiryo(r)
prs.save(OUT)
print("saved", OUT)
