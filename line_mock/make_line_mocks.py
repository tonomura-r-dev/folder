# -*- coding: utf-8 -*-
"""
LINEトーク画面モックアップ生成（しが就職・転職フェア 提案資料用）

HTMLで組んでヘッドレスChromeで3倍解像度スクショする方式。
画像生成AIのプロンプトの代わりに、同じ絵を毎回同じ形で出せるようにしたもの。

  cd line_mock
  python make_line_mocks.py            # 全部作る
  python make_line_mocks.py after s13  # 名前を指定して一部だけ

出力は line_mock/out/ 配下。
"""
import base64
import os
import sys

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
SCALE = 3  # 出力倍率（PPTX貼付なら3倍で十分きれい）

# ---- 配色（提案資料のトンマナに合わせる）-------------------------------
CHAT_BG = "#E9EDF2"      # トーク背景（純白にしない）
MY_GREEN = "#8DE055"     # 自分の吹き出し
LINE_GREEN = "#06C755"   # 選択ボタンの文字
NAVY = "#1F285A"         # 資料のメインカラー
PALE = "#F4F7FF"         # 資料のサブカラー

W, H = 390, 845          # 9:19.5 のスマホ画面

# ---- アイコン（すべてインラインSVG。外部画像は使わない）---------------
SVG_SIGNAL = ('<svg width="18" height="11" viewBox="0 0 18 11" fill="#1A1A1A">'
              '<rect x="0" y="7.5" width="3" height="3.5" rx="1"/>'
              '<rect x="5" y="5" width="3" height="6" rx="1"/>'
              '<rect x="10" y="2.5" width="3" height="8.5" rx="1"/>'
              '<rect x="15" y="0" width="3" height="11" rx="1"/></svg>')
SVG_WIFI = ('<svg width="16" height="12" viewBox="0 0 16 12" fill="#1A1A1A">'
            '<path d="M8 11.4 5.3 8.3a4.2 4.2 0 0 1 5.4 0L8 11.4Z"/>'
            '<path d="M8 4.6c1.6 0 3.1.6 4.2 1.6l1.4-1.6A8.6 8.6 0 0 0 8 2.3 8.6 8.6 0 0 0 2.4 4.6l1.4 1.6A6.3 6.3 0 0 1 8 4.6Z"/></svg>')
SVG_BATT = ('<svg width="26" height="12" viewBox="0 0 26 12">'
            '<rect x=".5" y=".5" width="22" height="11" rx="3.4" fill="none" stroke="#1A1A1A" stroke-opacity=".38"/>'
            '<rect x="2" y="2" width="19" height="8" rx="2.2" fill="#1A1A1A"/>'
            '<path d="M24 4.2v3.6c1-.3 1.6-1 1.6-1.8S25 4.5 24 4.2Z" fill="#1A1A1A" fill-opacity=".38"/></svg>')
SVG_BACK = ('<svg width="11" height="19" viewBox="0 0 11 19" fill="none">'
            '<path d="M9.4 1.2 1.6 9.5l7.8 8.3" stroke="#3A3F47" stroke-width="2" '
            'stroke-linecap="round" stroke-linejoin="round"/></svg>')
SVG_CALL = ('<svg width="21" height="21" viewBox="0 0 21 21" fill="none">'
            '<path d="M6.4 2.6 8.2 6 6.5 8a12 12 0 0 0 6.5 6.5l2-1.7 3.4 1.8-.6 3.4c-.2 1-1 1.6-2 1.5'
            'C8.9 18.9 2.1 12.1 1.5 5.2c-.1-1 .5-1.8 1.5-2l3.4-.6Z" stroke="#3A3F47" stroke-width="1.7" '
            'stroke-linejoin="round"/></svg>')
SVG_MENU = ('<svg width="20" height="14" viewBox="0 0 20 14" fill="none" stroke="#3A3F47" stroke-width="1.8" '
            'stroke-linecap="round"><path d="M1 1h18M1 7h18M1 13h18"/></svg>')
SVG_PERSON = ('<svg width="34" height="34" viewBox="0 0 34 34">'
              '<circle cx="17" cy="17" r="17" fill="#C6CDD6"/>'
              '<circle cx="17" cy="13.4" r="5.5" fill="#fff"/>'
              '<path d="M6.4 30.8a10.9 10.9 0 0 1 21.2 0A16.9 16.9 0 0 1 17 34c-4 0-7.6-1.2-10.6-3.2Z" fill="#fff"/></svg>')
SVG_PLUS = ('<svg width="22" height="22" viewBox="0 0 22 22" fill="none" stroke="#5B6470" stroke-width="1.7" '
            'stroke-linecap="round"><path d="M11 4v14M4 11h14"/></svg>')
SVG_CAMERA = ('<svg width="23" height="20" viewBox="0 0 23 20" fill="none" stroke="#5B6470" stroke-width="1.6">'
              '<rect x="1" y="4.2" width="21" height="14.8" rx="3.2"/>'
              '<path d="M7.6 4.2 9 1.4h5l1.4 2.8" stroke-linejoin="round"/><circle cx="11.5" cy="11.6" r="4.1"/></svg>')
SVG_MIC = ('<svg width="16" height="21" viewBox="0 0 16 21" fill="none" stroke="#5B6470" stroke-width="1.6">'
           '<rect x="4.8" y="1" width="6.4" height="11.4" rx="3.2"/>'
           '<path d="M1.4 10.2a6.6 6.6 0 0 0 13.2 0M8 16.8V20" stroke-linecap="round"/></svg>')
SVG_EMOJI = ('<svg width="21" height="21" viewBox="0 0 21 21" fill="none" stroke="#5B6470" stroke-width="1.6">'
             '<circle cx="10.5" cy="10.5" r="9.2"/><path d="M6.6 12.6a4.8 4.8 0 0 0 7.8 0" stroke-linecap="round"/>'
             '<circle cx="7.6" cy="8.2" r="1.1" fill="#5B6470" stroke="none"/>'
             '<circle cx="13.4" cy="8.2" r="1.1" fill="#5B6470" stroke="none"/></svg>')
SVG_CURSOR = ('<svg width="27" height="35" viewBox="0 0 27 35">'
              '<path d="M3 2.2 19.6 15.6l-7.4 1 4.2 8.6-3.7 1.8-4.2-8.6L3 23.9z" fill="#fff" stroke="#1A1A1A" '
              'stroke-width="1.7" stroke-linejoin="round"/></svg>')
SVG_BUILDING = ('<svg width="34" height="34" viewBox="0 0 34 34" fill="none" stroke="#1F285A" stroke-width="1.7">'
                '<path d="M4 30V8.5L15 4v26M15 30h15V13.5L15 9.5" stroke-linejoin="round"/>'
                '<path d="M8 13h3M8 18h3M8 23h3M20 18h5M20 23h5" stroke-linecap="round"/></svg>')
SVG_DOC = ('<svg width="52" height="52" viewBox="0 0 52 52" fill="none" stroke="#1F285A" stroke-width="2">'
           '<path d="M12 4h20l8 8v36H12z" stroke-linejoin="round"/><path d="M32 4v9h8" stroke-linejoin="round"/>'
           '<path d="M19 23h14M19 31h14M19 39h9" stroke-linecap="round"/></svg>')

# ---- 共通スタイル -------------------------------------------------------
CSS = f"""
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#FFFFFF; font-family:"Noto Sans JP","Yu Gothic UI",sans-serif;
        -webkit-font-smoothing:antialiased; }}
#canvas {{ background:#FFFFFF; overflow:hidden; }}

/* スマホ外枠：黒の細線のみ */
.phone {{ position:relative; display:flex; flex-direction:column;
          border:1px solid #1A1A1A; background:#fff; overflow:hidden; }}

/* ステータスバー */
.status {{ height:44px; flex:0 0 44px; display:flex; align-items:center;
           justify-content:space-between; padding:0 20px 0 24px; background:#fff; }}
.status .clock {{ font-size:14px; font-weight:700; color:#1A1A1A; letter-spacing:.2px; }}
.status .icons {{ display:flex; align-items:center; gap:5px; }}

/* ヘッダー */
.head {{ height:48px; flex:0 0 48px; display:flex; align-items:center; gap:12px;
         padding:0 16px; background:#fff; border-bottom:1px solid #E4E7EB; }}
.head .name {{ flex:1; font-size:16px; font-weight:700; color:#1A1A1A; letter-spacing:.2px; }}
.head .tools {{ display:flex; align-items:center; gap:16px; }}

/* トーク背景 */
.chat {{ flex:1; background:{CHAT_BG}; padding:14px 12px 10px; display:flex;
         flex-direction:column; gap:9px; overflow:hidden; }}

/* 日付の区切り */
.date {{ align-self:center; background:#C3CBD6; color:#fff; font-size:11px; font-weight:500;
         padding:3px 13px; border-radius:11px; margin:4px 0; position:relative; }}
.date .day {{ position:absolute; left:calc(100% + 10px); top:2px; white-space:nowrap;
              color:#9AA3AF; font-size:11px; font-weight:700; letter-spacing:.4px; }}

/* メッセージ行 */
.row {{ display:flex; align-items:flex-end; gap:6px; }}
.row.out {{ justify-content:flex-end; }}
.ava {{ flex:0 0 34px; align-self:flex-start; line-height:0; }}
.ava.blank {{ }}
.stack {{ display:flex; flex-direction:column; align-items:flex-start; gap:6px; max-width:250px; }}
.meta {{ font-size:10px; color:#8A93A0; line-height:1.3; white-space:nowrap; padding-bottom:2px; }}
.meta.r {{ display:flex; flex-direction:column; align-items:flex-end; }}

/* 吹き出し：影なし・フラット */
.bub {{ position:relative; font-size:14.5px; line-height:1.5; color:#1A1A1A;
        padding:9px 13px; border-radius:18px; word-break:break-all; }}
.bub.i {{ background:#FFFFFF; }}
.bub.i::before {{ content:""; position:absolute; left:-5px; top:11px; width:0; height:0;
                  border-right:8px solid #FFFFFF; border-bottom:8px solid transparent; }}
.bub.o {{ background:{MY_GREEN}; }}
.bub.o::before {{ content:""; position:absolute; right:-5px; top:11px; width:0; height:0;
                  border-left:8px solid {MY_GREEN}; border-bottom:8px solid transparent; }}

/* 選択ボタンのカード */
.card {{ width:250px; background:#fff; border:1px solid #E1E5EA; border-radius:14px; overflow:hidden; }}
.card .btn {{ height:46px; display:flex; align-items:center; justify-content:center;
              font-size:14.5px; font-weight:500; color:{LINE_GREEN}; border-top:1px solid #EDF0F3; }}
.card .btn:first-child {{ border-top:none; }}
.card .btn.sel {{ background:{NAVY}; color:#fff; font-weight:700; border-top-color:{NAVY}; }}

/* 濃紺のCTAボタン */
.cta {{ width:250px; height:48px; background:{NAVY}; color:#fff; border-radius:12px;
        display:flex; align-items:center; justify-content:center; font-size:15px; font-weight:700; }}
.cta.wide {{ width:100%; height:54px; font-size:16px; }}

/* カルーセル */
.carousel {{ display:flex; gap:9px; margin-left:40px; }}
.ccard {{ flex:0 0 172px; background:#fff; border:1px solid #E1E5EA; border-radius:14px; overflow:hidden; }}
.ccard .thumb {{ height:92px; background:{PALE}; display:flex; align-items:center; justify-content:center; }}
.ccard .tt {{ padding:10px 12px 2px; font-size:13px; font-weight:700; color:#1A1A1A; }}
.ccard .ds {{ padding:0 12px 10px; font-size:11.5px; color:#7C8593; line-height:1.5; }}
.ccard .cb {{ margin:0 12px 12px; height:36px; background:{NAVY}; color:#fff; border-radius:8px;
              display:flex; align-items:center; justify-content:center; font-size:12.5px; font-weight:700; }}

/* リッチメニュー */
.rich {{ flex:0 0 auto; display:grid; grid-template-columns:repeat(3,1fr); gap:2px; background:#fff; }}
.rich .cell {{ display:flex; align-items:center; justify-content:center; text-align:center;
               font-size:13px; font-weight:700; line-height:1.4; padding:0 6px; }}
.rich .cell.top {{ background:{PALE}; color:{NAVY}; }}
.rich .cell.bottom {{ background:{NAVY}; color:#fff; }}

/* 入力バー */
.input {{ height:52px; flex:0 0 52px; display:flex; align-items:center; gap:12px;
          padding:0 14px; background:#fff; border-top:1px solid #E4E7EB; }}
.input .field {{ flex:1; height:32px; background:#F1F3F6; border-radius:16px; }}

/* 骨組み（ステップ配信用のダミー吹き出し）*/
.sk {{ height:14px; border-radius:7px; background:#D3DAE3; }}
.sk.w {{ background:#fff; }}
.skb {{ padding:10px 13px; border-radius:18px; display:flex; flex-direction:column; gap:7px; }}

/* タップカーソル */
.cursor {{ position:absolute; z-index:5; }}
.ripple {{ position:absolute; z-index:4; width:52px; height:52px; border-radius:50%;
           background:rgba(31,40,90,.16); }}
"""


# ---- 部品 ---------------------------------------------------------------
def status_bar(clock="9:41"):
    return (f'<div class="status"><div class="clock">{clock}</div>'
            f'<div class="icons">{SVG_SIGNAL}{SVG_WIFI}{SVG_BATT}</div></div>')


def header(name):
    return (f'<div class="head">{SVG_BACK}<div class="name">{name}</div>'
            f'<div class="tools">{SVG_CALL}{SVG_MENU}</div></div>')


def input_bar():
    return (f'<div class="input">{SVG_PLUS}{SVG_CAMERA}{SVG_MIC}'
            f'<div class="field"></div>{SVG_EMOJI}</div>')


def date_pill(label, day=None):
    tag = f'<span class="day">{day}</span>' if day else ""
    return f'<div class="date">{label}{tag}</div>'


def msg_in(html, time="10:23", avatar=True, bubble=True):
    """相手のメッセージ。bubble=False なら中身（カード等）をそのまま吹き出しの位置に置く。"""
    ava = f'<div class="ava">{SVG_PERSON}</div>' if avatar else '<div class="ava blank"></div>'
    body = f'<div class="bub i">{html}</div>' if bubble else html
    meta = f'<div class="meta">{time}</div>' if time else ""
    return f'<div class="row in">{ava}<div class="stack">{body}</div>{meta}</div>'


def msg_out(text, time="10:24", read=True):
    r = '<span>既読</span>' if read else ""
    meta = f'<div class="meta r">{r}<span>{time}</span></div>' if time else ""
    return f'<div class="row out">{meta}<div class="bub o">{text}</div></div>'


def choice_card(items):
    """items: [(ラベル, 選択済みか), ...]"""
    btns = "".join(f'<div class="btn{" sel" if sel else ""}">{lb}</div>' for lb, sel in items)
    return f'<div class="card">{btns}</div>'


def cta(label, wide=False):
    return f'<div class="cta{" wide" if wide else ""}">{label}</div>'


def rich_menu(top, bottom, row_h=118):
    cells = "".join(f'<div class="cell top">{t}</div>' for t in top)
    cells += "".join(f'<div class="cell bottom">{b}</div>' for b in bottom)
    return f'<div class="rich" style="grid-auto-rows:{row_h}px">{cells}</div>'


def phone(chat_html, name="しが就職・転職フェア", clock="9:41", w=W, h=H,
          chat_bg=CHAT_BG, tail="", show_input=True, extra=""):
    """スマホ1台分のHTML。tail はトーク背景の下に置く要素（リッチメニュー等）。"""
    return (f'<div class="phone" style="width:{w}px;height:{h}px">'
            f'{status_bar(clock)}{header(name)}'
            f'<div class="chat" style="background:{chat_bg}">{chat_html}</div>'
            f'{tail}{input_bar() if show_input else ""}{extra}</div>')


def page(body, cw, ch, canvas_style=""):
    return (f'<!doctype html><html lang="ja"><head><meta charset="utf-8">'
            f'<style>{CSS}</style></head><body>'
            f'<div id="canvas" style="width:{cw}px;height:{ch}px;{canvas_style}">{body}</div>'
            f'</body></html>')


# ---- 各カット -----------------------------------------------------------
def cut_before():
    """Before：予約直後の1通きり。日付だけが過ぎて何も届かない。"""
    chat = (date_pill("8月2日")
            + msg_in("ご予約ありがとうございます。<br>当日お待ちしております。", time="10:12")
            + date_pill("8月3日") + date_pill("8月4日"))
    return page(phone(chat), W, H)


def cut_after():
    """After：予約後もやり取りが続いている。"""
    chat = (date_pill("8月4日")
            + msg_in("ご予約ありがとうございます", time="10:12")
            + msg_in("ご希望の職種を選んでください", time="10:12", avatar=False)
            + msg_in(choice_card([("事務", False), ("製造・軽作業", False),
                                  ("販売・サービス", False)]), time="10:12",
                     avatar=False, bubble=False)
            + msg_in("服装は自由、履歴書も不要です", time="12:30")
            + msg_in("明日開催です", time="18:00", avatar=False)
            + msg_in(cta("会場MAPを見る"), time="18:00", avatar=False, bubble=False))
    return page(phone(chat), W, H)


def cut_s13():
    """Slide13：属性取得アンケート。2番目を選択済みにしてタップカーソルを添える。"""
    chat = (date_pill("8月4日")
            + msg_in("希望職種を教えてください", time="10:23")
            + msg_in(choice_card([("事務", False), ("製造・軽作業", True),
                                  ("販売・サービス", False), ("営業", False)]),
                     time="", avatar=False, bubble=False)
            + msg_out("製造・軽作業", time="10:24"))
    # カーソルは選択済みボタン（上から2番目）の上に置く
    extra = ('<div class="ripple" style="left:223px;top:237px"></div>'
             f'<div class="cursor" style="left:236px;top:250px">{SVG_CURSOR}</div>')
    return page(phone(chat, extra=extra), W, H)


def cut_s14():
    """Slide14：リッチメニュー。画面下半分を6分割メニューが占める。"""
    chat = (date_pill("8月4日")
            + msg_in("ご登録ありがとうございます", time="9:02")
            + msg_in("下のメニューからいつでも確認できます", time="9:02", avatar=False))
    menu = rich_menu(["出展企業一覧", "会場MAP", "開催日程"],
                     ["事前予約", "キャリア相談", "よくある質問"], row_h=130)
    return page(phone(chat, tail=menu), W, H)


def cut_s15_sat():
    """Slide15（土曜夜）：カルーセルで比較検討。画面からはみ出させる。"""
    cards = ""
    for title, desc in [("注目企業", "正社員・年間休日120日"),
                        ("注目企業", "未経験歓迎・研修あり"),
                        ("注目企業", "地元勤務・転勤なし")]:
        cards += (f'<div class="ccard"><div class="thumb">{SVG_BUILDING}</div>'
                  f'<div class="tt">{title}</div><div class="ds">{desc}</div>'
                  f'<div class="cb">詳しく見る</div></div>')
    chat = (date_pill("8月8日")
            + msg_in("今週の注目企業をまとめました", time="20:00")
            + f'<div class="carousel">{cards}</div>')
    return page(phone(chat, clock="20:00"), W, H)


def cut_s15_sun():
    """Slide15（日曜夜）：行動を1つに絞る。"""
    chat = (date_pill("8月9日")
            + msg_in("明日で締切です", time="19:00")
            + msg_in(cta("予約する", wide=True), time="", avatar=False, bubble=False))
    return page(phone(chat, clock="19:00"), W, H)


def cut_s16():
    """Slide16：時間帯別配信。スマホ3台を等間隔で横一列（16:9）。"""
    morning = phone(date_pill("8月4日") + msg_in("おはようございます。<br>本日の新着求人です",
                                                time="7:30"),
                    clock="7:30", chat_bg="#F5F8FC")
    noon = phone(date_pill("8月4日")
                 + msg_in("お昼の読みもの", time="12:20")
                 + msg_in('<div class="ccard" style="width:214px"><div class="thumb">'
                          + SVG_BUILDING + '</div><div class="tt">面接で聞かれること</div>'
                          '<div class="ds">当日までに読んでおきたい3つ</div>'
                          '<div class="cb">記事を読む</div></div>',
                          time="12:20", avatar=False, bubble=False),
                 clock="12:20", chat_bg="#FBF7F1")
    night = phone(date_pill("8月4日")
                  + msg_in("本日の予約は今夜まで", time="21:00")
                  + msg_in(cta("予約する", wide=True), time="", avatar=False, bubble=False),
                  clock="21:00", chat_bg="#D9DEE6")

    cw, ch = 1280, 720
    scale = 0.72
    pw, ph = W * scale, H * scale
    gap = (cw - pw * 3) / 4
    body = ""
    for i, p in enumerate([morning, noon, night]):
        left = gap + i * (pw + gap)
        body += (f'<div style="position:absolute;left:{left:.0f}px;top:{(ch-ph)/2:.0f}px;'
                 f'width:{pw:.0f}px;height:{ph:.0f}px">'
                 f'<div style="transform:scale({scale});transform-origin:top left">{p}</div></div>')
    return page(body, cw, ch, canvas_style="position:relative")


def cut_s18():
    """Slide18：Day0-14のステップ配信。文字は入れず流れだけを見せる（9:24）。"""
    h = int(W * 24 / 9)
    days = [("Day 0", 2), ("Day 1", 1), ("Day 3", 2), ("Day 6", 1), ("Day 7", 2), ("Day 14", 1)]
    chat = ""
    for i, (day, lines) in enumerate(days):
        chat += date_pill("　　　", day=day)
        bars = "".join(f'<div class="sk w" style="width:{w}px"></div>'
                       for w in ([168, 118] if lines == 2 else [140]))
        chat += msg_in(f'<div class="skb" style="background:#fff">{bars}</div>',
                       time="", bubble=False)
        if i in (1, 4):  # 途中でユーザーからの反応も入れておく
            chat += ('<div class="row out"><div class="skb" style="background:' + MY_GREEN + '">'
                     '<div class="sk" style="width:92px;background:#6FC93E"></div></div></div>')
    return page(phone(chat, h=h, show_input=False), W, h)


def cut_s20():
    """Slide20：LP離脱防止ポップアップ（9:16）。"""
    h = int(W * 16 / 9)
    lp = (f'<div style="position:absolute;inset:0;background:#fff">'
          f'<div style="height:56px;background:{NAVY};display:flex;align-items:center;'
          f'padding:0 18px;color:#fff;font-size:14px;font-weight:700">しが就職・転職フェア</div>'
          f'<div style="height:190px;background:{PALE};display:flex;flex-direction:column;'
          f'align-items:center;justify-content:center;gap:10px">'
          f'<div style="font-size:20px;font-weight:700;color:{NAVY}">8/10（日）開催</div>'
          f'<div style="font-size:13px;color:#5C6577">滋賀県内 40社が出展／入場無料</div></div>'
          f'<div style="padding:22px 20px;display:flex;flex-direction:column;gap:11px">'
          + "".join(f'<div class="sk" style="width:{w}%"></div>' for w in [100, 92, 96, 68])
          + f'<div style="height:14px"></div>'
          + "".join(f'<div class="sk" style="width:{w}%"></div>' for w in [100, 88, 74])
          + '</div></div>')
    overlay = '<div style="position:absolute;inset:0;background:rgba(15,20,35,.55)"></div>'
    popup = (f'<div style="position:absolute;left:28px;right:28px;top:50%;transform:translateY(-50%);'
             f'background:#fff;border-radius:18px;padding:24px 20px 18px;display:flex;'
             f'flex-direction:column;align-items:center;gap:14px;'
             f'box-shadow:0 6px 20px rgba(0,0,0,.12)">'
             f'<div style="font-size:12px;font-weight:700;color:{LINE_GREEN};letter-spacing:1px">無料プレゼント</div>'
             f'{SVG_DOC}'
             f'<div style="font-size:15px;font-weight:700;color:#1A1A1A;text-align:center;line-height:1.5">'
             f'出展企業40社の<br>まとめ資料</div>'
             f'<div style="width:100%;height:50px;background:{NAVY};color:#fff;border-radius:12px;'
             f'display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:700">'
             f'LINEで受け取る</div>'
             f'<div style="font-size:12px;color:#9AA3AF">閉じる</div></div>')
    body = (f'<div style="position:relative;width:{W}px;height:{h}px;border:1px solid #1A1A1A;'
            f'overflow:hidden">{lp}{overlay}{popup}</div>')
    return page(body, W, h)


CUTS = [
    ("before", "01_before", cut_before),
    ("after", "02_after", cut_after),
    ("s13", "03_slide13_survey", cut_s13),
    ("s14", "04_slide14_richmenu", cut_s14),
    ("s15sat", "05_slide15_sat", cut_s15_sat),
    ("s15sun", "06_slide15_sun", cut_s15_sun),
    ("s16", "07_slide16_timeofday", cut_s16),
    ("s18", "08_slide18_step", cut_s18),
    ("s20", "09_slide20_popup", cut_s20),
]


# ---- レンダリング -------------------------------------------------------
def make_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--hide-scrollbars")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--force-color-profile=srgb")
    opts.add_argument("--font-render-hinting=none")
    return webdriver.Chrome(options=opts)


def shoot(driver, html, out_path):
    """#canvas のサイズちょうどで SCALE 倍のPNGを吐く。"""
    driver.get("data:text/html;charset=utf-8;base64," +
               base64.b64encode(html.encode("utf-8")).decode("ascii"))
    w, h = driver.execute_script(
        "const c=document.getElementById('canvas');"
        "return [Math.round(c.offsetWidth), Math.round(c.offsetHeight)];")
    driver.execute_cdp_cmd("Emulation.setDeviceMetricsOverride",
                           {"width": w, "height": h, "deviceScaleFactor": SCALE, "mobile": False})
    driver.get_screenshot_as_file(out_path)
    driver.execute_cdp_cmd("Emulation.clearDeviceMetricsOverride", {})
    return w * SCALE, h * SCALE


def main():
    want = [a.lower() for a in sys.argv[1:]]
    targets = [c for c in CUTS if not want or c[0] in want or c[1] in want]
    if not targets:
        print("該当なし。指定できる名前:", ", ".join(k for k, _, _ in CUTS))
        return

    os.makedirs(OUT_DIR, exist_ok=True)
    driver = make_driver()
    try:
        for key, fname, fn in targets:
            path = os.path.join(OUT_DIR, fname + ".png")
            px = shoot(driver, fn(), path)
            print(f"{fname}.png  {px[0]}x{px[1]}")
    finally:
        driver.quit()
    print(f"\n完了：{len(targets)}枚 → {OUT_DIR}")


if __name__ == "__main__":
    main()
