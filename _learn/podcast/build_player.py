#!/usr/bin/env python3
"""ep*.md から、ブラウザ読み上げで再生できる単一HTMLプレイヤーを生成する。

台本の単一情報源は ep01.md 〜 ep10.md。台本を直したらこれを実行し直す。

使い方:
    python3 _learn/podcast/build_player.py

出力:
    _learn/podcast/player.html          （ブラウザで直接開ける）
    さらに --out で任意の出力先を追加指定できる（Artifact公開用のコピー）
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SPEAKERS = ("ハル", "ソラ")

# 台本1本あたりの期待発話数。パース漏れの検出用。
EXPECTED = {
    "ep01": 60, "ep02": 56, "ep03": 60, "ep04": 56, "ep05": 58,
    "ep06": 58, "ep07": 62, "ep08": 76, "ep09": 72, "ep10": 75,
}

# 読み上げの崩れを直す置換。表示テキストは元のまま、読み上げ側だけ変換する。
# 追加したいときはここに足す。
# 語境界は \b ではなく「前後がASCII英字でない」で判定する。
# Python3 の \b は日本語文字も単語文字とみなすため、「私服でOK」のような並びで成立しない。
def _word(token: str) -> re.Pattern:
    return re.compile(r"(?<![A-Za-z])" + re.escape(token) + r"(?![A-Za-z])")


SPEECH_FIXES = [
    (re.compile(r"P-MAX"), "ピーマックス"),
    (re.compile(r"Advantage\+"), "アドバンテージプラス"),
    (_word("CAPI"), "キャピ"),
    (_word("ROAS"), "ロアス"),
    (_word("OK"), "オーケー"),
    (re.compile(r"1対1"), "一対一"),
    # 略語と数字が続くと1語に聞こえる（CPA1万円 → CPA 1万円）
    (re.compile(r"(?<=[A-Z])(?=\d)"), " "),
    # 桁区切りのカンマは不自然な間になるので落とす（8,000円 → 8000円）。
    # \b は使わない: Python3 では「円」も単語文字なので境界が成立しない。
    (re.compile(r"(?<=\d),(?=\d{3}(?!\d))"), ""),
]

LINE_RE = re.compile(r"^(%s)：(.+)$" % "|".join(SPEAKERS))


def to_speech(text: str) -> str:
    for pattern, repl in SPEECH_FIXES:
        text = pattern.sub(repl, text)
    return text


def parse_episode(path: Path) -> dict:
    raw = path.read_text(encoding="utf-8")
    head, _, body = raw.partition("\n---\n")
    if not body:
        raise SystemExit(f"{path.name}: 本文の区切り（---）が見つかりません")

    title_match = re.search(r"^#\s*EP(\d+)\s+(.+)$", head, re.MULTILINE)
    if not title_match:
        raise SystemExit(f"{path.name}: 見出し（# EPxx タイトル）が見つかりません")

    def field(label: str) -> str:
        m = re.search(rf"^{label}：(.+)$", head, re.MULTILINE)
        return m.group(1).strip() if m else ""

    lines = []
    for raw_line in body.splitlines():
        m = LINE_RE.match(raw_line.strip())
        if m:
            display = m.group(2).strip()
            lines.append([m.group(1), display, to_speech(display)])

    return {
        "id": path.stem,
        "num": title_match.group(1),
        "title": title_match.group(2).strip(),
        "day": field("対応").replace("`", ""),
        "length": field("想定尺"),
        "lines": lines,
    }


def build(episodes: list[dict]) -> str:
    payload = json.dumps(episodes, ensure_ascii=False, separators=(",", ":"))
    # </script> でテンプレートが割れないようにする
    payload = payload.replace("<", "\\u003c")
    total = sum(len(ep["lines"]) for ep in episodes)
    return TEMPLATE.replace("__EPISODES__", payload).replace("__TOTAL__", str(total))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", action="append", default=[],
                    help="追加の出力先（Artifact公開用のコピーなど）")
    args = ap.parse_args()

    paths = sorted(HERE.glob("ep*.md"))
    if not paths:
        raise SystemExit(f"{HERE} に ep*.md が見つかりません")

    episodes = []
    problems = []
    for path in paths:
        ep = parse_episode(path)
        episodes.append(ep)
        expected = EXPECTED.get(ep["id"])
        got = len(ep["lines"])
        flag = ""
        if expected is not None and expected != got:
            flag = f"  ← 期待 {expected} 行と不一致"
            problems.append(ep["id"])
        print(f"  {ep['id']}  {got:3d}行  {ep['title']}{flag}")

    html = build(episodes)

    targets = [HERE / "player.html"] + [Path(p) for p in args.out]
    for target in targets:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(html, encoding="utf-8")
        print(f"書き出し: {target}  ({len(html):,} bytes)")

    print(f"合計 {sum(len(e['lines']) for e in episodes)} 行 / {len(episodes)} 本")
    if problems:
        print(f"⚠ 発話数が期待と違う台本: {', '.join(problems)}", file=sys.stderr)
        print("  台本を編集したなら build_player.py の EXPECTED も更新してください。",
              file=sys.stderr)
        return 1
    return 0


TEMPLATE = r"""<title>Webマーケ学習 Podcast ── 広告運用 × LINE公式アカウント 全10話</title>

<style>
  :root {
    --bg:#FBFAF7; --surface:#F1EDE4; --line:#DFD8C9; --text:#23201B;
    --muted:#6E675C; --accent:#B0721A; --accent-2:#2F6F63;
    --rail:rgba(176,114,26,.14); --shadow:0 -1px 0 var(--line), 0 -12px 32px rgba(35,32,27,.06);
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg:#181613; --surface:#221F1A; --line:#35302A; --text:#EDE7DC;
      --muted:#9B9284; --accent:#E0A34B; --accent-2:#5FB3A2;
      --rail:rgba(224,163,75,.16); --shadow:0 -1px 0 var(--line), 0 -12px 32px rgba(0,0,0,.4);
    }
  }
  :root[data-theme="dark"] {
    --bg:#181613; --surface:#221F1A; --line:#35302A; --text:#EDE7DC;
    --muted:#9B9284; --accent:#E0A34B; --accent-2:#5FB3A2;
    --rail:rgba(224,163,75,.16); --shadow:0 -1px 0 var(--line), 0 -12px 32px rgba(0,0,0,.4);
  }
  :root[data-theme="light"] {
    --bg:#FBFAF7; --surface:#F1EDE4; --line:#DFD8C9; --text:#23201B;
    --muted:#6E675C; --accent:#B0721A; --accent-2:#2F6F63;
    --rail:rgba(176,114,26,.14); --shadow:0 -1px 0 var(--line), 0 -12px 32px rgba(35,32,27,.06);
  }

  * { box-sizing: border-box; }

  body {
    margin:0; background:var(--bg); color:var(--text);
    font-family:"Hiragino Sans","Hiragino Kaku Gothic ProN","Yu Gothic Medium",
                "Noto Sans JP","Meiryo",system-ui,sans-serif;
    line-height:1.9; -webkit-font-smoothing:antialiased; padding-bottom:152px;
  }
  .wrap { max-width:40rem; margin:0 auto; padding:0 1.25rem; }

  header { padding:2.75rem 0 1.5rem; }
  .eyebrow {
    font-size:.72rem; letter-spacing:.16em; text-transform:uppercase;
    color:var(--muted); margin:0 0 .8rem; font-variant-numeric:tabular-nums;
  }
  h1 {
    font-size:clamp(1.35rem,4.6vw,1.7rem); line-height:1.5; margin:0 0 .9rem;
    font-weight:700; letter-spacing:-.01em; text-wrap:balance;
  }
  .lede { margin:0; font-size:.85rem; color:var(--muted); line-height:1.85; }
  .lede code { font-family:ui-monospace,"SFMono-Regular",Menlo,monospace; font-size:.95em; color:var(--text); }

  .notice {
    margin:1.5rem 0 0; padding:.9rem 1rem; background:var(--surface);
    border:1px solid var(--line); border-radius:3px;
    font-size:.78rem; line-height:1.8; color:var(--muted);
  }
  .notice b { color:var(--text); font-weight:600; }

  /* ---- episode index ---- */
  .idx { border-top:1px solid var(--line); margin-top:1.75rem; }
  .idx h2 {
    font-size:.72rem; letter-spacing:.14em; text-transform:uppercase;
    color:var(--muted); font-weight:700; margin:1.5rem 0 .6rem;
  }
  .ep {
    display:flex; gap:.85rem; align-items:baseline; width:100%; text-align:left;
    appearance:none; background:none; border:0; border-top:1px solid var(--line);
    color:inherit; font:inherit; padding:.8rem .25rem; cursor:pointer;
    transition:background-color .15s ease;
  }
  .ep:first-of-type { border-top:0; }
  .ep:hover { background:var(--rail); }
  .ep:focus-visible { outline:2px solid var(--accent); outline-offset:-2px; }
  .ep[aria-current="true"] .ep-t { color:var(--accent); font-weight:700; }
  .ep-n {
    font-size:.74rem; color:var(--muted); font-variant-numeric:tabular-nums;
    letter-spacing:.06em; flex:none; width:2.2rem; padding-top:.15rem;
  }
  .ep-body { flex:1; min-width:0; }
  .ep-t { display:block; font-size:.92rem; line-height:1.6; }
  .ep-m { display:block; font-size:.72rem; color:var(--muted); line-height:1.7; }
  .ep-s {
    flex:none; font-size:.7rem; color:var(--muted); font-variant-numeric:tabular-nums;
    padding-top:.2rem; white-space:nowrap;
  }
  .ep-s.part { color:var(--accent); }

  /* ---- now playing ---- */
  .now { border-top:1px solid var(--line); margin-top:1.5rem; padding-top:1.75rem; }
  .now h2 { font-size:clamp(1.15rem,4vw,1.4rem); line-height:1.5; margin:.35rem 0 .7rem; font-weight:700; text-wrap:balance; }
  .cast { display:flex; gap:1.25rem; flex-wrap:wrap; padding-top:.3rem; }
  .cast div { display:flex; align-items:baseline; gap:.5rem; font-size:.8rem; color:var(--muted); }
  .dot { width:.5rem; height:.5rem; border-radius:50%; flex:none; transform:translateY(-1px); }
  .dot.a { background:var(--accent); }
  .dot.b { background:var(--accent-2); }
  .cast b { color:var(--text); font-weight:600; }

  /* ---- script ---- */
  .script { padding:1.5rem 0 0; }
  .ln {
    display:block; width:100%; text-align:left; appearance:none; background:none;
    border:0; border-left:2px solid transparent; color:inherit; font:inherit;
    line-height:1.9; padding:.55rem 0 .55rem 1rem; margin:0; cursor:pointer;
    opacity:.48; transition:opacity .25s ease, border-color .25s ease, background-color .25s ease;
  }
  .ln:hover { opacity:.78; }
  .ln:focus-visible { outline:2px solid var(--accent); outline-offset:3px; opacity:1; }
  .ln .who { display:block; font-size:.7rem; letter-spacing:.1em; font-weight:700; margin-bottom:.15rem; }
  .ln[data-who="ハル"] .who { color:var(--accent); }
  .ln[data-who="ソラ"] .who { color:var(--accent-2); }
  .ln.on { opacity:1; background:var(--rail); }
  .ln.on[data-who="ハル"] { border-left-color:var(--accent); }
  .ln.on[data-who="ソラ"] { border-left-color:var(--accent-2); }
  .ln.done { opacity:.3; }

  .endcap { padding:1.75rem 0 0; }
  .endcap p { font-size:.82rem; color:var(--muted); margin:0 0 .9rem; line-height:1.85; }

  /* ---- transport ---- */
  .bar {
    position:fixed; inset:auto 0 0 0; background:var(--bg);
    box-shadow:var(--shadow); padding:.85rem 0 max(.85rem, env(safe-area-inset-bottom));
  }
  .track { height:3px; background:var(--line); border-radius:2px; overflow:hidden; }
  .fill { height:100%; width:0%; background:var(--accent); transition:width .3s ease; }
  .row { display:flex; align-items:center; gap:.7rem; padding-top:.8rem; }
  .count {
    font-size:.74rem; color:var(--muted); font-variant-numeric:tabular-nums;
    letter-spacing:.04em; line-height:1.5; min-width:5.5rem;
  }
  .count b { display:block; color:var(--text); font-weight:600; }
  .spacer { flex:1; }

  button.ctl, select.ctl {
    appearance:none; background:none; border:1px solid var(--line); color:var(--text);
    border-radius:3px; cursor:pointer; font:inherit; font-size:.8rem; line-height:1;
    padding:.45rem .7rem; transition:border-color .15s ease, color .15s ease;
  }
  button.ctl:hover { border-color:var(--accent); color:var(--accent); }
  button.ctl:focus-visible, select.ctl:focus-visible { outline:2px solid var(--accent); outline-offset:2px; }
  button.ctl[disabled] { opacity:.35; cursor:not-allowed; }
  select.ctl {
    padding:.42rem 1.5rem .42rem .6rem;
    background-image:linear-gradient(45deg,transparent 50%,currentColor 50%),
                     linear-gradient(135deg,currentColor 50%,transparent 50%);
    background-position:calc(100% - 12px) 52%, calc(100% - 8px) 52%;
    background-size:4px 4px,4px 4px; background-repeat:no-repeat;
  }

  button.play {
    appearance:none; border:0; cursor:pointer; background:var(--accent); color:var(--bg);
    width:3rem; height:3rem; border-radius:50%; flex:none; display:grid; place-items:center;
  }
  button.play:focus-visible { outline:2px solid var(--accent); outline-offset:3px; }
  button.play[disabled] { opacity:.4; cursor:not-allowed; }
  button.play svg { width:1.1rem; height:1.1rem; fill:currentColor; }

  .status { font-size:.72rem; color:var(--muted); padding-top:.6rem; line-height:1.6; }

  footer {
    border-top:1px solid var(--line); margin-top:2.5rem; padding:1.5rem 0 2rem;
    font-size:.78rem; color:var(--muted); line-height:1.9;
  }
  footer code { font-family:ui-monospace,"SFMono-Regular",Menlo,monospace; font-size:.95em; }

  @media (prefers-reduced-motion: reduce) {
    * { transition:none !important; scroll-behavior:auto !important; }
  }
</style>

<div class="wrap">
  <header>
    <p class="eyebrow">Web Marketing Study — 全10話 / __TOTAL__ライン</p>
    <h1>広告運用 × LINE公式アカウント<br>音声で回す2週間</h1>
    <p class="lede">
      2週間集中プランの Day&nbsp;1〜10 と1対1で対応した対話台本を、
      端末の読み上げ機能で再生します。台本は <code>_learn/podcast/</code>。
    </p>
    <p class="notice">
      <b>音声ファイルではありません。</b>ブラウザの読み上げで喋らせているため、
      音質は端末のエンジン次第です。2話者に別々の日本語音声が自動で割り当たります
      （声が1つしかない端末では話速と高さで区別）。<b>iOSは画面ロックで停止します。</b>
    </p>
  </header>

  <nav class="idx">
    <h2>エピソード</h2>
    <div id="index"></div>
  </nav>

  <section class="now">
    <p class="eyebrow" id="nowNum">Episode 01</p>
    <h2 id="nowTitle">–</h2>
    <p class="lede" id="nowMeta">–</p>
    <div class="cast">
      <div><span class="dot a"></span><b>ハル</b>解説役</div>
      <div><span class="dot b"></span><b>ソラ</b>聞き手・新人</div>
    </div>
    <div class="script" id="script"></div>
    <div class="endcap" id="endcap" hidden>
      <p id="endcapText"></p>
      <button class="ctl" id="nextEp" type="button">次のエピソードへ →</button>
    </div>
  </section>

  <footer>
    行をタップするとそこから再生します。聴いた位置はこの端末に保存され、次に開いたとき続きから始まります。<br>
    自然な音声で聴きたいときは <code>_learn/podcast/README.md</code> の手順で NotebookLM へ。
    数値は2026年8月時点です。
  </footer>
</div>

<div class="bar">
  <div class="wrap">
    <div class="track"><div class="fill" id="fill"></div></div>
    <div class="row">
      <button class="play" id="play" aria-label="再生">
        <svg id="icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>
      </button>
      <span class="count"><b id="countEp">EP01</b><span id="countLn">– / –</span></span>
      <span class="spacer"></span>
      <button class="ctl" id="prev" aria-label="前の行">◀︎</button>
      <button class="ctl" id="next" aria-label="次の行">▶︎</button>
      <select class="ctl" id="rate" aria-label="再生速度">
        <option value="0.9">0.9×</option>
        <option value="1" selected>1.0×</option>
        <option value="1.2">1.2×</option>
        <option value="1.4">1.4×</option>
        <option value="1.6">1.6×</option>
      </select>
    </div>
    <p class="status" id="status">読み上げの準備をしています…</p>
  </div>
</div>

<script>
(function () {
  "use strict";

  var EPISODES = __EPISODES__;

  var PLAY  = '<path d="M8 5v14l11-7z"/>';
  var PAUSE = '<path d="M6 5h4v14H6zM14 5h4v14h-4z"/>';
  var STORE = "wmstudy.podcast.v1";

  var synth = window.speechSynthesis;
  var el = {
    index:   document.getElementById("index"),
    script:  document.getElementById("script"),
    nowNum:  document.getElementById("nowNum"),
    nowTitle:document.getElementById("nowTitle"),
    nowMeta: document.getElementById("nowMeta"),
    endcap:  document.getElementById("endcap"),
    endcapText: document.getElementById("endcapText"),
    nextEp:  document.getElementById("nextEp"),
    play:    document.getElementById("play"),
    icon:    document.getElementById("icon"),
    prev:    document.getElementById("prev"),
    next:    document.getElementById("next"),
    rate:    document.getElementById("rate"),
    countEp: document.getElementById("countEp"),
    countLn: document.getElementById("countLn"),
    fill:    document.getElementById("fill"),
    status:  document.getElementById("status")
  };

  var epIdx = 0;
  var idx = 0;
  var playing = false;
  var lineNodes = [];
  var epNodes = [];
  var voices = { "ハル": null, "ソラ": null };
  var singleVoice = false;
  var keepAlive = null;
  var booted = false;   // 初回描画では台本まで自動スクロールしない

  /* ---- persistence (sandbox may block storage; never let it break playback) ---- */
  function loadState() {
    try {
      var raw = window.localStorage.getItem(STORE);
      return raw ? JSON.parse(raw) : {};
    } catch (e) { return {}; }
  }
  function saveState(patch) {
    try {
      var s = loadState();
      for (var k in patch) { if (Object.prototype.hasOwnProperty.call(patch, k)) s[k] = patch[k]; }
      window.localStorage.setItem(STORE, JSON.stringify(s));
    } catch (e) { /* storage unavailable — session-only, that's fine */ }
  }
  function savePos() {
    var s = loadState();
    var pos = s.pos || {};
    pos[EPISODES[epIdx].id] = idx;
    saveState({ pos: pos, last: EPISODES[epIdx].id });
  }

  /* ---- episode index ---- */
  function buildIndex() {
    EPISODES.forEach(function (ep, i) {
      var b = document.createElement("button");
      b.className = "ep";
      b.type = "button";
      b.innerHTML =
        '<span class="ep-n">EP' + ep.num + '</span>' +
        '<span class="ep-body"><span class="ep-t"></span><span class="ep-m"></span></span>' +
        '<span class="ep-s"></span>';
      b.querySelector(".ep-t").textContent = ep.title;
      b.querySelector(".ep-m").textContent = ep.day + " ／ " + ep.length;
      b.addEventListener("click", function () { openEpisode(i, null); });
      el.index.appendChild(b);
      epNodes.push(b);
    });
  }

  function paintIndex() {
    var pos = loadState().pos || {};
    EPISODES.forEach(function (ep, i) {
      var node = epNodes[i];
      var badge = node.querySelector(".ep-s");
      var at = pos[ep.id];
      node.setAttribute("aria-current", i === epIdx ? "true" : "false");
      badge.classList.remove("part");
      if (at === undefined || at === null) {
        badge.textContent = ep.lines.length + "行";
      } else if (at >= ep.lines.length - 1) {
        badge.textContent = "聴了";
      } else {
        badge.textContent = (at + 1) + " / " + ep.lines.length;
        badge.classList.add("part");
      }
    });
  }

  /* ---- episode rendering ---- */
  function openEpisode(i, startAt) {
    synth.cancel();
    var wasPlaying = playing;
    epIdx = i;
    var ep = EPISODES[i];

    el.nowNum.textContent = "Episode " + ep.num;
    el.nowTitle.textContent = ep.title;
    el.nowMeta.textContent = ep.day + " ／ " + ep.length + " ／ 全" + ep.lines.length + "ライン";
    el.countEp.textContent = "EP" + ep.num;
    el.endcap.hidden = true;

    el.script.textContent = "";
    lineNodes = [];
    ep.lines.forEach(function (row, li) {
      var b = document.createElement("button");
      b.className = "ln";
      b.type = "button";
      b.setAttribute("data-who", row[0]);
      var who = document.createElement("span");
      who.className = "who";
      who.textContent = row[0];
      var txt = document.createElement("span");
      txt.textContent = row[1];
      b.appendChild(who);
      b.appendChild(txt);
      b.addEventListener("click", function () { jump(li); });
      el.script.appendChild(b);
      lineNodes.push(b);
    });

    if (startAt === null || startAt === undefined) {
      var saved = (loadState().pos || {})[ep.id];
      idx = (saved !== undefined && saved !== null && saved < ep.lines.length - 1) ? saved : 0;
    } else {
      idx = startAt;
    }

    paint();
    paintIndex();
    if (wasPlaying) { speak(idx); }
    else if (booted) { el.script.scrollIntoView({ block: "start", behavior: "smooth" }); }
    booted = true;
  }

  function paint() {
    var ep = EPISODES[epIdx];
    lineNodes.forEach(function (n, i) {
      n.classList.toggle("on", i === idx && playing);
      n.classList.toggle("done", i < idx);
    });
    el.countLn.textContent = (idx + 1) + " / " + ep.lines.length;
    el.fill.style.width = (idx / ep.lines.length * 100).toFixed(1) + "%";
    el.prev.disabled = idx === 0;
    el.next.disabled = idx >= ep.lines.length - 1;
  }

  /* ---- voices ---- */
  function pickVoices() {
    var ja = (synth.getVoices() || []).filter(function (v) { return /^ja/i.test(v.lang); });
    if (!ja.length) {
      el.status.textContent = "この端末には日本語の音声が入っていないため、読み上げできません。台本はそのまま読めます。";
      el.play.disabled = true;
      return false;
    }
    voices["ハル"] = ja[0];
    voices["ソラ"] = ja.length > 1 ? ja[1] : ja[0];
    singleVoice = ja.length < 2;
    el.status.textContent = singleVoice
      ? "日本語の音声が1つ（" + ja[0].name + "）のため、話速と高さで2人を区別します。"
      : "ハル：" + ja[0].name + " ／ ソラ：" + ja[1].name;
    el.play.disabled = false;
    return true;
  }

  /* ---- playback ---- */
  function speak(i) {
    var ep = EPISODES[epIdx];
    if (i >= ep.lines.length) { finish(); return; }
    idx = i;
    paint();
    savePos();
    lineNodes[i].scrollIntoView({ block: "center", behavior: "smooth" });

    var row = ep.lines[i];
    var u = new SpeechSynthesisUtterance(row[2]);
    u.lang = "ja-JP";
    if (voices[row[0]]) u.voice = voices[row[0]];
    var base = parseFloat(el.rate.value);
    if (singleVoice && row[0] === "ソラ") { u.pitch = 1.25; u.rate = base * 1.04; }
    else { u.pitch = 1; u.rate = base; }
    u.onend = function () { if (playing) speak(i + 1); };
    u.onerror = function (e) {
      if (e.error === "interrupted" || e.error === "canceled") return;
      el.status.textContent = "読み上げが中断されました（" + e.error + "）。再生し直してください。";
      stop();
    };
    synth.speak(u);
  }

  function start() {
    if (!voices["ハル"] && !pickVoices()) return;
    playing = true;
    el.icon.innerHTML = PAUSE;
    el.play.setAttribute("aria-label", "一時停止");
    el.endcap.hidden = true;
    clearInterval(keepAlive);
    // Chrome stalls long utterance queues; nudge it while playing.
    keepAlive = setInterval(function () {
      if (playing && synth.speaking && !synth.paused) { synth.pause(); synth.resume(); }
    }, 9000);
    speak(idx);
  }

  function stop() {
    playing = false;
    clearInterval(keepAlive);
    synth.cancel();
    el.icon.innerHTML = PLAY;
    el.play.setAttribute("aria-label", "再生");
    paint();
    paintIndex();
  }

  function finish() {
    var ep = EPISODES[epIdx];
    idx = ep.lines.length - 1;
    savePos();
    stop();
    el.fill.style.width = "100%";
    var hasNext = epIdx < EPISODES.length - 1;
    el.endcapText.textContent = hasNext
      ? "EP" + ep.num + " おつかれさまでした。宿題を先に片づけてから、次へ。"
      : "全10話おつかれさまでした。あとは週末に、提案を1本ゼロから作るだけです。";
    el.nextEp.hidden = !hasNext;
    if (hasNext) el.nextEp.textContent = "次は EP" + EPISODES[epIdx + 1].num + "  " + EPISODES[epIdx + 1].title;
    el.endcap.hidden = false;
    el.endcap.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  function jump(i) {
    var wasPlaying = playing;
    synth.cancel();
    idx = i;
    savePos();
    if (wasPlaying) { speak(i); } else { paint(); paintIndex(); }
  }

  /* ---- wiring ---- */
  el.play.addEventListener("click", function () { playing ? stop() : start(); });
  el.prev.addEventListener("click", function () { jump(Math.max(0, idx - 1)); });
  el.next.addEventListener("click", function () { jump(Math.min(EPISODES[epIdx].lines.length - 1, idx + 1)); });
  el.rate.addEventListener("change", function () { if (playing) jump(idx); });
  el.nextEp.addEventListener("click", function () {
    if (epIdx < EPISODES.length - 1) openEpisode(epIdx + 1, 0);
  });

  document.addEventListener("keydown", function (e) {
    var t = e.target.tagName;
    if (t === "SELECT" || t === "BUTTON") return;
    if (e.code === "Space") { e.preventDefault(); el.play.click(); }
  });

  window.addEventListener("beforeunload", function () { synth.cancel(); });

  buildIndex();

  var last = loadState().last;
  var startIdx = 0;
  EPISODES.forEach(function (ep, i) { if (ep.id === last) startIdx = i; });
  openEpisode(startIdx, null);

  if (!("speechSynthesis" in window)) {
    el.status.textContent = "このブラウザは読み上げに対応していません。台本はそのまま読めます。";
    el.play.disabled = true;
  } else {
    pickVoices();
    synth.addEventListener("voiceschanged", function () { if (!voices["ハル"]) pickVoices(); });
  }
})();
</script>
"""


if __name__ == "__main__":
    sys.exit(main())
