// デモ画面の本体：上の操作帯（ユーザー・時間・再生・表示）と、スマホ・裏側・資料用表示の切り替え。
// React を使わず素の DOM で描く。編集ツールのデモ画面と、Phase 5 の営業用HTMLの両方がこれをそのまま使う。
//
// 決まりごと
//   - 押した結果を画面側で決めない。すべてエンジンへの入力イベントにする（controller.ts）
//   - 表示は毎回「イベント列 → reduce → 状態」から作り直す
//   - 案件の文字は必ず esc() を通す

import type { Asset, Id, InputEventBody, Project, UserState } from '../../engine/src/index.ts';
import { MINUTES_PER_DAY } from '../../engine/src/index.ts';
import { backstage } from './backstage.ts';
import {
  addUser,
  allScripted,
  autoStep,
  dispatch,
  futureEvents,
  initialDemo,
  newDemo,
  nextStop,
  pendingInputs,
  prevStop,
  removeUser,
  resetDemo,
  seek,
  selectUser,
  worldOf,
  type Demo,
} from './controller.ts';
import { fmtTime, Labels } from './describe.ts';
import { fcDialog, journeyPane } from './journey.ts';
import { phone } from './phone.ts';
import { guideOf, summaryAt } from './timeline.ts';
import { SLIDE_H, SLIDE_W, slide, slide3 } from './slide.ts';
import { esc, pane, whoOf, type Mode, type Tab, type Ui, type View } from './view.ts';

export { esc } from './view.ts';

export type PlayerOptions = {
  project: Project;
  /** 続きから始めるとき（編集ツールで画面を切り替えて戻ってきたとき等） */
  demo?: Demo;
  personaId?: Id;
  /** 画像ファイルの URL（編集ツールは保存APIのURL、営業用HTMLは埋め込みデータ） */
  assetUrl?: (asset: Asset) => string;
  onChange?: (demo: Demo) => void;
};

export type PlayerHandle = {
  setProject(project: Project): void;
  reset(personaId?: Id): void;
  getDemo(): Demo;
  destroy(): void;
};

/** 自動再生で1コマ進める間隔（×1 のとき） */
const TICK_MS = 1200;

export function createPlayer(root: HTMLElement, opts: PlayerOptions): PlayerHandle {
  let project = opts.project;
  let demo = opts.demo ?? (opts.personaId ? newDemo(project, opts.personaId) : initialDemo(project));
  let view: View | null = null;
  const ui: Ui = { menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false };
  let timer: ReturnType<typeof setInterval> | null = null;
  /** 回答待ちで止まった再生を、答えたら続けるか */
  let resumeWhenAnswered = false;
  const assetUrl = opts.assetUrl ?? ((a: Asset) => a.path);

  root.classList.add('lsp');

  // ---------- 状態の更新 ----------

  const commit = (next: Demo, keepOpen = false) => {
    demo = next;
    if (!keepOpen) ui.open.clear();
    opts.onChange?.(demo);
    render();
  };

  const send = (body: InputEventBody) => {
    if (!demo.activeUserId) return;
    commit(dispatch(demo, body));
    // 回答待ちで止まっていた再生は、答えて待ちが無くなったら続ける
    if (resumeWhenAnswered && view && pendingInputs(project, view.world, demo).length === 0) startPlay();
  };

  /** 時間を動かす操作。再生中なら止める。tell なら、その時刻に誰に何が起きたかを知らせる（①②の人は切り替えない） */
  const moveTime = (to: number, tell = false) => {
    stopTimer();
    ui.playing = false;
    ui.status = '';
    resumeWhenAnswered = false;
    demo = seek(demo, to);
    if (tell) ui.status = summaryAt({ project, world: worldOf(project, demo), demo, ui, L: new Labels(project), assetUrl }, to);
    commit(demo);
  };

  // ---------- 自動再生 ----------

  function stopTimer() {
    if (timer) clearInterval(timer);
    timer = null;
  }

  function startPlay() {
    stopTimer();
    ui.playing = true;
    ui.status = '再生中';
    resumeWhenAnswered = false;
    timer = setInterval(tick, TICK_MS / ui.speed);
    render();
  }

  function pause(status: string) {
    stopTimer();
    ui.playing = false;
    ui.status = status;
    render();
  }

  function tick() {
    const r = autoStep(project, demo);
    if (r.stopped === 'pending' && r.pending) {
      // 回答を待っている人に①②を勝手に切り替えない。選ぶのは営業担当
      const name = whoOf(view?.world.users[r.pending.userId]);
      const what = r.pending.kind === 'survey' ? 'アンケートの回答' : 'ボタン選択';
      resumeWhenAnswered = true;
      pause(
        r.pending.userId === demo.activeUserId
          ? `${name}の${what}を待っています。①で答えると再生を続けます`
          : `${name}の${what}を待っています。上のユーザー一覧で「${name}」を選んで答えると、再生を続けます`,
      );
      return;
    }
    commit(r.demo);
    if (r.stopped === 'end') pause('最後まで再生しました');
  }

  // ---------- 描画 ----------

  function bar(v: View): string {
    const active = demo.activeUserId;
    const pending = new Set(pendingInputs(project, v.world, demo).map((p) => p.userId));
    const chips = demo.session.users
      .map((u) => {
        const s = v.world.users[u.id];
        return `<button type="button" class="lsp-chip${u.id === active ? ' on' : ''}" data-act="selectUser" data-user="${esc(u.id)}">${esc(s ? whoOf(s) : u.id)}${
          demo.scripted.includes(u.id) ? '<small>台本</small>' : ''
        }${pending.has(u.id) ? '<small class="wait">回答待ち</small>' : ''}</button>`;
      })
      .join('');
    const label = (p: (typeof project.personas)[number]) => whoOf({ name: p.name, ...(p.personName ? { personName: p.personName } : {}) });
    const manual = project.personas.map((p) => `<option value="m:${esc(p.id)}">${esc(label(p))}</option>`).join('');
    const scripted = project.personas
      .filter((p) => (p.script ?? []).length)
      .map((p) => `<option value="s:${esc(p.id)}">${esc(label(p))}</option>`)
      .join('');
    const day = Math.floor(v.world.now / MINUTES_PER_DAY);
    const time = fmtTime(v.world.now).split(' ')[2] ?? '00:00';
    const ahead = active && !demo.scripted.includes(active) ? futureEvents(demo, active) : 0;
    const mode = (m: Mode, text: string) => `<button type="button" class="${ui.mode === m ? 'on' : ''}" data-act="mode" data-mode="${m}">${text}</button>`;
    const speed = (s: 1 | 2 | 5) => `<button type="button" class="${ui.speed === s ? 'on' : ''}" data-act="speed" data-speed="${s}">×${s}</button>`;
    const me = v.world.users[active];
    // いま誰を見ているか・どう動くか・再生しているか（常に見える）
    const now = me
      ? `<span class="lsp-now-who"><span class="k">表示中</span><b>${esc(whoOf(me))}</b><span class="lsp-now-mode ${demo.scripted.includes(active) ? 'script' : 'manual'}">${demo.scripted.includes(active) ? '台本どおりに動く' : '手で操作'}</span><span class="lsp-now-mode ${ui.playing ? 'play' : ''}">${ui.playing ? `再生中 ×${ui.speed}` : '停止中'}</span></span>`
      : '';
    // 細かい操作（時刻の調整・速さ・初期化・ユーザーの追加・ほかの表示）は、開いたときだけ出す
    const more = ui.more
      ? `<div class="lsp-bar-more">
        <div class="lsp-bar-row"><span class="lbl">時刻を細かく動かす</span><span class="lsp-group"><button type="button" data-act="adv" data-min="-1440">−1日</button><button type="button" data-act="adv" data-min="-60">−1時間</button><button type="button" data-act="adv" data-min="60">＋1時間</button><button type="button" data-act="adv" data-min="1440">＋1日</button></span>
          <span class="lsp-group">Day <input type="number" min="0" max="365" value="${day}" data-jump="day" aria-label="日"><input type="text" value="${esc(time)}" data-jump="time" size="5" aria-label="時刻"><button type="button" data-act="jump">移動</button></span>
          <span class="lbl">再生の速さ</span><span class="lsp-group">${speed(1)}${speed(2)}${speed(5)}</span>
          <button type="button" data-act="reset" title="顔ぶれと台本は残し、手の操作を消して最初の時刻に戻す">最初に戻す</button></div>
        <div class="lsp-bar-row"><span class="lbl">ユーザー</span>
          <select data-change="addUser" aria-label="ユーザーを追加"><option value="">＋ ユーザーを追加…</option>${manual ? `<optgroup label="手で操作する">${manual}</optgroup>` : ''}${
            scripted ? `<optgroup label="台本どおりに動く">${scripted}</optgroup>` : `<option value="m:">（ペルソナなし）</option>`
          }</select>
          ${active ? '<button type="button" data-act="removeUser" title="選んでいるユーザーを外す">表示中の人を外す</button>' : ''}
          <span class="lbl">ほかの表示</span><span class="lsp-group">${mode('full', '画面＋裏側')}${mode('phone', '画面のみ')}${mode('slide', '資料用 16:9（2画面）')}</span></div>
      </div>`
      : '';
    // 資料用16:9（投影）では帯を1段に詰め、16:9の枠を大きく出す。ユーザーはプルダウンで選ぶ
    const slideMode = ui.mode === 'slide' || ui.mode === 'slide3';
    const picker = `<select class="lsp-userpick" data-change="selectUser" aria-label="表示するユーザー">${demo.session.users
      .map((u) => `<option value="${esc(u.id)}"${u.id === active ? ' selected' : ''}>${esc(whoOf(v.world.users[u.id]) || u.id)}</option>`)
      .join('')}</select>`;
    return `<div class="lsp-bar${slideMode ? ' compact' : ''}">
      ${
        slideMode
          ? ''
          : `<div class="lsp-bar-sec">
        <div class="lsp-cap">説明するお客様を選びます。「台本つきで全員を並べる」と、用意した全員が台本どおりに動きます。</div>
        <div class="lsp-bar-row">${chips || '<span class="lsp-none">ユーザーがいません（細かい操作から追加できます）</span>'}
          ${scripted ? '<button type="button" class="lsp-all" data-act="allScripted" title="台本のあるペルソナ全員を、台本どおりに動く設定で並べ直す">台本つきで全員を並べる</button>' : ''}</div>
      </div>`
      }
      <div class="lsp-bar-sec">
        <div class="lsp-cap">「次の出来事」で1つずつ進めます。「戻る」で1つ前へ、「再生」で最後まで自動で流れます。</div>
        <div class="lsp-bar-row main">
          <span class="lsp-clock">${esc(fmtTime(v.world.now))}</span>
          <button type="button" class="lsp-back-btn" data-act="prev" title="前に何かが起きた時点へ戻る">◀ 戻る</button>
          <button type="button" class="primary lsp-next-btn" data-act="next" title="次に誰かに何かが起きる時点へ進む">次の出来事 ▶</button>
          <button type="button" class="lsp-play-btn" data-act="play">${ui.playing ? '⏸ 停止' : '▶ 再生'}</button>
          ${slideMode && demo.session.users.length ? picker : now}
          <span class="lsp-spacer"></span>
          <span class="lsp-group lsp-view">${mode('three', '3画面')}${mode('slide3', '16:9（投影用）')}${
            ui.mode === 'full' ? mode('full', '画面＋裏側') : ui.mode === 'phone' ? mode('phone', '画面のみ') : ui.mode === 'slide' ? mode('slide', '資料用 16:9（2画面）') : ''
          }</span>
          <button type="button" class="ghost lsp-more-btn" data-act="more" aria-expanded="${!!ui.more}">${ui.more ? '細かい操作を閉じる ▲' : '細かい操作 ▼'}</button>
        </div>
        ${ui.status ? `<div class="lsp-playstate" role="status">${esc(ui.status)}</div>` : ''}
      </div>
      ${more}
      ${ahead ? `<div class="lsp-warn">この先に、${esc(whoOf(v.world.users[active]))}の操作が${ahead}件あります。いまの時点から操作すると、それより後の操作は消えます。</div>` : ''}
    </div>`;
  }

  function render() {
    let world;
    try {
      world = worldOf(project, demo);
    } catch (e) {
      root.innerHTML = `<div class="lsp-error">デモを表示できません：${esc(e instanceof Error ? e.message : e)}</div>`;
      return;
    }
    view = { project, world, demo, ui, L: new Labels(project), assetUrl };
    const c = project.brand.colors;
    root.style.setProperty('--lsp-primary', c.primary);
    root.style.setProperty('--lsp-accent', c.accent);
    if (c.chatBg) root.style.setProperty('--lsp-chat', c.chatBg);
    else root.style.removeProperty('--lsp-chat');

    const user = world.users[demo.activeUserId];
    let body: string;
    if (!user) body = '<div class="lsp-empty">「＋ ユーザーを追加」から、デモのユーザーを追加してください。</div>';
    else if (ui.mode === 'slide') body = slide(view, user);
    else if (ui.mode === 'slide3') body = slide3(view, user);
    else if (ui.mode === 'three')
      body = `<div class="lsp-stage three">${pane(1, 'ユーザーが見る画面', phone(view, user))}${pane(2, '裏側の処理', backstage(view, user))}${pane(3, 'CVまでの道筋', journeyPane(view, user))}</div>`;
    else body = `<div class="lsp-stage ${ui.mode}">${phone(view, user)}${ui.mode === 'full' ? backstage(view, user) : ''}</div>`;
    root.innerHTML = bar(view) + body + (user ? fcDialog(view, user) : '');
    root.style.setProperty('--lsp-bar-h', `${root.querySelector<HTMLElement>('.lsp-bar')?.offsetHeight ?? 0}px`);

    const scroller = root.querySelector<HTMLElement>('.lsp-chat');
    if (scroller) scroller.scrollTop = keepChat ?? scroller.scrollHeight;
    keepChat = null;
    fitSlide();
    if (ui.mode === 'slide' || ui.mode === 'slide3') showSlide();
    fitChart();
    centerHere();
    fitTimeline();
    showNow(tlScroll);
    tlScroll = null;
    if (user) highlight(view, user);
  }

  /** カルーセルを送ったときは、トークを一番下まで送らずに元の位置のままにする */
  let keepChat: number | null = null;

  /** 3画面の③：ページを下まで送ったとき、上に残る現在地カードの下に時系列が収まる高さにする */
  function fitTimeline() {
    const list = root.querySelector<HTMLElement>('.lsp-stage.three .lsp-tl-scroll');
    const tl = list?.closest<HTMLElement>('.lsp-tl');
    const where = root.querySelector<HTMLElement>('.lsp-stage.three .lsp-where');
    if (!list || !tl || !where) return;
    let sc: HTMLElement | null = root.parentElement;
    while (sc && !(/(auto|scroll)/.test(getComputedStyle(sc).overflowY) && sc.clientHeight < sc.scrollHeight)) sc = sc.parentElement;
    const viewH = sc ? sc.clientHeight : window.innerHeight;
    const barH = root.querySelector<HTMLElement>('.lsp-bar')?.offsetHeight ?? 0;
    const head = tl.offsetHeight - list.offsetHeight;
    list.style.maxHeight = `${Math.max(240, viewH - barH - where.offsetHeight - head - 28)}px`;
  }

  /** ③の時系列を「いま」の線が見える位置へ（行を開いた・絞り込んだときは位置を保つ） */
  let tlScroll: number | null = null;
  function showNow(keep: number | null) {
    const box = root.querySelector<HTMLElement>('.lsp-tl-scroll');
    const line = box?.querySelector<HTMLElement>('#lsp-tl-now');
    if (!box) return;
    if (keep !== null) {
      box.scrollTop = keep;
      return;
    }
    if (!line) return;
    const top = line.getBoundingClientRect().top - box.getBoundingClientRect().top + box.scrollTop;
    box.scrollTop = Math.max(0, top - box.clientHeight * 0.6);
  }

  /** 「次にすること」が指している①のボタン等を光らせる */
  function highlight(v: View, user: UserState) {
    const target = guideOf(v, user).target;
    if (!target) return;
    for (const el of root.querySelectorAll<HTMLElement>(`.lsp-phone ${target}`)) el.classList.add('lsp-hl');
  }

  /** 3画面の③で、「現在ここ」の箱が見える位置までスクロールする */
  function centerHere() {
    const box = root.querySelector<HTMLElement>('.lsp-stage.three .lsp-fc-scroll');
    const here = box?.querySelector<HTMLElement>('.lsp-fc-here');
    if (!box || !here) return;
    const top = here.getBoundingClientRect().top - box.getBoundingClientRect().top + box.scrollTop;
    box.scrollTop = Math.max(0, top - box.clientHeight / 3);
  }

  /** ③のフローチャートを、置き場所に収まる拡大率にする（資料用は常に、通常表示は「幅に合わせる」のとき） */
  function fitChart(widthOnly = false): number | null {
    const box = root.querySelector<HTMLElement>(widthOnly ? '.lsp-fc-scroll' : '.lsp-journey.compact .lsp-fc-scroll');
    const chart = box?.querySelector<HTMLElement>('.lsp-fc');
    if (!box || !chart) return null;
    chart.style.zoom = '1';
    const kw = box.clientWidth / Math.max(1, chart.scrollWidth);
    const kh = widthOnly ? 1 : box.clientHeight / Math.max(1, chart.scrollHeight);
    const k = Math.max(0.4, Math.min(1, kw, kh));
    chart.style.zoom = String(widthOnly ? Math.min(1.5, Math.max(0.4, kw)) : k);
    return widthOnly ? Math.min(1.5, Math.max(0.4, kw)) : k;
  }

  /** 縦にスクロールする入れ物（編集ツールの中なら作業領域、営業用HTMLならページ） */
  function scrollerOf(): HTMLElement | null {
    let sc: HTMLElement | null = root.parentElement;
    while (sc && !/(auto|scroll)/.test(getComputedStyle(sc).overflowY)) sc = sc.parentElement;
    return sc;
  }

  /**
   * 資料用表示（16:9）を、上の帯を除いた見える範囲に1画面で収める。幅と高さの小さい方に合わせ、
   * 大きな画面（1920×1080 等）では拡大する（最大1.5倍）。投影したときにスクロールしなくてよいように。
   */
  function fitSlide() {
    const wrap = root.querySelector<HTMLElement>('.lsp-slide-wrap');
    const canvas = root.querySelector<HTMLElement>('.lsp-slide');
    if (!wrap || !canvas) return;
    const sc = scrollerOf();
    const viewH = sc ? sc.clientHeight : window.innerHeight;
    const barH = root.querySelector<HTMLElement>('.lsp-bar')?.offsetHeight ?? 0;
    const k = Math.max(0.3, Math.min(1.5, wrap.clientWidth / SLIDE_W, (viewH - barH - 20) / SLIDE_H));
    canvas.style.transform = `scale(${k})`;
    canvas.style.left = `${Math.max(0, Math.round((wrap.clientWidth - SLIDE_W * k) / 2))}px`;
    wrap.style.height = `${Math.round(SLIDE_H * k)}px`;
  }

  /** 資料用表示では、上の帯のすぐ下に16:9の枠が来るようにする（編集ツールの見出しの分だけずれないように） */
  function showSlide() {
    const wrap = root.querySelector<HTMLElement>('.lsp-slide-wrap');
    const sc = scrollerOf();
    if (!wrap) return;
    const barH = root.querySelector<HTMLElement>('.lsp-bar')?.offsetHeight ?? 0;
    const top = wrap.getBoundingClientRect().top - (sc ? sc.getBoundingClientRect().top : 0);
    const by = top - barH - 10;
    if (Math.abs(by) > 2) (sc ?? document.scrollingElement ?? document.documentElement).scrollTop += by;
  }
  const resize = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(() => fitSlide());
  resize?.observe(root);

  // ---------- 操作 ----------

  function onClick(e: Event) {
    const el = (e.target as HTMLElement).closest<HTMLElement>('[data-act]');
    if (!el || !root.contains(el) || (el as HTMLButtonElement).disabled) return;
    const d = el.dataset;
    const pageId = d['page'] ?? '';
    const world = view?.world;
    const user = world?.users[demo.activeUserId];
    switch (d['act']) {
      // スマホの中
      case 'visit':
        return send({ type: 'visitPage', pageId });
      case 'addFriend':
        return send({ type: 'addFriend', routeId: d['route'] ?? '' });
      case 'back': {
        // 友だちでない人が、離脱防止ポップアップのあるページを初めて離れるときだけ「離脱」になる
        const p = project.webPages.find((x) => x.id === pageId);
        const dismissed = user?.behaviors.some((b) => b.type === 'dismissPopup' && b.detail['pageId'] === pageId);
        if (p?.exitPopup && user?.friendState !== 'friend' && !dismissed) return send({ type: 'exitIntent', pageId });
        return send({ type: 'closePage', pageId });
      }
      case 'dismissPopup':
        return send({ type: 'dismissPopup', pageId });
      case 'popupCta':
        return send({ type: 'tapPopupCta', pageId });
      case 'pageBtn':
        return send({ type: 'tapPageButton', pageId, blockIndex: Number(d['i']) });
      case 'listItem':
        return send({ type: 'tapPageButton', pageId, blockIndex: Number(d['i']), itemId: d['item'] ?? '' });
      case 'lineCta':
        return send({ type: 'tapLineCta', pageId });
      case 'tapButton':
        return send({ type: 'tapButton', buttonId: d['button'] ?? '', ...(d['node'] ? { nodeId: d['node'] } : {}) });
      case 'tapMenu':
        return send({ type: 'tapMenu', richMenuId: d['menu'] ?? '', cellId: d['cell'] ?? '' });
      case 'carGo':
        return viewCard(d['msg'] ?? '', d['card'] ?? '');
      case 'toggleMenu':
        ui.menuOpen = !ui.menuOpen;
        return render();
      case 'answer':
        return send({ type: 'answer', surveyId: d['survey'] ?? '', questionId: d['q'] ?? '', choiceIds: [d['choice'] ?? ''] });
      case 'pick': {
        const c = d['choice'] ?? '';
        ui.picked = ui.picked.includes(c) ? ui.picked.filter((x) => x !== c) : [...ui.picked, c];
        return render();
      }
      case 'answerMulti': {
        const choiceIds = ui.picked;
        ui.picked = [];
        return send({ type: 'answer', surveyId: d['survey'] ?? '', questionId: d['q'] ?? '', choiceIds });
      }
      case 'block':
        return send({ type: 'block' });
      case 'unblock':
        return send({ type: 'unblock' });
      case 'cv':
        return send({ type: 'cv', cvId: d['cv'] ?? '' });

      // ユーザー
      case 'selectUser':
        if (d['tab']) ui.tab = d['tab'] as Tab;
        ui.picked = [];
        return commit(selectUser(demo, d['user'] ?? ''));
      case 'removeUser':
        ui.picked = [];
        return commit(removeUser(demo, demo.activeUserId));
      case 'allScripted':
        stopTimer();
        ui.playing = false;
        ui.status = '台本どおりに動く人を並べました。「▶ 自動再生」で流せます';
        return commit(allScripted(project));

      // 時間
      case 'next': {
        const t = world ? nextStop(world, demo) : null;
        if (t !== null) moveTime(t, true);
        else pause('この先に起きることはありません');
        return;
      }
      case 'prev': {
        const t = world ? prevStop(world, demo) : null;
        if (t !== null) moveTime(t, true);
        return;
      }
      case 'adv':
        return moveTime(demo.now + Number(d['min'] ?? 0));
      case 'jump': {
        const day = Number(root.querySelector<HTMLInputElement>('[data-jump="day"]')?.value ?? 0);
        const m = /^(\d{1,2}):(\d{2})$/.exec(root.querySelector<HTMLInputElement>('[data-jump="time"]')?.value.trim() ?? '');
        if (!Number.isFinite(day) || !m) return pause('時刻は「09:00」の形で入れてください');
        return moveTime(Math.max(0, day) * MINUTES_PER_DAY + Number(m[1]) * 60 + Number(m[2]));
      }
      case 'reset':
        stopTimer();
        ui.playing = false;
        ui.status = '';
        resumeWhenAnswered = false;
        ui.picked = [];
        return commit(resetDemo(demo, project));

      // 再生
      case 'play':
        return ui.playing ? pause('一時停止しました') : startPlay();
      case 'speed':
        ui.speed = Number(d['speed']) as Ui['speed'];
        return ui.playing ? startPlay() : render();

      // 表示
      case 'mode':
        ui.mode = d['mode'] as Mode;
        return render();
      case 'tab':
        ui.tab = d['tab'] as Tab;
        return render();
      // ③フローチャート
      case 'fcOpen':
        ui.fcOpen = { step: d['step'] ?? '', ...(d['variant'] ? { variant: d['variant'] } : {}), ...(d['option'] ? { option: d['option'] } : {}) };
        return render();
      case 'fcClose':
        ui.fcOpen = null;
        return render();
      case 'fcStay':
        return;
      case 'fcZoom': {
        const z = d['z'];
        if (z === 'fit') ui.fcZoom = Math.round((fitChart(true) ?? 1) * 100) / 100;
        else ui.fcZoom = Math.min(1.5, Math.max(0.4, Math.round((ui.fcZoom + (z === 'in' ? 0.1 : -0.1)) * 10) / 10));
        return render();
      }
      case 'help':
        ui.help = !ui.help;
        return render();
      case 'more':
        ui.more = !ui.more;
        return render();
      // ③の時系列
      case 'tlOpen': {
        const key = d['key'] ?? '';
        const open = (ui.tlOpen ??= new Set());
        if (open.has(key)) open.delete(key);
        else open.add(key);
        tlScroll = root.querySelector<HTMLElement>('.lsp-tl-scroll')?.scrollTop ?? null;
        return render();
      }
      case 'tlJump':
        root.querySelector<HTMLElement>('.lsp-tl')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
      case 'tlFilter':
        ui.tlFilter = d['filter'] === 'me' ? 'me' : 'all';
        return render();
      case 'toggleWhy': {
        const i = Number(d['i']);
        if (ui.open.has(i)) ui.open.delete(i);
        else ui.open.add(i);
        return render();
      }
    }
  }

  function onSubmit(e: Event) {
    const form = e.target as HTMLFormElement;
    if (!root.contains(form) || !form.dataset['form']) return;
    e.preventDefault();
    if (form.dataset['form'] === 'send') {
      const text = String(new FormData(form).get('text') ?? '').trim();
      if (text) send({ type: 'sendText', text });
      return;
    }
    if (form.dataset['form'] === 'submit') {
      const values: Record<string, string> = {};
      new FormData(form).forEach((v, k) => (values[k] = String(v)));
      send({ type: 'submitForm', pageId: form.dataset['page'] ?? '', values });
    }
  }

  function onChange(e: Event) {
    const el = e.target as HTMLInputElement | HTMLSelectElement;
    if (!root.contains(el)) return;
    if (el.dataset['change'] === 'addUser' && el.value) {
      const [kind, personaId] = el.value.split(':');
      return commit(addUser(demo, project, personaId ?? '', kind === 's'));
    }
    if (el.dataset['change'] === 'fcOptions') {
      ui.fcOptions = (el as HTMLInputElement).checked;
      return render();
    }
    if (el.dataset['change'] === 'selectUser') return commit(selectUser(demo, el.value));
    if (el.dataset['change'] === 'filterList')
      return send({ type: 'filterList', pageId: el.dataset['page'] ?? '', blockIndex: Number(el.dataset['i']), facetId: el.dataset['facet'] ?? '', optionId: el.value });
    if (el.dataset['change'] === 'tlFuture' || el.dataset['change'] === 'tlMinor') {
      ui[el.dataset['change']] = (el as HTMLInputElement).checked;
      return render();
    }
    if (el.dataset['change'] === 'onlyNg') {
      ui.onlyNg = (el as HTMLInputElement).checked;
      return render();
    }
  }

  /** カルーセルのカードを切り替える。トークの位置は動かさない */
  function viewCard(messageId: string, cardId: string) {
    if (!messageId || !cardId) return;
    keepChat = root.querySelector<HTMLElement>('.lsp-chat')?.scrollTop ?? null;
    send({ type: 'viewCard', messageId, cardId });
  }

  /** カルーセルを指・マウスで左右に送る（40px 以上動かしたら1枚送る） */
  let drag: { x: number; el: HTMLElement; track: HTMLElement } | null = null;
  function onPointerDown(e: PointerEvent) {
    const el = (e.target as HTMLElement).closest<HTMLElement>('[data-drag="car"]');
    const track = el?.querySelector<HTMLElement>('.lsp-car-track');
    if (!el || !track || !root.contains(el) || (e.target as HTMLElement).closest('button')) return;
    drag = { x: e.clientX, el, track };
    el.setPointerCapture?.(e.pointerId);
  }
  function onPointerMove(e: PointerEvent) {
    if (!drag) return;
    drag.track.style.setProperty('--dx', `${e.clientX - drag.x}px`);
  }
  function onPointerUp(e: PointerEvent) {
    if (!drag) return;
    const { el, track, x } = drag;
    drag = null;
    track.style.removeProperty('--dx');
    const dx = e.clientX - x;
    if (Math.abs(dx) < 40) return;
    const btn = el.parentElement?.querySelector<HTMLButtonElement>(dx < 0 ? '.lsp-car-next' : '.lsp-car-prev');
    if (btn && !btn.disabled) viewCard(btn.dataset['msg'] ?? '', btn.dataset['card'] ?? '');
  }

  /** Esc で③の詳細を閉じる */
  function onKey(e: KeyboardEvent) {
    if (e.key === 'Escape' && ui.fcOpen) {
      ui.fcOpen = null;
      render();
    }
  }

  root.addEventListener('click', onClick);
  root.addEventListener('submit', onSubmit);
  root.addEventListener('change', onChange);
  root.addEventListener('pointerdown', onPointerDown);
  root.addEventListener('pointermove', onPointerMove);
  root.addEventListener('pointerup', onPointerUp);
  root.addEventListener('pointercancel', onPointerUp);
  document.addEventListener('keydown', onKey);

  const handle: PlayerHandle = {
    setProject(p) {
      project = p;
      render();
    },
    reset(personaId) {
      stopTimer();
      ui.playing = false;
      ui.status = '';
      ui.picked = [];
      ui.menuOpen = true;
      commit(personaId ? newDemo(project, personaId) : initialDemo(project));
    },
    getDemo: () => demo,
    destroy() {
      stopTimer();
      resize?.disconnect();
      root.removeEventListener('click', onClick);
      root.removeEventListener('submit', onSubmit);
      root.removeEventListener('change', onChange);
      root.removeEventListener('pointerdown', onPointerDown);
      root.removeEventListener('pointermove', onPointerMove);
      root.removeEventListener('pointerup', onPointerUp);
      root.removeEventListener('pointercancel', onPointerUp);
      document.removeEventListener('keydown', onKey);
      root.classList.remove('lsp');
      root.innerHTML = '';
    },
  };
  render();
  return handle;
}
