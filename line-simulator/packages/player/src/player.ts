// デモ画面（スマホ枠の LINE・Web と、横の裏側パネル）。React を使わず素の DOM で描く。
// 編集ツールのデモ画面と、Phase 5 の営業用HTMLの両方がこれをそのまま使う。
//
// 決まりごと
//   - 押した結果を画面側で決めない。すべてエンジンへの入力イベントにする（controller.ts の dispatch）
//   - 表示は毎回「イベント列 → reduce → 状態」から作り直す
//   - 案件の文字は必ず esc() を通す（< や & を含んでも崩れない）

import type { Asset, Bubble, Id, InputEventBody, Project, RichMenu, TalkItem, UserState, WebPage, WorldState } from '../../engine/src/index.ts';
import { MINUTES_PER_DAY } from '../../engine/src/index.ts';
import { advance, dispatch, newDemo, upcoming, worldOf, type Demo } from './controller.ts';
import { fmtTime, friendText, Labels } from './describe.ts';

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

const ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
export const esc = (v: unknown) => String(v ?? '').replace(/[&<>"']/g, (c) => ESC[c]!);

const hhmm = (min: number) => {
  const m = ((min % MINUTES_PER_DAY) + MINUTES_PER_DAY) % MINUTES_PER_DAY;
  return `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
};
const dayOf = (min: number) => Math.floor(min / MINUTES_PER_DAY);

const SYSTEM: Record<string, (p: Record<string, string>, L: Labels) => string> = {
  friendAdded: (p, L) => `友だち追加しました（${L.route(p['routeId'] ?? null)}）`,
  blocked: () => 'ブロックしました',
  unblocked: () => 'ブロックを解除しました',
  surveyShown: () => 'アンケートが届きました',
  pageOpened: () => 'ページを開きました',
};

export function createPlayer(root: HTMLElement, opts: PlayerOptions): PlayerHandle {
  let project = opts.project;
  let demo = opts.demo ?? newDemo(project, opts.personaId);
  let world: WorldState | null = null;
  let menuOpen = true;
  let picked: Id[] = []; // 複数選択アンケートで選び中の選択肢
  const assetUrl = opts.assetUrl ?? ((a: Asset) => a.path);

  root.classList.add('lsp');

  const commit = (next: Demo) => {
    demo = next;
    opts.onChange?.(demo);
    render();
  };
  const send = (body: InputEventBody) => commit(dispatch(demo, body));

  // ---------- 部品 ----------

  const L = () => new Labels(project);
  const asset = (id: Id | undefined) => (id ? project.assets.find((a) => a.id === id) : undefined);
  const img = (id: Id | undefined, alt = '') => {
    const a = asset(id);
    return a ? `<img src="${esc(assetUrl(a))}" alt="${esc(a.alt || alt)}">` : `<div class="lsp-noimg">画像（未登録）</div>`;
  };

  function bubble(b: Bubble, nodeId?: Id): string {
    const btns = (list: { id: Id; label: string }[]) =>
      `<div class="lsp-btns">${list
        .map((x) => `<button type="button" data-act="tapButton" data-button="${esc(x.id)}"${nodeId ? ` data-node="${esc(nodeId)}"` : ''}>${esc(x.label)}</button>`)
        .join('')}</div>`;
    switch (b.kind) {
      case 'text':
        return `<div class="lsp-b">${esc(b.text)}</div>`;
      case 'image':
        return `<div class="lsp-b img">${img(b.assetId)}</div>`;
      case 'buttons':
        return `<div class="lsp-b tmpl"><div class="lsp-tmpl-text">${esc(b.text)}</div>${btns(b.buttons)}</div>`;
      case 'card':
        return `<div class="lsp-b card lsp-card">${b.assetId ? img(b.assetId) : ''}<div class="lsp-card-title">${esc(b.title)}</div>${
          b.body ? `<div class="lsp-card-body">${esc(b.body)}</div>` : ''
        }${btns(b.buttons)}</div>`;
    }
  }

  function chat(user: UserState): string {
    const labels = L();
    const initial = esc([...project.brand.accountName][0] ?? 'L');
    const out: string[] = [];
    let lastDay = -1;
    let prev: TalkItem | undefined;
    for (const item of user.talk) {
      const day = dayOf(item.at);
      if (day !== lastDay) {
        out.push(`<div class="lsp-sys">Day ${day}</div>`);
        lastDay = day;
        prev = undefined;
      }
      if (item.side === 'system') {
        out.push(`<div class="lsp-sys">${esc(SYSTEM[item.code]?.(item.params ?? {}, labels) ?? item.code)}</div>`);
      } else if (item.side === 'user') {
        out.push(`<div class="lsp-row me"><span class="lsp-time">${hhmm(item.at)}</span>${bubble(item.bubble).replace('class="lsp-b', 'class="lsp-b me')}</div>`);
      } else {
        // 同じ配信の2つ目以降の吹き出しは、アイコンと名前を出さない
        const cont = prev?.side === 'bot' && prev.at === item.at && prev.messageId === item.messageId;
        out.push(
          `<div class="lsp-row bot"><div class="lsp-avatar${cont ? ' ghost' : ''}">${initial}</div><div class="lsp-col">${
            cont ? '' : `<div class="lsp-name">${esc(project.brand.accountName)}</div>`
          }${bubble(item.bubble, item.fromNodeId)}</div><span class="lsp-time">${hhmm(item.at)}</span></div>`,
        );
      }
      prev = item;
    }
    if (!user.talk.length) out.push('<div class="lsp-sys">まだメッセージはありません</div>');
    return out.join('');
  }

  function surveyPanel(user: UserState): string {
    if (user.surface.kind !== 'survey') return '';
    const { surveyId, questionId } = user.surface;
    const q = project.surveys.find((s) => s.id === surveyId)?.questions.find((x) => x.id === questionId);
    if (!q) return '';
    const data = `data-survey="${esc(surveyId)}" data-q="${esc(questionId)}"`;
    const chips = q.choices
      .map((c) =>
        q.mode === 'single'
          ? `<button type="button" data-act="answer" ${data} data-choice="${esc(c.id)}">${esc(c.label)}</button>`
          : `<button type="button" data-act="pick" data-choice="${esc(c.id)}" class="${picked.includes(c.id) ? 'on' : ''}" aria-pressed="${picked.includes(c.id)}">${esc(c.label)}</button>`,
      )
      .join('');
    const multi = q.mode === 'multiple' ? `<button type="button" class="lsp-send-multi" data-act="answerMulti" ${data}${picked.length ? '' : ' disabled'}>回答する</button>` : '';
    return `<div class="lsp-quick"><div class="lsp-quick-q">${esc(q.text)}${q.mode === 'multiple' ? '（複数選べます）' : ''}</div><div class="lsp-chips">${chips}</div>${multi}</div>`;
  }

  function menu(m: RichMenu): string {
    const bg = asset(m.imageAssetId);
    const cells = m.cells
      .map(
        (c) =>
          `<button type="button" data-act="tapMenu" data-menu="${esc(m.id)}" data-cell="${esc(c.id)}" aria-label="${esc(c.label)}" style="grid-column:${c.area.col + 1} / span ${
            c.area.colSpan ?? 1
          };grid-row:${c.area.row + 1} / span ${c.area.rowSpan ?? 1}">${esc(c.label)}</button>`,
      )
      .join('');
    const grid = `<div class="lsp-menu-grid${bg ? ' has-img' : ''}" style="grid-template-columns:repeat(${m.layout.cols},1fr);grid-template-rows:repeat(${m.layout.rows},auto)${
      bg ? `;background-image:url('${esc(assetUrl(bg))}');aspect-ratio:${bg.width}/${bg.height}` : ''
    }">${cells}</div>`;
    return `<div class="lsp-menu"><button type="button" class="lsp-menu-bar" data-act="toggleMenu" aria-expanded="${menuOpen}">${esc(m.label)} ${menuOpen ? '▼' : '▲'}</button>${menuOpen ? grid : ''}</div>`;
  }

  function talk(user: UserState): string {
    const blocked = user.friendState === 'blocked';
    const m = user.richMenuId ? project.richMenus.find((r) => r.id === user.richMenuId) : undefined;
    return `<div class="lsp-talk">
      <div class="lsp-talkbar"><span class="lsp-back">‹</span>${esc(project.brand.accountName)}</div>
      ${blocked ? '<div class="lsp-banner">ブロック中です。配信は届きません。<button type="button" data-act="unblock">ブロックを解除</button></div>' : ''}
      <div class="lsp-chat" aria-live="polite">${chat(user)}</div>
      ${blocked ? '' : surveyPanel(user)}
      ${m && !blocked ? menu(m) : ''}
      <form class="lsp-input" data-form="send"><input name="text" placeholder="メッセージを入力" autocomplete="off" aria-label="メッセージ"${blocked ? ' disabled' : ''}><button type="submit"${blocked ? ' disabled' : ''}>送信</button></form>
    </div>`;
  }

  function page(user: UserState, p: WebPage, popup: boolean): string {
    const id = esc(p.id);
    const blocks = p.blocks
      .map((b, i) => {
        if (b.kind === 'button') return `<button type="button" class="lsp-pbtn" data-act="pageBtn" data-page="${id}" data-i="${i}">${esc(b.label)}</button>`;
        if (b.kind === 'image') return img(b.assetId);
        return b.kind === 'heading' ? `<h2>${esc(b.text)}</h2>` : `<p>${esc(b.text)}</p>`;
      })
      .join('');
    const form = p.form
      ? `<form class="lsp-form" data-form="submit" data-page="${id}">${p.form.fields
          .map((f) => {
            const name = esc(f.id);
            const input =
              f.type === 'select'
                ? `<select name="${name}">${(f.options ?? []).map((o) => `<option>${esc(o)}</option>`).join('')}</select>`
                : `<input name="${name}" type="${f.type === 'text' ? 'text' : f.type}" autocomplete="off">`;
            return `<label>${esc(f.label)}</label>${input}`;
          })
          .join('')}<button type="submit" class="lsp-pbtn">送信する</button></form>`
      : '';
    const cta = p.lineCta ? `<button type="button" class="lsp-line-cta" data-act="lineCta" data-page="${id}">${esc(p.lineCta.label)}</button>` : '';
    const pop =
      popup && p.exitPopup
        ? `<div class="lsp-overlay" role="dialog" aria-modal="true" aria-label="${esc(p.exitPopup.title)}"><div class="lsp-popup">${
            p.exitPopup.assetId ? img(p.exitPopup.assetId) : ''
          }<div class="lsp-popup-title">${esc(p.exitPopup.title)}</div><div class="lsp-popup-body">${esc(p.exitPopup.body)}</div><button type="button" class="lsp-line-cta" data-act="popupCta" data-page="${id}">${esc(
            p.exitPopup.cta.label,
          )}</button><button type="button" class="lsp-popup-dismiss" data-act="dismissPopup" data-page="${id}">${esc(p.exitPopup.dismissLabel)}</button></div></div>`
        : '';
    const inLine = user.friendState === 'friend';
    return `<div class="lsp-browser">
      <div class="lsp-urlbar"><button type="button" data-act="back" data-page="${id}" title="${inLine ? 'LINEに戻る' : '前のページへ（離脱）'}" aria-label="戻る">‹</button><span class="lsp-url">${esc(p.url ?? p.title)}</span></div>
      <div class="lsp-page">${blocks}${form}${cta}<div class="lsp-fake">架空のページです</div></div>
      ${pop}
    </div>`;
  }

  function start(): string {
    const pages = project.webPages.filter((p) => p.kind !== 'thanks');
    return `<div class="lsp-start">
      <div class="lsp-start-title">はじめかたを選ぶ</div>
      <div class="lsp-start-sub">この人が、どこからやって来るかを選んでください。</div>
      ${pages.length ? '<div class="lsp-start-group">Webサイトから</div>' : ''}
      ${pages.map((p) => `<button type="button" data-act="visit" data-page="${esc(p.id)}">${esc(p.title)}<small>${esc(p.url ?? '')}</small></button>`).join('')}
      ${project.routes.length ? '<div class="lsp-start-group">LINEで直接友だち追加</div>' : ''}
      ${project.routes.map((r) => `<button type="button" class="line" data-act="addFriend" data-route="${esc(r.id)}">友だち追加する<small>経路：${esc(r.label)}</small></button>`).join('')}
    </div>`;
  }

  function phone(user: UserState, w: WorldState): string {
    const s = user.surface;
    const p = s.kind === 'page' || s.kind === 'popup' ? project.webPages.find((x) => x.id === s.pageId) : undefined;
    const body = p ? page(user, p, s.kind === 'popup') : user.friendState === 'not_friend' ? start() : talk(user);
    return `<div class="lsp-phone" aria-label="${esc(user.name)}のスマホ"><div class="lsp-status"><span>${esc(fmtTime(w.now))}</span><span>${esc(user.name)}</span></div><div class="lsp-screen">${body}</div></div>`;
  }

  function side(user: UserState, w: WorldState): string {
    const labels = L();
    const next = upcoming(w, user.id)[0];
    const personaId = demo.session.users[0]?.personaId ?? '';
    const personas = project.personas.map((p) => `<option value="${esc(p.id)}"${p.id === personaId ? ' selected' : ''}>${esc(p.name)}</option>`).join('');
    const tags = Object.keys(user.tags);
    const cvs = Object.entries(user.cvs);
    const log = [...user.deliveries]
      .reverse()
      .map(
        (d) =>
          `<li><span class="t">${esc(fmtTime(d.at))}</span><span>${esc(labels.message(d.messageId))}${d.count === 0 ? '<span class="reply">応答</span>' : ''}<span class="f">${esc(labels.flow(d.flowId))}</span></span></li>`,
      )
      .join('');
    const sent = user.deliveries.reduce((a, d) => a + d.count, 0);
    return `<div class="lsp-side">
      <div class="lsp-card-side"><h3>ユーザー</h3><div class="lsp-ctrls">${
        personas ? `<select data-change="persona" aria-label="ペルソナ">${personas}</select>` : ''
      }<button type="button" data-act="reset">最初から</button>${
        user.friendState === 'friend' ? '<button type="button" data-act="block">ブロックする</button>' : ''
      }</div></div>
      <div class="lsp-card-side"><h3>時間</h3><div class="lsp-clock">${esc(fmtTime(w.now))}</div>
        <div class="lsp-next">${next ? `次の予定：<b>${esc(fmtTime(next.fireAt))}</b>　${esc(labels.flow(next.flowId))}` : '次の予定はありません'}</div>
        <div class="lsp-ctrls"><button type="button" class="primary" data-act="next"${next ? '' : ' disabled'}>次の予定まで進める</button><button type="button" data-act="adv" data-min="60">＋1時間</button><button type="button" data-act="adv" data-min="1440">＋1日</button></div></div>
      <div class="lsp-card-side"><h3>いまの状態</h3><dl class="lsp-kv">
        <dt>友だち</dt><dd>${esc(friendText(user.friendState))}</dd>
        <dt>経路</dt><dd>${esc(labels.route(user.route))}</dd>
        <dt>CV</dt><dd>${cvs.length ? cvs.map(([id, v]) => `${esc(labels.cv(id))}×${v.count}`).join('、') : '<span class="lsp-none">なし</span>'}</dd>
        <dt>タグ</dt><dd>${tags.length ? `<span class="lsp-tags">${tags.map((t) => `<span class="lsp-tag">${esc(labels.tag(t))}</span>`).join('')}</span>` : '<span class="lsp-none">なし</span>'}</dd>
      </dl></div>
      <div class="lsp-card-side"><h3>配信履歴（${user.deliveries.length}件・通数 ${sent}）</h3>${log ? `<ul class="lsp-log">${log}</ul>` : '<div class="lsp-none">まだ配信はありません</div>'}</div>
    </div>`;
  }

  // ---------- 描画 ----------

  function render() {
    try {
      world = worldOf(project, demo);
    } catch (e) {
      root.innerHTML = `<div class="lsp-error">デモを表示できません：${esc(e instanceof Error ? e.message : e)}</div>`;
      return;
    }
    const user = world.users[demo.userId];
    const c = project.brand.colors;
    root.style.setProperty('--lsp-primary', c.primary);
    root.style.setProperty('--lsp-accent', c.accent);
    if (c.chatBg) root.style.setProperty('--lsp-chat', c.chatBg);
    else root.style.removeProperty('--lsp-chat');
    root.innerHTML = user ? `<div class="lsp-stage">${phone(user, world)}${side(user, world)}</div>` : '<div class="lsp-error">ユーザーがいません</div>';
    const scroller = root.querySelector<HTMLElement>('.lsp-chat');
    if (scroller) scroller.scrollTop = scroller.scrollHeight;
  }

  // ---------- 操作 ----------

  function onClick(e: Event) {
    const el = (e.target as HTMLElement).closest<HTMLElement>('[data-act]');
    if (!el || !root.contains(el) || (el as HTMLButtonElement).disabled) return;
    const d = el.dataset;
    const pageId = d['page'] ?? '';
    const user = world?.users[demo.userId];
    switch (d['act']) {
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
      case 'lineCta':
        return send({ type: 'tapLineCta', pageId });
      case 'tapButton':
        return send({ type: 'tapButton', buttonId: d['button'] ?? '', ...(d['node'] ? { nodeId: d['node'] } : {}) });
      case 'tapMenu':
        return send({ type: 'tapMenu', richMenuId: d['menu'] ?? '', cellId: d['cell'] ?? '' });
      case 'toggleMenu':
        menuOpen = !menuOpen;
        return render();
      case 'answer':
        return send({ type: 'answer', surveyId: d['survey'] ?? '', questionId: d['q'] ?? '', choiceIds: [d['choice'] ?? ''] });
      case 'pick': {
        const c = d['choice'] ?? '';
        picked = picked.includes(c) ? picked.filter((x) => x !== c) : [...picked, c];
        return render();
      }
      case 'answerMulti': {
        const choiceIds = picked;
        picked = [];
        return send({ type: 'answer', surveyId: d['survey'] ?? '', questionId: d['q'] ?? '', choiceIds });
      }
      case 'block':
        return send({ type: 'block' });
      case 'unblock':
        return send({ type: 'unblock' });
      case 'next': {
        const t = world ? upcoming(world, demo.userId)[0] : undefined;
        if (t) commit(advance(demo, t.fireAt));
        return;
      }
      case 'adv':
        return commit(advance(demo, demo.now + Number(d['min'] ?? 0)));
      case 'reset':
        return handle.reset(demo.session.users[0]?.personaId);
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
    const el = e.target as HTMLSelectElement;
    if (root.contains(el) && el.dataset['change'] === 'persona') handle.reset(el.value);
  }

  root.addEventListener('click', onClick);
  root.addEventListener('submit', onSubmit);
  root.addEventListener('change', onChange);

  const handle: PlayerHandle = {
    setProject(p) {
      project = p;
      render();
    },
    reset(personaId) {
      picked = [];
      menuOpen = true;
      commit(newDemo(project, personaId));
    },
    getDemo: () => demo,
    destroy() {
      root.removeEventListener('click', onClick);
      root.removeEventListener('submit', onSubmit);
      root.removeEventListener('change', onChange);
      root.classList.remove('lsp');
      root.innerHTML = '';
    },
  };
  render();
  return handle;
}
