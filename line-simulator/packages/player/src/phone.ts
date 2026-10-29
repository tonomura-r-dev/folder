// スマホ枠の中（はじめかた・Webページ・ポップアップ・LINEトーク）。選んでいるユーザー1人分を描く。

import type { Action, Bubble, CarouselCard, Id, RichMenu, TalkItem, UserState, WebBlock, WebPage } from '../../engine/src/index.ts';
import { fmtTime, type Labels } from './describe.ts';
import { dayOf, esc, hhmm, whoOf, type View } from './view.ts';

const SYSTEM: Record<string, (p: Record<string, string>, L: Labels) => string> = {
  friendAdded: (p, L) => `友だち追加しました（${L.route(p['routeId'] ?? null)}）`,
  blocked: () => 'ブロックしました',
  unblocked: () => 'ブロックを解除しました',
  surveyShown: () => 'アンケートが届きました',
  pageOpened: () => 'ページを開きました',
};

function img(v: View, id: Id | undefined, alt = ''): string {
  const a = id ? v.project.assets.find((x) => x.id === id) : undefined;
  return a ? `<img src="${esc(v.assetUrl(a))}" alt="${esc(a.alt || alt)}">` : `<div class="lsp-noimg">画像（未登録）</div>`;
}

/** カルーセルのいま表示しているカード。届いた後に見たカード（viewCard）の最後のもの。無ければ1枚目 */
function cardIndex(cards: CarouselCard[], user: UserState | undefined, messageId: Id | undefined, at: number): number {
  if (!user || !messageId) return 0;
  for (let i = user.behaviors.length - 1; i >= 0; i--) {
    const b = user.behaviors[i]!;
    if (b.at < at) break;
    if (b.type === 'viewCard' && b.detail['messageId'] === messageId) return Math.max(0, cards.findIndex((c) => c.id === b.detail['cardId']));
  }
  return 0;
}

function facts(c: CarouselCard): string {
  return c.facts?.length ? `<dl class="lsp-car-facts">${c.facts.map((f) => `<div><dt>${esc(f.label)}</dt><dd>${esc(f.value)}</dd></div>`).join('')}</dl>` : '';
}

/** 横に送れるカード。左右のボタン・点・ドラッグ（スワイプ）で切り替える。切り替えると viewCard を送る */
function carousel(v: View, cards: CarouselCard[], btns: (list: { id: Id; label: string }[]) => string, at?: { user: UserState; messageId: Id; at: number }): string {
  const i = cardIndex(cards, at?.user, at?.messageId, at?.at ?? 0);
  const data = at ? ` data-msg="${esc(at.messageId)}"` : '';
  const go = (k: number, label: string, cls: string, aria: string) =>
    at && k >= 0 && k < cards.length && k !== i
      ? `<button type="button" class="${cls}" data-act="carGo"${data} data-card="${esc(cards[k]!.id)}" aria-label="${esc(aria)}">${label}</button>`
      : `<button type="button" class="${cls}" disabled aria-label="${esc(aria)}">${label}</button>`;
  const card = (c: CarouselCard, k: number) =>
    `<div class="lsp-car-card${k === i ? ' on' : ''}" aria-hidden="${k !== i}"><div class="lsp-car-img">${c.assetId ? img(v, c.assetId) : ''}${
      c.badge ? `<span class="lsp-car-badge">${esc(c.badge)}</span>` : ''
    }</div><div class="lsp-car-title">${esc(c.title)}</div>${facts(c)}${c.body ? `<div class="lsp-car-body">${esc(c.body)}</div>` : ''}${btns(c.buttons)}</div>`;
  const dots = cards.map((c, k) => (at ? `<button type="button" class="lsp-car-dot${k === i ? ' on' : ''}" data-act="carGo"${data} data-card="${esc(c.id)}" aria-label="${k + 1}枚目"></button>` : '')).join('');
  return `<div class="lsp-b lsp-car"${data}><div class="lsp-car-view"${at ? ' data-drag="car"' : ''}><div class="lsp-car-track" style="--i:${i}">${cards.map(card).join('')}</div></div>${
    at
      ? `<div class="lsp-car-nav">${go(i - 1, '‹', 'lsp-car-prev', '前のカード')}<span class="lsp-car-dots">${dots}</span><span class="lsp-car-count">${i + 1} / ${cards.length}</span>${go(i + 1, '›', 'lsp-car-next', '次のカード')}</div>`
      : ''
  }</div>`;
}

function bubble(v: View, b: Bubble, nodeId?: Id, at?: { user: UserState; messageId: Id; at: number }): string {
  const btns = (list: { id: Id; label: string }[]) =>
    `<div class="lsp-btns">${list
      .map((x) => `<button type="button" data-act="tapButton" data-button="${esc(x.id)}"${nodeId ? ` data-node="${esc(nodeId)}"` : ''}>${esc(x.label)}</button>`)
      .join('')}</div>`;
  switch (b.kind) {
    case 'carousel':
      return carousel(v, b.cards, btns, at);
    case 'text':
      return `<div class="lsp-b">${esc(b.text)}</div>`;
    case 'image':
      return `<div class="lsp-b img">${img(v, b.assetId)}</div>`;
    case 'buttons':
      return `<div class="lsp-b tmpl"><div class="lsp-tmpl-text">${esc(b.text)}</div>${btns(b.buttons)}</div>`;
    case 'card':
      return `<div class="lsp-b card lsp-card">${b.assetId ? img(v, b.assetId) : ''}<div class="lsp-card-title">${esc(b.title)}</div>${
        b.body ? `<div class="lsp-card-body">${esc(b.body)}</div>` : ''
      }${btns(b.buttons)}</div>`;
  }
}

/** 押しても何も起きない吹き出し（③の文面確認用）。ボタンには押したときの動作を添える */
export function previewBubble(v: View, b: Bubble): string {
  const L = v.L;
  const what = (a: Action) =>
    a.kind === 'openPage'
      ? `「${L.page(a.pageId)}」を開く`
      : a.kind === 'answer'
        ? `アンケートに「${L.choices(a.surveyId, a.questionId, [a.choiceId])}」と回答`
        : a.kind === 'sendText'
          ? `「${a.text}」と送る`
          : a.kind === 'startFlow'
            ? `「${L.flow(a.flowId)}」を始める`
            : a.kind === 'addFriend'
              ? '友だち追加'
              : a.kind === 'switchMenu'
                ? 'メニューを切り替える'
                : '';
  const btns = (list: { label: string; action: Action }[]) =>
    `<div class="lsp-btns">${list.map((x) => `<span class="lsp-btn-ro">${esc(x.label)}${what(x.action) ? `<small>→ ${esc(what(x.action))}</small>` : ''}</span>`).join('')}</div>`;
  switch (b.kind) {
    case 'carousel':
      // 文面確認では全カードを横に並べ、スクロールで見る
      return `<div class="lsp-b lsp-car ro"><div class="lsp-car-ro">${b.cards
        .map(
          (c) =>
            `<div class="lsp-car-card"><div class="lsp-car-img">${c.assetId ? img(v, c.assetId) : ''}${c.badge ? `<span class="lsp-car-badge">${esc(c.badge)}</span>` : ''}</div><div class="lsp-car-title">${esc(
              c.title,
            )}</div>${facts(c)}${c.body ? `<div class="lsp-car-body">${esc(c.body)}</div>` : ''}${btns(c.buttons)}</div>`,
        )
        .join('')}</div></div>`;
    case 'text':
      return `<div class="lsp-b">${esc(b.text)}</div>`;
    case 'image':
      return `<div class="lsp-b img">${img(v, b.assetId)}</div>`;
    case 'buttons':
      return `<div class="lsp-b tmpl"><div class="lsp-tmpl-text">${esc(b.text)}</div>${btns(b.buttons)}</div>`;
    case 'card':
      return `<div class="lsp-b card lsp-card">${b.assetId ? img(v, b.assetId) : ''}<div class="lsp-card-title">${esc(b.title)}</div>${
        b.body ? `<div class="lsp-card-body">${esc(b.body)}</div>` : ''
      }${btns(b.buttons)}</div>`;
  }
}

function chat(v: View, user: UserState): string {
  const initial = esc([...v.project.brand.accountName][0] ?? 'L');
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
      // アンケートは質問の吹き出しで分かるので、「アンケートが届きました」の帯は出さない
      if (item.code === 'surveyShown') continue;
      out.push(`<div class="lsp-sys">${esc(SYSTEM[item.code]?.(item.params ?? {}, v.L) ?? item.code)}</div>`);
    } else if (item.side === 'user') {
      out.push(`<div class="lsp-row me"><span class="lsp-time">${hhmm(item.at)}</span>${bubble(v, item.bubble).replace('class="lsp-b', 'class="lsp-b me')}</div>`);
    } else {
      // 同じ配信の2つ目以降の吹き出しは、アイコンと名前を出さない
      const cont = prev?.side === 'bot' && prev.at === item.at && prev.messageId === item.messageId;
      out.push(
        `<div class="lsp-row bot"><div class="lsp-avatar${cont ? ' ghost' : ''}">${initial}</div><div class="lsp-col">${
          cont ? '' : `<div class="lsp-name">${esc(v.project.brand.accountName)}</div>`
        }${bubble(v, item.bubble, item.fromNodeId, { user, messageId: item.messageId, at: item.at })}</div><span class="lsp-time">${hhmm(item.at)}</span></div>`,
      );
    }
    prev = item;
  }
  if (!user.talk.length) out.push('<div class="lsp-sys">まだメッセージはありません</div>');
  return out.join('');
}

function surveyPanel(v: View, user: UserState): string {
  if (user.surface.kind !== 'survey') return '';
  const { surveyId, questionId } = user.surface;
  const q = v.project.surveys.find((s) => s.id === surveyId)?.questions.find((x) => x.id === questionId);
  if (!q) return '';
  const data = `data-survey="${esc(surveyId)}" data-q="${esc(questionId)}"`;
  const picked = v.ui.picked;
  const chips = q.choices
    .map((c) =>
      q.mode === 'single'
        ? `<button type="button" data-act="answer" ${data} data-choice="${esc(c.id)}">${esc(c.label)}</button>`
        : `<button type="button" data-act="pick" data-choice="${esc(c.id)}" class="${picked.includes(c.id) ? 'on' : ''}" aria-pressed="${picked.includes(c.id)}">${esc(c.label)}</button>`,
    )
    .join('');
  const multi = q.mode === 'multiple' ? `<button type="button" class="lsp-send-multi" data-act="answerMulti" ${data}${picked.length ? '' : ' disabled'}>回答する</button>` : '';
  return `<div class="lsp-quick"><div class="lsp-quick-q">▼ タップして回答${q.mode === 'multiple' ? '（複数選べます）' : ''}</div><div class="lsp-chips">${chips}</div>${multi}</div>`;
}

function menu(v: View, m: RichMenu): string {
  const bg = m.imageAssetId ? v.project.assets.find((a) => a.id === m.imageAssetId) : undefined;
  // タブのあるメニューだけ、「動作なし」の項目を「いま開いているタブ」として出す（タブの無いメニューの動作なし＝シナリオを始めるだけの項目）
  const tabbed = m.cells.some((c) => c.action.kind === 'switchMenu');
  const cells = m.cells
    .map(
      (c) =>
        // タブ（メニューの切り替え・いま開いているタブ）は低く出し、ほかの項目と見分ける
        `<button type="button"${c.action.kind === 'switchMenu' ? ' class="tab"' : tabbed && c.action.kind === 'none' ? ' class="tab on"' : c.emphasis ? ' class="em"' : ''} data-act="tapMenu" data-menu="${esc(m.id)}" data-cell="${esc(c.id)}" aria-label="${esc(c.label)}" style="grid-column:${c.area.col + 1} / span ${
          c.area.colSpan ?? 1
        };grid-row:${c.area.row + 1} / span ${c.area.rowSpan ?? 1}">${esc(c.label)}</button>`,
    )
    .join('');
  const grid = `<div class="lsp-menu-grid${bg ? ' has-img' : ''}" style="grid-template-columns:repeat(${m.layout.cols},1fr);grid-template-rows:repeat(${m.layout.rows},auto)${
    bg ? `;background-image:url('${esc(v.assetUrl(bg))}');aspect-ratio:${bg.width}/${bg.height}` : ''
  }">${cells}</div>`;
  const open = v.ui.menuOpen;
  return `<div class="lsp-menu"><button type="button" class="lsp-menu-bar" data-act="toggleMenu" aria-expanded="${open}">${esc(m.label)} ${open ? '▼' : '▲'}</button>${open ? grid : ''}</div>`;
}

/**
 * トークの入力欄に入っている文。LINE誘導のボタン（フォームの内容を引き継ぐ設定）を押した直後だけ。
 * 送る・ほかの操作をしたら消える（イベント列から決めるので、戻る・人の切り替えでもずれない）
 */
function prefillOf(v: View, user: UserState): string {
  if (user.friendState !== 'friend') return '';
  for (let i = user.behaviors.length - 1; i >= 0; i--) {
    const b = user.behaviors[i]!;
    if (b.type === 'tapLineCta') return v.project.webPages.find((p) => p.id === b.detail['pageId'])?.lineCta?.prefill ?? '';
    // 友だち追加・あいさつのアンケートに答えている間は残す
    if (b.type !== 'answer' && b.type !== 'addFriend' && b.type !== 'viewCard') return '';
  }
  return '';
}

function talk(v: View, user: UserState): string {
  const blocked = user.friendState === 'blocked';
  const pre = prefillOf(v, user);
  const m = user.richMenuId ? v.project.richMenus.find((r) => r.id === user.richMenuId) : undefined;
  return `<div class="lsp-talk">
    <div class="lsp-talkbar"><span class="lsp-back">‹</span>${esc(v.project.brand.accountName)}</div>
    ${blocked ? '<div class="lsp-banner">ブロック中です。配信は届きません。<button type="button" data-act="unblock">ブロックを解除</button></div>' : ''}
    <div class="lsp-chat" aria-live="polite">${chat(v, user)}</div>
    ${blocked ? '' : surveyPanel(v, user)}
    ${m && !blocked ? menu(v, m) : ''}
    ${pre ? '<div class="lsp-prefill-note">フォームの内容が入力欄に入っています。「送信」を押すと送ります</div>' : ''}
    <form class="lsp-input${pre ? ' prefilled' : ''}" data-form="send"><input name="text" placeholder="メッセージを入力" autocomplete="off" aria-label="メッセージ"${blocked ? ' disabled' : ''}${
      pre ? ` value="${esc(pre)}"` : ''
    }><button type="submit"${blocked ? ' disabled' : ''}>送信</button></form>
  </div>`;
}

/**
 * 一覧の絞り込みのいまの値（軸ごと）。このページを最後に開いてからの filterList が優先。
 * 無ければ、その人のタグに合う選択肢（希望条件）で開く。空文字は「すべて」
 */
export function listFilter(user: UserState, pageId: Id, blockIndex: number, b: Extract<WebBlock, { kind: 'list' }>): Record<Id, Id> {
  const out: Record<Id, Id> = {};
  for (const f of b.facets) out[f.id] = f.options.find((o) => o.tagId && o.tagId in user.tags)?.id ?? '';
  let opened = -1;
  user.behaviors.forEach((x, i) => {
    if (x.type === 'visitPage' && x.detail['pageId'] === pageId) opened = i;
  });
  user.behaviors.forEach((x, i) => {
    if (i > opened && x.type === 'filterList' && x.detail['pageId'] === pageId && Number(x.detail['blockIndex']) === blockIndex) out[x.detail['facetId'] ?? ''] = x.detail['optionId'] ?? '';
  });
  return out;
}

function listBlock(v: View, user: UserState, p: WebPage, i: number, b: Extract<WebBlock, { kind: 'list' }>): string {
  const sel = listFilter(user, p.id, i, b);
  // 締切を過ぎた項目は出さない（いま募集中のものだけ）
  const open = b.items.filter((it) => it.closesAt === undefined || v.world.now < it.closesAt);
  const hit = open.filter((it) => b.facets.every((f) => !sel[f.id] || (it.values[f.id] ?? []).includes(sel[f.id]!)));
  const data = `data-page="${esc(p.id)}" data-i="${i}"`;
  const facets = b.facets
    .map(
      (f) =>
        `<label class="lsp-list-f${sel[f.id] ? ' on' : ''}"><span>${esc(f.label)}</span><select data-change="filterList" ${data} data-facet="${esc(f.id)}"><option value="">すべて</option>${f.options
          .map((o) => `<option value="${esc(o.id)}"${sel[f.id] === o.id ? ' selected' : ''}>${esc(o.label)}</option>`)
          .join('')}</select></label>`,
    )
    .join('');
  const items = hit
    .map(
      (it) =>
        `<button type="button" class="lsp-list-item" data-act="listItem" ${data} data-item="${esc(it.id)}">${it.assetId ? `<span class="lsp-list-img">${img(v, it.assetId)}</span>` : ''}<span class="lsp-list-main">${
          it.badge ? `<span class="lsp-list-badge">${esc(it.badge)}</span>` : ''
        }<b>${esc(it.title)}</b>${it.facts?.length ? `<span class="lsp-list-facts">${it.facts.map((x) => `<span><i>${esc(x.label)}</i>${esc(x.value)}</span>`).join('')}</span>` : ''}${it.body ? `<span class="lsp-list-body">${esc(it.body)}</span>` : ''}</span></button>`,
    )
    .join('');
  return `<div class="lsp-list"><div class="lsp-list-fs">${facets}</div><div class="lsp-list-count">該当 <b>${hit.length}</b>件／募集中 ${open.length}件</div>${
    items || `<div class="lsp-list-empty">${esc(b.emptyText ?? '条件に合うものがありません。条件を変えてみてください。')}</div>`
  }</div>`;
}

function page(v: View, user: UserState, p: WebPage, popup: boolean): string {
  const id = esc(p.id);
  const blocks = p.blocks
    .map((b, i) => {
      if (b.kind === 'button') return `<button type="button" class="lsp-pbtn" data-act="pageBtn" data-page="${id}" data-i="${i}">${esc(b.label)}</button>`;
      if (b.kind === 'image') return img(v, b.assetId);
      if (b.kind === 'list') return listBlock(v, user, p, i, b);
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
          p.exitPopup.assetId ? img(v, p.exitPopup.assetId) : ''
        }<div class="lsp-popup-title">${esc(p.exitPopup.title)}</div><div class="lsp-popup-body">${esc(p.exitPopup.body)}</div><button type="button" class="lsp-line-cta" data-act="popupCta" data-page="${id}">${esc(
          p.exitPopup.cta.label,
        )}</button><button type="button" class="lsp-popup-dismiss" data-act="dismissPopup" data-page="${id}">${esc(p.exitPopup.dismissLabel)}</button></div></div>`
      : '';
  const inLine = user.friendState === 'friend';
  return `<div class="lsp-browser">
    <div class="lsp-urlbar"><button type="button" data-act="back" data-page="${id}" title="${inLine ? 'LINEに戻る' : '前のページへ（離脱）'}" aria-label="戻る">‹</button><span class="lsp-url">${esc(p.url ?? p.title)}</span></div>
    <div class="lsp-page">${blocks}${form}${cta}<div class="lsp-fake">デモ用のページです</div></div>
    ${pop}
  </div>`;
}

function start(v: View): string {
  const pages = v.project.webPages.filter((p) => p.kind !== 'thanks');
  const entries = (v.project.settings.demoEntries ?? []).filter((e) => v.project.webPages.some((p) => p.id === e.pageId));
  const pageBtn = (p: WebPage) => `<button type="button" data-act="visit" data-page="${esc(p.id)}">${esc(p.title)}<small>${esc(p.url ?? '')}</small></button>`;
  const routes = v.project.routes.map((r) => `<button type="button" class="line" data-act="addFriend" data-route="${esc(r.id)}">友だち追加する<small>経路：${esc(r.label)}</small></button>`).join('');
  if (entries.length) {
    // 営業デモの入口は案件で決めた数個だけ。ほかのページ・経路は消さずに折りたたむ
    const main = entries
      .map((e, i) => `<button type="button" data-act="visit" data-page="${esc(e.pageId)}"><span class="no">${i + 1}</span>${esc(e.label)}<small>${esc(v.L.page(e.pageId))}</small></button>`)
      .join('');
    const rest = pages.filter((p) => !entries.some((e) => e.pageId === p.id));
    return `<div class="lsp-start">
    <div class="lsp-start-title">はじめかたを選ぶ</div>
    <div class="lsp-start-sub">この人が、どこからやって来るかを選んでください。</div>
    <div class="lsp-start-main">${main}</div>
    <details class="lsp-start-more"><summary>ほかのページ・経路から始める（${rest.length + v.project.routes.length}件）</summary>
      ${rest.length ? '<div class="lsp-start-group">Webサイトの各ページ</div>' : ''}${rest.map(pageBtn).join('')}
      ${routes ? `<div class="lsp-start-group">LINEで直接友だち追加</div>${routes}` : ''}
    </details>
  </div>`;
  }
  return `<div class="lsp-start">
    <div class="lsp-start-title">はじめかたを選ぶ</div>
    <div class="lsp-start-sub">この人が、どこからやって来るかを選んでください。</div>
    <div class="lsp-start-main">
    ${pages.length ? '<div class="lsp-start-group">Webサイトから</div>' : ''}
    ${pages.map(pageBtn).join('')}
    ${routes ? `<div class="lsp-start-group">LINEで直接友だち追加</div>${routes}` : ''}
    </div>
  </div>`;
}

export function phone(v: View, user: UserState): string {
  const s = user.surface;
  const p = s.kind === 'page' || s.kind === 'popup' ? v.project.webPages.find((x) => x.id === s.pageId) : undefined;
  const body = p ? page(v, user, p, s.kind === 'popup') : user.friendState === 'not_friend' ? start(v) : talk(v, user);
  return `<div class="lsp-phone" aria-label="${esc(whoOf(user))}のスマホ"><div class="lsp-status"><span>${esc(fmtTime(v.world.now))}</span><span>${esc(user.personName ?? user.name)}</span></div><div class="lsp-screen">${body}</div></div>`;
}
