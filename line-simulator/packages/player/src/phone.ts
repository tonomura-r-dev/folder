// スマホ枠の中（はじめかた・Webページ・ポップアップ・LINEトーク）。選んでいるユーザー1人分を描く。

import type { Action, Bubble, Id, RichMenu, TalkItem, UserState, WebPage } from '../../engine/src/index.ts';
import { fmtTime, type Labels } from './describe.ts';
import { dayOf, esc, hhmm, type View } from './view.ts';

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

function bubble(v: View, b: Bubble, nodeId?: Id): string {
  const btns = (list: { id: Id; label: string }[]) =>
    `<div class="lsp-btns">${list
      .map((x) => `<button type="button" data-act="tapButton" data-button="${esc(x.id)}"${nodeId ? ` data-node="${esc(nodeId)}"` : ''}>${esc(x.label)}</button>`)
      .join('')}</div>`;
  switch (b.kind) {
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
        }${bubble(v, item.bubble, item.fromNodeId)}</div><span class="lsp-time">${hhmm(item.at)}</span></div>`,
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
  const cells = m.cells
    .map(
      (c) =>
        `<button type="button" data-act="tapMenu" data-menu="${esc(m.id)}" data-cell="${esc(c.id)}" aria-label="${esc(c.label)}" style="grid-column:${c.area.col + 1} / span ${
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

function talk(v: View, user: UserState): string {
  const blocked = user.friendState === 'blocked';
  const m = user.richMenuId ? v.project.richMenus.find((r) => r.id === user.richMenuId) : undefined;
  return `<div class="lsp-talk">
    <div class="lsp-talkbar"><span class="lsp-back">‹</span>${esc(v.project.brand.accountName)}</div>
    ${blocked ? '<div class="lsp-banner">ブロック中です。配信は届きません。<button type="button" data-act="unblock">ブロックを解除</button></div>' : ''}
    <div class="lsp-chat" aria-live="polite">${chat(v, user)}</div>
    ${blocked ? '' : surveyPanel(v, user)}
    ${m && !blocked ? menu(v, m) : ''}
    <form class="lsp-input" data-form="send"><input name="text" placeholder="メッセージを入力" autocomplete="off" aria-label="メッセージ"${blocked ? ' disabled' : ''}><button type="submit"${blocked ? ' disabled' : ''}>送信</button></form>
  </div>`;
}

function page(v: View, user: UserState, p: WebPage, popup: boolean): string {
  const id = esc(p.id);
  const blocks = p.blocks
    .map((b, i) => {
      if (b.kind === 'button') return `<button type="button" class="lsp-pbtn" data-act="pageBtn" data-page="${id}" data-i="${i}">${esc(b.label)}</button>`;
      if (b.kind === 'image') return img(v, b.assetId);
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
  return `<div class="lsp-start">
    <div class="lsp-start-title">はじめかたを選ぶ</div>
    <div class="lsp-start-sub">この人が、どこからやって来るかを選んでください。</div>
    ${pages.length ? '<div class="lsp-start-group">Webサイトから</div>' : ''}
    ${pages.map((p) => `<button type="button" data-act="visit" data-page="${esc(p.id)}">${esc(p.title)}<small>${esc(p.url ?? '')}</small></button>`).join('')}
    ${v.project.routes.length ? '<div class="lsp-start-group">LINEで直接友だち追加</div>' : ''}
    ${v.project.routes.map((r) => `<button type="button" class="line" data-act="addFriend" data-route="${esc(r.id)}">友だち追加する<small>経路：${esc(r.label)}</small></button>`).join('')}
  </div>`;
}

export function phone(v: View, user: UserState): string {
  const s = user.surface;
  const p = s.kind === 'page' || s.kind === 'popup' ? v.project.webPages.find((x) => x.id === s.pageId) : undefined;
  const body = p ? page(v, user, p, s.kind === 'popup') : user.friendState === 'not_friend' ? start(v) : talk(v, user);
  return `<div class="lsp-phone" aria-label="${esc(user.name)}のスマホ"><div class="lsp-status"><span>${esc(fmtTime(v.world.now))}</span><span>${esc(user.name)}</span></div><div class="lsp-screen">${body}</div></div>`;
}
