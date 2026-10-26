// 裏側の表示：説明カード／いまの状態／なぜ（判定の木）／配信履歴／誰に届くか。
// すべて「選んでいるユーザー」の分だけを出す（誰に届くかの表だけは全員を並べる）。

import type { Decision, FlowNode, Id, NarrationDef, Trace, UserState } from '../../engine/src/index.ts';
import { upcoming } from './controller.ts';
import { describeTraceNode, explainDecision, fmtTime, friendText, type Explanation } from './describe.ts';
import { esc, type Tab, type View } from './view.ts';

// ---------- 説明カード ----------

/** いちばん新しい「通ったノード」か「操作」に結びついた説明文 */
export function currentNarration(v: View, user: UserState): NarrationDef | undefined {
  let best: { at: number; rank: number; n: NarrationDef } | undefined;
  const consider = (at: number, rank: number, n: NarrationDef) => {
    if (!best || at > best.at || (at === best.at && rank >= best.rank)) best = { at, rank, n };
  };
  for (const n of v.project.narration) {
    if (n.on.kind === 'node') {
      const nodeId = n.on.nodeId;
      const hit = [...user.trail].reverse().find((t) => t.nodeId === nodeId);
      if (hit) consider(hit.at, 1, n);
    } else {
      const type = n.on.type;
      const hit = [...user.behaviors].reverse().find((b) => b.type === type);
      if (hit) consider(hit.at, 0, n);
    }
  }
  return best?.n;
}

export function narrationCard(v: View, user: UserState, big = false): string {
  const n = currentNarration(v, user);
  if (!n) return `<div class="lsp-narr${big ? ' big' : ''} empty"><div class="lsp-narr-title">操作すると、ここに説明が出ます</div></div>`;
  return `<div class="lsp-narr${big ? ' big' : ''}">${n.eyebrow ? `<div class="lsp-narr-eyebrow">${esc(n.eyebrow)}</div>` : ''}<div class="lsp-narr-title">${esc(n.title)}</div>${
    n.body ? `<div class="lsp-narr-body">${esc(n.body)}</div>` : ''
  }</div>`;
}

// ---------- いまの状態 ----------

function nodeLabel(v: View, flowId: Id, nodeId: Id | null): string {
  if (!nodeId) return '';
  const n: FlowNode | undefined = v.project.flows.find((f) => f.id === flowId)?.nodes.find((x) => x.id === nodeId);
  return n?.label ?? (n ? ({ entry: '入口', message: '文面', image: '画像', richMenu: 'メニュー切替', choice: 'ボタン選択', survey: 'アンケート', tag: 'タグ', branch: '分岐', wait: '待機', page: 'ページ', jump: '別シナリオへ', end: '終了', stopAll: '全停止' } as const)[n.type] : nodeId);
}

const FLOW_STATUS = { running: '実行中', waiting: '待機中', ended: '終了', stopped: '停止' } as const;

function stateTab(v: View, user: UserState): string {
  const L = v.L;
  const tags = Object.entries(user.tags);
  const answers = Object.entries(user.answers).flatMap(([sv, qs]) => Object.entries(qs).map(([q, ids]) => `${esc(L.questionText(sv, q))}：<b>${esc(L.choices(sv, q, ids))}</b>`));
  const attrs = Object.entries(user.attrs);
  const flows = v.project.flows
    .map((f) => ({ f, c: user.flows[f.id] }))
    .filter((x) => x.c)
    .map(({ f, c }) => {
      const next = upcoming(v.world, user.id).find((t) => t.flowId === f.id);
      return `<li><span class="lsp-pill s-${c!.status}">${FLOW_STATUS[c!.status]}</span><span>${esc(f.label)}${
        c!.status === 'waiting' || c!.status === 'running' ? `<span class="f">いま：${esc(nodeLabel(v, f.id, c!.nodeId))}${next ? `　次：${esc(fmtTime(next.fireAt))}` : ''}</span>` : ''
      }</span></li>`;
    })
    .join('');
  const cvButtons = v.project.cvs.map((c) => `<button type="button" data-act="cv" data-cv="${esc(c.id)}">「${esc(c.label)}」したことにする</button>`).join('');
  return `
    <dl class="lsp-kv">
      <dt>友だち</dt><dd>${esc(friendText(user.friendState))}</dd>
      <dt>経路</dt><dd>${esc(L.route(user.route))}</dd>
      <dt>CV</dt><dd>${Object.keys(user.cvs).length ? Object.entries(user.cvs).map(([id, x]) => `${esc(L.cv(id))}×${x.count}（${esc(fmtTime(x.at))}）`).join('、') : '<span class="lsp-none">なし</span>'}</dd>
      ${attrs.length ? `<dt>属性</dt><dd>${attrs.map(([k, a]) => `${esc(k)}=${esc(a)}`).join('、')}</dd>` : ''}
    </dl>
    <h4>タグ</h4>${tags.length ? `<div class="lsp-tags">${tags.map(([t, x]) => `<span class="lsp-tag" title="${esc(fmtTime(x.at))}に付いた">${esc(L.tag(t))}</span>`).join('')}</div>` : '<div class="lsp-none">なし</div>'}
    <h4>アンケートの回答</h4>${answers.length ? `<div class="lsp-answers">${answers.join('<br>')}</div>` : '<div class="lsp-none">まだ回答していません</div>'}
    <h4>シナリオの進み具合</h4>${flows ? `<ul class="lsp-flows">${flows}</ul>` : '<div class="lsp-none">まだどのシナリオにも入っていません</div>'}
    <h4>出来事を起こす</h4><div class="lsp-ctrls">${cvButtons}${
      user.friendState === 'friend' ? '<button type="button" data-act="block">ブロックする</button>' : user.friendState === 'blocked' ? '<button type="button" data-act="unblock">ブロックを解除</button>' : ''
    }</div>`;
}

// ---------- なぜ（判定の木） ----------

function traceTree(t: Trace, v: View): string {
  const kids = t.children?.length ? `<ul>${t.children.map((c) => traceTree(c, v)).join('')}</ul>` : '';
  return `<li class="${t.value ? 'yes' : 'no'}"><span class="mark">${t.value ? '✓' : '✗'}</span>${esc(describeTraceNode(t, v.L))}${kids}</li>`;
}

const TONE_MARK: Record<Explanation['tone'], string> = { ok: '✓', ng: '✗', stop: '■', info: '→' };

function whyTab(v: View, user: UserState): string {
  const list = v.world.decisions
    .map((d, i) => ({ d, i }))
    .filter(({ d }) => d.userId === user.id)
    .filter(({ d }) => !v.ui.onlyNg || d.outcome !== 'taken')
    .reverse();
  const items = list
    .map(({ d, i }) => {
      const x = explainDecision(d, v.L, v.project, v.world);
      const open = v.ui.open.has(i);
      const tree = d.trace ? `<ul class="lsp-tree">${traceTree(d.trace, v)}</ul>` : '';
      return `<li class="lsp-why ${x.tone}${open ? ' open' : ''}">
        <button type="button" class="lsp-why-head" data-act="toggleWhy" data-i="${i}" aria-expanded="${open}"${tree ? '' : ' disabled'}>
          <span class="mark">${TONE_MARK[x.tone]}</span><span class="t">${esc(fmtTime(d.at))}</span><span class="title">${esc(x.title)}${x.reason ? `<span class="reason">${esc(x.reason)}</span>` : ''}</span>${tree ? `<span class="caret">${open ? '▲' : '▼'}</span>` : ''}
        </button>${open ? tree : ''}</li>`;
    })
    .join('');
  return `<label class="lsp-check"><input type="checkbox" data-change="onlyNg"${v.ui.onlyNg ? ' checked' : ''}> 届かなかった・進まなかったものだけ</label>
    ${items ? `<ul class="lsp-whys">${items}</ul>` : '<div class="lsp-none">まだ判定はありません</div>'}
    <div class="lsp-hint">▼ を押すと、条件のどこで決まったかを木で表示します（✗ が当てはまらなかった条件）。</div>`;
}

// ---------- 配信履歴 ----------

function logTab(v: View, user: UserState): string {
  const L = v.L;
  const mine = user.deliveries.reduce((a, d) => a + d.count, 0);
  const all = Object.values(v.world.users).reduce((a, u) => a + u.deliveries.reduce((b, d) => b + d.count, 0), 0);
  const quota = v.project.settings.messageCounting?.freeQuota;
  const log = [...user.deliveries]
    .reverse()
    .map(
      (d) =>
        `<li><span class="t">${esc(fmtTime(d.at))}</span><span>${esc(L.message(d.messageId))}${d.count === 0 ? '<span class="reply">応答・通数に数えない</span>' : ''}<span class="f">${esc(L.flow(d.flowId))}</span></span></li>`,
    )
    .join('');
  return `<div class="lsp-count">この人 <b>${mine}</b>通　／　デモの全員 <b>${all}</b>通${quota ? `<span class="f">無料メッセージ ${quota}通のうち ${all}通（目安）</span>` : ''}</div>
    ${log ? `<ul class="lsp-log">${log}</ul>` : '<div class="lsp-none">まだ配信はありません</div>'}`;
}

// ---------- 誰に届くか ----------

/** シナリオから配信される文面（シナリオの順・ノードの順） */
function flowMessages(v: View): { flowId: Id; messageId: Id }[] {
  const out: { flowId: Id; messageId: Id }[] = [];
  const seen = new Set<string>();
  for (const f of v.project.flows)
    for (const n of f.nodes) {
      const id = n.type === 'message' || n.type === 'choice' ? n.messageId : n.type === 'survey' ? n.introMessageId : undefined;
      if (id && !seen.has(`${f.id}/${id}`)) {
        seen.add(`${f.id}/${id}`);
        out.push({ flowId: f.id, messageId: id });
      }
    }
  return out;
}

const SHORT: Record<string, string> = { blocked: 'ブロック', not_friend: '友だち前', duplicate: '重複', lower_priority: '優先度', stop_condition: '停止' };

function cell(v: View, u: UserState, flowId: Id, messageId: Id): { mark: string; cls: string; tip: string } {
  const got = u.deliveries.filter((d) => d.flowId === flowId && d.messageId === messageId);
  if (got.length)
    return {
      mark: `✓ ${fmtTime(got[0]!.at).replace('Day ', 'D')}${got.length > 1 ? ` ×${got.length}` : ''}`,
      cls: 'ok',
      tip: `${got.map((d) => fmtTime(d.at)).join('、')} に届きました`,
    };
  const miss: Decision | undefined = [...v.world.decisions].reverse().find((d) => d.userId === u.id && d.kind === 'deliver' && d.flowId === flowId && d.ref === messageId);
  if (miss) {
    const x = explainDecision(miss, v.L, v.project, v.world);
    return { mark: `✗ ${SHORT[miss.reason ?? ''] ?? '届かない'}`, cls: 'ng', tip: `${x.title}：${x.reason ?? ''}` };
  }
  const c = u.flows[flowId];
  if (!c) return { mark: '—', cls: 'none', tip: 'このシナリオに入っていません' };
  if (c.status === 'stopped') {
    const stop = [...v.world.decisions].reverse().find((d) => d.userId === u.id && d.kind === 'stopWhen' && d.flowId === flowId);
    const x = stop ? explainDecision(stop, v.L, v.project) : undefined;
    return { mark: '■ 停止', cls: 'stop', tip: x ? `${x.title}：${x.reason ?? ''}` : 'シナリオが止まりました' };
  }
  if (c.status === 'ended') return { mark: '— 別ルート', cls: 'none', tip: '分岐で別のルートに進んだため、この文面は通りませんでした' };
  return { mark: '… まだ', cls: 'wait', tip: 'この先で届く予定、またはまだ分かりません' };
}

function matrixTab(v: View): string {
  const users = v.demo.session.users.map((u) => v.world.users[u.id]).filter((u): u is UserState => !!u);
  if (users.length === 0) return '<div class="lsp-none">ユーザーがいません</div>';
  let lastFlow = '';
  const rows = flowMessages(v)
    .map(({ flowId, messageId }) => {
      const head = flowId !== lastFlow ? `<tr class="grp"><th colspan="${users.length + 1}">${esc(v.L.flow(flowId))}</th></tr>` : '';
      lastFlow = flowId;
      const cells = users
        .map((u) => {
          const c = cell(v, u, flowId, messageId);
          return `<td class="${c.cls}" title="${esc(c.tip)}"><button type="button" data-act="selectUser" data-user="${esc(u.id)}" data-tab="why">${esc(c.mark)}</button></td>`;
        })
        .join('');
      return `${head}<tr><th class="msg" title="${esc(messageId)}">${esc(v.L.message(messageId))}</th>${cells}</tr>`;
    })
    .join('');
  return `<div class="lsp-matrix"><table><thead><tr><th></th>${users
    .map((u) => `<th class="${u.id === v.demo.activeUserId ? 'active' : ''}">${esc(u.name)}</th>`)
    .join('')}</tr></thead><tbody>${rows}</tbody></table></div>
    <div class="lsp-hint">マスにマウスを乗せると理由が出ます。押すと、その人の「なぜ」を開きます。</div>`;
}

// ---------- まとめ ----------

const TABS: { id: Tab; label: string }[] = [
  { id: 'state', label: 'いまの状態' },
  { id: 'why', label: 'なぜ' },
  { id: 'log', label: '配信履歴' },
  { id: 'matrix', label: '誰に届く？' },
];

export function backstage(v: View, user: UserState): string {
  const body = v.ui.tab === 'state' ? stateTab(v, user) : v.ui.tab === 'why' ? whyTab(v, user) : v.ui.tab === 'log' ? logTab(v, user) : matrixTab(v);
  return `<div class="lsp-side">
    ${narrationCard(v, user)}
    <div class="lsp-card-side lsp-tabs-card">
      <div class="lsp-tabs" role="tablist">${TABS.map(
        (t) => `<button type="button" role="tab" aria-selected="${v.ui.tab === t.id}" class="${v.ui.tab === t.id ? 'on' : ''}" data-act="tab" data-tab="${t.id}">${t.label}</button>`,
      ).join('')}</div>
      <div class="lsp-tabbody">${body}</div>
    </div>
  </div>`;
}

/** 資料用表示に出す「なぜ」の直近（届かなかった・止めたを優先） */
export function recentExplanations(v: View, user: UserState, n: number): Explanation[] {
  const mine = v.world.decisions.filter((d) => d.userId === user.id && d.kind !== 'trigger' && d.kind !== 'autoReply');
  const ng = mine.filter((d) => d.outcome !== 'taken');
  const pick = (ng.length ? ng : mine).slice(-n).reverse();
  return pick.map((d) => explainDecision(d, v.L, v.project, v.world));
}
