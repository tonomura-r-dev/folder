// 裏側の表示：説明カード／いまの状態／なぜ（判定の木）／配信履歴／誰に届くか。
// すべて「選んでいるユーザー」の分だけを出す（誰に届くかの表だけは全員を並べる）。

import type { Decision, FlowNode, Id, NarrationDef, Trace, UserState } from '../../engine/src/index.ts';
import { bubbleButtons } from '../../engine/src/index.ts';
import { upcoming } from './controller.ts';
import { describeExpr, describeTraceNode, explainDecision, fmtTime, friendText, type Explanation } from './describe.ts';
import { nextAction } from './journey.ts';
import { flowConditions, formChoices, nextOwn, userHistoryHtml } from './timeline.ts';
import { esc, whoOf, type Tab, type View } from './view.ts';

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

/** 届いたカルーセルごとに「なぜこの中身か」（直前の分岐の条件）と、見た・押した回数 */
function carouselReasons(v: View, user: UserState): string {
  const L = v.L;
  const msgOf = (id: Id) => v.project.messages.find((m) => m.id === id);
  const rows = user.deliveries.filter((d) => msgOf(d.messageId)?.bubbles.some((b) => b.kind === 'carousel'));
  if (!rows.length) return '';
  const items = [...rows]
    .reverse()
    .map((d) => {
      const m = msgOf(d.messageId)!;
      const titles = m.bubbles.flatMap((b) => (b.kind === 'carousel' ? b.cards.map((c) => c.title) : []));
      // 同じ配信の判定（deliver）から遡って、同じシナリオで直前に通った分岐を探す
      const decisions = v.world.decisions;
      const di = decisions.findIndex((x) => x.userId === user.id && x.kind === 'deliver' && x.outcome === 'taken' && x.ref === d.messageId && x.at === d.at);
      let br: Decision | undefined;
      for (let i = di - 1; i >= 0 && d.flowId !== null; i--) {
        const x = decisions[i]!;
        if (x.at !== d.at) break;
        if (x.userId === user.id && x.flowId === d.flowId && x.kind === 'branch' && x.outcome === 'taken') {
          br = x;
          break;
        }
      }
      let why: string;
      if (d.flowId === null) why = '自動応答（メニュー・メッセージへの返信）';
      else if (br) {
        const node = v.project.flows.find((f) => f.id === br!.flowId)?.nodes.find((n) => n.id === br!.nodeId);
        const c = node?.type === 'branch' ? node.cases.find((x) => x.id === br!.ref) : undefined;
        why = c ? `「${node!.label ?? '分岐'}」で「${c.label}」：${describeExpr(c.when, L)}` : `「${node?.label ?? '分岐'}」で「それ以外」（ほかの条件に当てはまらない）`;
      } else why = `「${L.flow(d.flowId)}」の配信（この配信は全員に同じ内容）`;
      const ids = new Set(m.bubbles.flatMap((b) => bubbleButtons(b).map((x) => x.id)));
      // 1枚目は届いた時点で見えている
      const first = m.bubbles.find((b) => b.kind === 'carousel');
      const seen = new Set([first?.kind === 'carousel' ? first.cards[0]?.id : undefined, ...user.behaviors.filter((b) => b.type === 'viewCard' && b.detail['messageId'] === d.messageId && b.at >= d.at).map((b) => b.detail['cardId'])].filter(Boolean));
      const taps = user.behaviors.filter((b) => b.type === 'tapButton' && ids.has(b.detail['buttonId'] ?? '') && b.at >= d.at).length;
      return `<li><div class="lsp-carw-h"><span class="t">${esc(fmtTime(d.at))}</span><b>${esc(titles.join('／'))}</b></div><div class="lsp-carw-why">${esc(why)}</div><div class="lsp-carw-f">見たカード ${seen.size}/${titles.length}枚・ボタン ${taps}回</div></li>`;
    })
    .join('');
  return `<h4>届いたカルーセルと、その中身になった理由（${rows.length}）</h4><ul class="lsp-carw">${items}</ul>`;
}

function stateTab(v: View, user: UserState): string {
  const L = v.L;
  const tags = Object.entries(user.tags);
  const answers = Object.entries(user.answers).flatMap(([sv, qs]) => Object.entries(qs).map(([q, ids]) => `<li><span>${esc(L.questionText(sv, q))}</span><b>${esc(L.choices(sv, q, ids))}</b></li>`));
  const attrs = Object.entries(user.attrs);
  const plans = upcoming(v.world, user.id);
  const nextPlan = nextOwn(v, user) ?? nextAction(v.project, v.world, v.demo, user, L);
  // ユーザーの成果（CV）と、運営側が登録する段階（勤務確定など）は分けて出す
  const cvs = Object.entries(user.cvs).filter(([id]) => !L.isStage(id));
  const active = v.project.flows.map((f) => ({ f, c: user.flows[f.id] })).filter((x) => x.c && (x.c.status === 'running' || x.c.status === 'waiting'));
  const done = v.project.flows.map((f) => ({ f, c: user.flows[f.id] })).filter((x) => x.c && (x.c.status === 'ended' || x.c.status === 'stopped'));
  const activeHtml = active
    .map(({ f, c }) => {
      const next = plans.find((t) => t.flowId === f.id);
      const cond = flowConditions(v.project, L, f.id);
      return `<li class="lsp-fl"><div class="lsp-fl-h"><span class="lsp-pill s-${c!.status}">${FLOW_STATUS[c!.status]}</span><b>${esc(f.label)}</b></div>
        <dl class="lsp-fl-kv"><dt>いま</dt><dd>${esc(nodeLabel(v, f.id, c!.nodeId))}</dd>${next ? `<dt>次回</dt><dd class="strong">${esc(fmtTime(next.fireAt))}</dd>` : ''}<dt>配信条件</dt><dd>${esc(cond.trigger)}</dd>${
          cond.stop ? `<dt>停止条件</dt><dd>${esc(cond.stop)}</dd>` : ''
        }</dl></li>`;
    })
    .join('');
  const doneHtml = done
    .map(({ f, c }) => {
      const stop = c!.status === 'stopped' ? [...v.world.decisions].reverse().find((d) => d.userId === user.id && d.kind === 'stopWhen' && d.flowId === f.id) : undefined;
      const why = stop ? explainDecision(stop, L, v.project).reason : '';
      return `<li><span class="lsp-pill s-${c!.status}">${FLOW_STATUS[c!.status]}</span><span>${esc(f.label)}${c!.endedAt !== undefined ? `<span class="f">${esc(fmtTime(c!.endedAt))}${why ? `　${esc(why)}` : ''}</span>` : ''}</span></li>`;
    })
    .join('');
  // その段階を「完了した状態」に進める。済んだ段階は押せない（二重に記録しない）
  // 何度でも記録できる段階（再応募・イベントごとの勤務確定など）は、済みにしない
  const cvButtons = v.project.cvs
    .map((c) =>
      user.cvs[c.id] && !c.repeatable
        ? `<button type="button" class="done" disabled title="${esc(fmtTime(user.cvs[c.id]!.at))}に済み">✓ ${esc(c.label)}済み</button>`
        : c.demoButton === false
          ? `<button type="button" disabled title="①の画面で操作すると記録されます">${esc(c.label)}（①の画面で操作）</button>`
          : `<button type="button" data-act="cv" data-cv="${esc(c.id)}"${c.kind === 'stage' ? ' class="stage"' : ''}>${esc(c.demoAction ?? `${c.label}へ`)}${c.repeatable && user.cvs[c.id] ? `（${user.cvs[c.id]!.count + 1}回目）` : ''}</button>`,
    )
    .join('');
  // CVの記録・運営側の登録（何回目でも出す。フォームで選んだ求人などの選択項目も添える）
  const logOf = (stage: boolean) =>
    user.behaviors.flatMap((b) => {
      const cvId = b.type === 'cv' ? b.detail['cvId'] : b.type === 'submitForm' ? v.project.webPages.find((p) => p.id === b.detail['pageId'])?.form?.onSubmit.cvId : undefined;
      if (!cvId || L.isStage(cvId) !== stage) return [];
      const picked = b.type === 'submitForm' ? formChoices(v.project, b.detail) : '';
      return [`<li><span class="t">${esc(fmtTime(b.at))}</span><b>${esc(L.cv(cvId))}</b>${picked ? `<span>${esc(picked)}</span>` : ''}</li>`];
    });
  const cvLog = logOf(false);
  const stageLog = logOf(true);
  // タグは種類（グループ）ごとにまとめる（希望エリア・応募したイベント・勤務確定 など）
  const groups = new Map<string, string[]>();
  for (const t of v.project.tags) {
    const got = user.tags[t.id];
    if (!got) continue;
    const g = t.group ?? 'そのほか';
    (groups.get(g) ?? groups.set(g, []).get(g)!).push(`<span class="lsp-tag" title="${esc(fmtTime(got.at))}に付いた">${esc(L.tag(t.id))}</span>`);
  }
  const unknown = tags.filter(([t]) => !v.project.tags.some((x) => x.id === t)).map(([t, x]) => `<span class="lsp-tag" title="${esc(fmtTime(x.at))}に付いた">${esc(L.tag(t))}</span>`);
  if (unknown.length) (groups.get('そのほか') ?? groups.set('そのほか', []).get('そのほか')!).push(...unknown);
  const grouped = groups.size > 1 || (groups.size === 1 && !groups.has('そのほか'));
  const tagsHtml = !tags.length
    ? '<div class="lsp-none">なし</div>'
    : grouped
      ? `<dl class="lsp-taggrp">${[...groups].map(([g, list]) => `<dt>${esc(g)}</dt><dd>${list.join('')}</dd>`).join('')}</dl>`
      : `<div class="lsp-tags">${[...groups.values()].flat().join('')}</div>`;
  return `
    <div class="lsp-tiles">
      <div class="lsp-tile ${user.friendState === 'friend' ? 'ok' : user.friendState === 'blocked' ? 'ng' : ''}"><span class="k">友だち</span><b>${esc(friendText(user.friendState))}</b><small>${esc(L.route(user.route))}</small></div>
      <div class="lsp-tile ${cvs.length ? 'ok' : ''}"><span class="k">CV</span><b>${cvs.length ? cvs.map(([id]) => esc(L.cv(id))).join('・') : 'まだ'}</b><small>${cvs.length ? esc(fmtTime(Math.max(...cvs.map(([, x]) => x.at)))) : 'CV前'}</small></div>
      <div class="lsp-tile next"><span class="k">次回配信</span><b>${esc(nextPlan.when ?? '予定なし')}</b><small>${esc(nextPlan.what)}</small></div>
    </div>
    ${cvLog.length ? `<h4>CVの記録（${cvLog.length}）</h4><ul class="lsp-cvlog">${cvLog.join('')}</ul>` : ''}
    ${stageLog.length ? `<h4>${esc(v.project.brand.companyName.replace(/^株式会社|株式会社$/g, ''))}側の登録（${stageLog.length}）</h4><ul class="lsp-cvlog stage">${stageLog.join('')}</ul>` : ''}
    <h4>タグ（${tags.length}）</h4>${tagsHtml}
    <h4>アンケートの回答</h4>${answers.length ? `<ul class="lsp-answers">${answers.join('')}</ul>` : '<div class="lsp-none">まだ回答していません</div>'}
    ${attrs.length ? `<h4>属性</h4><div class="lsp-answers">${attrs.map(([k, a]) => `${esc(k)}=${esc(a)}`).join('、')}</div>` : ''}
    ${carouselReasons(v, user)}
    <h4>実行中のシナリオ（${active.length}）</h4>${activeHtml ? `<ul class="lsp-flows act">${activeHtml}</ul>` : '<div class="lsp-none">実行中のシナリオはありません</div>'}
    ${doneHtml ? `<details class="lsp-done"><summary>終了・停止したシナリオ（${done.length}）</summary><ul class="lsp-flows">${doneHtml}</ul></details>` : ''}
    <h4>次の段階へ進める</h4><div class="lsp-hint top">各段階に進んだ際のLINE配信や、裏側の処理を確認できます。押すと、その段階を完了した状態になります（フォームは開きません）。</div><div class="lsp-ctrls">${cvButtons}${
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
    .map((u) => `<th class="${u.id === v.demo.activeUserId ? 'active' : ''}">${esc(whoOf(u))}</th>`)
    .join('')}</tr></thead><tbody>${rows}</tbody></table></div>
    <div class="lsp-hint">マスにマウスを乗せると理由が出ます。押すと、その人の「なぜ」を開きます。</div>`;
}

// ---------- まとめ ----------

const TABS: { id: Tab; label: string }[] = [
  { id: 'state', label: 'いまの状態' },
  { id: 'events', label: '出来事' },
  { id: 'why', label: 'なぜ' },
  { id: 'log', label: '配信履歴' },
  { id: 'matrix', label: '誰に届く？' },
];

export function backstage(v: View, user: UserState): string {
  const body = v.ui.tab === 'state' ? stateTab(v, user) : v.ui.tab === 'events' ? `<div class="lsp-hint top">この人に起きたこと（新しい順）。行を押すと裏側の処理をすべて表示します。</div>${userHistoryHtml(v, user)}` : v.ui.tab === 'why' ? whyTab(v, user) : v.ui.tab === 'log' ? logTab(v, user) : matrixTab(v);
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
  // 分岐で選ばれなかった行き先は出さない（選ばれた行き先の方が、なぜこうなったかを表す）
  const mine = v.world.decisions.filter((d) => d.userId === user.id && d.kind !== 'trigger' && d.kind !== 'autoReply' && !(d.kind === 'branch' && d.outcome === 'skipped'));
  // 届かなかった・止まった判定を優先するのは、直近1日のものだけ（古い理由が出続けないように）
  const last = mine.at(-1)?.at ?? 0;
  const ng = mine.filter((d) => d.outcome !== 'taken' && d.at >= last - 1440);
  const pick = (ng.length ? ng : mine).slice(-n).reverse();
  return pick.map((d) => explainDecision(d, v.L, v.project, v.world));
}
