// イベント列 → 状態。状態を直接書き換える入口はここだけ。
// 同じ案件・同じイベント列なら必ず同じ結果になる（現在時刻・乱数を使わない）。
//
// 処理の順番（docs/DATA_MODEL.md §7.3）
//   1. イベントの時刻までに来るタイマーを先に処理する（時刻 → ユーザーID → フロー優先度 → ノードID の順）
//   2. イベントを処理する
//   3. その場で起きた後続処理（別フローの起動など）を、起きた順に処理しきる
//   4. 停止条件を満たしたフローを止める

import type {
  Action,
  Bubble,
  Decision,
  DecisionReason,
  Expr,
  Flow,
  FlowNode,
  Id,
  InputEvent,
  InputEventBody,
  Minutes,
  Project,
  Session,
  Timer,
  Trace,
  Trigger,
  UserState,
  WorldState,
} from './schema.ts';
import { MINUTES_PER_DAY } from './schema.ts';
import { evaluate } from './evaluate.ts';
import { buildIndex, byId, edgeKey, type ProjectIndex } from './indexing.ts';
import { renderBubble } from './text.ts';

/** 1回のイベント（またはタイマー1回分の時間進行）で実行してよいノード数の上限 */
export const STEP_BUDGET = 10_000;

type Ctx = {
  project: Project;
  index: ProjectIndex;
  world: WorldState;
  queue: (() => void)[];
  steps: number;
  /** 上限を超えた後は、そのイベントの残りを打ち切る */
  aborted: boolean;
};

export type ReduceOptions = { until?: Minutes };

export function reduce(project: Project, session: Session, options: ReduceOptions = {}): WorldState {
  const ctx: Ctx = {
    project,
    index: buildIndex(project),
    world: { now: 0, users: {}, timers: [], decisions: [], warnings: [] },
    queue: [],
    steps: 0,
    aborted: false,
  };
  initUsers(ctx, session);

  // 同時刻のイベントは入力順を保つ（安定ソート）
  const events = session.events.map((e, i) => ({ e, i })).sort((a, b) => a.e.at - b.e.at || a.i - b.i);
  for (const { e } of events) {
    if (options.until !== undefined && e.at > options.until) break;
    runTimersUntil(ctx, e.at);
    if (e.type === 'advanceTo') continue;
    begin(ctx);
    applyEvent(ctx, e);
    runTimersUntil(ctx, ctx.world.now);
  }
  if (options.until !== undefined) runTimersUntil(ctx, options.until);
  return ctx.world;
}

// ---------- 初期化 ----------

function initUsers(ctx: Ctx, session: Session) {
  const { project, world } = ctx;
  for (const u of [...session.users].sort((a, b) => byId(a.id, b.id))) {
    const persona = project.personas.find((p) => p.id === u.personaId);
    const friend = persona?.initial?.friendState ?? 'not_friend';
    const user: UserState = {
      id: u.id,
      personaId: u.personaId,
      name: u.displayName ?? persona?.name ?? u.id,
      attrs: { ...(persona?.attrs ?? {}) },
      friendState: friend,
      friendAddedAt: friend === 'not_friend' ? null : 0,
      route: null,
      tags: {},
      answers: {},
      cvs: {},
      behaviors: [],
      deliveries: [],
      flows: {},
      richMenuId: friend === 'not_friend' ? null : (project.settings.defaultRichMenuId ?? null),
      surface: { kind: 'talk' },
      talk: [],
      messageCount: 0,
    };
    for (const t of persona?.initial?.tags ?? []) user.tags[t] = { at: 0 };
    for (const c of persona?.initial?.cvs ?? []) user.cvs[c] = { at: 0, count: 1 };
    world.users[u.id] = user;
  }
}

/** 入力イベントから、ペルソナの台本どおりのセッションを作る */
export function sessionFromPersonas(project: Project, personaIds?: Id[]): Session {
  const personas = project.personas.filter((p) => !personaIds || personaIds.includes(p.id));
  const events: InputEvent[] = [];
  for (const p of personas)
    for (const s of p.script ?? []) events.push({ ...s.event, at: s.atMinutes, userId: p.id } as InputEvent);
  return {
    schemaVersion: 1,
    projectId: project.meta.id,
    users: personas.map((p) => ({ id: p.id, personaId: p.id })),
    events,
  };
}

// ---------- 時間進行 ----------

function begin(ctx: Ctx) {
  ctx.steps = 0;
  ctx.aborted = false;
}

function timerOrder(ctx: Ctx) {
  return (a: Timer, b: Timer) =>
    a.fireAt - b.fireAt ||
    byId(a.userId, b.userId) ||
    (ctx.index.flows.get(a.flowId)?.priority ?? 0) - (ctx.index.flows.get(b.flowId)?.priority ?? 0) ||
    byId(a.flowId, b.flowId) ||
    byId(a.nodeId, b.nodeId);
}

function runTimersUntil(ctx: Ctx, t: Minutes) {
  const { world } = ctx;
  const order = timerOrder(ctx);
  for (;;) {
    let next: Timer | undefined;
    for (const timer of world.timers) if (timer.fireAt <= t && (!next || order(timer, next) < 0)) next = timer;
    if (!next) break;
    world.timers.splice(world.timers.indexOf(next), 1);
    world.now = Math.max(world.now, next.fireAt);
    begin(ctx);
    fireTimer(ctx, next);
  }
  world.now = Math.max(world.now, t);
}

function fireTimer(ctx: Ctx, timer: Timer) {
  const user = ctx.world.users[timer.userId];
  const flow = ctx.index.flows.get(timer.flowId);
  const cursor = user?.flows[timer.flowId];
  if (!user || !flow || !cursor || cursor.status !== 'waiting' || cursor.nodeId !== timer.nodeId) return;
  dropTimers(ctx, user.id, flow.id);
  delete cursor.awaitEvent;
  const port = timer.kind === 'wait' ? 'out' : 'timeout';
  ctx.queue.push(() => follow(ctx, user, flow, timer.nodeId, port));
  settle(ctx, user);
}

function addTimer(ctx: Ctx, user: UserState, flow: Flow, node: FlowNode, fireAt: Minutes, kind: Timer['kind']) {
  ctx.world.timers.push({ userId: user.id, flowId: flow.id, nodeId: node.id, fireAt: Math.max(fireAt, ctx.world.now), kind });
}

function dropTimers(ctx: Ctx, userId: Id, flowId: Id) {
  ctx.world.timers = ctx.world.timers.filter((t) => !(t.userId === userId && t.flowId === flowId));
}

/** 後続処理を処理しきり、停止条件を確かめる */
function settle(ctx: Ctx, user: UserState) {
  drain(ctx);
  checkStops(ctx, user);
  drain(ctx);
}

function drain(ctx: Ctx) {
  while (ctx.queue.length) ctx.queue.shift()!();
}

// ---------- 判定記録 ----------

function decide(ctx: Ctx, user: UserState, d: Omit<Decision, 'at' | 'userId'>) {
  ctx.world.decisions.push({ at: ctx.world.now, userId: user.id, ...d });
}

function evalIn(ctx: Ctx, user: UserState, expr: Expr, flowId?: Id): Trace {
  return evaluate(expr, flowId === undefined
    ? { project: ctx.project, user, now: ctx.world.now }
    : { project: ctx.project, user, now: ctx.world.now, flowId });
}

// ---------- フローの実行 ----------

function startFlow(ctx: Ctx, user: UserState, flow: Flow) {
  const cur = user.flows[flow.id];
  if (cur) {
    const active = cur.status === 'running' || cur.status === 'waiting';
    if (flow.reentry === 'never' || (flow.reentry === 'afterEnd' && active)) {
      decide(ctx, user, { flowId: flow.id, nodeId: null, kind: 'trigger', outcome: 'skipped', reason: 'reentry' });
      return;
    }
    dropTimers(ctx, user.id, flow.id);
  }
  const entry = flow.nodes.find((n) => n.type === 'entry');
  if (!entry) {
    ctx.world.warnings.push({ at: ctx.world.now, code: 'runtime.unknownRef', userId: user.id, ref: flow.id });
    return;
  }
  user.flows[flow.id] = { status: 'running', nodeId: entry.id, startedAt: ctx.world.now, runs: (cur?.runs ?? 0) + 1 };
  decide(ctx, user, { flowId: flow.id, nodeId: entry.id, kind: 'trigger', outcome: 'taken' });
  run(ctx, user, flow, entry.id);
}

function follow(ctx: Ctx, user: UserState, flow: Flow, fromNodeId: Id, port: string) {
  const edge = ctx.index.edges.get(edgeKey(flow.id, fromNodeId, port));
  if (!edge) return finish(ctx, user, flow, 'ended');
  run(ctx, user, flow, edge.to);
}

function run(ctx: Ctx, user: UserState, flow: Flow, nodeId: Id) {
  const cursor = user.flows[flow.id]!;
  let current: Id | undefined = nodeId;
  while (current !== undefined) {
    if (ctx.aborted) return;
    if (++ctx.steps > STEP_BUDGET) {
      ctx.aborted = true;
      ctx.world.warnings.push({ at: ctx.world.now, code: 'runtime.stepBudget', userId: user.id, ref: `${flow.id}/${current}` });
      finish(ctx, user, flow, 'stopped');
      return;
    }
    cursor.status = 'running';
    cursor.nodeId = current;
    if (flow.stopWhen) {
      const trace = evalIn(ctx, user, flow.stopWhen, flow.id);
      if (trace.value) return stop(ctx, user, flow, 'stop_condition', trace);
    }
    const found = ctx.index.nodes.get(current);
    if (!found || found.flow.id !== flow.id) {
      ctx.world.warnings.push({ at: ctx.world.now, code: 'runtime.unknownRef', userId: user.id, ref: `${flow.id}/${current}` });
      return finish(ctx, user, flow, 'stopped');
    }
    const port = exec(ctx, user, flow, found.node);
    if (port === null) return;
    const edge = ctx.index.edges.get(edgeKey(flow.id, current, port));
    if (!edge) return finish(ctx, user, flow, 'ended');
    current = edge.to;
  }
}

function finish(ctx: Ctx, user: UserState, flow: Flow, status: 'ended' | 'stopped') {
  const cursor = user.flows[flow.id];
  if (!cursor) return;
  cursor.status = status;
  cursor.endedAt = ctx.world.now;
  delete cursor.awaitEvent;
  dropTimers(ctx, user.id, flow.id);
}

function stop(ctx: Ctx, user: UserState, flow: Flow, reason: DecisionReason, trace?: Trace) {
  const cursor = user.flows[flow.id];
  decide(ctx, user, {
    flowId: flow.id,
    nodeId: cursor?.nodeId ?? null,
    kind: 'stopWhen',
    outcome: 'stopped',
    reason,
    ...(trace ? { trace } : {}),
  });
  finish(ctx, user, flow, 'stopped');
}

/** 待機中のフローの停止条件を確かめる（応募した瞬間にタイマーを破棄するため） */
function checkStops(ctx: Ctx, user: UserState) {
  for (const flow of ctx.index.flowOrder) {
    const cursor = user.flows[flow.id];
    if (!flow.stopWhen || !cursor || cursor.status !== 'waiting') continue;
    const trace = evalIn(ctx, user, flow.stopWhen, flow.id);
    if (trace.value) stop(ctx, user, flow, 'stop_condition', trace);
  }
}

/** ノードを1つ実行し、次に進むポートを返す。待機・終了なら null */
function exec(ctx: Ctx, user: UserState, flow: Flow, node: FlowNode): string | null {
  const cursor = user.flows[flow.id]!;
  const now = ctx.world.now;
  switch (node.type) {
    case 'entry':
      return 'out';
    case 'message':
      deliver(ctx, user, flow, node, node.messageId);
      return 'out';
    case 'image':
      deliverImage(ctx, user, flow, node, node.assetId);
      return 'out';
    case 'richMenu':
      user.richMenuId = node.richMenuId;
      return 'out';
    case 'choice':
      deliver(ctx, user, flow, node, node.messageId);
      cursor.status = 'waiting';
      if (node.timeoutMinutes !== undefined) addTimer(ctx, user, flow, node, now + node.timeoutMinutes, 'timeout');
      return null;
    case 'survey': {
      const survey = ctx.index.surveys.get(node.surveyId);
      const answers = user.answers[node.surveyId] ?? {};
      const next = survey?.questions.find((q) => !(q.id in answers));
      if (!next) return 'done';
      if (node.introMessageId) deliver(ctx, user, flow, node, node.introMessageId);
      cursor.status = 'waiting';
      if (user.friendState === 'friend') {
        user.surface = { kind: 'survey', surveyId: node.surveyId, questionId: next.id };
        user.talk.push({ at: now, side: 'system', code: 'surveyShown', params: { surveyId: node.surveyId } });
      }
      if (node.timeoutMinutes !== undefined) addTimer(ctx, user, flow, node, now + node.timeoutMinutes, 'timeout');
      return null;
    }
    case 'tag':
      for (const t of node.remove ?? []) delete user.tags[t];
      for (const t of node.add ?? []) addTag(ctx, user, t, node.id);
      return 'out';
    case 'branch': {
      for (const c of node.cases) {
        const trace = evalIn(ctx, user, c.when, flow.id);
        decide(ctx, user, {
          flowId: flow.id,
          nodeId: node.id,
          kind: 'branch',
          outcome: trace.value ? 'taken' : 'skipped',
          ref: c.id,
          trace,
          ...(trace.value ? {} : { reason: 'condition_false' as const }),
        });
        if (trace.value) return `case:${c.id}`;
      }
      decide(ctx, user, { flowId: flow.id, nodeId: node.id, kind: 'branch', outcome: 'taken', ref: 'else' });
      return 'else';
    }
    case 'wait': {
      const w = node.wait;
      cursor.status = 'waiting';
      if (w.kind === 'duration') addTimer(ctx, user, flow, node, now + w.minutes, 'wait');
      else if (w.kind === 'dayAt') {
        const base = w.from === 'flowStart' || user.friendAddedAt === null ? cursor.startedAt : user.friendAddedAt;
        addTimer(ctx, user, flow, node, dayAt(base, w.day, w.time), 'wait');
      } else {
        cursor.awaitEvent = w.event;
        addTimer(ctx, user, flow, node, now + w.timeoutMinutes, 'timeout');
      }
      return null;
    }
    case 'page':
      user.surface = { kind: 'page', pageId: node.pageId };
      user.talk.push({ at: now, side: 'system', code: 'pageOpened', params: { pageId: node.pageId } });
      return 'out';
    case 'jump': {
      finish(ctx, user, flow, 'ended');
      const target = ctx.index.flows.get(node.flowId);
      if (target) ctx.queue.push(() => startFlow(ctx, user, target));
      return null;
    }
    case 'end':
      finish(ctx, user, flow, 'ended');
      return null;
    case 'stopAll':
      for (const f of ctx.index.flowOrder) {
        const c = user.flows[f.id];
        if (f.id !== flow.id && c && (c.status === 'running' || c.status === 'waiting')) stop(ctx, user, f, 'stop_all');
      }
      finish(ctx, user, flow, 'ended');
      return null;
  }
}

export function dayAt(base: Minutes, day: number, time: string): Minutes {
  const m = /^(\d{1,2}):(\d{2})$/.exec(time);
  const hm = m ? Number(m[1]) * 60 + Number(m[2]) : 0;
  return (Math.floor(base / MINUTES_PER_DAY) + day) * MINUTES_PER_DAY + hm;
}

// ---------- 配信 ----------

/** 配信の判定順：友だちか → 重複防止 → 優先度。通ったら配信する */
function gate(ctx: Ctx, user: UserState, flow: Flow | null, node: FlowNode | null, key: string): boolean {
  const base = { flowId: flow?.id ?? null, nodeId: node?.id ?? null, kind: 'deliver' as const, ref: key };
  if (user.friendState !== 'friend') {
    decide(ctx, user, { ...base, outcome: 'excluded', reason: user.friendState === 'blocked' ? 'blocked' : 'not_friend' });
    return false;
  }
  const dedupe = ctx.project.settings.dedupeWindowMinutes;
  if (dedupe > 0 && user.deliveries.some((d) => d.messageId === key && ctx.world.now - d.at < dedupe)) {
    decide(ctx, user, { ...base, outcome: 'deduped', reason: 'duplicate' });
    return false;
  }
  if (flow && flow.concurrency === 'exclusive') {
    const blocked = user.deliveries.some((d) => {
      if (d.at !== ctx.world.now || d.flowId === null || d.flowId === flow.id) return false;
      const other = ctx.index.flows.get(d.flowId);
      return !!other && other.concurrency === 'exclusive' && other.priority < flow.priority;
    });
    if (blocked) {
      decide(ctx, user, { ...base, outcome: 'deferred', reason: 'lower_priority' });
      return false;
    }
  }
  return true;
}

function record(ctx: Ctx, user: UserState, flow: Flow | null, node: FlowNode | null, key: string, bubbles: Bubble[], count: number) {
  const now = ctx.world.now;
  user.deliveries.push({ at: now, flowId: flow?.id ?? null, nodeId: node?.id ?? null, messageId: key, count });
  for (const b of bubbles)
    user.talk.push({ at: now, side: 'bot', bubble: renderBubble(b, ctx.project, user), messageId: key, ...(node ? { fromNodeId: node.id } : {}) });
  user.messageCount += count;
  decide(ctx, user, { flowId: flow?.id ?? null, nodeId: node?.id ?? null, kind: 'deliver', outcome: 'taken', ref: key });
}

function deliver(ctx: Ctx, user: UserState, flow: Flow | null, node: FlowNode | null, messageId: Id, count?: number): boolean {
  const msg = ctx.index.messages.get(messageId);
  if (!msg) {
    ctx.world.warnings.push({ at: ctx.world.now, code: 'runtime.unknownRef', userId: user.id, ref: messageId });
    return false;
  }
  if (!gate(ctx, user, flow, node, messageId)) return false;
  record(ctx, user, flow, node, messageId, msg.bubbles, count ?? msg.countAs ?? 1);
  return true;
}

function deliverImage(ctx: Ctx, user: UserState, flow: Flow, node: FlowNode, assetId: Id) {
  const key = `image:${assetId}`;
  if (gate(ctx, user, flow, node, key)) record(ctx, user, flow, node, key, [{ kind: 'image', assetId }], 1);
}

// ---------- トリガー・タグ・CV ----------

function fire(ctx: Ctx, user: UserState, match: (t: Trigger) => boolean) {
  for (const flow of ctx.index.flowOrder) if (match(flow.trigger)) ctx.queue.push(() => startFlow(ctx, user, flow));
}

function addTag(ctx: Ctx, user: UserState, tagId: Id, byNodeId?: Id) {
  if (tagId in user.tags) return;
  user.tags[tagId] = byNodeId === undefined ? { at: ctx.world.now } : { at: ctx.world.now, byNodeId };
  fire(ctx, user, (t) => t.kind === 'tagAdded' && t.tagId === tagId);
}

function doCv(ctx: Ctx, user: UserState, cvId: Id) {
  const prev = user.cvs[cvId];
  user.cvs[cvId] = { at: ctx.world.now, count: (prev?.count ?? 0) + 1 };
  fire(ctx, user, (t) => t.kind === 'cv' && t.cvId === cvId);
}

// ---------- 入力イベント ----------

function detailOf(ev: InputEventBody): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(ev)) {
    if (k === 'type' || k === 'at' || k === 'userId') continue;
    if (typeof v === 'string') out[k] = v;
    else if (typeof v === 'number') out[k] = String(v);
    else if (Array.isArray(v)) out[k] = v.join(',');
  }
  return out;
}

function applyEvent(ctx: Ctx, ev: InputEvent) {
  const user = ctx.world.users[ev.userId];
  if (!user) {
    ctx.world.warnings.push({ at: ev.at, code: 'runtime.unknownRef', userId: ev.userId, ref: ev.userId });
    return;
  }
  ctx.world.now = Math.max(ctx.world.now, ev.at);
  handle(ctx, user, ev);

  // このイベントを待っていたフロー（untilEvent）を進める
  for (const flow of ctx.index.flowOrder) {
    const cursor = user.flows[flow.id];
    if (!cursor || cursor.status !== 'waiting' || cursor.awaitEvent !== ev.type || cursor.nodeId === null) continue;
    const nodeId = cursor.nodeId;
    dropTimers(ctx, user.id, flow.id);
    delete cursor.awaitEvent;
    ctx.queue.push(() => follow(ctx, user, flow, nodeId, 'out'));
  }
  settle(ctx, user);
}

/** 入力イベント1件を処理する。アクション経由（ボタン→ページ等）でも同じ関数を通す */
function handle(ctx: Ctx, user: UserState, ev: InputEventBody): void {
  const { index } = ctx;
  const now = ctx.world.now;
  user.behaviors.push({ at: now, type: ev.type, detail: detailOf(ev) });

  switch (ev.type) {
    case 'visitPage':
      if (index.pages.has(ev.pageId)) user.surface = { kind: 'page', pageId: ev.pageId };
      return;
    case 'exitIntent': {
      if (index.pages.get(ev.pageId)?.exitPopup) user.surface = { kind: 'popup', pageId: ev.pageId };
      const pageId = ev.pageId;
      fire(ctx, user, (t) => t.kind === 'pageExit' && t.pageId === pageId);
      return;
    }
    case 'dismissPopup':
      user.surface = { kind: 'page', pageId: ev.pageId };
      return;
    case 'closePage':
      user.surface = { kind: 'talk' };
      return;
    case 'tapPopupCta': {
      const popup = index.pages.get(ev.pageId)?.exitPopup;
      if (!popup) return;
      // ポップアップを閉じてから動作する（友だち追加ならトーク画面に切り替わる）
      user.surface = { kind: 'page', pageId: ev.pageId };
      return act(ctx, user, popup.cta.action);
    }
    case 'tapPageButton': {
      const block = index.pages.get(ev.pageId)?.blocks[ev.blockIndex];
      if (block?.kind === 'button') act(ctx, user, block.action);
      return;
    }
    case 'tapLineCta': {
      const cta = index.pages.get(ev.pageId)?.lineCta;
      if (cta) handle(ctx, user, { type: 'addFriend', routeId: cta.routeId });
      return;
    }
    case 'submitForm': {
      const page = index.pages.get(ev.pageId);
      const onSubmit = page?.form?.onSubmit;
      if (onSubmit?.nextPageId) user.surface = { kind: 'page', pageId: onSubmit.nextPageId };
      if (onSubmit?.cvId) doCv(ctx, user, onSubmit.cvId);
      const pageId = ev.pageId;
      fire(ctx, user, (t) => t.kind === 'formSubmit' && t.pageId === pageId);
      return;
    }
    case 'addFriend': {
      if (user.friendState === 'friend') return;
      if (user.friendState === 'blocked') {
        user.friendState = 'friend';
        user.talk.push({ at: now, side: 'system', code: 'unblocked' });
        return;
      }
      user.friendState = 'friend';
      user.friendAddedAt = now;
      user.route = ev.routeId;
      user.richMenuId = ctx.project.settings.defaultRichMenuId ?? null;
      user.surface = { kind: 'talk' };
      user.talk.push({ at: now, side: 'system', code: 'friendAdded', params: { routeId: ev.routeId } });
      const routeId = ev.routeId;
      fire(ctx, user, (t) => t.kind === 'friendAdd' && (!t.routeIds || t.routeIds.includes(routeId)));
      return;
    }
    case 'block':
      if (user.friendState !== 'friend') return;
      user.friendState = 'blocked';
      user.talk.push({ at: now, side: 'system', code: 'blocked' });
      return;
    case 'unblock':
      if (user.friendState !== 'blocked') return;
      user.friendState = 'friend';
      user.talk.push({ at: now, side: 'system', code: 'unblocked' });
      return;
    case 'tapMenu': {
      const cell = index.richMenus.get(ev.richMenuId)?.cells.find((c) => c.id === ev.cellId);
      if (cell) act(ctx, user, cell.action);
      const { richMenuId, cellId } = ev;
      fire(ctx, user, (t) => t.kind === 'menuTap' && t.richMenuId === richMenuId && t.cellId === cellId);
      return;
    }
    case 'tapButton': {
      const button = index.buttons.get(ev.buttonId);
      if (button) act(ctx, user, button.action);
      // ボタン選択ノードで待っているなら、押したボタンの出口へ進む
      const found = ev.nodeId === undefined ? undefined : index.nodes.get(ev.nodeId);
      if (found && found.node.type === 'choice') {
        const { flow, node } = found;
        const cursor = user.flows[flow.id];
        if (cursor && cursor.status === 'waiting' && cursor.nodeId === node.id) {
          dropTimers(ctx, user.id, flow.id);
          ctx.queue.push(() => follow(ctx, user, flow, node.id, `btn:${ev.buttonId}`));
        }
      }
      return;
    }
    case 'sendText':
      user.talk.push({ at: now, side: 'user', bubble: { kind: 'text', text: ev.text } });
      autoReply(ctx, user, ev.text);
      return;
    case 'answer':
      answer(ctx, user, ev.surveyId, ev.questionId, ev.choiceIds);
      return;
    case 'cv':
      doCv(ctx, user, ev.cvId);
      return;
  }
}

/** ボタン・メニュー等の動作。ユーザーの操作として行動履歴にも残す */
function act(ctx: Ctx, user: UserState, action: Action): void {
  switch (action.kind) {
    case 'sendText':
      return handle(ctx, user, { type: 'sendText', text: action.text });
    case 'openPage':
      return handle(ctx, user, { type: 'visitPage', pageId: action.pageId });
    case 'answer':
      return handle(ctx, user, { type: 'answer', surveyId: action.surveyId, questionId: action.questionId, choiceIds: [action.choiceId] });
    case 'addFriend':
      return handle(ctx, user, { type: 'addFriend', routeId: action.routeId });
    case 'startFlow': {
      const flow = ctx.index.flows.get(action.flowId);
      if (flow) ctx.queue.push(() => startFlow(ctx, user, flow));
      return;
    }
    case 'switchMenu':
      user.richMenuId = action.richMenuId;
      return;
    case 'none':
      return;
  }
}

function autoReply(ctx: Ctx, user: UserState, text: string) {
  const rules = [...ctx.project.autoReplies].sort((a, b) => a.priority - b.priority || byId(a.id, b.id));
  const rule = rules.find((r) =>
    r.match.keywords.some((k) => (r.match.mode === 'exact' ? text.trim() === k : text.includes(k))),
  );
  if (!rule) return;
  // 応答メッセージは通数に数えない（LINEの応答メッセージは無料のため）
  let replied = false;
  for (const [i, c] of rule.cases.entries()) {
    const trace = evalIn(ctx, user, c.when);
    decide(ctx, user, {
      flowId: null,
      nodeId: null,
      kind: 'autoReply',
      outcome: trace.value ? 'taken' : 'skipped',
      ref: `${rule.id}#${i}`,
      trace,
      ...(trace.value ? {} : { reason: 'condition_false' as const }),
    });
    if (trace.value) {
      deliver(ctx, user, null, null, c.messageId, 0);
      replied = true;
      break;
    }
  }
  if (!replied && rule.fallbackMessageId) {
    decide(ctx, user, { flowId: null, nodeId: null, kind: 'autoReply', outcome: 'taken', ref: `${rule.id}#fallback` });
    deliver(ctx, user, null, null, rule.fallbackMessageId, 0);
  }
  const ruleId = rule.id;
  fire(ctx, user, (t) => t.kind === 'keyword' && t.autoReplyId === ruleId);
}

function answer(ctx: Ctx, user: UserState, surveyId: Id, questionId: Id, choiceIds: Id[]) {
  const survey = ctx.index.surveys.get(surveyId);
  const question = survey?.questions.find((q) => q.id === questionId);
  if (!survey || !question) return;
  const chosen = question.mode === 'single' ? choiceIds.slice(0, 1) : [...choiceIds];
  (user.answers[surveyId] ??= {})[questionId] = chosen;

  const labels = chosen.map((id) => question.choices.find((c) => c.id === id)?.label ?? id);
  user.talk.push({ at: ctx.world.now, side: 'user', bubble: { kind: 'text', text: labels.join(' / ') } });

  for (const c of question.choices) {
    if (!chosen.includes(c.id)) continue;
    for (const t of c.removeTags ?? []) delete user.tags[t];
    for (const t of c.addTags ?? []) addTag(ctx, user, t);
  }

  const answers = user.answers[surveyId]!;
  const next = survey.questions.find((q) => !(q.id in answers));
  if (user.surface.kind === 'survey' && user.surface.surveyId === surveyId)
    user.surface = next ? { kind: 'survey', surveyId, questionId: next.id } : { kind: 'talk' };
  if (next) return;

  // 全問答えたら、このアンケートで待っていたフローを進める
  for (const flow of ctx.index.flowOrder) {
    const cursor = user.flows[flow.id];
    if (!cursor || cursor.status !== 'waiting' || cursor.nodeId === null) continue;
    const node = ctx.index.nodes.get(cursor.nodeId)?.node;
    if (node?.type !== 'survey' || node.surveyId !== surveyId) continue;
    const nodeId = node.id;
    dropTimers(ctx, user.id, flow.id);
    ctx.queue.push(() => follow(ctx, user, flow, nodeId, 'done'));
  }
}
