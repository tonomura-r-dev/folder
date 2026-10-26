// 条件式の評価。必ず判定の木（Trace）を返す。真偽だけを返す関数は作らない。
// and/or は打ち切らずに全部評価する（「何が足りなかったか」を全部見せるため）。

import type { Expr, Id, Minutes, Project, Trace, UserState } from './schema.ts';

export type EvalContext = {
  project: Project;
  user: UserState;
  now: Minutes;
  /** elapsed(since: flowStart) の基準になるフロー */
  flowId?: Id;
};

const leaf = (expr: Expr, value: boolean, actual?: unknown): Trace =>
  actual === undefined ? { op: expr.op, value, expr } : { op: expr.op, value, expr, actual };

const within = (at: Minutes, now: Minutes, withinMinutes?: Minutes) =>
  withinMinutes === undefined || now - at <= withinMinutes;

function compare(actual: unknown, cmp: string, expected: unknown): boolean {
  switch (cmp) {
    case 'eq':
      return actual === expected;
    case 'neq':
      return actual !== expected;
    case 'in':
      return Array.isArray(expected) && expected.includes(actual);
    case 'gte':
      return typeof actual === 'number' && typeof expected === 'number' && actual >= expected;
    case 'lte':
      return typeof actual === 'number' && typeof expected === 'number' && actual <= expected;
    default:
      return false;
  }
}

export function evaluate(expr: Expr, ctx: EvalContext): Trace {
  const { user, now } = ctx;
  switch (expr.op) {
    case 'and':
    case 'or': {
      const children = expr.children.map((c) => evaluate(c, ctx));
      const value = expr.op === 'and' ? children.every((c) => c.value) : children.some((c) => c.value);
      return { op: expr.op, value, expr, children };
    }
    case 'not': {
      const child = evaluate(expr.child, ctx);
      return { op: 'not', value: !child.value, expr, children: [child] };
    }
    case 'always':
      return leaf(expr, true);
    case 'hasTag':
      return leaf(expr, expr.tagId in user.tags, Object.keys(user.tags));
    case 'answer': {
      const got = user.answers[expr.surveyId]?.[expr.questionId] ?? [];
      const hit = expr.choiceIds.filter((c) => got.includes(c));
      const value =
        expr.match === 'any' ? hit.length > 0 : expr.match === 'all' ? hit.length === expr.choiceIds.length : hit.length === 0;
      return leaf(expr, value, got);
    }
    case 'answered': {
      const answers = user.answers[expr.surveyId] ?? {};
      if (expr.questionId !== undefined) return leaf(expr, expr.questionId in answers, Object.keys(answers));
      const survey = ctx.project.surveys.find((s) => s.id === expr.surveyId);
      const value = !!survey && survey.questions.every((q) => q.id in answers);
      return leaf(expr, value, Object.keys(answers));
    }
    case 'attr': {
      const actual = user.attrs[expr.key];
      return leaf(expr, compare(actual, expr.cmp, expr.value), actual ?? null);
    }
    case 'route':
      return leaf(expr, user.route !== null && expr.routeIds.includes(user.route), user.route);
    case 'friend':
      return leaf(expr, user.friendState === expr.state, user.friendState);
    case 'cv': {
      const cv = user.cvs[expr.cvId];
      const done = !!cv && within(cv.at, now, expr.withinMinutes);
      return leaf(expr, done === expr.done, cv ? cv.at : null);
    }
    case 'did': {
      const hits = user.behaviors.filter(
        (b) =>
          b.type === expr.event &&
          within(b.at, now, expr.withinMinutes) &&
          Object.entries(expr.where ?? {}).every(([k, v]) => b.detail[k] === v),
      );
      const value = expr.cmp === 'exists' ? hits.length > 0 : hits.length >= (expr.value ?? 1);
      return leaf(expr, value, hits.length);
    }
    case 'elapsed': {
      let base: Minutes | null = null;
      if (expr.since === 'friendAdd') base = user.friendAddedAt;
      else if (expr.since === 'flowStart') base = ctx.flowId ? (user.flows[ctx.flowId]?.startedAt ?? null) : null;
      else base = user.deliveries.at(-1)?.at ?? null;
      if (base === null) return leaf(expr, false, null);
      const elapsed = now - base;
      return leaf(expr, expr.cmp === 'gte' ? elapsed >= expr.minutes : elapsed < expr.minutes, elapsed);
    }
    case 'delivered': {
      const t = expr.target;
      const hits = user.deliveries.filter(
        (d) =>
          (t.messageId === undefined || d.messageId === t.messageId) &&
          (t.flowId === undefined || d.flowId === t.flowId) &&
          (t.nodeId === undefined || d.nodeId === t.nodeId) &&
          within(d.at, now, expr.withinMinutes),
      );
      return leaf(expr, hits.length > 0 === expr.done, hits.length);
    }
  }
}

/** 判定の木から、結果を決めた「偽の葉」を集める（除外理由の1行目に使う） */
export function failingLeaves(trace: Trace): Trace[] {
  if (!trace.children) return trace.value ? [] : [trace];
  if (trace.op === 'not') {
    // not の中身が真だったことが原因。中身そのものを理由として返す
    return trace.value ? [] : [trace];
  }
  if (trace.value) return [];
  return trace.children.flatMap(failingLeaves);
}
