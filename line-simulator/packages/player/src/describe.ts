// エンジンの判定結果を、案件のラベルを使って日本語にする。
// エンジンは日本語を持たないので、画面・CLI に出す言葉はすべてここで作る。

import type { Decision, Expr, Id, Minutes, Project, Trace, UserState } from '../../engine/src/index.ts';
import { failingLeaves, MINUTES_PER_DAY } from '../../engine/src/index.ts';

export function fmtTime(min: Minutes): string {
  const day = Math.floor(min / MINUTES_PER_DAY);
  const m = min - day * MINUTES_PER_DAY;
  return `Day ${day} ${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
}

export function fmtDuration(min: Minutes): string {
  if (min % MINUTES_PER_DAY === 0) return `${min / MINUTES_PER_DAY}日`;
  if (min % 60 === 0) return `${min / 60}時間`;
  return `${min}分`;
}

export class Labels {
  private readonly p: Project;
  constructor(project: Project) {
    this.p = project;
  }

  tag = (id: Id) => this.p.tags.find((t) => t.id === id)?.label ?? id;
  cv = (id: Id) => this.p.cvs.find((c) => c.id === id)?.label ?? id;
  route = (id: Id | null) => (id === null ? '—' : (this.p.routes.find((r) => r.id === id)?.label ?? id));
  flow = (id: Id | null) => (id === null ? '自動応答' : (this.p.flows.find((f) => f.id === id)?.label ?? id));
  page = (id: Id) => this.p.webPages.find((w) => w.id === id)?.title ?? id;
  question(surveyId: Id, questionId: Id) {
    return this.p.surveys.find((s) => s.id === surveyId)?.questions.find((q) => q.id === questionId);
  }
  questionText = (surveyId: Id, questionId: Id) => this.question(surveyId, questionId)?.text.replace(/を教えてください$/, '') ?? questionId;
  choices = (surveyId: Id, questionId: Id, ids: Id[]) =>
    ids.map((id) => this.question(surveyId, questionId)?.choices.find((c) => c.id === id)?.label ?? id).join('・');
  /** メッセージの1行目（一覧で中身を思い出せる程度） */
  message(id: Id): string {
    if (id.startsWith('image:')) return '（画像）';
    const m = this.p.messages.find((x) => x.id === id);
    const b = m?.bubbles.find((x) => x.kind === 'text' || x.kind === 'buttons' || x.kind === 'card');
    const text = !b ? id : b.kind === 'card' ? b.title : b.kind === 'text' || b.kind === 'buttons' ? b.text : id;
    const line = text.replace(/\{\{[^}]+\}\}/g, '〇〇').split('\n').find((l) => l.trim()) ?? id;
    return line.length > 28 ? `${line.slice(0, 28)}…` : line;
  }
  node(id: Id | null): string {
    if (id === null) return '';
    for (const f of this.p.flows) {
      const n = f.nodes.find((x) => x.id === id);
      if (n) return n.label ?? id;
    }
    return id;
  }
  branchCase(nodeId: Id | null, caseId: string | undefined): string {
    if (caseId === 'else') return 'それ以外';
    for (const f of this.p.flows) {
      const n = f.nodes.find((x) => x.id === nodeId);
      if (n?.type === 'branch') return n.cases.find((c) => c.id === caseId)?.label ?? caseId ?? '';
    }
    return caseId ?? '';
  }
}

const FRIEND = { friend: '友だち', not_friend: '友だちではない', blocked: 'ブロック中' } as const;
const EVENT = {
  visitPage: 'ページを見た',
  exitIntent: '離脱しようとした',
  dismissPopup: 'ポップアップを閉じた',
  submitForm: 'フォームを送信した',
  addFriend: '友だち追加した',
  block: 'ブロックした',
  unblock: 'ブロック解除した',
  tapMenu: 'メニューを押した',
  tapButton: 'ボタンを押した',
  sendText: 'メッセージを送った',
  answer: 'アンケートに答えた',
  cv: 'CVした',
  closePage: 'ページを閉じた',
  tapPopupCta: 'ポップアップのボタンを押した',
  tapPageButton: 'ページのボタンを押した',
  tapLineCta: 'LINE誘導のボタンを押した',
} as const;

export const friendText = (s: UserState['friendState']) => FRIEND[s];

/** 条件式を「〜である」の文にする */
export function describeExpr(e: Expr, L: Labels): string {
  switch (e.op) {
    case 'and':
      return e.children.map((c) => describeExpr(c, L)).join(' かつ ');
    case 'or':
      return `（${e.children.map((c) => describeExpr(c, L)).join(' または ')}）`;
    case 'not':
      return `「${describeExpr(e.child, L)}」ではない`;
    case 'always':
      return '常に';
    case 'hasTag':
      return `タグ「${L.tag(e.tagId)}」が付いている`;
    case 'answer': {
      const q = L.questionText(e.surveyId, e.questionId);
      const c = L.choices(e.surveyId, e.questionId, e.choiceIds);
      return e.match === 'none' ? `${q}が ${c} ではない` : e.match === 'all' ? `${q}に ${c} をすべて選んだ` : `${q}が ${c}`;
    }
    case 'answered':
      return e.questionId ? `${L.questionText(e.surveyId, e.questionId)}に回答済み` : 'アンケートに回答済み';
    case 'attr':
      return `属性 ${e.key} ${{ eq: '＝', neq: '≠', in: '∈', gte: '≧', lte: '≦' }[e.cmp]} ${JSON.stringify(e.value)}`;
    case 'route':
      return `経路が ${e.routeIds.map((r) => L.route(r)).join('・')}`;
    case 'friend':
      return FRIEND[e.state];
    case 'cv':
      return `${e.done ? '' : 'まだ'}${L.cv(e.cvId)}${e.done ? '済み' : 'していない'}${e.withinMinutes ? `（${fmtDuration(e.withinMinutes)}以内）` : ''}`;
    case 'did':
      return `「${EVENT[e.event]}」${e.cmp === 'exists' ? 'ことがある' : `が${e.value ?? 1}回以上`}`;
    case 'elapsed': {
      const since = { friendAdd: '友だち追加', flowStart: 'このシナリオの開始', lastDelivery: '最後の配信' }[e.since];
      return `${since}から${fmtDuration(e.minutes)}${e.cmp === 'gte' ? '以上' : '未満'}`;
    }
    case 'delivered':
      return `${e.target.messageId ? `「${L.message(e.target.messageId)}」` : e.target.flowId ? L.flow(e.target.flowId) : 'その配信'}を${e.done ? '受け取った' : 'まだ受け取っていない'}`;
  }
}

/** 偽になった葉1つを「なぜ当てはまらなかったか」の文にする（実際の値つき） */
function whyFalse(t: Trace, L: Labels): string {
  const e = t.expr;
  switch (e.op) {
    case 'hasTag':
      return `タグ「${L.tag(e.tagId)}」が付いていない`;
    case 'answer': {
      const got = (t.actual as Id[] | undefined) ?? [];
      const q = L.questionText(e.surveyId, e.questionId);
      return got.length === 0
        ? `${q}が未回答`
        : `${q}が ${L.choices(e.surveyId, e.questionId, e.choiceIds)} ではない（回答：${L.choices(e.surveyId, e.questionId, got)}）`;
    }
    case 'cv':
      return e.done ? `まだ${L.cv(e.cvId)}していない` : `すでに${L.cv(e.cvId)}している`;
    case 'friend':
      return `${FRIEND[e.state]}ではない（いま：${FRIEND[(t.actual as UserState['friendState']) ?? 'not_friend']}）`;
    case 'not':
      return `「${describeExpr(e.child, L)}」に当てはまる`;
    default:
      return `「${describeExpr(e, L)}」に当てはまらない`;
  }
}

export function describeFailure(trace: Trace, L: Labels): string {
  return failingLeaves(trace).map((t) => whyFalse(t, L)).join('／');
}

const REASON = {
  blocked: 'ブロック中',
  not_friend: 'まだ友だちではない',
  lower_priority: '同じ時刻に優先度の高い配信があった',
  duplicate: '重複防止期間内に同じ配信を受け取っている',
  stop_condition: '停止条件に当てはまった',
  condition_false: '条件に当てはまらない',
  reentry: 'このシナリオには再度入れない設定',
  stop_all: '別のシナリオで全停止された',
} as const;

/** 判定記録1件を「何が・なぜ」の1行にする */
export function describeDecision(d: Decision, L: Labels, project: Project): string {
  const flow = L.flow(d.flowId);
  switch (d.kind) {
    case 'deliver': {
      const what = `「${L.message(d.ref ?? '')}」`;
      if (d.outcome === 'taken') return `届いた　${what}（${flow}）`;
      return `届かない　${what}（${flow}）理由：${d.reason ? REASON[d.reason] : d.outcome}`;
    }
    case 'branch': {
      const name = L.node(d.nodeId);
      const c = L.branchCase(d.nodeId, d.ref);
      if (d.outcome === 'taken') return `分岐「${name}」→ ${c}（${flow}）`;
      return `分岐「${name}」の「${c}」に進まなかった　理由：${d.trace ? describeFailure(d.trace, L) : REASON.condition_false}`;
    }
    case 'stopWhen': {
      const stopWhen = project.flows.find((f) => f.id === d.flowId)?.stopWhen;
      const cond = stopWhen ? describeExpr(stopWhen, L) : '';
      return d.reason === 'stop_all' ? `停止　${flow}（${REASON.stop_all}）` : `停止　${flow}　理由：停止条件「${cond}」に当てはまった`;
    }
    case 'trigger':
      return d.outcome === 'taken' ? `開始　${flow}` : `開始しない　${flow}　理由：${d.reason ? REASON[d.reason] : d.outcome}`;
    case 'autoReply':
      return d.outcome === 'taken'
        ? `自動応答　${d.ref?.endsWith('#fallback') ? 'どの条件にも当てはまらないので既定の返信' : '条件に当てはまった返信'}`
        : `自動応答の候補を見送り　理由：${d.trace ? describeFailure(d.trace, L) : REASON.condition_false}`;
  }
}
