// デモ画面の③「CVまでの道筋」。営業先に見せる説明用の表示で、編集用のフローチャートではない。
// 案件の journey（無ければシナリオ一覧から作る簡易版）を、選んでいるユーザーの状態に当てはめて
// 「どこまで来たか・いまどこか・次に何が起きるか・この人はどのルートか」を出す。

import type { Expr, FlowNode, Id, JourneyStep, Minutes, Project, UserState, WorldState } from '../../engine/src/index.ts';
import { evaluate } from '../../engine/src/index.ts';
import { pendingInputs, upcoming, type Demo } from './controller.ts';
import { fmtTime, type Labels } from './describe.ts';
import { esc, type View } from './view.ts';

export type StepStatus = 'done' | 'current' | 'skipped' | 'todo';
export type JourneyRow = {
  step: JourneyStep;
  status: StepStatus;
  /** 来た時刻（条件式だけで判定した段は分からないので無し） */
  at?: Minutes;
  /** この人のルート（variants のうち当てはまったもの） */
  variantId?: Id;
  isGoal: boolean;
};

/** 案件の道筋。無ければシナリオの並び順とCVから簡易版を作る */
export function journeySteps(project: Project): { steps: JourneyStep[]; auto: boolean } {
  if (project.journey?.length) return { steps: project.journey, auto: false };
  const steps: JourneyStep[] = project.flows.map((f) => ({ id: `auto:${f.id}`, label: f.label, nodeIds: f.nodes.map((n) => n.id) }));
  for (const c of project.cvs) steps.push({ id: `auto:cv:${c.id}`, label: c.label, cvId: c.id });
  return { steps, auto: true };
}

const holds = (expr: Expr, project: Project, user: UserState, now: Minutes) => evaluate(expr, { project, user, now }).value;

function reachedAt(step: JourneyStep, project: Project, user: UserState, now: Minutes): { reached: boolean; at?: Minutes } {
  let at: Minutes | undefined;
  const ids = new Set(step.nodeIds ?? []);
  if (ids.size) at = user.trail.find((t) => ids.has(t.nodeId))?.at;
  if (step.cvId !== undefined) {
    const cv = user.cvs[step.cvId];
    if (cv && (at === undefined || cv.at < at)) at = cv.at;
  }
  if (at !== undefined) return { reached: true, at };
  return { reached: !!step.reachedWhen && holds(step.reachedWhen, project, user, now) };
}

/** 段ごとの状態。いちばん先まで来た段が「いまここ」、そこより手前で来ていない段は「通らず」 */
export function journeyOf(project: Project, world: WorldState, user: UserState): JourneyRow[] {
  const { steps } = journeySteps(project);
  const goalIndex = steps.reduce((g, s, i) => (s.cvId !== undefined ? i : g), -1);
  const reach = steps.map((s) => reachedAt(s, project, user, world.now));
  const current = reach.reduce((c, r, i) => (r.reached ? i : c), -1);
  return steps.map((step, i) => {
    const r = reach[i]!;
    const status: StepStatus = i === current ? 'current' : i < current ? (r.reached ? 'done' : 'skipped') : 'todo';
    const variantId = step.variants?.find((x) => holds(x.when, project, user, world.now))?.id;
    return { step, status, isGoal: i === goalIndex, ...(r.at !== undefined ? { at: r.at } : {}), ...(variantId ? { variantId } : {}) };
  });
}

const QUIET_END = 'このシナリオが終わります';
const QUIET_WAIT = 'さらに待機';

/** 待機の次に通るノードをたどって、何が起きるかを一言にする */
function whatComesNext(project: Project, L: Labels, flowId: Id, nodeId: Id, port: string): string {
  const flow = project.flows.find((f) => f.id === flowId);
  if (!flow) return '';
  let from = nodeId;
  let p = port;
  for (let i = 0; i < 20; i++) {
    const edge = flow.edges.find((e) => e.from.nodeId === from && e.from.port === p);
    const n: FlowNode | undefined = edge && flow.nodes.find((x) => x.id === edge.to);
    if (!n) return QUIET_END;
    if (n.type === 'message' || n.type === 'choice') return `「${L.message(n.messageId)}」を配信`;
    if (n.type === 'image') return '画像を配信';
    if (n.type === 'survey') return 'アンケートを表示';
    if (n.type === 'branch') return `「${n.label ?? '分岐'}」で行き先を判定`;
    if (n.type === 'jump') return `「${L.flow(n.flowId)}」へ移る`;
    if (n.type === 'end' || n.type === 'stopAll') return QUIET_END;
    if (n.type === 'wait') return QUIET_WAIT;
    from = n.id;
    p = 'out';
  }
  return '';
}

/** 次に実行される施策（この人の） */
export function nextAction(project: Project, world: WorldState, demo: Demo, user: UserState, L: Labels): { when?: string; what: string } {
  if (user.friendState === 'blocked') return { what: 'ブロック中のため、この先の配信は届きません' };
  const pending = pendingInputs(project, world, demo).find((p) => p.userId === user.id);
  if (pending) return { what: pending.kind === 'survey' ? 'アンケートの回答待ち（答えると、回答に合わせて次へ進みます）' : 'ボタンの選択待ち（押したボタンで次が決まります）' };
  // 「シナリオが終わるだけ」「さらに待つだけ」の予定は飛ばし、お客様に見せる意味のある次の動きを出す
  const plans = upcoming(world, user.id).map((t) => ({ t, what: whatComesNext(project, L, t.flowId, t.nodeId, t.kind === 'timeout' ? 'timeout' : 'out') }));
  const p = plans.find((x) => x.what !== QUIET_END && x.what !== QUIET_WAIT) ?? plans[0];
  if (!p) return { what: user.friendState === 'friend' ? '予定されている配信はありません' : 'まだ友だちではありません' };
  return { when: fmtTime(p.t.fireAt), what: `${L.flow(p.t.flowId)}：${p.what}` };
}

const BADGE: Record<StepStatus, string> = { done: '済', current: 'いまここ', skipped: '通らず', todo: 'これから' };

function variantsHtml(row: JourneyRow, compact: boolean): string {
  const vs = row.step.variants ?? [];
  if (!vs.length) return '';
  const items = vs
    .map((x) => {
      const mine = x.id === row.variantId;
      return compact
        ? `<span class="lsp-lane${mine ? ' mine' : ''}">${mine ? '▶ ' : ''}${esc(x.label)}</span>`
        : `<li class="${mine ? 'mine' : ''}"><span class="lsp-v-l">${esc(x.label)}${mine ? '<b>この人</b>' : ''}</span>${x.purpose ? `<span class="lsp-v-p">${esc(x.purpose)}</span>` : ''}</li>`;
    })
    .join('');
  return compact ? `<div class="lsp-lanes">${items}</div>` : `<ul class="lsp-variants">${items}</ul>`;
}

/** ③の本体。compact は資料用16:9（文字を大きく、説明は「いまここ」の段だけ） */
export function journeyPane(v: View, user: UserState, compact = false): string {
  const { auto } = journeySteps(v.project);
  const rows = journeyOf(v.project, v.world, user);
  const running = Object.entries(user.flows)
    .filter(([, c]) => c.status === 'running' || c.status === 'waiting')
    .map(([id]) => v.L.flow(id));
  const next = nextAction(v.project, v.world, v.demo, user, v.L);
  const goal = rows.find((r) => r.isGoal);
  const reachedGoal = goal && (goal.status === 'current' || goal.status === 'done');

  const list = rows
    .map((r, i) => {
      const showPurpose = !compact || r.status === 'current';
      return `<li class="lsp-st ${r.status}${r.isGoal ? ' goal' : ''}">
        <span class="lsp-st-dot">${r.isGoal ? '★' : i + 1}</span>
        <div class="lsp-st-body">
          <div class="lsp-st-lbl"><span>${esc(r.step.label)}</span><span class="lsp-st-badge">${BADGE[r.status]}${r.at !== undefined && r.status !== 'todo' ? ` ${esc(fmtTime(r.at))}` : ''}</span></div>
          ${showPurpose && r.step.purpose ? `<div class="lsp-st-pur">${esc(r.step.purpose)}</div>` : ''}
          ${variantsHtml(r, compact)}
          ${r.status === 'current' && running.length ? `<div class="lsp-st-now">動いているシナリオ：${esc(running.join('、'))}</div>` : ''}
        </div>
      </li>`;
    })
    .join('');

  const banner =
    user.friendState === 'blocked'
      ? '<div class="lsp-jbanner ng">ブロック中：この先の配信は届きません</div>'
      : reachedGoal
        ? `<div class="lsp-jbanner ok">ゴール「${esc(goal!.step.label)}」に到達しました</div>`
        : '';
  return `<div class="lsp-journey${compact ? ' compact' : ''}">
    ${goal ? `<div class="lsp-jgoal">ゴール：<b>${esc(goal.step.label)}</b></div>` : ''}
    ${banner}
    <div class="lsp-jnext"><span class="lsp-jn-k">次に起きること</span>${next.when ? `<span class="lsp-jn-t">${esc(next.when)}</span>` : ''}<span class="lsp-jn-w">${esc(next.what)}</span></div>
    ${rows.length ? `<ol class="lsp-steps">${list}</ol>` : '<div class="lsp-none">シナリオがありません</div>'}
    ${auto && !compact ? '<div class="lsp-hint">シナリオの一覧から自動で並べた簡易表示です。案件の「CVまでの道筋」で、段の名前・目的・ルートの違いを設定できます。</div>' : ''}
  </div>`;
}
