// デモ画面の③「CVまでの道筋」。営業先に見せる説明用のフローチャートで、編集用の詳しいフロー図ではない。
// 案件の journey（無ければシナリオ一覧から作る簡易版）を、選んでいるユーザーの状態に当てはめて
// START → 施策（四角）→ 分岐（ひし形）→ … → GOAL を描き、「現在ここ」と通った経路を示す。
// 施策の文面は journey に持たず、シナリオのノード・文面を参照する（①と同じ文面が出る）。

import type { Expr, Feasibility, Flow, FlowNode, Id, JourneyOption, JourneyStep, JourneyVariant, Minutes, Project, UserState, WorldState } from '../../engine/src/index.ts';
import { bubbleButtons, evaluate } from '../../engine/src/index.ts';
import { bubbleTexts, renderBubble } from '../../engine/src/text.ts';
import { pendingInputs, upcoming, type Demo } from './controller.ts';
import { describeExpr, describeFailure, describeTrigger, explainDecision, fmtDuration, fmtTime, type Labels } from './describe.ts';
import { previewBubble } from './phone.ts';
import { guideHtml, guideOf, nowHtml, timelineHtml } from './timeline.ts';
import { esc, whoOf, type View } from './view.ts';

export type StepStatus = 'done' | 'current' | 'skipped' | 'todo';
export type JourneyRow = {
  step: JourneyStep;
  status: StepStatus;
  /** 来た時刻（条件式だけで判定した段は分からないので無し） */
  at?: Minutes;
  /** この人のルート（variants のうち、いま条件に当てはまるもの） */
  variantId?: Id;
  /** 実際に通った行き先（variants の nodeIds を通ったもの） */
  traversed: Id[];
  isStart: boolean;
  isGoal: boolean;
};

export const FEASIBILITY: Record<Feasibility, { mark: string; text: string }> = {
  lstep: { mark: '◎', text: 'Lステップ単体で可' },
  conditional: { mark: '○', text: 'プラン・設定に条件あり' },
  external: { mark: '△', text: '外部ツール・Webサイト側の実装が必要' },
  unknown: { mark: '？', text: '実現可否は未確認' },
};

/** 案件の道筋。無ければシナリオの並び順とCVから簡易版を作る */
export function journeySteps(project: Project): { steps: JourneyStep[]; auto: boolean } {
  if (project.journey?.length) return { steps: project.journey, auto: false };
  const steps: JourneyStep[] = project.flows.map((f) => ({ id: `auto:${f.id}`, label: f.label, nodeIds: f.nodes.map((n) => n.id) }));
  for (const c of project.cvs) steps.push({ id: `auto:cv:${c.id}`, label: c.label, cvId: c.id });
  return { steps, auto: true };
}

const holds = (expr: Expr, project: Project, user: UserState, now: Minutes) => evaluate(expr, { project, user, now }).value;
const firstPass = (user: UserState, ids: Id[] | undefined) => {
  const set = new Set(ids ?? []);
  return set.size ? user.trail.find((t) => set.has(t.nodeId))?.at : undefined;
};

function reachedAt(step: JourneyStep, project: Project, user: UserState, now: Minutes): { reached: boolean; at?: Minutes } {
  const times = [firstPass(user, step.nodeIds), ...(step.variants ?? []).map((x) => firstPass(user, x.nodeIds))];
  if (step.cvId !== undefined) times.push(user.cvs[step.cvId]?.at);
  const known = times.filter((t): t is Minutes => t !== undefined);
  if (known.length) return { reached: true, at: Math.min(...known) };
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
    // 通った順（最初に通った時刻の順）
    const traversed = (step.variants ?? [])
      .map((x) => ({ id: x.id, at: firstPass(user, x.nodeIds) }))
      .filter((x): x is { id: Id; at: Minutes } => x.at !== undefined)
      .sort((a, b) => a.at - b.at)
      .map((x) => x.id);
    // ノードを持たない行き先（例：友だち追加の経路）は、段に来ていて条件に当てはまれば「通った」とみなす
    const matched = step.variants?.find((x) => x.id === variantId);
    if (r.reached && matched && !matched.nodeIds?.length && !traversed.includes(matched.id)) traversed.push(matched.id);
    return { step, status, traversed, isStart: i === 0, isGoal: i === goalIndex, ...(r.at !== undefined ? { at: r.at } : {}), ...(variantId ? { variantId } : {}) };
  });
}

// ---------- 次に起きること ----------

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

// ---------- 施策の中身（シナリオの文面を参照する） ----------

export type Content = {
  /** 文面（シナリオのノード、または自動応答） */
  messageId?: Id;
  surveyId?: Id;
  nodeId?: Id;
  /** いつ届くか */
  timing: string;
  /** どのシナリオ・自動応答から、誰に */
  target: string;
  /** 止める条件 */
  stop?: string;
};

function waitText(w: Extract<FlowNode, { type: 'wait' }>['wait']): string {
  if (w.kind === 'dayAt') return `${w.from === 'flowStart' ? 'シナリオ開始' : '友だち追加'}の${w.day === 0 ? '当日' : `${w.day}日後`} ${w.time}`;
  if (w.kind === 'duration') return `${fmtDuration(w.minutes)}後`;
  if (w.kind === 'at') return `Day ${w.day} ${w.time}（全員共通の日時）`;
  return `操作を待って（最長${fmtDuration(w.timeoutMinutes)}）`;
}

/** ノードの直前をたどって「いつ届くか」を文にする */
function timingOf(project: Project, L: Labels, flow: Flow, nodeId: Id, depth = 0): string {
  const texts = flow.edges
    .filter((e) => e.to === nodeId)
    .map((e) => {
      const from = flow.nodes.find((n) => n.id === e.from.nodeId);
      const port = e.from.port;
      if (!from) return '';
      switch (from.type) {
        case 'wait':
          return waitText(from.wait);
        case 'entry':
          return `シナリオ開始時（${describeTrigger(flow.trigger, project)}）`;
        case 'branch':
          return `「${from.label ?? '分岐'}」が「${L.branchCase(from.id, port === 'else' ? 'else' : port.slice(5))}」のとき`;
        case 'choice':
          return port === 'timeout' ? 'ボタンが押されないまま時間が過ぎたとき' : `ボタン「${buttonLabel(project, port.slice(4))}」を押したとき`;
        case 'survey':
          return port === 'timeout' ? 'アンケートの回答待ちが切れたとき' : 'アンケートに答えたあと';
        case 'message':
          return '直前の配信に続けて';
        default:
          return depth < 5 ? timingOf(project, L, flow, from.id, depth + 1) : '';
      }
    })
    .filter(Boolean);
  return [...new Set(texts)].join('／') || 'シナリオの途中';
}

function buttonLabel(project: Project, id: Id): string {
  for (const m of project.messages) for (const b of m.bubbles) for (const x of bubbleButtons(b)) if (x.id === id) return x.label;
  return id;
}

const flowTarget = (project: Project, flow: Flow) => `シナリオ「${flow.label}」（始まり：${describeTrigger(flow.trigger, project)}）`;

/** 段・行き先に紐づく文面とアンケート（重複なし・登録した順） */
export function contentsOf(project: Project, L: Labels, nodeIds: Id[] = [], messageIds: Id[] = []): Content[] {
  const out: Content[] = [];
  for (const id of nodeIds) {
    const flow = project.flows.find((f) => f.nodes.some((n) => n.id === id));
    const node = flow?.nodes.find((n) => n.id === id);
    if (!flow || !node) continue;
    const base = { nodeId: id, timing: timingOf(project, L, flow, id), target: flowTarget(project, flow), ...(flow.stopWhen ? { stop: describeExpr(flow.stopWhen, L) } : {}) };
    if (node.type === 'message' || node.type === 'choice') out.push({ ...base, messageId: node.messageId });
    else if (node.type === 'survey') out.push({ ...base, surveyId: node.surveyId });
  }
  for (const id of messageIds) {
    const ar = project.autoReplies.find((a) => a.fallbackMessageId === id || a.cases.some((c) => c.messageId === id));
    if (!ar) {
      out.push({ messageId: id, timing: '—', target: '（どこからも配信されていない文面）' });
      continue;
    }
    const kw = ar.match.keywords.join('・');
    const hits = (text: string) => ar.match.keywords.some((k) => (ar.match.mode === 'exact' ? text.trim() === k : text.includes(k)));
    const cell = project.richMenus.flatMap((m) => m.cells).find((c) => c.action.kind === 'sendText' && hits(c.action.text));
    const when = ar.cases.find((c) => c.messageId === id)?.when;
    out.push({
      messageId: id,
      timing: `「${kw}」と送られたとき${cell ? `（リッチメニュー「${cell.label}」からも）` : ''}`,
      target: `自動応答${when ? `：${describeExpr(when, L)}の人に` : '：どの条件にも当てはまらない人に'}（通数に数えない）`,
    });
  }
  return out;
}

const PLACEHOLDER = /\{\{[^}]+\}\}/g;

/** 箱に出す文面のさわり（最初の文面の最初の数行） */
function previewOf(project: Project, contents: Content[]): string {
  const c = contents[0];
  if (!c) return '';
  if (c.surveyId) {
    const s = project.surveys.find((x) => x.id === c.surveyId);
    return s ? `質問：${s.questions.map((q) => q.text).join('／')}` : '';
  }
  const m = project.messages.find((x) => x.id === c.messageId);
  return (m?.bubbles.flatMap(bubbleTexts).join('\n') ?? '').replace(PLACEHOLDER, '〇〇').trim().slice(0, 90);
}

// ---------- フローチャート ----------

type Box = { step: JourneyStep; variant?: JourneyVariant };

const nodeIdsOf = (b: Box) => (b.variant ? b.variant.nodeIds : b.step.nodeIds) ?? [];
const messageIdsOf = (b: Box) => (b.variant ? b.variant.messageIds : b.step.messageIds) ?? [];
const feasibilityOf = (b: Box) => (b.variant ? b.variant.feasibility : b.step.feasibility);
const optionsOf = (b: Box) => (b.variant ? b.variant.options : b.step.options) ?? [];
const attr = (b: Box, o?: JourneyOption) => `data-act="fcOpen" data-step="${esc(b.step.id)}"${b.variant ? ` data-variant="${esc(b.variant.id)}"` : ''}${o ? ` data-option="${esc(o.id)}"` : ''}`;

/** 追加候補。分岐の行き先（幅が狭い）と資料用は件数だけ出し、押すと行き先の詳細で一覧を見せる */
function optionsHtml(v: View, b: Box, compact: boolean): string {
  const opts = optionsOf(b);
  if (!opts.length || !v.ui.fcOptions) return '';
  if (compact || b.variant) return `<button type="button" class="lsp-fc-optcount" ${attr(b)}>＋追加候補 ${opts.length}</button>`;
  return `<div class="lsp-fc-opts"><span class="lsp-fc-opts-k">追加候補</span>${opts
    .map((o) => `<button type="button" class="lsp-fc-opt" ${attr(b, o)}>${esc(o.label)}${o.feasibility ? `<i>${FEASIBILITY[o.feasibility].mark}</i>` : ''}</button>`)
    .join('')}</div>`;
}

function nodeHtml(v: View, b: Box, row: JourneyRow, state: string, here: boolean, compact: boolean): string {
  const contents = contentsOf(v.project, v.L, nodeIdsOf(b), messageIdsOf(b));
  const title = b.variant ? (b.variant.title ?? b.variant.label) : b.step.label;
  const preview = compact ? '' : previewOf(v.project, contents);
  const f = feasibilityOf(b);
  const kind = !b.variant && row.isStart ? ' start' : !b.variant && row.isGoal ? ' goal' : '';
  return `<button type="button" class="lsp-fc-node${kind} ${state}" ${attr(b)}>
      ${kind === ' start' ? '<span class="lsp-fc-tag">START</span>' : kind === ' goal' ? '<span class="lsp-fc-tag">GOAL</span>' : ''}
      <span class="lsp-fc-title">${esc(title)}</span>
      ${preview ? `<span class="lsp-fc-prev">${esc(preview)}</span>` : ''}
      <span class="lsp-fc-meta">${contents.length ? `<span>文面 ${contents.length}</span>` : ''}${f ? `<span title="${esc(FEASIBILITY[f].text)}">${FEASIBILITY[f].mark}${compact || b.variant ? '' : ` ${esc(FEASIBILITY[f].text)}`}</span>` : ''}${
        state === 'passed' ? '<span class="ok">✓ 通過</span>' : state === 'lane' ? '<span class="lane">この人のルート</span>' : state === 'skipped' ? '<span>通らず</span>' : ''
      }</span>
      ${here ? '<span class="lsp-fc-here">現在ここ</span>' : ''}
    </button>`;
}

const STEP_STATE: Record<StepStatus, string> = { done: 'passed', current: 'passed', skipped: 'skipped', todo: 'todo' };

/** 資料用16:9で、段が多いときに出す段：START・GOAL・最後の段と、「現在ここ」の前後1段。ほかは「⋮ ほか○段」にまとめる */
export const COMPACT_MAX_ROWS = 7;
export function compactVisible(rows: JourneyRow[]): boolean[] {
  if (rows.length <= COMPACT_MAX_ROWS) return rows.map(() => true);
  const cur = Math.max(0, rows.findIndex((r) => r.status === 'current'));
  return rows.map((r, i) => i === 0 || r.isGoal || i === rows.length - 1 || Math.abs(i - cur) <= 1);
}

function chartHtml(v: View, rows: JourneyRow[], compact: boolean): string {
  const show = compact ? compactVisible(rows) : rows.map(() => true);
  let hidden = 0;
  const gap = () => {
    const n = hidden;
    hidden = 0;
    return n ? `<div class="lsp-fc-gap">⋮ ほか${n}段</div>` : '';
  };
  const html = rows
    .map((row, i) => {
      if (!show[i]) {
        hidden++;
        return '';
      }
      const skippedBefore = hidden > 0;
      const arrow = (skippedBefore ? gap() : '') + (i === 0 ? '' : `<div class="lsp-fc-arrow${row.status === 'done' || row.status === 'current' ? ' on' : ''}"></div>`);
      const vs = row.step.variants ?? [];
      // くり返しの流れ（例：再応募 → 勤務確定へ戻る）。この人が何回くり返したかも添える
      const back = rows.find((r) => r.step.id === row.step.loopTo);
      const times = row.step.cvId ? (v.world.users[v.demo.activeUserId]?.cvs[row.step.cvId]?.count ?? 0) : 0;
      const loop = back
        ? `<div class="lsp-fc-loop${row.status === 'done' || row.status === 'current' ? ' on' : ''}">↺ 「${esc(back.step.label)}」へ戻る${row.step.loopLabel ? `：${esc(row.step.loopLabel)}` : ''}${times ? `<b>この人 ${times}回</b>` : ''}</div>`
        : '';
      if (!vs.length) {
        const here = row.status === 'current';
        return `${arrow}<div class="lsp-fc-row">${nodeHtml(v, { step: row.step }, row, STEP_STATE[row.status], here, compact)}${optionsHtml(v, { step: row.step }, compact)}</div>${loop}`;
      }
      // 分岐：いまの行き先は「当てはまっている行き先」、無ければ最後に通った行き先
      const currentVariant = row.status === 'current' ? (row.variantId ?? row.traversed.at(-1)) : undefined;
      const reached = row.status === 'done' || row.status === 'current';
      const dia = `<button type="button" class="lsp-fc-dia ${reached ? 'passed' : row.status}" ${attr({ step: row.step })}><span>${esc(row.step.decision ?? row.step.label)}</span>${
        row.status === 'current' && !currentVariant ? '<span class="lsp-fc-here">現在ここ</span>' : ''
      }</button>`;
      const branches = vs
        .map((x) => {
          const passed = row.traversed.includes(x.id);
          const state = passed ? 'passed' : x.id === row.variantId ? 'lane' : reached ? 'other' : 'todo';
          const on = passed || (x.id === row.variantId && reached);
          return `<div class="lsp-fc-branch${on ? ' on' : ''}"><span class="lsp-fc-elabel">${esc(x.label)}</span>${nodeHtml(v, { step: row.step, variant: x }, row, state, x.id === currentVariant, compact)}${optionsHtml(
            v,
            { step: row.step, variant: x },
            compact,
          )}</div>`;
        })
        .join('');
      return `${arrow}<div class="lsp-fc-row">${dia}<div class="lsp-fc-line${reached ? ' on' : ''}"></div></div><div class="lsp-fc-branches${i < rows.length - 1 ? ' merge' : ''}" style="--n:${vs.length}">${branches}</div>`;
    })
    .join('');
  return html + gap();
}

/** 資料用16:9の③：フローチャートを縮めず、現在地の前後とGOALだけを1行ずつ並べる */
function stripHtml(v: View, rows: JourneyRow[]): string {
  const cur = rows.findIndex((r) => r.status === 'current');
  const goal = rows.findIndex((r) => r.isGoal);
  const pick = new Set<number>([Math.max(0, cur - 1), cur, cur + 1, goal].filter((i) => i >= 0 && i < rows.length));
  const out: string[] = [];
  let last = -1;
  for (const i of [...pick].sort((a, b) => a - b)) {
    if (last >= 0 && i > last + 1) out.push(`<li class="gap">⋮ ${i - last - 1}段</li>`);
    const r = rows[i]!;
    const variant = r.step.variants?.find((x) => x.id === (r.variantId ?? r.traversed.at(-1)));
    const name = variant ? `${r.step.label}：${variant.title ?? variant.label}` : r.step.label;
    const mark = r.status === 'current' ? '●' : r.status === 'done' ? '✓' : r.status === 'skipped' ? '－' : '○';
    out.push(`<li class="${r.status}${r.isGoal ? ' goal' : ''}${r.isStart ? ' start' : ''}"><span class="m">${mark}</span><span class="n">${esc(name)}</span>${r.status === 'current' ? '<span class="lsp-fc-here">現在ここ</span>' : r.isGoal ? '<span class="g">GOAL</span>' : ''}</li>`);
    last = i;
  }
  void v;
  return `<ol class="lsp-strip">${out.join('')}</ol>`;
}

/** ③の上：選んでいる人の現在地（いまの段階・CVの状況・次にすること・この人の次の予定） */
function whereHtml(v: View, user: UserState, rows: JourneyRow[], compact: boolean): string {
  const goal = rows.find((r) => r.isGoal);
  const goalAt = goal?.step.cvId ? user.cvs[goal.step.cvId]?.at : undefined;
  const ci = rows.findIndex((r) => r.status === 'current');
  const stage = ci >= 0 ? `${rows[ci]!.step.label}` : user.friendState === 'not_friend' ? 'まだ友だちになっていません' : '—';
  const cv =
    goalAt !== undefined
      ? `<span class="lsp-wh-cv ok">✓ ゴール「${esc(goal!.step.label)}」達成（${esc(fmtTime(goalAt))}）</span>`
      : goal
        ? `<span class="lsp-wh-cv">ゴール「${esc(goal.step.label)}」まで${ci >= 0 ? `あと${Math.max(0, rows.indexOf(goal) - ci)}段` : ''}</span>`
        : '';
  // 次の予定は「次にすること」の文に入るので、ここでは出さない（上に残るカードを低く保つ）
  return `<div class="lsp-where">
    <div class="lsp-wh-top"><span class="lsp-wh-who">${esc(whoOf(user))}</span>${v.demo.scripted.includes(user.id) ? '<span class="lsp-wh-badge">台本</span>' : '<span class="lsp-wh-badge manual">手で操作</span>'}${cv}</div>
    <div class="lsp-wh-stage"><span class="k">現在地</span><b>${esc(stage)}</b>${ci >= 0 && rows.length ? `<small>${ci + 1} / ${rows.length} 段目</small>` : ''}</div>
    ${guideHtml(guideOf(v, user))}
  </div>`;
}

/** ③の本体：上＝現在地と次にすること／中＝CVまでのフロー／下＝全ユーザーの時系列。compact は資料用16:9 */
export function journeyPane(v: View, user: UserState, compact = false): string {
  const { auto } = journeySteps(v.project);
  const rows = journeyOf(v.project, v.world, user);
  const z = v.ui.fcZoom;
  const tools = compact
    ? ''
    : `<div class="lsp-fc-tools"><span class="lsp-group"><button type="button" data-act="fcZoom" data-z="out" aria-label="縮小">−</button><span class="lsp-fc-z">${Math.round(z * 100)}%</span><button type="button" data-act="fcZoom" data-z="in" aria-label="拡大">＋</button><button type="button" data-act="fcZoom" data-z="fit">幅に合わせる</button></span><label class="lsp-check"><input type="checkbox" data-change="fcOptions"${v.ui.fcOptions ? ' checked' : ''}> 追加候補を表示</label></div>`;
  const legend = `<div class="lsp-fc-legend"><span><b class="solid on"></b>通過済み</span><span><b class="solid"></b>未通過</span><span><b class="solid off"></b>通らず・別ルート</span><span><b class="dash"></b>追加候補（未採用）</span>${(Object.keys(FEASIBILITY) as Feasibility[])
    .map((k) => `<span>${FEASIBILITY[k].mark} ${FEASIBILITY[k].text}</span>`)
    .join('')}</div>`;
  const chart = compact && rows.length ? stripHtml(v, rows) : rows.length ? `<div class="lsp-fc-scroll"><div class="lsp-fc" style="zoom:${compact ? 1 : z}">${chartHtml(v, rows, compact)}</div></div>` : '<div class="lsp-none">シナリオがありません</div>';
  return `<div class="lsp-journey${compact ? ' compact' : ''}">
    ${whereHtml(v, user, rows, compact)}
    ${compact ? '' : nowHtml(v)}
    <div class="lsp-jsec flow">
      ${compact ? '' : '<div class="lsp-jsec-h">CVまでのフロー<small>箱を押すと文面・タイミング・対象を表示</small></div>'}
      ${tools}
      ${chart}
      ${user.friendState === 'not_friend' && !compact ? '<div class="lsp-hint">友だち追加すると、いまいる場所に「現在ここ」が出ます。</div>' : ''}
      ${compact ? '' : legend}
      ${auto && !compact ? '<div class="lsp-hint">シナリオの一覧から自動で並べた簡易表示です。案件の「CVまでの道筋」で、START・分岐・施策名・GOAL・説明を設定できます。</div>' : ''}
    </div>
    ${timelineHtml(v, compact)}
  </div>`;
}

// ---------- 箱を押したときの詳細（全文・タイミング・対象・目的） ----------

function deliveredText(v: View, user: UserState, c: Content): { tone: string; text: string } {
  if (c.nodeId) {
    const got = user.deliveries.filter((d) => d.nodeId === c.nodeId);
    if (got.length) return { tone: 'ok', text: `この人には ${got.map((d) => fmtTime(d.at)).join('、')} に届きました` };
    const miss = [...v.world.decisions].reverse().find((d) => d.userId === user.id && d.kind === 'deliver' && d.nodeId === c.nodeId);
    if (miss) {
      const x = explainDecision(miss, v.L, v.project, v.world);
      return { tone: 'ng', text: `この人には届いていません：${x.reason ?? x.title}` };
    }
    if (c.surveyId) return user.trail.some((t) => t.nodeId === c.nodeId) ? { tone: 'ok', text: 'この人には表示されました' } : { tone: 'none', text: 'この人にはまだ表示されていません' };
    return { tone: 'none', text: 'この人にはまだ届いていません' };
  }
  const got = user.deliveries.filter((d) => d.flowId === null && d.messageId === c.messageId);
  return got.length ? { tone: 'ok', text: `この人には ${got.map((d) => fmtTime(d.at)).join('、')} に届きました（自動応答）` } : { tone: 'none', text: 'この人にはまだ届いていません' };
}

function contentHtml(v: View, user: UserState, c: Content): string {
  let body = '';
  if (c.surveyId) {
    const s = v.project.surveys.find((x) => x.id === c.surveyId);
    body = s
      ? s.questions
          .map(
            (q) =>
              `<div class="lsp-fc-q"><b>${esc(q.text)}</b>${q.choices
                .map((ch) => `<span>${esc(ch.label)}${ch.addTags?.length ? `<small>→ タグ ${esc(ch.addTags.map(v.L.tag).join('・'))}</small>` : ''}</span>`)
                .join('')}</div>`,
          )
          .join('')
      : '';
  } else {
    const m = v.project.messages.find((x) => x.id === c.messageId);
    body = m ? `<div class="lsp-fc-bubbles">${m.bubbles.map((b) => previewBubble(v, renderBubble(b, v.project, user))).join('')}</div>` : '<div class="lsp-none">文面がありません</div>';
  }
  const d = deliveredText(v, user, c);
  return `<div class="lsp-fc-content">
    <dl class="lsp-kv"><dt>タイミング</dt><dd>${esc(c.timing)}</dd><dt>配信対象</dt><dd>${esc(c.target)}</dd>${c.stop ? `<dt>止める条件</dt><dd>${esc(c.stop)}</dd>` : ''}</dl>
    ${body}
    <div class="lsp-fc-st ${d.tone}">${esc(d.text)}</div>
  </div>`;
}

function feasibilityHtml(f: Feasibility | undefined, note: string | undefined): string {
  if (!f && !note) return '';
  return `<div class="lsp-fc-feas">${f ? `<b>${FEASIBILITY[f].mark} ${esc(FEASIBILITY[f].text)}</b>` : ''}${note ? `<span>${esc(note)}</span>` : ''}</div>`;
}

/** ③の箱を押したときに出す詳細。画面の外にはみ出さないよう、デモ全体の上に重ねる */
export function fcDialog(v: View, user: UserState): string {
  const open = v.ui.fcOpen;
  if (!open) return '';
  const step = journeySteps(v.project).steps.find((s) => s.id === open.step);
  const variant = step?.variants?.find((x) => x.id === open.variant);
  if (!step || (open.variant && !variant)) return '';
  const box: Box = variant ? { step, variant } : { step };
  const option = open.option ? optionsOf(box).find((o) => o.id === open.option) : undefined;
  const rows = journeyOf(v.project, v.world, user);
  const row = rows.find((r) => r.step.id === step.id)!;

  let kind: string;
  let title: string;
  let body: string;
  if (option) {
    kind = '<span class="lsp-fc-kind opt">追加候補（未採用）</span>';
    title = option.label;
    body = `<div class="lsp-fc-note">提案の候補として表示しているだけで、シナリオには入っておらず、配信もされません。</div>
      ${option.summary ? `<p>${esc(option.summary)}</p>` : ''}
      ${feasibilityHtml(option.feasibility, option.note)}
      ${option.sample ? `<h4>文面の例</h4><div class="lsp-fc-bubbles"><div class="lsp-b">${esc(option.sample)}</div></div>` : ''}
      <div class="lsp-fc-where">追加する場所：${esc(variant ? `${step.decision ?? step.label} → ${variant.title ?? variant.label}` : step.label)}</div>`;
  } else if (!variant && step.variants?.length) {
    // ひし形：行き先ごとの条件と、この人がどれに当てはまるか（②の判定と同じ評価）
    kind = '<span class="lsp-fc-kind">分岐</span>';
    title = step.decision ?? step.label;
    const items = step.variants
      .map((x) => {
        const t = evaluate(x.when, { project: v.project, user, now: v.world.now });
        return `<li class="${t.value ? 'yes' : 'no'}"><b>${esc(x.label)} → ${esc(x.title ?? x.label)}</b><span>条件：${esc(describeExpr(x.when, v.L))}</span><span class="who">${
          t.value ? '✓ この人は当てはまります' : `✗ ${esc(describeFailure(t, v.L))}`
        }</span></li>`;
      })
      .join('');
    body = `${step.purpose ? `<p>${esc(step.purpose)}</p>` : ''}<ul class="lsp-fc-cases">${items}</ul><div class="lsp-hint">上から順に、最初に当てはまった行き先に進みます。</div>`;
  } else {
    const isStart = !variant && row.isStart;
    const isGoal = !variant && row.isGoal;
    kind = `<span class="lsp-fc-kind${isStart ? ' start' : isGoal ? ' goal' : ''}">${isStart ? 'START' : isGoal ? 'GOAL' : '採用施策'}</span>`;
    title = variant ? (variant.title ?? variant.label) : step.label;
    const purpose = variant?.purpose ?? step.purpose;
    const contents = contentsOf(v.project, v.L, nodeIdsOf(box), messageIdsOf(box));
    const opts = optionsOf(box);
    body = `${variant ? `<div class="lsp-fc-where">分岐「${esc(step.decision ?? step.label)}」で「${esc(variant.label)}」の人（条件：${esc(describeExpr(variant.when, v.L))}）</div>` : ''}
      ${purpose ? `<h4>目的</h4><p>${esc(purpose)}</p>` : ''}
      ${isGoal && step.cvId ? `<div class="lsp-fc-where">CV地点：${esc(v.L.cv(step.cvId))}${user.cvs[step.cvId] ? `（この人は ${esc(fmtTime(user.cvs[step.cvId]!.at))} に到達）` : ''}</div>` : ''}
      ${feasibilityHtml(variant ? variant.feasibility : step.feasibility, variant ? variant.feasibilityNote : step.feasibilityNote)}
      <h4>配信する文面（${contents.length}件・シナリオに登録された文面）</h4>
      ${contents.length ? contents.map((c) => contentHtml(v, user, c)).join('') : '<div class="lsp-none">この施策に紐づく文面はありません</div>'}
      ${opts.length ? `<h4>追加候補（未採用）</h4><div class="lsp-fc-opts">${opts.map((o) => `<button type="button" class="lsp-fc-opt" ${attr(box, o)}>${esc(o.label)}</button>`).join('')}</div>` : ''}`;
  }
  return `<div class="lsp-modal" data-act="fcClose"><div class="lsp-modal-box" role="dialog" aria-modal="true" aria-label="${esc(title)}" data-act="fcStay">
    <div class="lsp-modal-head">${kind}<h3>${esc(title)}</h3><button type="button" class="lsp-modal-x" data-act="fcClose" aria-label="閉じる">×</button></div>
    <div class="lsp-modal-body">${body}</div>
  </div></div>`;
}
