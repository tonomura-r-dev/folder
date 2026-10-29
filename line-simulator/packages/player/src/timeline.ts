// ③の「全ユーザーの時系列」と「次にすること」。
// 固定の台本は持たず、エンジンの実行結果（行動・配信・通ったノード・判定・タグ・CV）から作る。
// これから先は「デモの終わりまで reduce した結果」から作り、いまより後を「予定」として分ける
// （操作がない場合・台本どおりに進んだ場合の予定。手で操作すると作り直される）。

import type { FlowNode, Id, InputEvent, Minutes, Project, UserState, WorldState } from '../../engine/src/index.ts';
import { bubbleButtons, MINUTES_PER_DAY, reduce } from '../../engine/src/index.ts';
import { DEMO_END_DAY, nextStop, pendingInputs, upcoming, type Demo } from './controller.ts';
import { describeExpr, describeTrigger, explainDecision, fmtTime, type Labels } from './describe.ts';
import { journeyOf, nextAction } from './journey.ts';
import { esc, hhmm, whoOf, type View } from './view.ts';

/** 出来事の種類：①ユーザー操作 ②LINEメッセージ ③タグ・状態 ④シナリオ ⑤CV ⑥停止・対象外 */
export type TlKind = 'op' | 'msg' | 'state' | 'flow' | 'cv' | 'stop';
export type TlItem = { kind: TlKind; side: 'front' | 'back'; text: string; minor?: boolean; op?: string };
export type TlEvent = {
  key: string;
  at: Minutes;
  userId: Id;
  userName: string;
  kind: TlKind;
  title: string;
  /** 表側（LINE・Web画面で見えること）。空なら「裏側の処理のみ」 */
  front: string[];
  /** 裏側（タグ・シナリオ・判定） */
  back: string[];
  /** 展開したときに出す全部（細かい判定を含む） */
  items: TlItem[];
  /** 細かい処理だけ（既定では出さない） */
  minor: boolean;
};

export const KIND: Record<TlKind, { mark: string; label: string }> = {
  op: { mark: '👆', label: 'ユーザー操作' },
  msg: { mark: '✉', label: 'LINEメッセージ' },
  state: { mark: '🏷', label: 'タグ・状態' },
  flow: { mark: '⚙', label: 'シナリオ' },
  cv: { mark: '★', label: 'CV' },
  stop: { mark: '■', label: '停止・対象外' },
};
const RANK: TlKind[] = ['cv', 'op', 'msg', 'stop', 'state', 'flow'];

// ---------- 操作を言葉にする ----------

function buttonLabel(project: Project, id: Id): string {
  for (const m of project.messages) for (const b of m.bubbles) for (const x of bubbleButtons(b)) if (x.id === id) return x.label;
  return id;
}
function cellLabel(project: Project, menuId: Id, cellId: Id): string {
  const menus = project.richMenus.filter((m) => !menuId || m.id === menuId);
  for (const m of menus) for (const c of m.cells) if (c.id === cellId) return c.label;
  return cellId;
}

/** 操作1件（行動履歴・台本のイベント）を「何をしたか」の一文にする */
export function describeOp(project: Project, L: Labels, type: string, d: Record<string, string>): string {
  const page = project.webPages.find((p) => p.id === d['pageId']);
  switch (type) {
    case 'visitPage':
      return `ページ「${L.page(d['pageId'] ?? '')}」を開く`;
    case 'exitIntent':
      return page?.exitPopup ? 'ページを離れようとする → 離脱防止ポップアップが出る' : 'ページを離れようとする';
    case 'dismissPopup':
      return `ポップアップの「${page?.exitPopup?.dismissLabel ?? '閉じる'}」を押す`;
    case 'tapPopupCta':
      return `ポップアップの「${page?.exitPopup?.cta.label ?? 'ボタン'}」を押す`;
    case 'tapPageButton': {
      const b = page?.blocks[Number(d['blockIndex'])];
      return `ページのボタン「${b?.kind === 'button' ? b.label : 'ボタン'}」を押す`;
    }
    case 'tapLineCta':
      return `「${page?.lineCta?.label ?? 'LINEで受け取る'}」を押す`;
    case 'submitForm':
      return `「${L.page(d['pageId'] ?? '')}」を送信する`;
    case 'closePage':
      return 'ページを閉じてLINEに戻る';
    case 'addFriend':
      return `LINEで友だち追加する（${L.route(d['routeId'] ?? null)}）`;
    case 'tapMenu':
      return `リッチメニューの「${cellLabel(project, d['richMenuId'] ?? '', d['cellId'] ?? '')}」を押す`;
    case 'tapButton':
      return `ボタン「${buttonLabel(project, d['buttonId'] ?? '')}」を押す`;
    case 'sendText':
      return `「${d['text'] ?? ''}」とメッセージを送る`;
    case 'answer': {
      const sv = d['surveyId'] ?? '';
      const q = d['questionId'] ?? '';
      return `アンケート「${L.questionText(sv, q)}」に「${L.choices(sv, q, (d['choiceIds'] ?? '').split(',').filter(Boolean))}」と回答`;
    }
    case 'cv':
      return `「${L.cv(d['cvId'] ?? '')}」を記録する（店舗・管理画面側の操作）`;
    case 'viewCard':
      return `カルーセルのカード「${L.card(d['messageId'] ?? '', d['cardId'] ?? '')}」を見る`;
    case 'block':
      return 'ブロックする';
    case 'unblock':
      return 'ブロックを解除する';
    default:
      return type;
  }
}

/** 入力イベントの中身を行動履歴と同じ形（文字列の辞書）にする */
export function detailOfEvent(ev: InputEvent): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(ev)) {
    if (k === 'type' || k === 'at' || k === 'userId') continue;
    if (typeof v === 'string') out[k] = v;
    else if (typeof v === 'number') out[k] = String(v);
    else if (Array.isArray(v)) out[k] = v.join(',');
  }
  return out;
}

const OP_TITLE: Record<string, string> = {
  visitPage: 'Webページ閲覧',
  exitIntent: 'ページ離脱',
  dismissPopup: 'ポップアップを閉じる',
  tapPopupCta: 'ポップアップから友だち追加へ',
  tapPageButton: 'ページのボタン操作',
  tapLineCta: 'LINE誘導ボタン',
  submitForm: 'フォーム送信',
  closePage: 'LINEに戻る',
  addFriend: '友だち追加',
  tapMenu: 'リッチメニュー操作',
  tapButton: 'ボタン操作',
  sendText: 'メッセージ送信',
  answer: 'アンケート回答',
  cv: 'CVの記録',
  block: 'ブロック',
  unblock: 'ブロック解除',
  viewCard: 'カルーセルのカードを見る',
};

// ---------- 時系列を作る ----------

function nodeOf(project: Project, flowId: Id, nodeId: Id): FlowNode | undefined {
  return project.flows.find((f) => f.id === flowId)?.nodes.find((n) => n.id === nodeId);
}

/** 1人分の出来事。同じ人・同じ時刻のものを1件にまとめる */
function eventsOfUser(project: Project, L: Labels, world: WorldState, user: UserState): TlEvent[] {
  const groups = new Map<Minutes, TlItem[]>();
  const add = (at: Minutes, item: TlItem) => {
    const list = groups.get(at) ?? [];
    if (!list.some((x) => x.text === item.text && x.side === item.side)) list.push(item);
    groups.set(at, list);
  };

  // ① ユーザーの操作（Web・LINE）
  for (const b of user.behaviors) add(b.at, { kind: b.type === 'cv' ? 'cv' : 'op', side: 'front', text: describeOp(project, L, b.type, b.detail), op: b.type });
  // 友だち追加・ブロック（トークの帯）
  for (const t of user.talk) {
    if (t.side !== 'system') continue;
    if (t.code === 'friendAdded') add(t.at, { kind: 'state', side: 'back', text: `友だちになる（経路：${L.route(t.params?.['routeId'] ?? null)}）` });
    if (t.code === 'blocked') add(t.at, { kind: 'stop', side: 'back', text: 'ブロック：この人への配信はすべて止まる' });
    if (t.code === 'unblocked') add(t.at, { kind: 'state', side: 'back', text: 'ブロック解除：配信の対象に戻る' });
  }
  // LINEに届いたもの（配信・自動応答・アンケートの質問）は、トークに並んだ順に
  let prev = '';
  for (const t of user.talk) {
    if (t.side !== 'bot') continue;
    const key = `${t.at}/${t.messageId}`;
    if (t.messageId.startsWith('survey:')) {
      if (t.bubble.kind === 'text') add(t.at, { kind: 'msg', side: 'front', text: `アンケートの質問「${t.bubble.text.split('\n')[0]}」が届く` });
    } else if (key !== prev) add(t.at, { kind: 'msg', side: 'front', text: `「${L.message(t.messageId)}」を受信` });
    prev = key;
  }

  // ② 判定（シナリオの開始・分岐・配信・停止）。記録された順に並ぶ
  for (const d of world.decisions) {
    if (d.userId !== user.id) continue;
    const flow = L.flow(d.flowId);
    switch (d.kind) {
      case 'trigger':
        if (d.outcome === 'taken') add(d.at, { kind: 'flow', side: 'back', text: `シナリオ「${flow}」を開始` });
        else add(d.at, { kind: 'stop', side: 'back', text: `シナリオ「${flow}」には入らない（${explainDecision(d, L, project).reason ?? ''}）`, minor: true });
        break;
      case 'branch':
        if (d.outcome === 'taken') add(d.at, { kind: 'flow', side: 'back', text: `分岐「${L.node(d.nodeId)}」→「${L.branchCase(d.nodeId, d.ref)}」へ` });
        break;
      case 'deliver':
        if (d.outcome === 'taken') {
          add(d.at, { kind: 'msg', side: 'back', text: d.flowId ? `「${flow}」の配信（${L.node(d.nodeId) || '文面'}）を実行` : '自動応答で返信（通数に数えない）' });
        } else {
          const x = explainDecision(d, L, project, world);
          add(d.at, { kind: 'stop', side: 'back', text: `「${L.message(d.ref ?? '')}」は送らない：${x.reason ?? ''}` });
        }
        break;
      case 'stopWhen': {
        const x = explainDecision(d, L, project);
        add(d.at, { kind: 'stop', side: 'back', text: `シナリオ「${flow}」を停止：${x.reason ?? ''}` });
        break;
      }
      case 'autoReply':
        if (d.outcome === 'taken') add(d.at, { kind: 'flow', side: 'back', text: '自動応答のルールに当てはまった' });
        break;
    }
  }

  // ③ 通ったノード（タグ・待機・ページ・メニュー切替・終了）
  user.trail.forEach((t, i) => {
    const n = nodeOf(project, t.flowId, t.nodeId);
    if (!n) return;
    const flow = L.flow(t.flowId);
    if (n.type === 'tag') {
      for (const id of n.add ?? []) add(t.at, { kind: 'state', side: 'back', text: `タグ「${L.tag(id)}」を付ける` });
      for (const id of n.remove ?? []) add(t.at, { kind: 'state', side: 'back', text: `タグ「${L.tag(id)}」を外す` });
    } else if (n.type === 'wait') {
      // 待機が明ける時刻：いまの予定（タイマー）か、同じシナリオで次に通ったノードの時刻
      const timer = world.timers.find((x) => x.userId === user.id && x.flowId === t.flowId && x.nodeId === t.nodeId && x.fireAt >= t.at);
      const resumed = user.trail.slice(i + 1).find((x) => x.flowId === t.flowId)?.at;
      const until = timer?.fireAt ?? resumed;
      if (until !== undefined && until > t.at) add(t.at, { kind: 'flow', side: 'back', text: `「${flow}」：${fmtTime(until)} まで待機` });
      else {
        // 待っている間に止まった（予約・ブロック等）
        const stopped = world.decisions.find((d) => d.userId === user.id && d.kind === 'stopWhen' && d.flowId === t.flowId && d.at >= t.at);
        if (stopped) add(t.at, { kind: 'flow', side: 'back', text: `「${flow}」：次の配信まで待機（${fmtTime(stopped.at)} に停止）` });
      }
    } else if (n.type === 'page') add(t.at, { kind: 'op', side: 'front', text: `ページ「${L.page(n.pageId)}」が開く` });
    else if (n.type === 'richMenu') add(t.at, { kind: 'state', side: 'front', text: `リッチメニューが「${project.richMenus.find((m) => m.id === n.richMenuId)?.label ?? n.richMenuId}」に切り替わる` });
    else if (n.type === 'end') add(t.at, { kind: 'flow', side: 'back', text: `シナリオ「${flow}」が終わる`, minor: true });
  });
  // アンケート回答で付く・外れるタグ
  for (const b of user.behaviors) {
    if (b.type !== 'answer') continue;
    const q = L.question(b.detail['surveyId'] ?? '', b.detail['questionId'] ?? '');
    for (const c of q?.choices ?? []) {
      if (!(b.detail['choiceIds'] ?? '').split(',').includes(c.id)) continue;
      for (const id of c.addTags ?? []) add(b.at, { kind: 'state', side: 'back', text: `タグ「${L.tag(id)}」を付ける` });
      for (const id of c.removeTags ?? []) add(b.at, { kind: 'state', side: 'back', text: `タグ「${L.tag(id)}」を外す` });
    }
  }
  // ④ CV（フォーム送信・記録）
  // 状態には最後の1回しか残らないので、2回目以降も出るよう操作（フォーム送信・記録）から拾う
  for (const b of user.behaviors) {
    const cvId = b.type === 'cv' ? b.detail['cvId'] : b.type === 'submitForm' ? project.webPages.find((p) => p.id === b.detail['pageId'])?.form?.onSubmit.cvId : undefined;
    if (cvId) add(b.at, { kind: 'cv', side: 'back', text: `CV「${L.cv(cvId)}」を記録` });
  }

  return [...groups.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([at, items]) => toEvent(user, at, items));
}

function toEvent(user: UserState, at: Minutes, items: TlItem[]): TlEvent {
  const main = items.filter((x) => !x.minor);
  const kinds = new Set(main.map((x) => x.kind));
  const kind = RANK.find((k) => kinds.has(k)) ?? 'flow';
  const front = main.filter((x) => x.side === 'front').map((x) => x.text);
  const back = main.filter((x) => x.side === 'back').map((x) => x.text);
  return { key: `${user.id}@${at}`, at, userId: user.id, userName: whoOf(user), kind, title: titleOf(main, kind), front, back, items, minor: main.length === 0 };
}

function titleOf(items: TlItem[], kind: TlKind): string {
  const has = (s: string) => items.some((x) => x.text.includes(s));
  if (kind === 'cv') return items.find((x) => x.kind === 'cv' && x.side === 'back')?.text.replace(/^CV「(.+)」を記録$/, 'CV：$1') ?? 'CV';
  if (has('友だちになる')) return '友だち追加';
  if (has('ブロック：')) return 'ブロック';
  if (kind === 'op') return OP_TITLE[items.find((x) => x.op && x.kind === 'op')?.op ?? ''] ?? 'ユーザー操作';
  if (kind === 'msg') return items.some((x) => x.side === 'front' && x.text.startsWith('アンケートの質問')) && !items.some((x) => x.text.endsWith('を受信')) ? 'アンケート表示' : 'メッセージ受信';
  if (kind === 'stop') return has('を停止') ? '配信停止' : '対象外の判定';
  if (kind === 'state') return 'タグ・状態の変更';
  if (has('を開始')) return 'シナリオ開始';
  if (has('まで待機')) return '次の配信まで待機';
  return '裏側の処理';
}

/** 全員の出来事（時刻順・同じ時刻はユーザーの並び順） */
export function buildTimeline(project: Project, L: Labels, world: WorldState, userOrder: Id[]): TlEvent[] {
  const order = new Map(userOrder.map((id, i) => [id, i]));
  return userOrder
    .map((id) => world.users[id])
    .filter((u): u is UserState => !!u)
    .flatMap((u) => eventsOfUser(project, L, world, u))
    .sort((a, b) => a.at - b.at || (order.get(a.userId) ?? 0) - (order.get(b.userId) ?? 0));
}

// ---------- デモの終わりまでの見通し（予定） ----------

const cache = new WeakMap<object, { project: Project; end: Minutes; world: WorldState }>();

/** デモの終わり（いまより後ならそこ）まで進めた結果。同じイベント列なら作り直さない */
export function horizonOf(project: Project, demo: Demo): WorldState {
  const end = Math.max(demo.now, (project.settings.demoEndDay ?? DEMO_END_DAY) * MINUTES_PER_DAY);
  const hit = cache.get(demo.session);
  if (hit && hit.project === project && hit.end === end) return hit.world;
  const world = reduce(project, demo.session, { until: end });
  cache.set(demo.session, { project, end, world });
  return world;
}

/** 過去・いま・予定に分けた全員の時系列 */
export function timelineOf(v: View): TlEvent[] {
  return buildTimeline(v.project, v.L, horizonOf(v.project, v.demo), v.demo.session.users.map((u) => u.id));
}

/** 「次の出来事 ▶」で進む時刻と、そこで起きること（全員） */
export function nextGlobal(v: View, events: TlEvent[] = timelineOf(v)): { at: Minutes; events: TlEvent[] } | null {
  const t = nextStop(v.world, v.demo);
  if (t === null) return null;
  return { at: t, events: events.filter((e) => e.at === t) };
}

/** その時刻に起きたことの一言（「次の出来事」を押したあとの知らせに使う） */
export function summaryAt(v: View, at: Minutes): string {
  const evs = timelineOf(v).filter((e) => e.at === at);
  if (!evs.length) return '';
  const active = v.demo.activeUserId;
  const others = evs.filter((e) => e.userId !== active);
  const mine = evs.filter((e) => e.userId === active);
  const what = (e: TlEvent) => `${e.userName}：${e.title}${e.front.length ? '' : '（裏側の処理のみ）'}`;
  if (mine.length) return `${fmtTime(at)}　${mine.map(what).join('／')}${others.length ? `　ほか ${others.map((e) => e.userName).join('・')}にも出来事` : ''}`;
  return `${fmtTime(at)}　別のユーザーの出来事です（${others.map(what).join('／')}）。①②は表示中の${whoOf(v.world.users[active])}のまま。③の時系列で確認できます`;
}

// ---------- 次にすること（選んでいる人） ----------

/** この人の次の出来事（シミュレーションの結果。操作がない場合・台本どおりの場合） */
export function nextOwn(v: View, user: UserState): { at: Minutes; when: string; what: string } | null {
  const e = timelineOf(v).find((x) => x.userId === user.id && x.at > v.world.now && !x.minor);
  if (!e) return null;
  return { at: e.at, when: fmtTime(e.at), what: e.front[0] ?? `${e.back[0] ?? e.title}（裏側の処理のみ）` };
}

export type Guide = {
  /** manual=手で操作／script=台本が操作／wait=配信待ち／done=完了／stopped=停止・対象外／start=はじめかた */
  kind: 'manual' | 'script' | 'wait' | 'done' | 'stopped' | 'start';
  text: string;
  /** ①の中で光らせる要素（CSSセレクタ） */
  target?: string;
};

export function guideOf(v: View, user: UserState): Guide {
  const { project, world, demo, L } = v;
  const scripted = demo.scripted.includes(user.id);
  const nextScript = scripted ? demo.session.events.filter((e): e is InputEvent => e.userId === user.id && e.at > demo.now).sort((a, b) => a.at - b.at)[0] : undefined;
  const goal = journeyOf(project, world, user).find((r) => r.isGoal);
  const goalDone = !!goal?.step.cvId && !!user.cvs[goal.step.cvId];
  const own = nextOwn(v, user);
  const plan = own ?? nextAction(project, world, demo, user, L);

  if (user.friendState === 'blocked')
    return { kind: 'stopped', text: 'ブロック中のため、この人への配信はすべて止まっています（理由：ユーザーがブロック）。②の「ブロックを解除」で戻せます。' };
  if (nextScript) {
    const op = describeOp(project, L, nextScript.type, detailOfEvent(nextScript));
    // 台本の操作より先に配信があるなら、まず配信を待つ
    if (own && own.at < nextScript.at)
      return { kind: 'wait', text: `「次の出来事 ▶」で、この人の次の出来事（${plan.when}：${plan.what}）まで進めてください。台本の次の操作は ${fmtTime(nextScript.at)} に「${op}」です。` };
    return { kind: 'script', text: `台本の次の操作：${fmtTime(nextScript.at)} に「${op}」。「次の出来事 ▶」で進めると自動で行われます。` };
  }
  if (scripted && !nextScript && demo.now < (project.settings.demoEndDay ?? DEMO_END_DAY) * MINUTES_PER_DAY && plan.when)
    return { kind: 'wait', text: `台本の操作は終わりました。「次の出来事 ▶」で次の出来事（${plan.when}：${plan.what}）まで進めてください。` };

  const s = user.surface;
  if (!scripted) {
    if (s.kind === 'popup') {
      const p = project.webPages.find((x) => x.id === s.pageId);
      return { kind: 'manual', text: `①のポップアップの「${p?.exitPopup?.cta.label ?? 'ボタン'}」を押してください（友だち追加に進みます）。`, target: '[data-act="popupCta"]' };
    }
    if (s.kind === 'page') {
      const p = project.webPages.find((x) => x.id === s.pageId);
      if (p?.form) return { kind: 'manual', text: `①の「${p.title}」で「送信する」を押してください。`, target: '.lsp-form .lsp-pbtn' };
      if (p?.lineCta && user.friendState !== 'friend') return { kind: 'manual', text: `①の「${p.lineCta.label}」を押してください（LINEの友だち追加に進みます）。`, target: '[data-act="lineCta"]' };
      if (p?.exitPopup && user.friendState !== 'friend')
        return { kind: 'manual', text: '①の左上「‹」でページを離れようとすると、離脱防止ポップアップが出ます。ページ内のボタンも押せます。', target: '[data-act="back"]' };
      return { kind: 'manual', text: '①の左上「‹」を押すと、LINEのトークに戻ります。', target: '[data-act="back"]' };
    }
    if (user.friendState === 'not_friend') return { kind: 'start', text: '①の「はじめかた」から、この人がやって来る入口を選んでください。', target: '.lsp-start-main button' };
    if (s.kind === 'survey') {
      const q = project.surveys.find((x) => x.id === s.surveyId)?.questions.find((x) => x.id === s.questionId);
      return { kind: 'manual', text: `①のアンケート「${q?.text.split('\n')[0] ?? ''}」で、答えを選んでください（${q?.choices.map((c) => c.label).join('／') ?? ''}）。`, target: '.lsp-quick' };
    }
    if (pendingInputs(project, world, demo).some((p) => p.userId === user.id && p.kind === 'choice')) {
      let labels = '';
      for (const t of [...user.talk].reverse())
        if (t.side === 'bot' && (t.bubble.kind === 'buttons' || t.bubble.kind === 'card')) {
          labels = t.bubble.buttons.map((b) => b.label).join('／');
          break;
        }
      return { kind: 'manual', text: `①のメッセージのボタン（${labels}）から1つ押してください。`, target: '.lsp-chat .lsp-row.bot:last-of-type .lsp-btns' };
    }
  }
  if (plan.when) return { kind: 'wait', text: `「次の出来事 ▶」で、この人の次の出来事（${plan.when}：${plan.what}）まで進めてください。${goalDone ? '' : '①のメニューやボタンを押して反応を確かめることもできます。'}` };
  if (goalDone) return { kind: 'done', text: `ゴール「${goal!.step.label}」に到達しました（${fmtTime(user.cvs[goal!.step.cvId!]!.at)}）。この人の主要な配信は完了です。` };
  const stopDecision = [...world.decisions].reverse().find((d) => d.userId === user.id && d.kind === 'stopWhen');
  if (stopDecision) {
    const x = explainDecision(stopDecision, L, project);
    return { kind: 'stopped', text: `予定されている配信はありません。最後に止まったのは${x.title.replace(/を止めました$/, '')}（${x.reason ?? ''}）。` };
  }
  return { kind: 'stopped', text: '予定されている配信はありません。①のメニューやボタンを押すと、反応する配信を確かめられます。' };
}

// ---------- 描画 ----------

const GUIDE_LABEL: Record<Guide['kind'], string> = { manual: '手で操作', script: '台本が操作', wait: '配信を待つ', done: '完了', stopped: '停止中', start: 'はじめかた' };

export function guideHtml(g: Guide): string {
  return `<div class="lsp-guide g-${g.kind}"><span class="lsp-guide-k">次にすること<i>${GUIDE_LABEL[g.kind]}</i></span><span class="lsp-guide-t">${esc(g.text)}</span></div>`;
}

/** 同じ人の行は強く、ほかの人は控えめに。いまより後は「予定」 */
function rowHtml(v: View, e: TlEvent, open: boolean): string {
  const now = v.world.now;
  const when = e.at < now ? 'past' : e.at === now ? 'now' : 'future';
  const mine = e.userId === v.demo.activeUserId;
  const lines = (list: string[], max: number) => list.slice(0, open ? list.length : max).map(esc).join('<br>') + (!open && list.length > max ? `<span class="more">ほか${list.length - max}件</span>` : '');
  const detail = open
    ? `<ul class="lsp-tl-items">${e.items.map((x) => `<li class="k-${x.kind}${x.minor ? ' minor' : ''}"><span>${KIND[x.kind].mark}</span>${x.side === 'front' ? '表' : '裏'}：${esc(x.text)}</li>`).join('')}</ul>${
        mine ? '' : `<button type="button" class="lsp-tl-switch" data-act="selectUser" data-user="${esc(e.userId)}">この人を①②に表示する</button>`
      }`
    : '';
  return `<li class="lsp-tl-row ${when}${mine ? ' mine' : ''} k-${e.kind}" data-at="${e.at}">
    <button type="button" class="lsp-tl-head" data-act="tlOpen" data-key="${esc(e.key)}" aria-expanded="${open}">
      <span class="lsp-tl-time">${hhmm(e.at)}${when === 'future' ? '<small>予定</small>' : when === 'now' ? '<small>いま</small>' : ''}</span>
      <span class="lsp-tl-kind" title="${KIND[e.kind].label}">${KIND[e.kind].mark}</span>
      <span class="lsp-tl-main"><span class="lsp-tl-who">${esc(e.userName)}</span><b>${esc(e.title)}</b>
        <span class="lsp-tl-f"><i>表</i>${e.front.length ? lines(e.front, 2) : '<span class="none">裏側の処理のみ。LINE画面に変化なし</span>'}</span>
        ${e.back.length ? `<span class="lsp-tl-b"><i>裏</i>${lines(e.back, 2)}</span>` : ''}
      </span>
    </button>${detail}</li>`;
}

export function timelineHtml(v: View, compact = false): string {
  const all = timelineOf(v);
  const ui = v.ui;
  const active = v.demo.activeUserId;
  const filter = ui.tlFilter ?? 'all';
  const shown = all.filter((e) => (ui.tlMinor || !e.minor) && (filter === 'all' || e.userId === active) && (ui.tlFuture !== false || e.at <= v.world.now));
  const next = nextGlobal(v, all);
  // 選んでいる人の出来事を先に。資料用16:9は2件まで（残りは人数だけ）
  const nextEvents = next ? [...next.events].sort((a, b) => Number(b.userId === active) - Number(a.userId === active)) : [];
  const listed = compact ? nextEvents.slice(0, 2) : nextEvents;
  const nextHtml = next
    ? `<div class="lsp-tl-next"><span class="k">全体の次の出来事</span><b>${esc(fmtTime(next.at))}</b>${
        listed.length
          ? listed.map((e) => `<span>${esc(e.userName)}：${esc(e.title)}${e.userId === active ? '' : '（別ユーザー）'}</span>`).join('') +
            (nextEvents.length > listed.length ? `<span>ほか${nextEvents.length - listed.length}人</span>` : '')
          : '<span>裏側の処理のみ（待機の切り替わり）</span>'
      }</div>`
    : '<div class="lsp-tl-next"><span class="k">全体の次の出来事</span><span>この先に予定はありません</span></div>';
  if (compact) {
    const past = shown.filter((e) => e.at <= v.world.now).slice(-1);
    return `<div class="lsp-tl compact"><div class="lsp-tl-title">直前の出来事（全員）</div><ol class="lsp-tl-list">${past.map((e) => rowHtml(v, e, false)).join('') || '<li class="lsp-none">まだ出来事はありません</li>'}</ol>${nextHtml}</div>`;
  }
  let day = -1;
  let nowMarked = false;
  const rows: string[] = [];
  for (const e of shown) {
    if (!nowMarked && e.at > v.world.now) {
      rows.push(`<li class="lsp-tl-nowline" id="lsp-tl-now"><span>いま ${esc(fmtTime(v.world.now))}　ここから下は予定（操作がない場合・台本どおりの場合）</span></li>`);
      nowMarked = true;
    }
    const d = Math.floor(e.at / MINUTES_PER_DAY);
    if (d !== day) {
      rows.push(`<li class="lsp-tl-day">Day ${d}</li>`);
      day = d;
    }
    rows.push(rowHtml(v, e, !!ui.tlOpen?.has(e.key)));
  }
  if (!nowMarked) rows.push(`<li class="lsp-tl-nowline" id="lsp-tl-now"><span>いま ${esc(fmtTime(v.world.now))}　この先の予定はありません</span></li>`);
  const opt = (val: 'all' | 'me', label: string) => `<button type="button" class="${filter === val ? 'on' : ''}" data-act="tlFilter" data-filter="${val}">${label}</button>`;
  return `<div class="lsp-tl">
    <div class="lsp-tl-headbar"><span class="lsp-tl-title">全ユーザーの時系列</span><span class="lsp-group">${opt('all', '全員')}${opt('me', 'この人だけ')}</span>
      <label class="lsp-check"><input type="checkbox" data-change="tlFuture"${ui.tlFuture !== false ? ' checked' : ''}> 予定も表示</label>
      <label class="lsp-check"><input type="checkbox" data-change="tlMinor"${ui.tlMinor ? ' checked' : ''}> 細かい判定も表示</label></div>
    ${nextHtml}
    <div class="lsp-tl-legend">${(Object.keys(KIND) as TlKind[]).map((k) => `<span class="k-${k}">${KIND[k].mark} ${KIND[k].label}</span>`).join('')}<span class="mine">太字＝表示中の人</span></div>
    <ol class="lsp-tl-list lsp-tl-scroll">${rows.join('')}</ol>
  </div>`;
}

/** ③の「いま起きたこと」：いまの時刻（無ければ直前）に起きた出来事を、全員分まとめて大きく出す */
export function nowHtml(v: View): string {
  const all = timelineOf(v).filter((e) => !e.minor && e.at <= v.world.now);
  const last = all.at(-1)?.at;
  const active = v.demo.activeUserId;
  if (last === undefined) return `<div class="lsp-nowcard empty"><div class="lsp-nowcard-h"><b>いま起きたこと</b></div><div class="lsp-none">まだ出来事はありません。「次の出来事 ▶」で進めてください。</div></div>`;
  const evs = all.filter((e) => e.at === last).sort((a, b) => Number(b.userId === active) - Number(a.userId === active));
  const lines = (list: string[], cls: string, mark: string) =>
    list.length ? `<div class="lsp-nc-${cls}"><i>${mark}</i><span>${list.slice(0, 3).map(esc).join('<br>')}${list.length > 3 ? `<span class="more">ほか${list.length - 3}件</span>` : ''}</span></div>` : '';
  const items = evs
    .map((e) => {
      const mine = e.userId === active;
      return `<li class="${mine ? 'mine' : 'other'} k-${e.kind}">
        <div class="lsp-nc-who"><span class="lsp-tl-kind">${KIND[e.kind].mark}</span><b>${esc(e.userName)}</b><span class="t">${esc(e.title)}</span>${
          mine ? '<span class="tag">表示中の人</span>' : `<button type="button" class="lsp-tl-switch" data-act="selectUser" data-user="${esc(e.userId)}">この人を表示</button>`
        }</div>
        ${e.front.length ? lines(e.front, 'f', '表') : '<div class="lsp-nc-f none"><i>表</i><span>LINE画面に変化なし（裏側の処理のみ）</span></div>'}
        ${lines(e.back, 'b', '裏')}
      </li>`;
    })
    .join('');
  return `<div class="lsp-nowcard"><div class="lsp-nowcard-h"><b>${last === v.world.now ? 'いま起きたこと' : '直前の出来事'}</b><span>${esc(fmtTime(last))}・${evs.length}人</span><button type="button" class="ghost" data-act="tlJump">全員の時系列へ ↓</button></div><ul>${items}</ul></div>`;
}

/** ②の「この人の出来事」：その人のいままでの出来事を新しい順に */
export function userHistoryHtml(v: View, user: UserState): string {
  const mine = timelineOf(v).filter((e) => e.userId === user.id && e.at <= v.world.now && !e.minor).reverse();
  if (!mine.length) return '<div class="lsp-none">まだ出来事はありません</div>';
  return `<ol class="lsp-tl-list hist">${mine.map((e) => rowHtml(v, e, !!v.ui.tlOpen?.has(e.key))).join('')}</ol>`;
}

/** ②に出す、実行中のシナリオの「配信条件」「停止条件」 */
export function flowConditions(project: Project, L: Labels, flowId: Id): { trigger: string; stop?: string } {
  const f = project.flows.find((x) => x.id === flowId);
  if (!f) return { trigger: '' };
  return { trigger: describeTrigger(f.trigger, project), ...(f.stopWhen ? { stop: describeExpr(f.stopWhen, L) } : {}) };
}

export { upcoming };
