// デモの進行（画面を持たない部分）。画面のボタンはすべてここを通してエンジンへの入力イベントにする。
// 状態は「イベント列」と「いまの時刻」だけ。画面に出すものは毎回 reduce で作り直す（DATA_MODEL §1）。
//
// 複数ユーザー
//   - 全員が同じ仮想時刻を共有する。ユーザーごとの状態はエンジンが userId で分けて持つ
//   - 台本つきのユーザーは、台本の操作を最初からイベント列に入れておく（時間を進めると台本どおりに動く）
//   - 手で操作すると、その人は台本から外れる（その人のこの先の操作は消す）
// 時間
//   - 進める・戻すは「いまの時刻」を動かすだけ。同じイベント列なら同じ結果になる
//   - 過去の時点から手で操作すると、その人のそれより後の操作は消える（分かれた未来は残さない）

import type { Id, InputEvent, InputEventBody, Minutes, Project, Session, Timer, WorldState } from '../../engine/src/index.ts';
import { MINUTES_PER_DAY, reduce } from '../../engine/src/index.ts';

/** デモを始める時刻（Day 0 の 10:00） */
export const DEMO_START: Minutes = 600;
/** 自動再生を止める日 */
export const DEMO_END_DAY = 30;

export type Demo = {
  session: Session;
  now: Minutes;
  activeUserId: Id;
  /** 台本どおりに動くユーザー（自動再生で回答待ちでも止まらない） */
  scripted: Id[];
};

const nextUserId = (demo: Demo) => {
  const used = new Set(demo.session.users.map((u) => u.id));
  for (let i = 1; ; i++) if (!used.has(`u${i}`)) return `u${i}`;
};

/** 同じペルソナを2人目以降に追加したときの表示名 */
function displayNameFor(demo: Demo, project: Project, personaId: Id): string | undefined {
  const persona = project.personas.find((p) => p.id === personaId);
  if (!persona) return 'あなた';
  const same = demo.session.users.filter((u) => u.personaId === personaId).length;
  return same === 0 ? undefined : `${persona.name}（${same + 1}）`;
}

function scriptEvents(project: Project, personaId: Id, userId: Id): InputEvent[] {
  const persona = project.personas.find((p) => p.id === personaId);
  return (persona?.script ?? []).map((s) => ({ ...s.event, at: s.atMinutes, userId }) as InputEvent);
}

export function emptyDemo(project: Project, start: Minutes = DEMO_START): Demo {
  return { session: { schemaVersion: 1, projectId: project.meta.id, users: [], events: [] }, now: start, activeUserId: '', scripted: [] };
}

/** 手で操作するユーザー1人で始める（Phase 3 と同じ始まり方） */
export function newDemo(project: Project, personaId?: Id, start: Minutes = DEMO_START): Demo {
  const persona = project.personas.find((p) => p.id === personaId) ?? project.personas[0];
  return addUser(emptyDemo(project, start), project, persona?.id ?? '', false);
}

/** 案件設定（settings.demoStart）どおりの始まり方。既定は1人を手で操作 */
export function initialDemo(project: Project): Demo {
  const s = project.settings.demoStart;
  if (s?.mode === 'scripted' && project.personas.some((p) => (p.script ?? []).length > 0)) return allScripted(project);
  return newDemo(project, s?.personaId);
}

export function addUser(demo: Demo, project: Project, personaId: Id, scripted: boolean): Demo {
  const id = nextUserId(demo);
  const name = displayNameFor(demo, project, personaId);
  return {
    ...demo,
    session: {
      ...demo.session,
      users: [...demo.session.users, { id, personaId, ...(name ? { displayName: name } : {}) }],
      events: scripted ? [...demo.session.events, ...scriptEvents(project, personaId, id)] : demo.session.events,
    },
    activeUserId: id,
    scripted: scripted ? [...demo.scripted, id] : demo.scripted,
  };
}

/** 台本のあるペルソナ全員を、台本つきで並べる（いまのユーザーは外す） */
export function allScripted(project: Project, start: Minutes = DEMO_START): Demo {
  let demo = emptyDemo(project, start);
  for (const p of project.personas) if ((p.script ?? []).length) demo = addUser(demo, project, p.id, true);
  return { ...demo, activeUserId: demo.session.users[0]?.id ?? '' };
}

export function removeUser(demo: Demo, userId: Id): Demo {
  const users = demo.session.users.filter((u) => u.id !== userId);
  return {
    ...demo,
    session: { ...demo.session, users, events: demo.session.events.filter((e) => e.userId !== userId) },
    activeUserId: demo.activeUserId === userId ? (users[0]?.id ?? '') : demo.activeUserId,
    scripted: demo.scripted.filter((id) => id !== userId),
  };
}

export const selectUser = (demo: Demo, userId: Id): Demo => (demo.session.users.some((u) => u.id === userId) ? { ...demo, activeUserId: userId } : demo);

/** 初期化：同じ顔ぶれのまま、手の操作を消して最初の時刻に戻す（台本は残す） */
export function resetDemo(demo: Demo, project: Project, start: Minutes = DEMO_START): Demo {
  const events = demo.session.users.flatMap((u) => (demo.scripted.includes(u.id) ? scriptEvents(project, u.personaId, u.id) : []));
  return { ...demo, session: { ...demo.session, events }, now: start };
}

/**
 * ユーザーの操作。1回の操作で1分進める（同じ時刻に操作が重ならないように）。
 * その人のこの先の操作（台本や、時間を戻す前の操作）は消し、その人は台本から外れる。
 */
export function dispatch(demo: Demo, body: InputEventBody, userId: Id = demo.activeUserId): Demo {
  const at = demo.now + 1;
  const kept = demo.session.events.filter((e) => !(e.userId === userId && e.at > demo.now));
  return {
    ...demo,
    now: at,
    session: { ...demo.session, events: [...kept, { ...body, at, userId } as InputEvent] },
    scripted: demo.scripted.filter((id) => id !== userId),
  };
}

/** 時刻を進める（戻さない） */
export const advance = (demo: Demo, to: Minutes): Demo => (to > demo.now ? { ...demo, now: to } : demo);
/** 時刻を動かす（進める・戻す） */
export const seek = (demo: Demo, to: Minutes): Demo => ({ ...demo, now: Math.max(0, Math.round(to)) });

export const worldOf = (project: Project, demo: Demo): WorldState => reduce(project, demo.session, { until: demo.now });

/** このユーザーに、このあと予定されている待機・時間切れ（早い順） */
export function upcoming(world: WorldState, userId: Id): Timer[] {
  return world.timers.filter((t) => t.userId === userId && t.fireAt > world.now).sort((a, b) => a.fireAt - b.fireAt);
}

/** この先に残っている、その人の操作の数（台本・時間を戻す前の操作） */
export const futureEvents = (demo: Demo, userId: Id) => demo.session.events.filter((e) => e.userId === userId && e.at > demo.now).length;

/** 次に何かが起きる時刻（誰かの配信予定か、誰かの操作）。無ければ null */
export function nextStop(world: WorldState, demo: Demo): Minutes | null {
  let t = Infinity;
  for (const timer of world.timers) if (timer.fireAt > demo.now && timer.fireAt < t) t = timer.fireAt;
  for (const e of demo.session.events) if (e.at > demo.now && e.at < t) t = e.at;
  return t === Infinity ? null : t;
}

/** 前に何かが起きた時刻（巻き戻しの「前の出来事へ」） */
export function prevStop(world: WorldState, demo: Demo): Minutes | null {
  let t = -Infinity;
  const consider = (x: number) => {
    if (x < demo.now && x > t) t = x;
  };
  for (const u of Object.values(world.users)) {
    for (const d of u.deliveries) consider(d.at);
    for (const b of u.behaviors) consider(b.at);
  }
  return t === -Infinity ? null : t;
}

export type Pending = { userId: Id; kind: 'survey' | 'choice' };

/**
 * 手で操作するユーザーのうち、回答を待たれている人（アンケートの回答・ボタン選択）。
 * 自動再生はここで止まる。台本つきのユーザーは台本が答えるので含めない。
 */
export function pendingInputs(project: Project, world: WorldState, demo: Demo): Pending[] {
  const out: Pending[] = [];
  for (const u of demo.session.users) {
    if (demo.scripted.includes(u.id)) continue;
    const state = world.users[u.id];
    if (!state || state.friendState !== 'friend') continue;
    if (state.surface.kind === 'survey') {
      out.push({ userId: u.id, kind: 'survey' });
      continue;
    }
    const waitingChoice = Object.entries(state.flows).some(([flowId, c]) => {
      if (c.status !== 'waiting' || !c.nodeId) return false;
      const node = project.flows.find((f) => f.id === flowId)?.nodes.find((n) => n.id === c.nodeId);
      return node?.type === 'choice';
    });
    if (waitingChoice) out.push({ userId: u.id, kind: 'choice' });
  }
  return out;
}

export type StepResult = { demo: Demo; stopped: null | 'pending' | 'end'; pending?: Pending };

/**
 * 自動再生の1コマ：回答待ちの人がいれば止まる。いなければ次に何かが起きる時刻まで進める。
 * 再生の速さはこの呼び出しの間隔だけで決まるので、速さを変えても結果は変わらない。
 */
export function autoStep(project: Project, demo: Demo, endDay = project.settings.demoEndDay ?? DEMO_END_DAY): StepResult {
  const world = worldOf(project, demo);
  const pending = pendingInputs(project, world, demo)[0];
  if (pending) return { demo, stopped: 'pending', pending };
  const t = nextStop(world, demo);
  const end = endDay * MINUTES_PER_DAY;
  if (t === null || t > end) return { demo: advance(demo, end), stopped: 'end' };
  return { demo: seek(demo, t), stopped: null };
}
