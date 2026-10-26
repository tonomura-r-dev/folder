// デモの進行（画面を持たない部分）。画面のボタンはすべてここを通してエンジンへの入力イベントにする。
// 状態はイベント列と「いまの時刻」だけ。画面に出すものは毎回 reduce で作り直す（DATA_MODEL §1）。

import type { Id, InputEventBody, Minutes, Project, Session, Timer, WorldState } from '../../engine/src/index.ts';
import { reduce } from '../../engine/src/index.ts';

/** デモを始める時刻（Day 0 の 10:00） */
export const DEMO_START: Minutes = 600;
export const DEMO_USER = 'demo';

export type Demo = { session: Session; now: Minutes; userId: Id };

export function newDemo(project: Project, personaId?: Id, start: Minutes = DEMO_START): Demo {
  const persona = project.personas.find((p) => p.id === personaId) ?? project.personas[0];
  return {
    session: {
      schemaVersion: 1,
      projectId: project.meta.id,
      users: [{ id: DEMO_USER, personaId: persona?.id ?? '', ...(persona ? {} : { displayName: 'あなた' }) }],
      events: [],
    },
    now: start,
    userId: DEMO_USER,
  };
}

/** ユーザーの操作。1回の操作で1分進める（同じ時刻に操作が重ならないように） */
export function dispatch(demo: Demo, body: InputEventBody): Demo {
  const at = demo.now + 1;
  return { ...demo, now: at, session: { ...demo.session, events: [...demo.session.events, { ...body, at, userId: demo.userId }] } };
}

/** 時刻を進める（戻すのは Phase 4） */
export const advance = (demo: Demo, to: Minutes): Demo => (to > demo.now ? { ...demo, now: to } : demo);

export const worldOf = (project: Project, demo: Demo): WorldState => reduce(project, demo.session, { until: demo.now });

/** このユーザーに、このあと予定されている待機・時間切れ（早い順） */
export function upcoming(world: WorldState, userId: Id): Timer[] {
  return world.timers.filter((t) => t.userId === userId && t.fireAt > world.now).sort((a, b) => a.fireAt - b.fireAt);
}
