// デモの進行（packages/player/src/controller.ts）。画面のボタンと同じ関数で、ペルソナの台本どおりに操作する。
// 「画面で同じ操作をしたときの配信履歴が、CLI（Phase 2a）の結果と一致する」の確認。

import { describe, expect, it } from 'vitest';
import { reduce, sessionFromPersonas } from '../packages/engine/src/index.ts';
import type { InputEventBody } from '../packages/engine/src/index.ts';
import { advance, dispatch, newDemo, upcoming, worldOf, type Demo } from '../packages/player/src/controller.ts';
import { DAY, sample } from './helpers.ts';

const project = sample();
const cli = reduce(project, sessionFromPersonas(project), { until: 30 * DAY });

/** 台本の時刻ちょうどに操作するよう、1分前まで時間を進めてから dispatch する（画面の「時間を進める」→ボタン と同じ） */
function play(personaId: string): Demo {
  let demo = newDemo(project, personaId, 0);
  for (const step of project.personas.find((p) => p.id === personaId)!.script ?? []) {
    demo = advance(demo, step.atMinutes - 1);
    demo = dispatch(demo, step.event);
  }
  return advance(demo, 30 * DAY);
}

describe('画面の操作と CLI の結果が一致する', () => {
  for (const p of project.personas) {
    it(`${p.name}：配信履歴・タグ・CVが同じ`, () => {
      const demo = play(p.id);
      const ui = worldOf(project, demo).users[demo.activeUserId]!;
      const ref = cli.users[p.id]!;
      expect(ui.deliveries).toEqual(ref.deliveries);
      expect(Object.keys(ui.tags).sort()).toEqual(Object.keys(ref.tags).sort());
      expect(ui.cvs).toEqual(ref.cvs);
      expect(ui.friendState).toBe(ref.friendState);
    });
  }
});

describe('デモの進行', () => {
  it('操作のたびに1分進み、イベントがその時刻で記録される', () => {
    let d = newDemo(project, 'p_tokyo');
    d = dispatch(d, { type: 'visitPage', pageId: 'page_lp' });
    d = dispatch(d, { type: 'exitIntent', pageId: 'page_lp' });
    expect(d.session.events.map((e) => e.at)).toEqual([601, 602]);
    expect(worldOf(project, d).users[d.activeUserId]!.surface).toEqual({ kind: 'popup', pageId: 'page_lp' });
  });

  it('友だち追加すると、次に届く予定（Day 1 09:00 のステップ）が見える', () => {
    let d = newDemo(project, 'p_tokyo');
    const steps: InputEventBody[] = [
      { type: 'visitPage', pageId: 'page_lp' },
      { type: 'exitIntent', pageId: 'page_lp' },
      { type: 'tapPopupCta', pageId: 'page_lp' },
      { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] },
      { type: 'answer', surveyId: 'sv_hope', questionId: 'q_job', choiceIds: ['c_office'] },
    ];
    for (const ev of steps) d = dispatch(d, ev);
    const w = worldOf(project, d);
    expect(upcoming(w, d.activeUserId)[0]).toMatchObject({ flowId: 'flow_step_tokyo', fireAt: DAY + 9 * 60 });
    // 予定の時刻まで進めると届く
    const later = worldOf(project, advance(d, DAY + 9 * 60));
    expect(later.users[d.activeUserId]!.deliveries.at(-1)!.messageId).toBe('msg_t_d1');
  });

  it('ペルソナが無い案件でも始められる', () => {
    const p = { ...project, personas: [] };
    const d = newDemo(p);
    expect(worldOf(p, d).users[d.activeUserId]!.name).toBe('あなた');
  });

  it('壊れた条件式（編集途中）でも止まらない', () => {
    const p = structuredClone(project);
    const n = p.flows[0]!.nodes.find((x) => x.id === 'n_w_area')!;
    if (n.type === 'branch') n.cases[0]!.when = { op: 'xor' } as never;
    let d = newDemo(p, 'p_tokyo');
    d = dispatch(d, { type: 'addFriend', routeId: 'route_exit_popup' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] });
    expect(() => worldOf(p, dispatch(d, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_job', choiceIds: ['c_office'] }))).not.toThrow();
  });
});
