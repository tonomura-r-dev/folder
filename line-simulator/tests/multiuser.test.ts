// Phase 4：複数ユーザー・時間の前進／巻き戻し／初期化・自動再生（packages/player/src/controller.ts）。

import { describe, expect, it } from 'vitest';
import type { InputEventBody, Project } from '../packages/engine/src/index.ts';
import { MINUTES_PER_DAY, personaCoverage, reduce, sessionFromPersonas } from '../packages/engine/src/index.ts';
import {
  addUser,
  allScripted,
  autoStep,
  dispatch,
  futureEvents,
  newDemo,
  nextStop,
  pendingInputs,
  removeUser,
  resetDemo,
  seek,
  worldOf,
  type Demo,
} from '../packages/player/src/controller.ts';
import { DAY, sample } from './helpers.ts';

const project = sample();
const cli = reduce(project, sessionFromPersonas(project), { until: 30 * DAY });

/** 自動再生を最後まで（止まったら終わり） */
function playToEnd(demo: Demo, p: Project = project): { demo: Demo; steps: number; stopped: string } {
  let steps = 0;
  for (;;) {
    const r = autoStep(p, demo);
    demo = r.demo;
    if (r.stopped) return { demo, steps, stopped: r.stopped };
    if (++steps > 10_000) throw new Error('終わらない');
  }
}

const personaOf = (demo: Demo, userId: string) => demo.session.users.find((u) => u.id === userId)!.personaId;

describe('複数ユーザー', () => {
  it('台本つきの全員（5人）を自動再生で最後まで流すと、一人ずつ CLI と同じ結果になる（状態が混ざらない）', () => {
    const { demo, stopped } = playToEnd(allScripted(project));
    expect(stopped).toBe('end');
    const w = worldOf(project, demo);
    expect(Object.keys(w.users)).toHaveLength(5);
    for (const [id, u] of Object.entries(w.users)) {
      const ref = cli.users[personaOf(demo, id)]!;
      expect(u.deliveries).toEqual(ref.deliveries);
      expect(Object.keys(u.tags).sort()).toEqual(Object.keys(ref.tags).sort());
      expect(u.flows).toEqual(ref.flows);
    }
  });

  it('同じペルソナを2人入れても、名前が分かれ、状態は別々', () => {
    let d = newDemo(project, 'p_tokyo');
    d = addUser(d, project, 'p_tokyo', false);
    expect(d.session.users.map((u) => u.displayName ?? '')).toEqual(['', '東京希望さん（2）']);
    d = dispatch(d, { type: 'addFriend', routeId: 'route_exit_popup' }, 'u1');
    const w = worldOf(project, d);
    expect(w.users['u1']!.friendState).toBe('friend');
    expect(w.users['u2']!.friendState).toBe('not_friend');
  });

  it('ユーザーを外すと、その人の操作だけ消え、ほかの人は変わらない', () => {
    const end = playToEnd(allScripted(project)).demo;
    const before = worldOf(project, end);
    const after = worldOf(project, removeUser(end, 'u2'));
    expect(after.users['u2']).toBeUndefined();
    expect(after.users['u1']!.deliveries).toEqual(before.users['u1']!.deliveries);
    expect(after.users['u3']!.deliveries).toEqual(before.users['u3']!.deliveries);
  });
});

describe('時間の巻き戻し・初期化', () => {
  it('Day 7 から Day 3 に戻すと Day 3 の状態になり、また進めると同じ Day 7 になる', () => {
    let d = allScripted(project);
    while (d.now < 7 * DAY) {
      const r = autoStep(project, d);
      if (r.stopped) break;
      d = r.demo;
    }
    d = seek(d, 7 * DAY);
    const day7 = worldOf(project, d);
    const back = seek(d, 3 * DAY);
    const day3 = worldOf(project, back);
    for (const u of Object.values(day3.users)) expect(u.deliveries.every((x) => x.at <= 3 * DAY)).toBe(true);
    expect(day3).toEqual(reduce(project, d.session, { until: 3 * DAY }));
    expect(worldOf(project, seek(back, 7 * DAY))).toEqual(day7);
  });

  it('過去に戻って手で操作すると、その人のそれより後の操作だけが消える', () => {
    let d = playToEnd(allScripted(project)).demo;
    const u1Future = d.session.events.filter((e) => e.userId === 'u1').length;
    d = seek(d, 0);
    expect(futureEvents(d, 'u1')).toBe(u1Future);
    d = dispatch(d, { type: 'visitPage', pageId: 'page_lp' }, 'u1');
    expect(d.session.events.filter((e) => e.userId === 'u1')).toHaveLength(1);
    expect(d.scripted).not.toContain('u1'); // 手で操作した人は台本から外れる
    expect(d.session.events.filter((e) => e.userId === 'u2').length).toBeGreaterThan(0); // ほかの人の台本は残る
  });

  it('初期化すると、顔ぶれと台本は残り、手の操作は消えて最初の時刻に戻る', () => {
    let d = allScripted(project);
    d = addUser(d, project, 'p_osaka', false);
    d = dispatch(d, { type: 'addFriend', routeId: 'route_exit_popup' });
    const r = resetDemo(d, project);
    expect(r.session.users).toHaveLength(6);
    expect(r.session.events.some((e) => e.userId === 'u6')).toBe(false);
    expect(r.session.events.filter((e) => e.userId === 'u1')).toEqual(d.session.events.filter((e) => e.userId === 'u1'));
    expect(r.now).toBe(600);
  });
});

describe('自動再生', () => {
  it('手で操作する人がアンケートで回答待ちになると止まり、答えると先へ進む', () => {
    let d = newDemo(project, 'p_tokyo');
    d = dispatch(d, { type: 'addFriend', routeId: 'route_exit_popup' });
    const r1 = autoStep(project, d);
    expect(r1.stopped).toBe('pending');
    expect(r1.pending).toEqual({ userId: 'u1', kind: 'survey' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_job', choiceIds: ['c_office'] });
    const r2 = autoStep(project, d);
    expect(r2.stopped).toBeNull();
    expect(r2.demo.now).toBe(DAY + 9 * 60); // 次の配信（Day 1 09:00）まで進む
  });

  it('ボタン選択で待っているときも止まる（Day 7 の「応募する／今回は見送る」）', () => {
    let d = newDemo(project, 'p_tokyo');
    const steps: InputEventBody[] = [
      { type: 'addFriend', routeId: 'route_exit_popup' },
      { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] },
      { type: 'answer', surveyId: 'sv_hope', questionId: 'q_job', choiceIds: ['c_office'] },
    ];
    for (const body of steps) d = dispatch(d, body);
    const { stopped, demo } = playToEnd(d);
    expect(stopped).toBe('pending');
    expect(pendingInputs(project, worldOf(project, demo), demo)[0]).toEqual({ userId: 'u1', kind: 'choice' });
    expect(demo.now).toBe(7 * DAY + 12 * 60);
  });

  it('台本つきの人は回答待ちでも止まらない', () => {
    const { stopped, steps } = playToEnd(allScripted(project));
    expect(stopped).toBe('end');
    expect(steps).toBeGreaterThan(10);
  });

  it('再生の刻み方（速さ）を変えても結果は同じ：1コマずつ進めても、最後へ一気に飛んでも一致する', () => {
    const stepped = worldOf(project, playToEnd(allScripted(project)).demo);
    const jumped = worldOf(project, seek(allScripted(project), 30 * MINUTES_PER_DAY));
    expect(stepped).toEqual(jumped);
  });

  it('次に起きることが無ければ null', () => {
    const d = seek(allScripted(project), 60 * DAY);
    expect(nextStop(worldOf(project, d), d)).toBeNull();
  });
});

describe('ペルソナで流したときの未到達チェック', () => {
  it('検証案件は全シナリオに誰かが入る', () => {
    expect(personaCoverage(project)).toEqual([]);
  });

  it('誰も入らないシナリオは警告になる', () => {
    const p = structuredClone(project);
    p.flows.push({
      id: 'flow_orphan', label: '孤立', trigger: { kind: 'tagAdded', tagId: 'job_light' }, priority: 50, concurrency: 'parallel', reentry: 'never',
      nodes: [{ id: 'n_orphan_entry', type: 'entry', position: { x: 0, y: 0 } }, { id: 'n_orphan_msg', type: 'message', messageId: p.messages[0]!.id, position: { x: 0, y: 150 } }],
      edges: [{ id: 'e_orphan', from: { nodeId: 'n_orphan_entry', port: 'out' }, to: 'n_orphan_msg' }],
    });
    // 軽作業タグが付くのは大阪希望さんだけなので、そのタグを付けないようにすると誰も入らない
    p.personas = p.personas.filter((x) => x.id !== 'p_osaka' && x.id !== 'p_late');
    const issues = personaCoverage(p);
    // 大阪を選ぶ人（大阪希望さん・途中で回答さん）がいないので、大阪のステップ配信にも誰も入らない
    expect(issues.map((i) => i.params?.['id'])).toEqual(['flow_step_osaka', 'flow_orphan']);
    expect(issues.every((i) => i.level === 'warning' && i.code === 'persona.noneMatch')).toBe(true);
  });
});
