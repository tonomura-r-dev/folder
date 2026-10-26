// 条件式・配信制御・時間進行・差分適用・エンジン検査の単体テスト。

import { describe, expect, it } from 'vitest';
import type { Project, Session, UserState } from '../packages/engine/src/index.ts';
import { applyPatch, approve, evaluate, failingLeaves, reduce, STEP_BUDGET, validate } from '../packages/engine/src/index.ts';
import { scanEngineSource } from '../tools/engine-scan.ts';
import { baseProject, chain, DAY, edge, node, sample } from './helpers.ts';

const friendAt = (p: Project, at = 0, extra: Session['events'] = []): Session => ({
  schemaVersion: 1,
  projectId: p.meta.id,
  users: [{ id: 'u1', personaId: 'u1' }],
  events: [{ at, userId: 'u1', type: 'addFriend', routeId: 'r1' }, ...extra],
});

describe('条件式', () => {
  const user = { tags: { tag_a: { at: 0 } }, answers: {}, cvs: {}, behaviors: [], deliveries: [], flows: {}, attrs: { age: 30 }, friendState: 'friend', friendAddedAt: 0, route: 'r1' } as unknown as UserState;
  const ctx = { project: baseProject(), user, now: 2 * DAY };

  it('AND は打ち切らずに全部評価し、偽の葉だけを理由として返す', () => {
    const t = evaluate(
      { op: 'and', children: [{ op: 'hasTag', tagId: 'tag_a' }, { op: 'cv', cvId: 'cv1', done: true }, { op: 'attr', key: 'age', cmp: 'gte', value: 40 }] },
      ctx,
    );
    expect(t.value).toBe(false);
    expect(t.children).toHaveLength(3);
    expect(failingLeaves(t).map((l) => l.op)).toEqual(['cv', 'attr']);
  });

  it('NOT が偽のときは、中身が真だったことを理由にする', () => {
    const t = evaluate({ op: 'not', child: { op: 'hasTag', tagId: 'tag_a' } }, ctx);
    expect(t.value).toBe(false);
    expect(failingLeaves(t)[0]!.op).toBe('not');
  });

  it('OR・経過時間・経路', () => {
    expect(evaluate({ op: 'or', children: [{ op: 'cv', cvId: 'cv1', done: true }, { op: 'route', routeIds: ['r1'] }] }, ctx).value).toBe(true);
    expect(evaluate({ op: 'elapsed', since: 'friendAdd', cmp: 'gte', minutes: DAY }, ctx).value).toBe(true);
    expect(evaluate({ op: 'elapsed', since: 'friendAdd', cmp: 'lt', minutes: DAY }, ctx).value).toBe(false);
  });
});

describe('配信制御', () => {
  it('優先順位：同時刻に排他フローが2本配信しようとすると、優先度の低い方は見送る', () => {
    const p = baseProject({
      flows: [
        chain('hi', [node({ id: 'hi_m', type: 'message', messageId: 'm1' })], { priority: 1 }),
        chain('lo', [node({ id: 'lo_m', type: 'message', messageId: 'm2' })], { priority: 2 }),
      ],
    });
    const w = reduce(p, friendAt(p));
    expect(w.users['u1']!.deliveries.map((d) => d.messageId)).toEqual(['m1']);
    expect(w.decisions.find((d) => d.ref === 'm2' && d.kind === 'deliver')).toMatchObject({ outcome: 'deferred', reason: 'lower_priority' });
  });

  it('並行配信を許可したフローは見送られない', () => {
    const p = baseProject({
      flows: [
        chain('hi', [node({ id: 'hi_m', type: 'message', messageId: 'm1' })], { priority: 1 }),
        chain('lo', [node({ id: 'lo_m', type: 'message', messageId: 'm2' })], { priority: 2, concurrency: 'parallel' }),
      ],
    });
    expect(reduce(p, friendAt(p)).users['u1']!.deliveries.map((d) => d.messageId)).toEqual(['m1', 'm2']);
  });

  it('重複防止：期間内に同じメッセージは2回届かない', () => {
    const p = baseProject({
      settings: { dedupeWindowMinutes: DAY },
      flows: [
        chain('f', [
          node({ id: 'a', type: 'message', messageId: 'm1' }),
          node({ id: 'w', type: 'wait', wait: { kind: 'duration', minutes: 60 } }),
          node({ id: 'b', type: 'message', messageId: 'm1' }),
          node({ id: 'w2', type: 'wait', wait: { kind: 'duration', minutes: DAY } }),
          node({ id: 'c', type: 'message', messageId: 'm1' }),
        ]),
      ],
    });
    const w = reduce(p, friendAt(p), { until: 3 * DAY });
    expect(w.users['u1']!.deliveries.map((d) => d.at)).toEqual([0, 60 + DAY]);
    expect(w.decisions.find((d) => d.nodeId === 'b' && d.kind === 'deliver')).toMatchObject({ outcome: 'deduped' });
  });

  it('待機中に CV すると、その場でタイマーが破棄される', () => {
    const p = baseProject({
      flows: [
        chain('f', [node({ id: 'w', type: 'wait', wait: { kind: 'duration', minutes: DAY } }), node({ id: 'm', type: 'message', messageId: 'm1' })], {
          stopWhen: { op: 'cv', cvId: 'cv1', done: true },
        }),
      ],
    });
    const w = reduce(p, friendAt(p, 0, [{ at: 60, userId: 'u1', type: 'cv', cvId: 'cv1' }]), { until: 3 * DAY });
    expect(w.users['u1']!.flows['f']!.status).toBe('stopped');
    expect(w.timers).toEqual([]);
    expect(w.users['u1']!.deliveries).toEqual([]);
  });

  it('「N日目のHH:MM」は友だち追加の日を起点にする', () => {
    const p = baseProject({
      flows: [chain('f', [node({ id: 'w', type: 'wait', wait: { kind: 'dayAt', day: 2, time: '09:30' } }), node({ id: 'm', type: 'message', messageId: 'm1' })])],
    });
    const w = reduce(p, friendAt(p, 22 * 60), { until: 5 * DAY });
    expect(w.users['u1']!.deliveries[0]!.at).toBe(2 * DAY + 9 * 60 + 30);
  });

  it('ボタン選択：押したボタンの出口に進む', () => {
    const p = baseProject({
      messages: [
        { id: 'q', bubbles: [{ kind: 'buttons', text: 'q', buttons: [{ id: 'yes', label: 'y', action: { kind: 'none' } }, { id: 'no', label: 'n', action: { kind: 'none' } }] }] },
        { id: 'm_yes', bubbles: [{ kind: 'text', text: 'y' }] },
        { id: 'm_no', bubbles: [{ kind: 'text', text: 'n' }] },
      ],
      flows: [chain('f', [node({ id: 'c', type: 'choice', messageId: 'q' })])],
    });
    const f = p.flows[0]!;
    f.nodes.push(node({ id: 'y', type: 'message', messageId: 'm_yes' }), node({ id: 'n', type: 'message', messageId: 'm_no' }));
    f.edges.push(edge('c', 'y', 'btn:yes'), edge('c', 'n', 'btn:no'));
    const w = reduce(p, friendAt(p, 0, [{ at: 5, userId: 'u1', type: 'tapButton', buttonId: 'no', nodeId: 'c' }]));
    expect(w.users['u1']!.deliveries.map((d) => d.messageId)).toEqual(['q', 'm_no']);
  });

  it('時間が進まないループは、実行時に上限で止めて報告する（固まらない）', () => {
    const p = baseProject({ flows: [chain('f', [node({ id: 'a', type: 'tag', add: ['tag_a'] })])] });
    p.flows[0]!.edges.push(edge('a', 'a'));
    const w = reduce(p, friendAt(p));
    expect(w.warnings[0]).toMatchObject({ code: 'runtime.stepBudget' });
    expect(w.users['u1']!.flows['f']!.status).toBe('stopped');
    expect(STEP_BUDGET).toBe(10_000);
  });
});

describe('シナリオ差分（AI連携）', () => {
  const patch = {
    schemaVersion: 1 as const,
    summary: 'test',
    ops: [{ op: 'add' as const, path: '/tags/-', value: { id: 'tag_new', label: '新' }, prov: { state: 'inferred' as const, rationale: 'メニューから推定' } }],
  };

  it('推定で足したものは、承認するまで重大エラー', () => {
    const base = sample();
    const { project, failed, issues } = applyPatch(base, patch);
    expect(failed).toEqual([]);
    expect(issues).toContainEqual(expect.objectContaining({ code: 'prov.unapproved', path: `/tags/${base.tags.length}` }));
    // 元の案件は変わらない
    expect(base.tags.find((t) => t.id === 'tag_new')).toBeUndefined();
    const ok = approve(project, `/tags/${base.tags.length}`);
    expect(validate(ok).filter((i) => i.code === 'prov.unapproved')).toEqual([]);
  });

  it('AIが「承認済み」を付けた差分は受け付けない', () => {
    const { issues } = applyPatch(sample(), { ...patch, ops: [{ ...patch.ops[0]!, prov: { state: 'approved' } }] });
    expect(issues).toContainEqual(expect.objectContaining({ code: 'patch.approvedByAi' }));
  });

  it('存在しない場所への差分は failed に入る', () => {
    const { failed } = applyPatch(sample(), { ...patch, ops: [{ op: 'replace', path: '/flows/99/label', value: 'x', prov: { state: 'observed' } }] });
    expect(failed).toHaveLength(1);
  });
});

describe('エンジンの汎用性検査', () => {
  it('文字列の日本語・禁止API・外部 import を見つけ、コメントの日本語は許す', () => {
    const src = [
      '// 日本語のコメントは問題ない',
      "import { x } from '../../tools/a.ts';",
      "const s = '応募';",
      'const t = Date.now();',
      'const ok = "apply";',
    ].join('\n');
    const problems = scanEngineSource(src, 'x.ts');
    expect(problems).toHaveLength(3);
    expect(problems.join('\n')).toMatch(/外の import[\s\S]*日本語[\s\S]*Date/);
  });
});

describe('自動応答と重複防止（Phase 6 で見つけた不具合の修正）', () => {
  it('少し前に配信した文面でも、ユーザーが頼んだ自動応答は届く。配信の側は重複防止が効く', () => {
    const p = baseProject({
      settings: { dedupeWindowMinutes: DAY },
      autoReplies: [{ id: 'ar1', priority: 1, match: { mode: 'exact', keywords: ['クーポン'] }, cases: [], fallbackMessageId: 'm1' }],
      flows: [
        chain('f', [
          node({ id: 'a', type: 'message', messageId: 'm1' }),
          node({ id: 'w', type: 'wait', wait: { kind: 'duration', minutes: 120 } }),
          node({ id: 'b', type: 'message', messageId: 'm1' }),
        ]),
      ],
    });
    const w = reduce(p, friendAt(p, 0, [{ at: 60, userId: 'u1', type: 'sendText', text: 'クーポン' }]), { until: DAY });
    // 0分：配信 → 60分：頼まれた返信は届く → 180分：同じ文面の配信は重複防止で止まる
    expect(w.users['u1']!.deliveries.map((d) => [d.at, d.count])).toEqual([[0, 1], [60, 0]]);
    expect(w.decisions.find((d) => d.nodeId === 'b' && d.kind === 'deliver')).toMatchObject({ outcome: 'deduped' });
  });
});
