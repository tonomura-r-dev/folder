// Phase 2a の検収条件（docs/ROADMAP.md）を、エンジンの状態で直接確かめる。

import { describe, expect, it } from 'vitest';
import { dayAt, failingLeaves, reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { summarize } from '../tools/summary.ts';
import { DAY, sample } from './helpers.ts';

const project = sample();
const world = reduce(project, sessionFromPersonas(project), { until: 30 * DAY });
const user = (id: string) => world.users[id]!;
const delivered = (id: string) => user(id).deliveries.map((d) => d.messageId);

describe('検証案件 sample-staff', () => {
  it('重大エラーも警告もない', () => {
    expect(validate(project)).toEqual([]);
  });

  it('東京希望さん：東京のステップが届き、大阪・共通は届かない', () => {
    const got = delivered('p_tokyo');
    expect(got).toEqual(expect.arrayContaining(['msg_welcome', 'msg_t_d1', 'msg_t_office', 'msg_t_d7']));
    expect(got.some((m) => m.startsWith('msg_o_') || m.startsWith('msg_c_'))).toBe(false);
    // 事務を選んだので軽作業の求人は届かない
    expect(got).not.toContain('msg_t_light');
    // Day 1 の 09:00 に届く
    expect(user('p_tokyo').deliveries.find((d) => d.messageId === 'msg_t_d1')!.at).toBe(dayAt(0, 1, '09:00'));
  });

  it('大阪希望さん：東京向けから外れた理由が「東京のタグが無い」として残る', () => {
    expect(delivered('p_osaka').some((m) => m.startsWith('msg_t_'))).toBe(false);
    expect(delivered('p_osaka')).toEqual(expect.arrayContaining(['msg_o_d1', 'msg_o_d3', 'msg_o_d7']));
    const skipped = world.decisions.find((d) => d.userId === 'p_osaka' && d.nodeId === 'n_w_area' && d.ref === 'tokyo');
    expect(skipped?.outcome).toBe('skipped');
    const leaves = failingLeaves(skipped!.trace!);
    expect(leaves.map((l) => l.expr)).toEqual([{ op: 'hasTag', tagId: 'area_tokyo' }]);
  });

  it('大阪希望さん：同じ「求人」でもタグに応じて大阪の返信になる', () => {
    expect(delivered('p_osaka')).toContain('msg_r_osaka');
    expect(delivered('p_tokyo')).toContain('msg_r_tokyo');
    // 応答メッセージは通数に数えない
    expect(user('p_osaka').deliveries.find((d) => d.messageId === 'msg_r_osaka')!.count).toBe(0);
  });

  it('応募済みさん：東京ステップが停止条件で止まり、求人は1通も届かない', () => {
    expect(user('p_applied').flows['flow_step_tokyo']!.status).toBe('stopped');
    const stop = world.decisions.find((d) => d.userId === 'p_applied' && d.flowId === 'flow_step_tokyo' && d.kind === 'stopWhen');
    expect(stop?.reason).toBe('stop_condition');
    expect(delivered('p_applied').some((m) => m.startsWith('msg_t_'))).toBe(false);
    expect(delivered('p_applied')).toContain('msg_welcome_applied');
  });

  it('応募済みさん：再応募で応募後フォローがもう一度始まる', () => {
    expect(user('p_applied').cvs['cv_apply']!.count).toBe(2);
    expect(user('p_applied').flows['flow_after_apply']!.runs).toBe(2);
  });

  it('未応募さん：未回答のまま共通ステップに流れ、ブロック後の配信は「ブロック中」で除外される', () => {
    expect(user('p_notyet').answers).toEqual({});
    expect(user('p_notyet').flows['flow_step_common']).toBeDefined();
    expect(delivered('p_notyet')).toContain('msg_c_ask');
    expect(delivered('p_notyet')).not.toContain('msg_c_general');
    const excluded = world.decisions.find((d) => d.userId === 'p_notyet' && d.ref === 'msg_c_general');
    expect(excluded).toMatchObject({ outcome: 'excluded', reason: 'blocked' });
  });

  it('友だちになる前は配信されない（経路Bは応募→友だち追加の順）', () => {
    const applied = user('p_applied');
    expect(applied.friendAddedAt).toBe(611);
    expect(applied.deliveries.every((d) => d.at >= 611)).toBe(true);
  });

  it('同じ入力を2回流すと、結果が完全に一致する（決定論）', () => {
    const again = reduce(project, sessionFromPersonas(project), { until: 30 * DAY });
    expect(JSON.stringify(again)).toBe(JSON.stringify(world));
  });

  it('Day 3 まで流した結果は、Day 30 まで流した結果の途中と一致する（任意のDayへ戻れる）', () => {
    const day3 = reduce(project, sessionFromPersonas(project), { until: 3 * DAY });
    for (const u of Object.values(day3.users)) {
      const full = user(u.id).deliveries.filter((d) => d.at <= 3 * DAY);
      expect(u.deliveries).toEqual(full);
    }
  });

  it('4人・Day 30 が 1秒以内', () => {
    const t = performance.now();
    reduce(project, sessionFromPersonas(project), { until: 30 * DAY });
    expect(performance.now() - t).toBeLessThan(1000);
  });

  it('通しの結果がスナップショットと一致する', async () => {
    await expect(JSON.stringify(summarize(project, world), null, 2)).toMatchFileSnapshot('__snapshots__/sample-staff.day30.json');
  });
});
