// 提案資料から作った企業別デモ案件（projects/client-*/）の検証。
// 案件データは顧客の資料にもとづくのでリポジトリに入れていない。手元に無いときは、このテストは飛ばす。
// テストには顧客名を書かず、ID だけで確かめる。

import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { Project, WorldState } from '../packages/engine/src/index.ts';
import { personaCoverage, reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { allScripted, autoStep, dispatch, newDemo, seek, worldOf, type Demo } from '../packages/player/src/controller.ts';
import { journeyOf } from '../packages/player/src/journey.ts';
import { buildExport } from '../tools/export-lib.ts';
import { loadProject } from '../tools/project-io.ts';
import { DAY } from './helpers.ts';

const has = (slug: string) => existsSync(join(import.meta.dirname, '..', 'projects', slug, 'project.json'));
const at = (day: number, hh = 0, mm = 0) => day * DAY + hh * 60 + mm;
const run = (p: Project, until: number) => reduce(p, sessionFromPersonas(p), { until });
const user = (w: WorldState, id: string) => w.users[id]!;
const got = (w: WorldState, id: string) => user(w, id).deliveries.map((d) => [d.messageId, d.at] as const);
const ids = (w: WorldState, id: string) => user(w, id).deliveries.map((d) => d.messageId);
const me = (p: Project, d: Demo) => worldOf(p, d).users[d.activeUserId]!;
function playToEnd(p: Project) {
  let d = allScripted(p);
  for (let i = 0; i < 10_000; i++) {
    const r = autoStep(p, d);
    d = r.demo;
    if (r.stopped) break;
  }
  return d;
}

async function exportChecks(p: Project) {
  const r = await buildExport(p, { readAsset: () => null, confirmFictional: true });
  expect(r.ok).toBe(true);
  // 見積・料金や社内情報は入れない
  for (const s of ['見積', '概算', '265,000', '260,000', '月額', '_internal', '"prov"']) expect(r.html).not.toContain(s);
  expect(r.html).toContain("default-src 'none'");
  expect(r.html!.length).toBeLessThan(5 * 1024 * 1024);
}

describe.skipIf(!has('client-a-house'))('企業別デモ①（注文住宅・来場予約）', () => {
  const p = has('client-a-house') ? loadProject('client-a-house') : (null as unknown as Project);
  const w = has('client-a-house') ? run(p, 30 * DAY) : (null as unknown as WorldState);

  it('検証で重大エラー・警告が無く、全シナリオに誰かが入る', () => {
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
  });

  it('14日間・全5通のステップ配信（当日／1日後／3日後／6日後／12日後）', () => {
    const step = user(w, 'p_consider').deliveries.filter((d) => d.flowId === 'flow_step').map((d) => [d.messageId, d.at]);
    expect(step).toEqual([
      ['msg_s1_cases', at(0, 11, 3)],
      ['msg_s2_land', at(1, 19)],
      ['msg_s3_loan', at(3, 19)],
      ['msg_s4_flat', at(6, 19)],
      ['msg_s5_winter', at(12, 19)],
    ]);
  });

  it('回答と行動で配信が切り替わる：平屋／2階建て／未回答、資料請求済み、費用・間取りへの関心', () => {
    expect(ids(w, 'p_consider')).toContain('msg_s4_flat');
    expect(ids(w, 'p_plan')).toContain('msg_s4_two');
    expect(ids(w, 'p_doc')).toContain('msg_s2_land_doc');
    expect(ids(w, 'p_doc')).not.toContain('msg_s2_land');
    expect(ids(w, 'p_cost')).toContain('msg_k_cost');
    expect(ids(w, 'p_plan')).toContain('msg_k_plan');
    expect(w.decisions.find((d) => d.userId === 'p_block' && d.ref === 'msg_s4_both')).toMatchObject({ outcome: 'excluded', reason: 'blocked' });
  });

  it('来場予約で来場促進（ステップ・企画）が止まり、予約確認・3日前・前日・来場後に切り替わる', () => {
    const cost = user(w, 'p_cost');
    const reserved = cost.cvs['cv_reserve']!.at;
    expect(cost.deliveries.filter((d) => d.at > reserved && (d.flowId === 'flow_step' || d.flowId === 'flow_k_event'))).toEqual([]);
    expect(cost.flows['flow_step']!.status).toBe('stopped');
    expect(got(w, 'p_cost').filter(([m]) => m.startsWith('msg_r_') || m === 'msg_v_thanks' || m === 'msg_k_after')).toEqual([
      ['msg_r_confirm', at(6, 21, 5)],
      ['msg_r_3days', at(10, 10)],
      ['msg_r_daybefore', at(12, 18)],
      ['msg_v_thanks', at(13, 17)],
      ['msg_k_after', at(16, 12)],
    ]);
    // 契約したら、契約前の後押しは止まる（ここでは届いた後に契約）
    expect(cost.flows['flow_visit']!.status).toBe('ended');
  });

  it('来場予約してから友だちになった人：予約確認は「友だち前」で届かず、予約済み向けのあいさつ。ステップは最初から止まる', () => {
    expect(w.decisions.find((d) => d.userId === 'p_reserved' && d.ref === 'msg_r_confirm')).toMatchObject({ outcome: 'excluded', reason: 'not_friend' });
    expect(ids(w, 'p_reserved')).toContain('msg_welcome_res');
    expect(user(w, 'p_reserved').deliveries.some((d) => d.flowId === 'flow_step')).toBe(false);
  });

  it('③：資料請求は中間CV、来場予約が GOAL。入口・関心のルートは②のタグ・経路と一致する', () => {
    const rows = (id: string) => journeyOf(p, w, user(w, id));
    expect(rows('p_cost').find((r) => r.isGoal)!.step.id).toBe('j_reserve');
    expect(rows('p_doc').find((r) => r.step.id === 'j_doc')!.status).toBe('done');
    expect(rows('p_cost').find((r) => r.step.id === 'j_entry')!.traversed).toEqual(['v_button']);
    for (const id of ['p_consider', 'p_cost', 'p_plan']) {
      const seg = rows(id).find((r) => r.step.id === 'j_segment')!;
      const tags = user(w, id).tags;
      expect(seg.variantId).toBe('int_cost' in tags ? 'v_cost' : 'int_plan' in tags ? 'v_plan' : 'v_all');
    }
    expect(rows('p_cost').at(-1)!.status).toBe('current');
  });

  it('過去に戻って別の答えを選ぶと、その先の配信は前の答えを引き継がず計算し直される', () => {
    let d = newDemo(p, 'p_consider');
    d = dispatch(d, { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_status', choiceIds: ['c_compare'] });
    const before = d.now;
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_type', choiceIds: ['c_flat'] });
    d = seek(d, at(7));
    expect(me(p, d).deliveries.map((x) => x.messageId)).toContain('msg_s4_flat');
    // 回答の前に戻って、2階建てを選び直す
    d = seek(d, before);
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_type', choiceIds: ['c_two'] });
    d = seek(d, at(7));
    const now = me(p, d);
    expect(now.deliveries.map((x) => x.messageId)).toContain('msg_s4_two');
    expect(now.deliveries.map((x) => x.messageId)).not.toContain('msg_s4_flat');
    expect('house_flat' in now.tags).toBe(false);
  });

  it('自動再生で全員を最後まで流すと CLI と同じ結果になり、営業用HTMLを書き出せる', async () => {
    const d = playToEnd(p);
    const ww = worldOf(p, d);
    for (const u of Object.values(ww.users)) expect(u.deliveries).toEqual(user(w, d.session.users.find((x) => x.id === u.id)!.personaId).deliveries);
    await exportChecks(p);
  }, 30_000);
});

describe.skipIf(!has('client-b-kids'))('企業別デモ②（保育・見学予約）', () => {
  const p = has('client-b-kids') ? loadProject('client-b-kids') : (null as unknown as Project);
  const w = has('client-b-kids') ? run(p, 95 * DAY) : (null as unknown as WorldState);

  it('検証で重大エラー・警告が無く、全シナリオに誰かが入る。自動再生は Day 95 まで', () => {
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
    expect(p.settings.demoEndDay).toBe(95);
  });

  it('ステップ配信10本：Day0 挨拶・アンケート、Day1〜Day30（未予約の人には全部届く）', () => {
    const step = user(w, 'p_info').deliveries.filter((d) => d.flowId === 'flow_step').map((d) => Math.floor(d.at / DAY));
    expect(step).toEqual([1, 3, 5, 7, 10, 14, 18, 23, 30]);
    expect(ids(w, 'p_info')[0]).toBe('msg_welcome');
  });

  it('Q4「今すぐ」は予約への案内が直接届き、ほかは Q3 の重視点に合う案内が届く', () => {
    expect(ids(w, 'p_now')).toContain('msg_tour_now');
    expect(ids(w, 'p_now').some((m) => m.startsWith('msg_f_'))).toBe(false);
    expect(ids(w, 'p_later')).toContain('msg_f_iq');
    expect(ids(w, 'p_info')).toContain('msg_f_care');
    expect(ids(w, 'p_other')).toContain('msg_f_sports');
    // 簡易アンケートに答えない人には、本アンケートも出さない
    expect(ids(w, 'p_skip')).not.toContain('msg_main_intro');
  });

  it('見学予約で予約促進（ステップ・休眠B）が止まり、来場前のご案内とメニューの切り替えになる', () => {
    const later = user(w, 'p_later');
    const t = later.cvs['cv_tour']!.at;
    expect(later.deliveries.filter((d) => d.at > t && (d.flowId === 'flow_step' || d.flowId === 'flow_dormant_b'))).toEqual([]);
    expect(ids(w, 'p_later')).toEqual(expect.arrayContaining(['msg_t_confirm', 'msg_t_daybefore']));
    // 予約中は予約済みのメニュー、見学のあとは元のメニュー
    const mid = run(p, at(20));
    expect(user(mid, 'p_later').richMenuId).toBe('rm_reserved');
    expect(user(w, 'p_later').richMenuId).toBe('rm_main');
  });

  it('休眠掘り起こし 対象A：見学の3日後・2週間後・45日後。気になる点で出し分け、入園で止まる', () => {
    const a = (id: string) => user(w, id).deliveries.filter((d) => d.flowId === 'flow_dormant_a').map((d) => d.messageId);
    expect(a('p_later')).toEqual(['msg_a1', 'msg_a2_bus', 'msg_a3']);
    expect(a('p_now')).toEqual(['msg_a1', 'msg_a2_cost']);
    expect(user(w, 'p_now').flows['flow_dormant_a']!.status).toBe('stopped');
  });

  it('休眠掘り起こし 対象B：30日反応なしで近況、60日で検討状況、90日は「検討中」だけに再見学の案内', () => {
    const b = (id: string) => user(w, id).deliveries.filter((d) => d.flowId === 'flow_dormant_b').map((d) => [d.messageId, Math.floor(d.at / DAY)]);
    expect(b('p_info')).toEqual([['msg_b1', 31], ['msg_b2', 60], ['msg_b3', 90]]);
    // Day 20 にメニューを押した人は、Day 31 の時点では「反応あり」なので近況は送らない。他園に決定なら再見学の案内もしない
    expect(b('p_other')).toEqual([['msg_b2', 60]]);
    expect(w.decisions.find((d) => d.userId === 'p_skip' && d.ref === 'msg_b1')).toMatchObject({ outcome: 'excluded', reason: 'blocked' });
  });

  it('③：見学予約が GOAL。見学・入園はその後の成果地点。回答のルートは②のタグと一致する', () => {
    const rows = (id: string) => journeyOf(p, w, user(w, id));
    expect(rows('p_now').find((r) => r.isGoal)!.step.id).toBe('j_tour');
    expect(rows('p_now').at(-1)).toMatchObject({ status: 'current' });
    expect(rows('p_now').find((r) => r.step.id === 'j_branch')!.traversed).toEqual(['v_now']);
    expect(rows('p_skip').find((r) => r.step.id === 'j_branch')!.traversed).toEqual(['v_none']);
    expect(rows('p_info').find((r) => r.step.id === 'j_dormant')!.traversed).toEqual(['v_b']);
    expect(rows('p_later').find((r) => r.step.id === 'j_dormant')!.traversed).toEqual(['v_a']);
  });

  it('過去に戻って Q4 を「今すぐ」に変えると、重視点の案内ではなく予約の案内に計算し直される', () => {
    let d = newDemo(p, 'p_later');
    d = dispatch(d, { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_quick', questionId: 'q_quick', choiceIds: ['c_qk_edu'] });
    for (const [q, c] of [['q_age', 'c_age3'], ['q_enter', 'c_any'], ['q_focus', 'c_iq']]) d = dispatch(d, { type: 'answer', surveyId: 'sv_main', questionId: q!, choiceIds: [c!] });
    const before = d.now;
    d = dispatch(d, { type: 'answer', surveyId: 'sv_main', questionId: 'q_tour', choiceIds: ['c_later'] });
    expect(me(p, d).deliveries.map((x) => x.messageId)).toContain('msg_f_iq');
    d = seek(d, before);
    d = dispatch(d, { type: 'answer', surveyId: 'sv_main', questionId: 'q_tour', choiceIds: ['c_now'] });
    const now = me(p, d);
    expect(now.deliveries.map((x) => x.messageId)).toContain('msg_tour_now');
    expect(now.deliveries.map((x) => x.messageId)).not.toContain('msg_f_iq');
    expect('tour_later' in now.tags).toBe(false);
  });

  it('自動再生で全員を最後まで流すと CLI と同じ結果になり、営業用HTMLを書き出せる', async () => {
    const d = playToEnd(p);
    expect(d.now).toBe(95 * DAY);
    const ww = worldOf(p, d);
    for (const u of Object.values(ww.users)) expect(u.deliveries).toEqual(user(w, d.session.users.find((x) => x.id === u.id)!.personaId).deliveries);
    await exportChecks(p);
  }, 30_000);
});
