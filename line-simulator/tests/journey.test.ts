// 3画面の③「CVまでの道筋」（packages/player/src/journey.ts）と、検証案件の「エリア未回答」の導線。

import { describe, expect, it } from 'vitest';
import type { Project, UserState, WorldState } from '../packages/engine/src/index.ts';
import { reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { allScripted, dispatch, newDemo, seek, worldOf, type Demo } from '../packages/player/src/controller.ts';
import { explainDecision, Labels } from '../packages/player/src/describe.ts';
import { journeyOf, journeyPane, journeySteps, nextAction } from '../packages/player/src/journey.ts';
import type { Ui, View } from '../packages/player/src/view.ts';
import { loadProject } from '../tools/project-io.ts';
import { buildExport, sanitize } from '../tools/export-lib.ts';
import { DAY, sample } from './helpers.ts';

const project = sample();
const L = new Labels(project);
const at = (until: number, p: Project = project) => reduce(p, sessionFromPersonas(p), { until });
const w30 = at(30 * DAY);
const u = (w: WorldState, id: string) => w.users[id]!;
const got = (w: WorldState, id: string) => u(w, id).deliveries.map((d) => d.messageId);
/** 段ごとの「状態(ルート)」 */
const path = (p: Project, w: WorldState, user: UserState) => journeyOf(p, w, user).map((r) => `${r.step.id}:${r.status}${r.variantId ? `(${r.variantId})` : ''}`);

const ui: Ui = { menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set(), onlyNg: false, playing: false, speed: 1, status: '' };
const viewOf = (p: Project, demo: Demo): View => ({ project: p, world: worldOf(p, demo), demo, ui, L: new Labels(p), assetUrl: (a) => a.path });

describe('③ CVまでの道筋：どこまで来たか・いまどこか・この人のルート', () => {
  it('東京希望さん：全段を通って応募（ゴール）に到達。ルートは「離脱ポップアップ」「東京」', () => {
    expect(path(project, w30, u(w30, 'p_tokyo'))).toEqual([
      'j_visit:done',
      'j_friend:done(v_popup)',
      'j_survey:done(v_answered)',
      'j_jobs:done(v_tokyo)',
      'j_push:done',
      'j_apply:current',
    ]);
  });

  it('応募済みさん：先に応募したので、求人紹介と後押しは「通らず」でゴールにいる', () => {
    expect(path(project, w30, u(w30, 'p_applied'))).toEqual([
      'j_visit:done',
      'j_friend:done(v_thanks)',
      'j_survey:done(v_answered)',
      'j_jobs:skipped(v_tokyo)',
      'j_push:skipped',
      'j_apply:current',
    ]);
  });

  it('未応募さん：エリアを選ばないまま、求人紹介の段で「未選択」のルートにいる', () => {
    const day1 = at(DAY + 11 * 60);
    expect(path(project, day1, u(day1, 'p_notyet')).slice(2, 4)).toEqual(['j_survey:current(v_none)', 'j_jobs:todo(v_area_none)']);
    expect(path(project, w30, u(w30, 'p_notyet')).slice(3, 4)).toEqual(['j_jobs:current(v_area_none)']);
  });

  it('途中で回答さん：回答した時点で、ルートが「未選択」から「大阪」に変わる', () => {
    const before = at(DAY + 19 * 60);
    const after = at(DAY + 21 * 60);
    expect(journeyOf(project, before, u(before, 'p_late'))[3]!.variantId).toBe('v_area_none');
    expect(journeyOf(project, after, u(after, 'p_late'))[3]!.variantId).toBe('v_osaka');
  });

  it('次に起きること：この人の次の配信予定を、シナリオ名と文面で出す。ブロック中はそう言う', () => {
    const demo = allScripted(project);
    const day1 = at(DAY + 11 * 60);
    expect(nextAction(project, day1, { ...demo, now: DAY + 11 * 60 }, u(day1, 'p_notyet'), L)).toEqual({
      when: 'Day 1 12:00',
      what: 'エリア未選択のリマインド・再案内：「〇〇、希望エリアがまだ選ばれていません。」を配信',
    });
    // アンケートの時間切れ（シナリオが終わるだけ）は飛ばし、次の配信（Day 3 の再案内）を出す
    const later = at(DAY + 12 * 60 + 30);
    expect(nextAction(project, later, { ...demo, now: DAY + 12 * 60 + 30 }, u(later, 'p_notyet'), L).when).toBe('Day 3 12:00');
    expect(nextAction(project, w30, { ...demo, now: 30 * DAY }, u(w30, 'p_notyet'), L).what).toBe('ブロック中のため、この先の配信は届きません');
  });

  it('画面の③：「いまここ」「この人」のルート・ゴールを表示し、手で操作すると更新される', () => {
    let demo = newDemo(project, 'p_notyet');
    demo = dispatch(demo, { type: 'addFriend', routeId: 'route_exit_popup' });
    let html = journeyPane(viewOf(project, demo), worldOf(project, demo).users[demo.activeUserId]!);
    expect(html).toContain('ゴール：<b>求人応募（CV）</b>');
    expect(html).toMatch(/lsp-st current">[\s\S]*希望エリアのアンケート/);
    expect(html).toMatch(/<li class="mine"><span class="lsp-v-l">まだ選んでいない<b>この人<\/b>/);
    // アンケートで東京を選ぶと、ルートが「回答した」「東京」に変わる
    demo = dispatch(demo, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] });
    html = journeyPane(viewOf(project, demo), worldOf(project, demo).users[demo.activeUserId]!);
    expect(html).toMatch(/<li class="mine"><span class="lsp-v-l">回答した<b>この人<\/b>/);
    expect(html).toMatch(/<li class="mine"><span class="lsp-v-l">東京<b>この人<\/b>/);
    expect(html).toContain('アンケートの回答待ち');
    // 2問目にも答えると、次は東京の求人配信
    demo = dispatch(demo, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_job', choiceIds: ['c_office'] });
    html = journeyPane(viewOf(project, demo), worldOf(project, demo).users[demo.activeUserId]!);
    expect(html).toContain('<span class="lsp-jn-t">Day 1 09:00</span><span class="lsp-jn-w">東京希望のステップ配信：「【東京エリア】今週の新着求人です。」を配信</span>');
  });

  it('道筋が無い案件（既存データ）は、シナリオの一覧とCVから簡易表示を作る', () => {
    const p = sample();
    delete p.journey;
    const { steps, auto } = journeySteps(p);
    expect(auto).toBe(true);
    expect(steps.map((s) => s.label)).toEqual([...p.flows.map((f) => f.label), '応募']);
    const rows = journeyOf(p, w30, u(w30, 'p_tokyo'));
    expect(rows.at(-1)).toMatchObject({ status: 'current', isGoal: true });
  });

  it('別の業種（美容サロン）でも同じ部品で動く：来店回数でルートが分かれ、予約がゴール', () => {
    const beauty = loadProject('sample-beauty');
    expect(validate(beauty)).toEqual([]);
    const w = at(30 * DAY, beauty);
    const first = Object.values(w.users).find((x) => x.personaId === 'p_first_cut')!;
    const repeat = Object.values(w.users).find((x) => x.personaId === 'p_repeat_color')!;
    expect(path(beauty, w, first)).toEqual(['j_friend:done(v_popup)', 'j_survey:done', 'j_follow:done(v_first)', 'j_book:current', 'j_return:todo']);
    expect(path(beauty, w, repeat)).toEqual(['j_friend:done(v_default)', 'j_survey:done', 'j_follow:done(v_repeat)', 'j_book:skipped', 'j_return:current']);
  });

  it('検証：道筋の参照切れは重大エラー、「来たとみなす条件」が無い段は警告', () => {
    const p = sample();
    p.journey!.push({ id: 'j_x', label: 'x', nodeIds: ['n_nowhere'], cvId: 'cv_none' }, { id: 'j_y', label: 'y' }, { id: 'j_x', label: 'dup', cvId: 'cv_apply' });
    const issues = validate(p);
    expect(issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ level: 'error', code: 'ref.missing', params: { kind: 'node', id: 'n_nowhere' } }),
        expect.objectContaining({ level: 'error', code: 'ref.missing', params: { kind: 'cv', id: 'cv_none' } }),
        expect.objectContaining({ level: 'error', code: 'id.duplicate', path: '/journey/8' }),
        expect.objectContaining({ level: 'warning', code: 'optional.empty', params: { reason: 'journeyReach' } }),
      ]),
    );
  });

  it('営業用HTMLにも道筋が入り、同じ部品で表示される', async () => {
    const r = await buildExport(project, { readAsset: () => null, confirmFictional: true });
    expect(r.ok).toBe(true);
    expect(sanitize(project).journey).toHaveLength(6);
    expect(r.html).toContain('"j_apply"');
    expect(r.html).toContain('CVまでの道筋');
  }, 30_000);
});

describe('エリア未回答の導線（検証案件）', () => {
  it('友だち追加の時点で「未選択」タグが付き、東京・大阪を選んだら外れる', () => {
    const day0 = at(13 * 60 + 4);
    expect('area_none' in u(day0, 'p_tokyo').tags).toBe(true);
    expect('area_none' in u(w30, 'p_tokyo').tags).toBe(false);
    expect('area_none' in u(w30, 'p_notyet').tags).toBe(true);
  });

  it('未回答のままの人には Day 1 にリマインド、Day 3 に再案内が届き、求人の配信は届かない', () => {
    const d = u(w30, 'p_notyet').deliveries.filter((x) => x.flowId === 'flow_step_common');
    expect(d.map((x) => [x.messageId, x.at])).toEqual([
      ['msg_c_remind', DAY + 12 * 60],
      ['msg_c_ask', 3 * DAY + 12 * 60],
    ]);
    expect(got(w30, 'p_notyet').some((m) => m.startsWith('msg_t_') || m.startsWith('msg_o_'))).toBe(false);
  });

  it('ブロック中の人には、その後の案内は届かない（ブロック中として除外）', () => {
    const excluded = w30.decisions.find((x) => x.userId === 'p_notyet' && x.ref === 'msg_c_general');
    expect(excluded).toMatchObject({ outcome: 'excluded', reason: 'blocked' });
    expect(u(w30, 'p_notyet').messageCount).toBe(3);
  });

  it('回答した人・応募済みの人には、リマインドが1通も届かない（その場で止まる）', () => {
    for (const id of ['p_tokyo', 'p_osaka', 'p_applied']) {
      expect(got(w30, id).some((m) => m.startsWith('msg_c_'))).toBe(false);
      expect(u(w30, id).flows['flow_step_common']!.status).toBe('stopped');
    }
    const stop = w30.decisions.find((x) => x.userId === 'p_tokyo' && x.kind === 'stopWhen' && x.flowId === 'flow_step_common')!;
    expect(explainDecision(stop, L, project, w30)).toEqual({ tone: 'stop', title: '「エリア未選択のリマインド・再案内」を止めました', reason: 'タグ「希望エリア：未選択」が外れたため' });
    const applied = w30.decisions.find((x) => x.userId === 'p_applied' && x.kind === 'stopWhen' && x.flowId === 'flow_step_common')!;
    expect(explainDecision(applied, L, project, w30).reason).toBe('「応募済み」に当てはまったため');
  });

  it('途中で回答した人：リマインドが止まり、回答した日を起点に大阪の求人が始まる（過去の分をまとめて送らない）', () => {
    const late = u(w30, 'p_late');
    expect(late.deliveries.map((x) => [x.messageId, x.at])).toEqual([
      ['msg_welcome', 663],
      ['msg_c_remind', DAY + 12 * 60],
      ['msg_o_d1', 2 * DAY + 9 * 60],
      ['msg_o_d3', 4 * DAY + 9 * 60],
      ['msg_o_d7', 8 * DAY + 12 * 60],
    ]);
    expect(late.flows['flow_step_common']!.status).toBe('stopped');
  });

  it('友だち追加直後に答えた人の求人配信は、これまでと同じ日時（Day 1 09:00 から）', () => {
    expect(u(w30, 'p_tokyo').deliveries.find((x) => x.messageId === 'msg_t_d1')!.at).toBe(DAY + 9 * 60);
    expect(u(w30, 'p_osaka').deliveries.find((x) => x.messageId === 'msg_o_d7')!.at).toBe(7 * DAY + 12 * 60);
  });

  it('アンケート以外（「求人」と送ったときの返信のボタン）から選んでも、同じように地域の配信へ移る', () => {
    let demo = newDemo(project, 'p_notyet');
    demo = dispatch(demo, { type: 'addFriend', routeId: 'route_exit_popup' });
    demo = seek(demo, 2 * DAY);
    demo = dispatch(demo, { type: 'sendText', text: '求人' });
    expect(worldOf(project, demo).users[demo.activeUserId]!.deliveries.at(-1)!.messageId).toBe('msg_r_ask_area');
    demo = dispatch(demo, { type: 'tapButton', buttonId: 'btn_r_osaka' });
    demo = seek(demo, 10 * DAY);
    const me = worldOf(project, demo).users[demo.activeUserId]!;
    expect(me.flows['flow_step_osaka']!.status).toBe('ended');
    expect(me.flows['flow_step_common']!.status).toBe('stopped');
    expect(me.deliveries.filter((x) => x.messageId === 'msg_c_ask')).toEqual([]);
  });

  it('同じエリアを2回選んでも、求人の配信は1回だけ（重複しない）', () => {
    let demo = newDemo(project, 'p_notyet');
    demo = dispatch(demo, { type: 'addFriend', routeId: 'route_exit_popup' });
    demo = dispatch(demo, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] });
    demo = seek(demo, DAY + 10 * 60);
    demo = dispatch(demo, { type: 'tapButton', buttonId: 'btn_r_tokyo' });
    demo = seek(demo, 10 * DAY);
    const me = worldOf(project, demo).users[demo.activeUserId]!;
    expect(me.deliveries.filter((x) => x.messageId === 'msg_t_d1')).toHaveLength(1);
    expect(me.flows['flow_step_tokyo']!.runs).toBe(1);
  });
});
