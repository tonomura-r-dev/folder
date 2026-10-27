// 3画面の③「CVまでの道筋」（packages/player/src/journey.ts）のフローチャートと、検証案件の「エリア未回答」の導線。

import { describe, expect, it } from 'vitest';
import type { Project, ScenarioPatch, UserState, WorldState } from '../packages/engine/src/index.ts';
import { applyPatch, approve, reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { allScripted, dispatch, newDemo, seek, worldOf, type Demo } from '../packages/player/src/controller.ts';
import { explainDecision, Labels } from '../packages/player/src/describe.ts';
import { contentsOf, fcDialog, journeyOf, journeyPane, journeySteps, nextAction } from '../packages/player/src/journey.ts';
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
/** 段ごとの「状態(この人のルート)[通った行き先]」 */
const path = (p: Project, w: WorldState, user: UserState) =>
  journeyOf(p, w, user).map((r) => `${r.step.id}:${r.status}${r.variantId ? `(${r.variantId})` : ''}${r.traversed.length ? `[${r.traversed.join(',')}]` : ''}`);

const ui = (over: Partial<Ui> = {}): Ui => ({ menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, ...over });
const viewOf = (p: Project, demo: Demo, over: Partial<Ui> = {}): View => ({ project: p, world: worldOf(p, demo), demo, ui: ui(over), L: new Labels(p), assetUrl: (a) => a.path });
const me = (p: Project, demo: Demo) => worldOf(p, demo).users[demo.activeUserId]!;
/** 「現在ここ」が付いている箱の名前 */
const hereOf = (html: string) => [...html.matchAll(/<(?:button)[^>]*class="lsp-fc-(?:node|dia)[^"]*"[^>]*>([\s\S]*?)<\/button>/g)].filter((m) => m[1]!.includes('現在ここ')).map((m) => /<span(?: class="lsp-fc-title")?>([^<]+)<\/span>/.exec(m[1]!)?.[1]);

describe('③ CVまでの道筋：段の状態・この人のルート・通った経路', () => {
  it('東京希望さん：START から全段を通り、GOAL（応募）にいる。分岐は「東京」を通った', () => {
    expect(path(project, w30, u(w30, 'p_tokyo'))).toEqual([
      'j_friend:done',
      'j_survey:done',
      'j_area:done(v_tokyo)[v_tokyo]',
      'j_detail:done',
      'j_push:done',
      'j_apply:current',
    ]);
    const rows = journeyOf(project, w30, u(w30, 'p_tokyo'));
    expect([rows[0]!.isStart, rows.at(-1)!.isGoal]).toEqual([true, true]);
  });

  it('応募済みさん：先に応募したので、地域別・求人詳細・応募促進は「通らず」でゴールにいる', () => {
    expect(path(project, w30, u(w30, 'p_applied'))).toEqual(['j_friend:done', 'j_survey:done', 'j_area:skipped(v_tokyo)', 'j_detail:skipped', 'j_push:skipped', 'j_apply:current']);
  });

  it('未応募さん：アンケートの次は「未回答」の行き先（地域選択リマインド）を通る', () => {
    const day1 = at(DAY + 11 * 60);
    expect(path(project, day1, u(day1, 'p_notyet')).slice(1, 3)).toEqual(['j_survey:current', 'j_area:todo(v_none)']);
    expect(path(project, w30, u(w30, 'p_notyet'))[2]).toBe('j_area:current(v_none)[v_none]');
  });

  it('途中で回答さん：「未回答」を通ったあと、回答した時点でルートが「大阪」に変わり、両方を通った経路が残る', () => {
    const before = at(DAY + 19 * 60);
    const after = at(DAY + 21 * 60);
    expect(path(project, before, u(before, 'p_late'))[2]).toBe('j_area:current(v_none)[v_none]');
    expect(path(project, after, u(after, 'p_late'))[2]).toBe('j_area:current(v_osaka)[v_none]');
    expect(path(project, w30, u(w30, 'p_late'))[2]).toBe('j_area:done(v_osaka)[v_none,v_osaka]');
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

  it('道筋が無い案件（既存データ）は、シナリオの一覧とCVから簡易表示を作る', () => {
    const p = sample();
    delete p.journey;
    const { steps, auto } = journeySteps(p);
    expect(auto).toBe(true);
    expect(steps.map((s) => s.label)).toEqual([...p.flows.map((f) => f.label), '応募']);
    const rows = journeyOf(p, w30, u(w30, 'p_tokyo'));
    expect(rows.at(-1)).toMatchObject({ status: 'current', isGoal: true });
    expect(journeyPane(viewOf(p, allScripted(p)), u(w30, 'p_tokyo'))).toContain('自動で並べた簡易表示');
  });

  it('別の業種（美容サロン）でも同じ部品で動く：来店回数のひし形で分かれ、予約が GOAL', () => {
    const beauty = loadProject('sample-beauty');
    expect(validate(beauty)).toEqual([]);
    const w = at(30 * DAY, beauty);
    const first = Object.values(w.users).find((x) => x.personaId === 'p_first_cut')!;
    const repeat = Object.values(w.users).find((x) => x.personaId === 'p_repeat_color')!;
    expect(path(beauty, w, first)).toEqual(['j_friend:done', 'j_survey:done', 'j_follow:done(v_first)[v_first]', 'j_book:current', 'j_return:todo']);
    expect(path(beauty, w, repeat)).toEqual(['j_friend:done', 'j_survey:done', 'j_follow:current(v_repeat)[v_repeat]', 'j_book:todo', 'j_return:todo']);
    const html = journeyPane(viewOf(beauty, allScripted(beauty)), first);
    expect(html).toContain('来店は2回目以上？');
    expect(html).toMatch(/lsp-fc-node goal[^>]*>\s*<span class="lsp-fc-tag">GOAL<\/span>\s*<span class="lsp-fc-title">予約（CV）/);
  });
});

describe('③ フローチャートの表示と操作', () => {
  it('START・GOAL は色付き、分岐はひし形と行き先の矢印ラベル、採用施策と追加候補（点線）は別の見た目', () => {
    const demo = allScripted(project);
    const html = journeyPane(viewOf(project, demo), u(w30, 'p_tokyo'));
    expect(html).toMatch(/class="lsp-fc-node start passed"[\s\S]*?START[\s\S]*?LINE友だち追加/);
    expect(html).toMatch(/class="lsp-fc-node goal passed"[\s\S]*?GOAL[\s\S]*?求人応募（CV）/);
    expect(html).toContain('<span>地域選択済み？</span>');
    expect([...html.matchAll(/<span class="lsp-fc-elabel">([^<]+)<\/span>/g)].map((m) => m[1])).toEqual(['東京', '大阪', '未回答']);
    // 段の追加候補は点線の箱で並べ、分岐の行き先（幅が狭い）は件数の点線ボタン → 押すと一覧
    expect([...html.matchAll(/class="lsp-fc-opt"[^>]*>([^<]+)</g)].map((m) => m[1])).toEqual(['応募前の不安を解消するQ&amp;A', '面接日程をLINEで予約']);
    expect(html).toContain('class="lsp-fc-optcount" data-act="fcOpen" data-step="j_area" data-variant="v_none">＋追加候補 3</button>');
    const list = fcDialog(viewOf(project, demo, { fcOpen: { step: 'j_area', variant: 'v_none' } }), u(w30, 'p_tokyo'));
    expect([...list.matchAll(/class="lsp-fc-opt"[^>]*>([^<]+)</g)].map((m) => m[1])).toEqual(['希望条件の再回答促進', '求人の魅力を伝える配信', 'リッチメニューからの条件登録誘導']);
    // 追加候補を隠す設定
    const hidden = journeyPane(viewOf(project, demo, { fcOptions: false }), u(w30, 'p_tokyo'));
    expect(hidden).not.toContain('lsp-fc-opt"');
    expect(hidden).not.toContain('lsp-fc-optcount');
  });

  it('箱には施策名と文面の1〜3行のさわり（シナリオの文面から）を出す', () => {
    const html = journeyPane(viewOf(project, allScripted(project)), u(w30, 'p_tokyo'));
    expect(html).toContain('<span class="lsp-fc-title">地域選択リマインド</span>');
    expect(html).toContain('<span class="lsp-fc-prev">〇〇、希望エリアがまだ選ばれていません。');
    expect(html).toContain('質問：希望の勤務エリアを教えてください／希望の職種を教えてください');
  });

  it('①で操作すると「現在ここ」が動き、②と同じタグで分岐の行き先が決まる', () => {
    let demo = newDemo(project, 'p_notyet');
    expect(hereOf(journeyPane(viewOf(project, demo), me(project, demo)))).toEqual([]);
    demo = dispatch(demo, { type: 'addFriend', routeId: 'route_exit_popup' });
    // 友だち追加と同時にアンケートが出るので、現在地はアンケート。START は通過済み
    let html = journeyPane(viewOf(project, demo), me(project, demo));
    expect(hereOf(html)).toEqual(['希望地域アンケート']);
    expect(html).toMatch(/lsp-fc-node start passed/);
    // ②のタグは「未選択」→ ③は「未回答」の行き先がこの人のルート
    expect(Object.keys(me(project, demo).tags)).toEqual(['area_none']);
    expect(journeyPane(viewOf(project, demo), me(project, demo))).toMatch(/lsp-fc-node lane"[\s\S]*?地域選択リマインド/);
    demo = seek(demo, DAY + 12 * 60 + 5);
    expect(hereOf(journeyPane(viewOf(project, demo), me(project, demo)))).toEqual(['地域選択リマインド']);
    // リマインドのボタンで東京を選ぶ → ②のタグが東京に、③の「現在ここ」が東京向け求人配信に
    demo = dispatch(demo, { type: 'tapButton', buttonId: 'btn_c_remind_tokyo' });
    expect(Object.keys(me(project, demo).tags).sort()).toEqual(['area_selected', 'area_tokyo']);
    html = journeyPane(viewOf(project, demo), me(project, demo));
    expect(hereOf(html)).toEqual(['東京向け求人配信']);
    expect(html).toMatch(/lsp-fc-node passed"[\s\S]*?地域選択リマインド/);
  });

  it('箱を押すと、文面の全文・タイミング・配信対象・止める条件・目的・この人への配信状況が出る', () => {
    const demo = { ...allScripted(project), activeUserId: 'p_notyet' };
    const v = viewOf(project, demo, { fcOpen: { step: 'j_area', variant: 'v_none' } });
    const html = fcDialog({ ...v, world: w30 }, u(w30, 'p_notyet'));
    expect(html).toContain('採用施策');
    expect(html).toContain('希望エリアがまだ選ばれていません。\nエリアを選ぶと、そのエリアの新着求人だけをお届けします。');
    expect(html).toContain('<dt>タイミング</dt><dd>友だち追加の1日後 12:00</dd>');
    expect(html).toContain('シナリオ「エリア未選択のリマインド・再案内」（始まり：タグ「希望エリア：未選択」が付いた）');
    expect(html).toContain('<dt>止める条件</dt>');
    expect(html).toContain('この人には Day 1 12:00 に届きました');
    expect(html).toContain('この人には届いていません：この人がブロックしているため');
    // ボタンは押せない見本として、押したときの動作つきで出す
    expect(html).toContain('<span class="lsp-btn-ro">東京<small>→ アンケートに「東京」と回答</small></span>');
    expect(html).toContain('◎ Lステップ単体で可');
  });

  it('自動応答の文面も参照できる（「求人」と送ったとき・リッチメニューからも）', () => {
    const c = contentsOf(project, L, [], ['msg_r_tokyo']);
    expect(c).toEqual([{ messageId: 'msg_r_tokyo', timing: '「求人」と送られたとき（リッチメニュー「求人を探す」からも）', target: '自動応答：タグ「希望エリア：東京」が付いているの人に（通数に数えない）' }]);
  });

  it('文面は③に二重登録しない：シナリオの文面を直すと、①のトークと③の詳細の両方が変わる', () => {
    const p = sample();
    p.messages.find((m) => m.id === 'msg_c_remind')!.bubbles[0] = { kind: 'text', text: '【新しい文面】エリアを選んでください' };
    let demo = newDemo(p, 'p_notyet');
    demo = dispatch(demo, { type: 'addFriend', routeId: 'route_exit_popup' });
    demo = seek(demo, DAY + 13 * 60);
    expect(JSON.stringify(me(p, demo).talk)).toContain('【新しい文面】');
    expect(fcDialog(viewOf(p, demo, { fcOpen: { step: 'j_area', variant: 'v_none' } }), me(p, demo))).toContain('【新しい文面】エリアを選んでください');
    expect(JSON.stringify(p.journey)).not.toContain('【新しい文面】');
  });

  it('追加候補は「未採用」と分かる詳細を出すだけで、シナリオにも配信にも影響しない', () => {
    const demo = allScripted(project);
    const html = fcDialog(viewOf(project, demo, { fcOpen: { step: 'j_area', variant: 'v_none', option: 'o_reanswer' } }), u(w30, 'p_notyet'));
    expect(html).toContain('追加候補（未採用）');
    expect(html).toContain('シナリオには入っておらず、配信もされません');
    expect(html).toContain('【1分で完了】');
    const noOptions = sample();
    for (const s of noOptions.journey!) {
      delete s.options;
      for (const x of s.variants ?? []) delete x.options;
    }
    expect(JSON.stringify(at(30 * DAY, noOptions))).toBe(JSON.stringify(w30));
  });

  it('ひし形を押すと、行き先ごとの条件と、この人がどれに当てはまるかを②と同じ判定で出す', () => {
    const demo = allScripted(project);
    const html = fcDialog(viewOf(project, demo, { fcOpen: { step: 'j_area' } }), u(w30, 'p_osaka'));
    expect(html).toContain('<li class="no"><b>東京 → 東京向け求人配信</b>');
    expect(html).toContain('✗ タグ「希望エリア：東京」が付いていない');
    expect(html).toContain('<li class="yes"><b>大阪 → 大阪向け求人配信</b>');
    // ②の「なぜ」（希望エリアの分岐で東京に進まなかった理由）と同じ言葉
    const skipped = w30.decisions.find((d) => d.userId === 'p_osaka' && d.nodeId === 'n_w_area' && d.ref === 'tokyo')!;
    expect(explainDecision(skipped, L, project, w30).reason).toBe('タグ「希望エリア：東京」が付いていないため');
  });

  it('Lステップでの実現性：単体で可・条件あり・外部実装が必要・未確認を区別し、未確認は「可」と出さない', () => {
    const html = journeyPane(viewOf(project, allScripted(project)), u(w30, 'p_tokyo'));
    expect(html).toContain('△ 外部ツール・Webサイト側の実装が必要');
    expect(html).toContain('◎ Lステップ単体で可');
    expect(html).toContain('<span>？ 実現可否は未確認</span>');
    const dlg = fcDialog(viewOf(project, allScripted(project), { fcOpen: { step: 'j_area', variant: 'v_none', option: 'o_menu' } }), u(w30, 'p_tokyo'));
    expect(dlg).toContain('？ 実現可否は未確認');
    expect(dlg).toContain('要確認');
    const p = sample();
    (p.journey![0] as { feasibility: string }).feasibility = 'maybe';
    expect(validate(p)).toContainEqual(expect.objectContaining({ level: 'error', code: 'journey.feasibility', path: '/journey/0/feasibility' }));
  });

  it('AI提案で足した施策（追加候補）も、承認するまでは書き出せず、承認後は同じ仕組みで表示される', () => {
    const patch: ScenarioPatch = {
      schemaVersion: 1,
      summary: 'test',
      ops: [{ op: 'add', path: '/journey/4/options/-', value: { id: 'o_ai', label: 'AIの候補：先輩の声', summary: 'x', feasibility: 'lstep' }, prov: { state: 'inferred', rationale: 'スクショから推定' } }],
    };
    const { project: p, failed } = applyPatch(project, patch);
    expect(failed).toEqual([]);
    expect(validate(p)).toContainEqual(expect.objectContaining({ code: 'prov.unapproved', path: '/journey/4/options/2' }));
    const ok = approve(p, '/journey/4/options/2');
    expect(validate(ok)).toEqual([]);
    expect(journeyPane(viewOf(ok, allScripted(ok)), u(w30, 'p_tokyo'))).toContain('AIの候補：先輩の声');
  });

  it('検証：道筋の参照切れは重大エラー、「来たとみなす条件」が無い段は警告', () => {
    const p = sample();
    p.journey!.push(
      { id: 'j_x', label: 'x', nodeIds: ['n_nowhere'], cvId: 'cv_none', messageIds: ['msg_none'] },
      { id: 'j_y', label: 'y' },
      { id: 'j_x', label: 'dup', cvId: 'cv_apply', variants: [{ id: 'v', label: 'v', when: { op: 'always' }, nodeIds: ['n_gone'] }] },
    );
    const issues = validate(p);
    expect(issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ level: 'error', code: 'ref.missing', params: { kind: 'node', id: 'n_nowhere' } }),
        expect.objectContaining({ level: 'error', code: 'ref.missing', params: { kind: 'cv', id: 'cv_none' } }),
        expect.objectContaining({ level: 'error', code: 'ref.missing', params: { kind: 'message', id: 'msg_none' } }),
        expect.objectContaining({ level: 'error', code: 'ref.missing', params: { kind: 'node', id: 'n_gone' } }),
        expect.objectContaining({ level: 'error', code: 'id.duplicate', path: '/journey/8' }),
        expect.objectContaining({ level: 'warning', code: 'optional.empty', params: { reason: 'journeyReach' } }),
      ]),
    );
  });

  it('営業用HTMLにも道筋が入り、同じ部品（フローチャート・詳細）で表示される', async () => {
    const r = await buildExport(project, { readAsset: () => null, confirmFictional: true });
    expect(r.ok).toBe(true);
    expect(sanitize(project).journey).toHaveLength(6);
    expect(r.html).toContain('"j_apply"');
    expect(r.html).toContain('"o_reanswer"');
    expect(r.html).toContain('lsp-fc-dia');
    expect(r.html).toContain('追加候補（未採用）');
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
