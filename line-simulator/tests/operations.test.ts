// 運用寄りの機能：全員共通の日時の待ち・アンケートの聞き直し・絞り込みつき一覧・運営側の登録（段階）・LINE誘導の入力欄への引き継ぎ。

import { describe, expect, it } from 'vitest';
import type { Project } from '../packages/engine/src/index.ts';
import { reduce, validate } from '../packages/engine/src/index.ts';
import { dispatch, newDemo, seek, worldOf } from '../packages/player/src/controller.ts';
import { backstage } from '../packages/player/src/backstage.ts';
import { Labels } from '../packages/player/src/describe.ts';
import { journeyOf } from '../packages/player/src/journey.ts';
import { listFilter, phone } from '../packages/player/src/phone.ts';
import { buildTimeline } from '../packages/player/src/timeline.ts';
import { baseProject, chain, codes, DAY, edge, node } from './helpers.ts';

const view = (p: Project, d: ReturnType<typeof newDemo>) => {
  const world = worldOf(p, d);
  return { project: p, world, demo: d, ui: { menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' } as never;
};
const session = (events: object[]) => ({ schemaVersion: 1 as const, projectId: 't', users: [{ id: 'u1', personaId: 'u1' }], events: events.map((e) => ({ userId: 'u1', ...e })) as never });

describe('全員共通の日時（wait: at）', () => {
  const project = () => {
    const f = chain('f_at', [node({ id: 'w', type: 'wait', wait: { kind: 'at', day: 2, time: '12:00' } }), node({ id: 'm', type: 'message', messageId: 'm1' })], { stopWhen: { op: 'friend', state: 'blocked' } });
    f.nodes.push(node({ id: 'late', type: 'end', reason: '配信日のあとで友だち追加' }));
    f.edges.push(edge('w', 'late', 'passed'));
    return baseProject({ flows: [f] });
  };

  it('友だち追加の時刻に関係なく、決めた日時に届く', () => {
    const w = reduce(project(), session([{ at: 600, type: 'addFriend', routeId: 'r1' }]), { until: 3 * DAY });
    expect(w.users['u1']!.deliveries.map((d) => d.at)).toEqual([2 * DAY + 720]);
  });

  it('日時を過ぎてから友だちになった人には送らず、「過ぎていた」の出口へ進む', () => {
    const w = reduce(project(), session([{ at: 2 * DAY + 800, type: 'addFriend', routeId: 'r1' }]), { until: 4 * DAY });
    expect(w.users['u1']!.deliveries).toEqual([]);
    expect(w.users['u1']!.trail.map((t) => t.nodeId)).toContain('late');
  });

  it('検証：日・時刻の形を確かめ、出口は out と passed', () => {
    const p = project();
    expect(codes(validate(p), 'error')).toEqual([]);
    const w = p.flows[0]!.nodes.find((n) => n.id === 'w')!;
    if (w.type === 'wait') w.wait = { kind: 'at', day: -1, time: '25:00' };
    expect(codes(validate(p), 'error')).toEqual(['node.requiredMissing', 'node.requiredMissing']);
  });
});

describe('アンケートの聞き直し（reask）', () => {
  const project = () =>
    baseProject({
      tags: [{ id: 't_n', label: '北' }, { id: 't_s', label: '南' }],
      surveys: [{ id: 'sv', label: '希望', questions: [{ id: 'q', text: 'エリアは？', mode: 'single', choices: [{ id: 'c_n', label: '北', addTags: ['t_n'], removeTags: ['t_s'] }, { id: 'c_s', label: '南', addTags: ['t_s'], removeTags: ['t_n'] }] }] }],
      richMenus: [{ id: 'rm', label: 'メニュー', layout: { cols: 1, rows: 1 }, cells: [{ id: 'cell', label: '希望を変える', area: { col: 0, row: 0 }, action: { kind: 'none' } }] }],
      flows: [
        chain('f_first', [node({ id: 's1', type: 'survey', surveyId: 'sv' })]),
        chain('f_change', [node({ id: 's2', type: 'survey', surveyId: 'sv', reask: true })], { trigger: { kind: 'menuTap', richMenuId: 'rm', cellId: 'cell' }, reentry: 'always' }),
      ],
      settings: { dedupeWindowMinutes: 0, defaultRichMenuId: 'rm' },
    });

  it('前の回答を消して1問目から聞き直し、選び直した答えでタグが入れ替わる', () => {
    const p = project();
    let d = dispatch(newDemo(p, 'u1'), { type: 'addFriend', routeId: 'r1' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv', questionId: 'q', choiceIds: ['c_n'] });
    expect(Object.keys(worldOf(p, d).users['u1']!.tags)).toEqual(['t_n']);
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm', cellId: 'cell' });
    const asked = worldOf(p, d).users['u1']!;
    expect(asked.surface).toEqual({ kind: 'survey', surveyId: 'sv', questionId: 'q' });
    expect(asked.answers['sv']).toBeUndefined();
    d = dispatch(d, { type: 'answer', surveyId: 'sv', questionId: 'q', choiceIds: ['c_s'] });
    const after = worldOf(p, d).users['u1']!;
    expect(Object.keys(after.tags)).toEqual(['t_s']);
    expect(after.flows['f_change']!.status).toBe('ended');
  });
});

describe('絞り込みつきの一覧（list）', () => {
  const project = () =>
    baseProject({
      tags: [{ id: 't_n', label: '北' }],
      webPages: [
        {
          id: 'pg_list', kind: 'lp', title: '一覧',
          blocks: [
            { kind: 'heading', text: '一覧' },
            {
              kind: 'list',
              facets: [
                { id: 'area', label: 'エリア', options: [{ id: 'n', label: '北', tagId: 't_n' }, { id: 's', label: '南' }] },
                { id: 'day', label: '曜日', options: [{ id: 'we', label: '土日' }, { id: 'wd', label: '平日' }] },
              ],
              items: [
                { id: 'i1', title: '北の土日', values: { area: ['n'], day: ['we'] }, action: { kind: 'openPage', pageId: 'pg_d1' } },
                { id: 'i2', title: '南の平日', values: { area: ['s'], day: ['wd'] }, action: { kind: 'openPage', pageId: 'pg_d1' } },
                { id: 'i3', title: '北の平日', values: { area: ['n'], day: ['wd'] }, action: { kind: 'openPage', pageId: 'pg_d1' } },
              ],
            },
          ],
        },
        { id: 'pg_d1', kind: 'lp', title: '詳細', blocks: [{ kind: 'heading', text: '詳細' }] },
      ],
    });

  it('検証：軸・選択肢・項目のIDと参照先、台本の絞り込みと項目を確かめる', () => {
    const p = project();
    expect(codes(validate(p), 'error')).toEqual([]);
    const bad = project();
    const b = bad.webPages[0]!.blocks[1]!;
    if (b.kind !== 'list') throw new Error('list');
    b.items[0]!.values = { area: ['x'] };
    b.items[1]!.action = { kind: 'openPage', pageId: 'nope' };
    b.facets[0]!.options[0]!.tagId = 'no_tag';
    bad.personas[0]!.script = [
      { atMinutes: 600, event: { type: 'filterList', pageId: 'pg_list', blockIndex: 1, facetId: 'area', optionId: 'zz' } },
      { atMinutes: 601, event: { type: 'tapPageButton', pageId: 'pg_list', blockIndex: 1, itemId: 'zz' } },
    ];
    expect(codes(validate(bad), 'error').filter((c) => c === 'ref.missing')).toHaveLength(5);
  });

  it('タグに合う選択肢で開き、絞り込むと件数が変わる。項目を押すと動作し、ページを開き直すと絞り込みは戻る', () => {
    const p = project();
    let d = newDemo(p, 'u1');
    d = dispatch(d, { type: 'visitPage', pageId: 'pg_list' });
    let w = worldOf(p, d);
    const list = p.webPages[0]!.blocks[1]!;
    if (list.kind !== 'list') throw new Error('list');
    // タグが無いので最初は全件
    expect(listFilter(w.users['u1']!, 'pg_list', 1, list)).toEqual({ area: '', day: '' });
    expect(phone(view(p, d), w.users['u1']!)).toContain('該当 <b>3</b>件');
    d = dispatch(d, { type: 'filterList', pageId: 'pg_list', blockIndex: 1, facetId: 'day', optionId: 'wd' });
    w = worldOf(p, d);
    const html = phone(view(p, d), w.users['u1']!);
    expect(html).toContain('該当 <b>2</b>件');
    expect(html).toContain('南の平日');
    expect(html).not.toContain('北の土日');
    // ③の時系列に「絞り込んだ」が残る
    const tl = buildTimeline(p, new Labels(p), worldOf(p, d), ['u1']).flatMap((e) => e.front);
    expect(tl).toContain('一覧を「曜日：平日」で絞り込む');
    d = dispatch(d, { type: 'tapPageButton', pageId: 'pg_list', blockIndex: 1, itemId: 'i3' });
    w = worldOf(p, d);
    expect(w.users['u1']!.surface).toEqual({ kind: 'page', pageId: 'pg_d1' });
    // 開き直すと、絞り込みは最初の状態（タグに合う選択肢）に戻る
    w.users['u1']!.tags['t_n'] = { at: 0 };
    d = dispatch(d, { type: 'visitPage', pageId: 'pg_list' });
    const again = worldOf(p, d).users['u1']!;
    again.tags['t_n'] = { at: 0 };
    expect(listFilter(again, 'pg_list', 1, list)).toEqual({ area: 'n', day: '' });
  });
});

describe('運営側の登録（段階）', () => {
  const project = () =>
    baseProject({
      brand: { companyName: '株式会社サンプル', accountName: 'a', colors: { primary: '#000', accent: '#0f0' } },
      cvs: [{ id: 'cv_apply', label: '応募' }, { id: 'st_fix', label: '勤務確定', kind: 'stage', repeatable: true, demoAction: '勤務確定を登録' }],
      flows: [chain('f_fix', [node({ id: 'm', type: 'message', messageId: 'm1' })], { trigger: { kind: 'cv', cvId: 'st_fix' }, reentry: 'always' })],
    });

  it('CVと分けて記録し、何度でも登録できる。③では「運営側の登録」として出る', () => {
    const p = project();
    let d = dispatch(newDemo(p, 'u1'), { type: 'addFriend', routeId: 'r1' });
    d = dispatch(d, { type: 'cv', cvId: 'cv_apply' });
    d = seek(d, 700);
    d = dispatch(d, { type: 'cv', cvId: 'st_fix' });
    d = seek(d, 800);
    d = dispatch(d, { type: 'cv', cvId: 'st_fix' });
    const v = view(p, d);
    const w = worldOf(p, d);
    expect(w.users['u1']!.cvs['st_fix']!.count).toBe(2);
    expect(w.users['u1']!.deliveries).toHaveLength(2);
    const html = backstage(v, w.users['u1']!);
    expect(html).toContain('<h4>CVの記録（1）</h4>');
    expect(html).toContain('<h4>サンプル側の登録（2）</h4>');
    // 何度でも登録できる段階は「済み」にならない
    expect(html).toContain('勤務確定を登録（3回目）');
    expect(html).not.toContain('勤務確定済み');
    const events = buildTimeline(p, new Labels(p), w, ['u1']);
    const stage = events.filter((e) => e.kind === 'stage');
    expect(stage).toHaveLength(2);
    expect(stage[0]!.title).toBe('サンプルが「勤務確定」を登録');
  });
});

describe('LINE誘導の入力欄への引き継ぎ（prefill）', () => {
  const project = () =>
    baseProject({
      webPages: [{ id: 'pg_thanks', kind: 'thanks', title: '完了', blocks: [{ kind: 'heading', text: '完了' }], lineCta: { label: 'LINEで連絡', routeId: 'r1', prefill: '【応募しました】サンプル' } }],
    });

  it('押した直後はトークの入力欄に文が入り、送るか別の操作をすると消える', () => {
    const p = project();
    let d = dispatch(newDemo(p, 'u1'), { type: 'visitPage', pageId: 'pg_thanks' });
    d = dispatch(d, { type: 'tapLineCta', pageId: 'pg_thanks' });
    let html = phone(view(p, d), worldOf(p, d).users['u1']!);
    expect(html).toContain('value="【応募しました】サンプル"');
    d = dispatch(d, { type: 'sendText', text: '【応募しました】サンプル' });
    html = phone(view(p, d), worldOf(p, d).users['u1']!);
    expect(html).not.toContain('value="【応募しました】サンプル"');
  });
});

describe('予約フォーム：選んだ値の差し込み・選択肢の締切・確認画面', () => {
  const project = () =>
    baseProject({
      surveys: [{ id: 'sv', label: '診断', questions: [{ id: 'q', text: '学年は？', mode: 'single', choices: [{ id: 'c1', label: '中1' }] }] }],
      messages: [{ id: 'm1', bubbles: [{ kind: 'text', text: '日時：{{form.f_slot}}／学年：{{answer.sv.q|未回答}}' }] }, { id: 'm2', bubbles: [{ kind: 'text', text: 'm2' }] }],
      webPages: [
        {
          id: 'pg_rsv', kind: 'form', title: '予約', blocks: [{ kind: 'text', text: '診断の学年：{{answer.sv.q|未回答}}' }],
          form: { fields: [{ id: 'f_slot', label: '日時', type: 'select', options: ['1日目', '3日目'], closesAt: [DAY, null] }], onSubmit: { cvId: 'cv1', nextPageId: 'pg_done' }, confirm: { title: '予約内容の確認', submitLabel: 'この内容で予約する' } },
        },
        { id: 'pg_done', kind: 'thanks', title: '完了', blocks: [{ kind: 'heading', text: '完了' }, { kind: 'button', label: 'LINEに戻る', action: { kind: 'backToLine' } }] },
      ],
      flows: [chain('f_ok', [node({ id: 'm', type: 'message', messageId: 'm1' })], { trigger: { kind: 'cv', cvId: 'cv1' }, reentry: 'always' })],
    });

  it('{{form.項目}} はいちばん最近に送信したフォームの値、{{キー|既定}} は値が無いときの文字になる', () => {
    const p = project();
    expect(validate(p).filter((i) => i.code === 'text.unknownVar')).toEqual([]);
    let d = dispatch(newDemo(p, 'u1'), { type: 'addFriend', routeId: 'r1' });
    d = dispatch(d, { type: 'visitPage', pageId: 'pg_rsv' });
    expect(phone(view(p, d), worldOf(p, d).users['u1']!)).toContain('診断の学年：未回答');
    d = dispatch(d, { type: 'submitForm', pageId: 'pg_rsv', values: { f_slot: '3日目' } });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'pg_done', blockIndex: 1 });
    const u = worldOf(p, d).users['u1']!;
    expect(u.surface).toEqual({ kind: 'talk' });
    const last = u.talk.at(-1)!;
    expect(last.side === 'bot' ? last.bubble : null).toEqual({ kind: 'text', text: '日時：3日目／学年：未回答' });
    // 知らない項目は検証で知らせる
    const bad = project();
    bad.messages[1]!.bubbles = [{ kind: 'text', text: '{{form.nope}}' }];
    expect(validate(bad).filter((i) => i.code === 'text.unknownVar')).toHaveLength(1);
  });

  it('締切を過ぎた選択肢は出さない。締切の数が選択肢と合わなければ検証で知らせる', () => {
    const p = project();
    let d = dispatch(newDemo(p, 'u1'), { type: 'visitPage', pageId: 'pg_rsv' });
    expect(phone(view(p, d), worldOf(p, d).users['u1']!)).toContain('<option>1日目</option>');
    d = seek(d, DAY + 10);
    const html = phone(view(p, d), worldOf(p, d).users['u1']!);
    expect(html).not.toContain('1日目');
    expect(html).toContain('<option>3日目</option>');
    const bad = project();
    bad.webPages[0]!.form!.fields[0]!.closesAt = [DAY];
    expect(codes(validate(bad), 'error')).toEqual(['node.requiredMissing']);
  });

  it('確認画面：選んだ値を並べ、「この内容で予約する」と「修正する」。送信ボタンの文字も変わる', () => {
    const p = project();
    const d = dispatch(newDemo(p, 'u1'), { type: 'visitPage', pageId: 'pg_rsv' });
    const v = view(p, d) as { ui: Record<string, unknown> };
    const u = worldOf(p, d).users['u1']!;
    expect(phone(v as never, u)).toContain('>入力内容を確認する</button>');
    v.ui['form'] = { userId: 'u1', pageId: 'pg_rsv', values: { f_slot: '3日目' }, confirming: true };
    const html = phone(v as never, u);
    expect(html).toContain('<h2>予約内容の確認</h2>');
    expect(html).toContain('<dt>日時</dt><dd>3日目</dd>');
    expect(html).toContain('data-act="formOk"');
    expect(html).toContain('data-act="formBack"');
    // 「修正する」で戻ると、選んでいた値のまま
    v.ui['form'] = { userId: 'u1', pageId: 'pg_rsv', values: { f_slot: '3日目' }, confirming: false };
    expect(phone(v as never, u)).toContain('<option selected>3日目</option>');
  });

  it('demoButton:false のCVは、②に手で記録するボタンを出さない', () => {
    const p = project();
    p.cvs = [{ id: 'cv1', label: '体験予約', demoButton: false }];
    const d = dispatch(newDemo(p, 'u1'), { type: 'addFriend', routeId: 'r1' });
    const html = backstage(view(p, d), worldOf(p, d).users['u1']!);
    expect(html).toContain('体験予約（①の画面で操作）');
    expect(html).not.toContain('data-cv="cv1"');
  });

  it('③の GOAL は最後のCV地点。運営側の登録（stage）は GOAL にしない', () => {
    const p = project();
    p.cvs = [{ id: 'cv1', label: '体験予約' }, { id: 'st', label: '入会', kind: 'stage' }];
    p.journey = [{ id: 'j1', label: '予約', cvId: 'cv1' }, { id: 'j2', label: '入会', cvId: 'st' }];
    const d = newDemo(p, 'u1');
    expect(journeyOf(p, worldOf(p, d), worldOf(p, d).users['u1']!).filter((r) => r.isGoal).map((r) => r.step.id)).toEqual(['j1']);
  });
});
