// ③の「全ユーザーの時系列」と「次にすること」（packages/player/src/timeline.ts）、はじめかたの入口（settings.demoEntries）。
// 時系列は固定の台本ではなく、エンジンの実行結果（行動・配信・判定・タグ・CV）から作られていることを確かめる。

import { describe, expect, it } from 'vitest';
import type { Project } from '../packages/engine/src/index.ts';
import { validate } from '../packages/engine/src/index.ts';
import { allScripted, dispatch, newDemo, nextStop, seek, selectUser, worldOf, type Demo } from '../packages/player/src/controller.ts';
import { Labels } from '../packages/player/src/describe.ts';
import { phone } from '../packages/player/src/phone.ts';
import { buildTimeline, guideOf, horizonOf, nextGlobal, summaryAt, timelineHtml, timelineOf } from '../packages/player/src/timeline.ts';
import type { Ui, View } from '../packages/player/src/view.ts';
import { DAY, sample } from './helpers.ts';

const project = sample();
const ui = (over: Partial<Ui> = {}): Ui => ({ menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, ...over });
const viewOf = (p: Project, demo: Demo, over: Partial<Ui> = {}): View => ({ project: p, world: worldOf(p, demo), demo, ui: ui(over), L: new Labels(p), assetUrl: (a) => a.path });
const me = (v: View) => v.world.users[v.demo.activeUserId]!;

describe('全ユーザーの時系列：エンジンの結果から作る', () => {
  const demo = seek(allScripted(project), 30 * DAY);
  const world = worldOf(project, demo);
  const tl = buildTimeline(project, new Labels(project), world, demo.session.users.map((u) => u.id));

  it('全員の出来事が時刻順に並び、全員分ある', () => {
    expect(tl.map((e) => e.at)).toEqual([...tl.map((e) => e.at)].sort((a, b) => a - b));
    expect(new Set(tl.map((e) => e.userId))).toEqual(new Set(demo.session.users.map((u) => u.id)));
  });

  it('届いた配信・操作・CV・停止は、その人のその時刻の出来事に必ず入る（漏れがない）', () => {
    for (const u of Object.values(world.users)) {
      const at = (t: number) => tl.find((e) => e.userId === u.id && e.at === t);
      for (const d of u.deliveries) expect(at(d.at)?.front.some((x) => x.endsWith('を受信')), `${u.name} ${d.at} ${d.messageId}`).toBe(true);
      for (const b of u.behaviors) expect(at(b.at), `${u.name} ${b.type}`).toBeTruthy();
      for (const [, c] of Object.entries(u.cvs)) expect(at(c.at)?.kind).toBe('cv');
      for (const t of Object.values(u.tags)) if (t.at > 0) expect(at(t.at)?.back.some((x) => x.startsWith('タグ「')), `${u.name} tag@${t.at}`).toBe(true);
    }
    for (const d of world.decisions.filter((x) => x.kind === 'stopWhen')) {
      const e = tl.find((x) => x.userId === d.userId && x.at === d.at);
      expect(e?.back.some((x) => x.includes('を停止')), `${d.userId} ${d.flowId}`).toBe(true);
    }
  });

  it('LINE画面に何も出ない出来事は、表側が空（画面では「裏側の処理のみ」と出す）', () => {
    const backOnly = tl.filter((e) => !e.minor && e.front.length === 0);
    expect(backOnly.length).toBeGreaterThan(0);
    const v = viewOf(project, demo, { tlFuture: true });
    expect(timelineHtml(v)).toContain('裏側の処理のみ。LINE画面に変化なし');
  });

  it('同じ人・同じ時刻のものは1件にまとまる（重複しない）', () => {
    const keys = tl.map((e) => e.key);
    expect(new Set(keys).size).toBe(keys.length);
  });
});

describe('過去・いま・予定を分ける', () => {
  it('いまより後は「予定」として出し、いまの線を引く。予定はデモの終わりまで見通した結果と同じ', () => {
    const demo = seek(allScripted(project), 600 + 30);
    const v = viewOf(project, demo);
    const all = timelineOf(v);
    expect(all.some((e) => e.at > v.world.now)).toBe(true);
    const html = timelineHtml(v);
    expect(html).toContain('id="lsp-tl-now"');
    expect(html).toMatch(/lsp-tl-row future/);
    // 予定の出来事は、実際に時間を進めたときに起きることと一致する
    const later = seek(demo, 3 * DAY);
    const after = buildTimeline(project, new Labels(project), worldOf(project, later), later.session.users.map((u) => u.id)).filter((e) => e.at <= 3 * DAY);
    const planned = all.filter((e) => e.at <= 3 * DAY);
    expect(planned.map((e) => [e.key, e.front, e.back])).toEqual(after.map((e) => [e.key, e.front, e.back]));
  });

  it('予定を隠すと、いまより後の行は出ない', () => {
    const demo = seek(allScripted(project), 600 + 30);
    const html = timelineHtml(viewOf(project, demo, { tlFuture: false }));
    expect(html).not.toMatch(/lsp-tl-row future/);
  });

  it('時間を戻して答えを変えると、時系列も作り直される（古い回答・古い配信が残らない）', () => {
    let demo = newDemo(project, 'p_tokyo', 600);
    demo = dispatch(demo, { type: 'visitPage', pageId: 'page_lp' });
    demo = dispatch(demo, { type: 'exitIntent', pageId: 'page_lp' });
    demo = dispatch(demo, { type: 'tapPopupCta', pageId: 'page_lp' });
    const before = demo.now;
    const tokyo = dispatch(demo, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] });
    const t1 = timelineOf(viewOf(project, tokyo));
    expect(t1.some((e) => e.front.some((x) => x.includes('東京')))).toBe(true);
    const osaka = dispatch(seek(tokyo, before), { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_osaka'] });
    const t2 = timelineOf(viewOf(project, osaka));
    expect(t2.some((e) => e.front.some((x) => x.includes('「東京」と回答')))).toBe(false);
    expect(t2.some((e) => e.front.some((x) => x.includes('「大阪」と回答')))).toBe(true);
  });

  it('同じイベント列なら見通しを作り直さない（時間を動かすだけなら速い）', () => {
    const demo = allScripted(project);
    expect(horizonOf(project, demo)).toBe(horizonOf(project, seek(demo, 2 * DAY)));
  });
});

describe('「次の出来事」：別のユーザーの出来事でも①②は切り替えず、③で分かる', () => {
  it('次に進む時刻と、そこで誰に何が起きるかが分かる。選んでいる人は変わらない', () => {
    let demo = allScripted(project);
    demo = selectUser(demo, demo.session.users.find((u) => u.personaId === 'p_tokyo')!.id);
    const active = demo.activeUserId;
    const v = viewOf(project, demo);
    const next = nextGlobal(v)!;
    expect(next.at).toBe(nextStop(v.world, demo));
    expect(next.events.length).toBeGreaterThan(0);
    const moved = seek(demo, next.at);
    expect(moved.activeUserId).toBe(active);
    // 東京希望さんより先に、応募済みさん（10:00〜）の出来事がある
    const other = next.events.find((e) => e.userId !== active);
    expect(other).toBeTruthy();
    expect(summaryAt(viewOf(project, moved), next.at)).toMatch(/別のユーザーの出来事です.*①②は表示中の東京希望さんのまま/);
  });
});

describe('次にすること', () => {
  const start = (personaId: string) => newDemo(project, personaId, 600);

  it('友だち前：はじめかたを選ぶ（①の入口を光らせる）', () => {
    const v = viewOf(project, start('p_tokyo'));
    const g = guideOf(v, me(v));
    expect(g.kind).toBe('start');
    expect(g.target).toBe('.lsp-start-main button');
  });

  it('ページ → ポップアップ → アンケート → 配信待ち と、操作のたびに変わる', () => {
    let demo = dispatch(start('p_tokyo'), { type: 'visitPage', pageId: 'page_lp' });
    let v = viewOf(project, demo);
    expect(guideOf(v, me(v)).target).toBe('[data-act="back"]');
    demo = dispatch(demo, { type: 'exitIntent', pageId: 'page_lp' });
    v = viewOf(project, demo);
    expect(guideOf(v, me(v))).toMatchObject({ kind: 'manual', target: '[data-act="popupCta"]' });
    expect(guideOf(v, me(v)).text).toContain('LINEで受け取る');
    demo = dispatch(demo, { type: 'tapPopupCta', pageId: 'page_lp' });
    v = viewOf(project, demo);
    expect(guideOf(v, me(v))).toMatchObject({ kind: 'manual', target: '.lsp-quick' });
    demo = dispatch(demo, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_area', choiceIds: ['c_tokyo'] });
    demo = dispatch(demo, { type: 'answer', surveyId: 'sv_hope', questionId: 'q_job', choiceIds: ['c_office'] });
    v = viewOf(project, demo);
    const g = guideOf(v, me(v));
    expect(g.kind).toBe('wait');
    expect(g.text).toMatch(/「次の出来事 ▶」で、この人の次の出来事（Day \d+ \d\d:\d\d：/);
  });

  it('台本の人は「台本が操作」、ブロックした人は停止の理由を出す', () => {
    const demo = allScripted(project);
    const v = viewOf(project, demo);
    const tokyo = Object.values(v.world.users).find((u) => u.personaId === 'p_tokyo')!;
    expect(guideOf(v, tokyo).kind).toBe('script');
    const late = seek(demo, 10 * DAY);
    const vl = viewOf(project, late);
    const blocked = Object.values(vl.world.users).find((u) => u.personaId === 'p_notyet')!;
    expect(blocked.friendState).toBe('blocked');
    expect(guideOf(vl, blocked)).toMatchObject({ kind: 'stopped' });
    expect(guideOf(vl, blocked).text).toContain('ブロック');
  });
});

describe('はじめかたの入口（settings.demoEntries）', () => {
  const withEntries: Project = { ...project, settings: { ...project.settings, demoEntries: [{ pageId: 'page_lp', label: 'Webサイトから開始' }] } };

  it('指定した入口だけを大きく出し、ほかのページ・経路は折りたたみに残す（消さない）', () => {
    const v = viewOf(withEntries, newDemo(withEntries, 'p_tokyo', 600));
    const html = phone(v, me(v));
    const main = /<div class="lsp-start-main">([\s\S]*?)<\/div>/.exec(html)![1]!;
    expect(main.match(/data-act="visit"/g)?.length).toBe(1);
    expect(main).toContain('Webサイトから開始');
    expect(html).toContain('ほかのページ・経路から始める');
    expect(html).toContain('data-page="page_form"');
    expect(html).toContain('data-route="route_thanks"');
  });

  it('指定が無ければ、これまでどおり全ページを並べる', () => {
    const v = viewOf(project, newDemo(project, 'p_tokyo', 600));
    expect(phone(v, me(v))).not.toContain('ほかのページ・経路から始める');
  });

  it('存在しないページ・空の名前は検査で止める', () => {
    const bad: Project = { ...project, settings: { ...project.settings, demoEntries: [{ pageId: 'page_none', label: '' }] } };
    const errs = validate(bad).filter((i) => i.level === 'error').map((i) => i.path);
    expect(errs).toContain('/settings/demoEntries/0/pageId');
    expect(errs).toContain('/settings/demoEntries/0/label');
  });
});

describe('人物名と役割を分ける（persona.personName）', () => {
  const named: Project = { ...project, personas: project.personas.map((x) => (x.id === 'p_tokyo' ? { ...x, personName: '架空 一郎' } : x)) };

  it('文面の {{user.name}} は人物名、画面は「人物名＠役割」。人物名が無い人はこれまでどおり', async () => {
    const { whoOf } = await import('../packages/player/src/view.ts');
    const d = seek(allScripted(named), 2 * DAY);
    const w = worldOf(named, d);
    const tokyo = Object.values(w.users).find((u) => u.personaId === 'p_tokyo')!;
    expect(tokyo.personName).toBe('架空 一郎');
    expect(whoOf(tokyo)).toBe('架空 一郎＠東京希望');
    const texts = tokyo.talk.flatMap((t) => (t.side === 'bot' && t.bubble.kind === 'text' ? [t.bubble.text] : []));
    expect(texts.some((t) => t.includes('架空 一郎'))).toBe(true);
    expect(texts.some((t) => t.includes('東京希望さん'))).toBe(false);
    const osaka = Object.values(w.users).find((u) => u.personaId === 'p_osaka')!;
    expect(whoOf(osaka)).toBe(osaka.name);
    // 時系列・①②③で同じ呼び名
    const tl = timelineOf(viewOf(named, d));
    expect(tl.filter((e) => e.userId === tokyo.id).every((e) => e.userName === '架空 一郎＠東京希望')).toBe(true);
  });
});

describe('②「次の段階へ進める」のボタン', () => {
  it('名前は「○○へ」（案件で変えられる）。済んだ段階は押せず、二重に記録しない', async () => {
    const { backstage } = await import('../packages/player/src/backstage.ts');
    const p: Project = { ...project, cvs: project.cvs.map((c, i) => (i === 0 ? { ...c, demoAction: '応募後へ' } : c)) };
    let d = dispatch(newDemo(p, 'p_tokyo', 600), { type: 'addFriend', routeId: 'route_exit_popup' });
    let v = viewOf(p, d);
    let html = backstage(v, me(v));
    expect(html).toContain('各段階に進んだ際のLINE配信や、裏側の処理を確認できます。');
    expect(html).toContain('>応募後へ</button>');
    expect(html).not.toContain('したことにする');
    d = dispatch(d, { type: 'cv', cvId: p.cvs[0]!.id });
    v = viewOf(p, d);
    html = backstage(v, me(v));
    expect(html).toMatch(/class="done" disabled[^>]*>✓ [^<]+済み<\/button>/);
    expect(html).not.toContain(`data-cv="${p.cvs[0]!.id}"`);
  });
});

describe('③「いま起きたこと」', () => {
  it('いまの時刻の出来事を全員分出し、表示中の人を先頭に。ほかの人は「この人を表示」で切り替えられる', async () => {
    const { nowHtml } = await import('../packages/player/src/timeline.ts');
    let demo = allScripted(project);
    const v0 = viewOf(project, demo);
    const t = nextGlobal(v0)!.at;
    demo = seek(demo, t);
    const v = viewOf(project, demo);
    const html = nowHtml(v);
    expect(html).toContain('いま起きたこと');
    const evs = timelineOf(v).filter((e) => e.at === t && !e.minor);
    for (const e of evs) expect(html).toContain(e.userName);
    if (evs.some((e) => e.userId !== demo.activeUserId)) expect(html).toContain('この人を表示');
    // 時刻を少し進めると「直前の出来事」になる
    expect(nowHtml(viewOf(project, seek(demo, t + 1)))).toContain('直前の出来事');
  });
});
