// 提案資料から作った企業別デモ案件（projects/client-*/）の検証。
// 案件データは顧客の資料にもとづくのでリポジトリに入れていない。手元に無いときは、このテストは飛ばす。
// テストには顧客名を書かず、ID だけで確かめる。

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { InputEventBody, Project, UserState, WorldState } from '../packages/engine/src/index.ts';
import { bubbleButtons, personaCoverage, reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { allScripted, autoStep, dispatch, newDemo, seek, worldOf, type Demo } from '../packages/player/src/controller.ts';
import { journeyOf } from '../packages/player/src/journey.ts';
import { Labels } from '../packages/player/src/describe.ts';
import { buildTimeline } from '../packages/player/src/timeline.ts';
import { buildExport } from '../tools/export-lib.ts';
import { loadProject } from '../tools/project-io.ts';
import { DAY } from './helpers.ts';

const has = (slug: string) => existsSync(join(import.meta.dirname, '..', 'projects', slug, 'project.json'));
const at = (day: number, hh = 0, mm = 0) => day * DAY + hh * 60 + mm;
const run = (p: Project, until: number) => reduce(p, sessionFromPersonas(p), { until });
const user = (w: WorldState, id: string) => w.users[id]!;
const ids = (w: WorldState, id: string) => user(w, id).deliveries.map((d) => d.messageId);
const flowMsgs = (w: WorldState, id: string, flowId: string) => user(w, id).deliveries.filter((d) => d.flowId === flowId).map((d) => [d.messageId, Math.floor(d.at / DAY)] as const);
const me = (p: Project, d: Demo) => worldOf(p, d).users[d.activeUserId]!;
/** トークの吹き出しを「公式：…／ユーザー：…」の列にする（アンケートの表示確認用） */
const chatLines = (u: UserState) =>
  u.talk.filter((t) => t.side !== 'system').map((t) => `${t.side === 'bot' ? '公式' : 'ユーザー'}：${t.bubble.kind === 'text' ? t.bubble.text.split('\n')[0] : t.bubble.kind}`);

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
  const file = (path: string) => join(import.meta.dirname, '..', 'projects', p.meta.slug, path);
  const r = await buildExport(p, { readAsset: (path: string) => (existsSync(file(path)) ? readFileSync(file(path)) : null), confirmFictional: true });
  // 代替画像も埋め込まれている
  expect(r.html).toContain('data:image/');
  expect(r.ok).toBe(true);
  for (const s of ['見積', '概算', '265,000', '260,000', '月額', '_internal', '"prov"', 'すべて架空です']) expect(r.html).not.toContain(s);
  expect(r.html).toContain("default-src 'none'");
  expect(r.html).toContain('登場するユーザー（人物）は架空です');
  expect(r.html!.length).toBeLessThan(5 * 1024 * 1024);
}

/** 全ボタン・メニュー・ページのボタンを、友だちのユーザーで1つずつ押し、何かが変わる（無反応でない）ことを確かめる */
function deadButtons(p: Project, personaId: string, routeId: string) {
  const base = dispatch(newDemo(p, personaId), { type: 'addFriend', routeId });
  const snap = (d: Demo) => {
    const u = me(p, d);
    return JSON.stringify([u.talk.length, u.surface, u.deliveries.length, Object.keys(u.tags).sort(), u.richMenuId]);
  };
  const dead: string[] = [];
  const tryIt = (label: string, ...events: InputEventBody[]) => {
    let d = base;
    for (const e of events.slice(0, -1)) d = dispatch(d, e);
    const before = snap(d);
    const after = snap(dispatch(d, events.at(-1)!));
    if (before === after) dead.push(label);
  };
  for (const m of p.messages) for (const b of m.bubbles) for (const x of bubbleButtons(b)) tryIt(`button:${x.id}`, { type: 'tapButton', buttonId: x.id });
  // いま開いているタブ（動作なし）は押しても変わらないのが正しい
  for (const r of p.richMenus) {
    // 既定でないメニュー（2つ目のタブ等）は、そのタブに切り替えてから押す
    const open = p.richMenus.flatMap((m) => m.cells.filter((c) => c.action.kind === 'switchMenu' && c.action.richMenuId === r.id).map((c): InputEventBody => ({ type: 'tapMenu', richMenuId: m.id, cellId: c.id })));
    const pre = r.id !== p.settings.defaultRichMenuId && open[0] ? [open[0]] : [];
    for (const c of r.cells) if (c.action.kind !== 'none') tryIt(`menu:${c.id}`, ...pre, { type: 'tapMenu', richMenuId: r.id, cellId: c.id });
  }
  for (const pg of p.webPages) {
    pg.blocks.forEach((b, i) => b.kind === 'button' && tryIt(`page:${pg.id}#${i}`, { type: 'visitPage', pageId: pg.id }, { type: 'tapPageButton', pageId: pg.id, blockIndex: i }));
    if (pg.lineCta) tryIt(`lineCta:${pg.id}`, { type: 'visitPage', pageId: pg.id }, { type: 'tapLineCta', pageId: pg.id });
  }
  return dead;
}

/** ボタンの文言と行き先のページが合っているか（キーワード → ページ） */
function mismatched(p: Project, rules: [RegExp, string][]) {
  const out: string[] = [];
  const check = (label: string, action: { kind: string; pageId?: string }) => {
    for (const [re, pageId] of rules) if (re.test(label) && action.kind === 'openPage' && action.pageId !== pageId) out.push(`${label} → ${action.pageId}`);
  };
  for (const m of p.messages) for (const b of m.bubbles) if (b.kind === 'buttons' || b.kind === 'card') for (const x of b.buttons) check(x.label, x.action);
  for (const pg of p.webPages) for (const b of pg.blocks) if (b.kind === 'button') check(b.label, b.action);
  return out;
}

/** シナリオの配信（自動応答以外）で、同じ文面が同じ人に2回届いていないか */
const repeats = (w: WorldState) =>
  Object.values(w.users).flatMap((u) => {
    const seen = new Set<string>();
    return u.deliveries.filter((d) => d.flowId !== null && (seen.has(d.messageId) || !seen.add(d.messageId))).map((d) => `${u.id}:${d.messageId}`);
  });

describe.skipIf(!has('client-a-house'))('企業別デモ①（注文住宅・来場予約）', () => {
  const p = has('client-a-house') ? loadProject('client-a-house') : (null as unknown as Project);
  const w = has('client-a-house') ? run(p, 35 * DAY) : (null as unknown as WorldState);

  it('検証で重大エラー・警告が無く、全シナリオに誰かが入る', () => {
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
  });

  it('全CTA・リッチメニュー・ページのボタンが、押すと反応する（友だち追加済みでLINE追加を押すとトークへ戻る）', () => {
    expect(deadButtons(p, 'p_consider', 'route_popup')).toEqual([]);
    let d = dispatch(newDemo(p, 'p_consider'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'visitPage', pageId: 'page_lp' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_lp', blockIndex: 2 });
    expect(me(p, d).surface).toEqual({ kind: 'talk' });
  });

  it('ボタンの文言と行き先が合っている（施工事例・ルームツアー・予約・間取り・住宅性能・総予算）', () => {
    expect(
      mismatched(p, [[/施工事例を見る/, 'page_cases'], [/予約/, 'page_reserve'], [/間取り(の例|プラン集)を見る/, 'page_plans'], [/住宅性能/, 'page_perf'], [/総予算/, 'page_budget'], [/地図/, 'page_map']]).filter((x) => !x.includes('次回の打ち合わせ')),
    ).toEqual([]);
    // 記事のボタンは、関心タグを付けて内容の画面を開く
    let d = dispatch(newDemo(p, 'p_consider'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_s3_cost' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_loan' });
    expect('int_cost' in me(p, d).tags).toBe(true);
    // 同じボタンをもう一度押しても無反応にならない
    d = dispatch(d, { type: 'closePage', pageId: 'page_loan' });
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_s3_cost' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_loan' });
  });

  it('リッチメニューに「LINEで相談」があり、費用・土地・間取り・来場予約を選べる', () => {
    const cell = p.richMenus[0]!.cells.find((c) => c.id === 'cell_consult')!;
    expect(cell.label).toBe('LINEで相談');
    let d = dispatch(newDemo(p, 'p_consider'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_main', cellId: 'cell_consult' });
    expect(me(p, d).deliveries.at(-1)!.messageId).toBe('msg_r_consult');
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_c_plan' });
    expect(me(p, d).deliveries.at(-1)!.messageId).toBe('msg_r_plan');
  });

  it('アンケートは「公式の質問 → ユーザーの回答 → 次の質問」の順にトークへ残る', () => {
    const lines = chatLines(user(w, 'p_consider'));
    const i = lines.indexOf('公式：家づくりは、いまどの段階ですか？');
    expect(lines.slice(i, i + 4)).toEqual(['公式：家づくりは、いまどの段階ですか？', 'ユーザー：情報収集を始めた', '公式：平屋と2階建て、どちらをお考えですか？', 'ユーザー：平屋']);
  });

  it('簡易診断に全問答えた人にだけ、お礼を1回送る（来場予約済みは文面を変える／途中・期限切れ・ステップ配信には影響しない）', () => {
    const thanks = (id: string) => ids(w, id).filter((m) => m === 'msg_diag_thanks' || m === 'msg_diag_thanks_res');
    for (const id of ['p_consider', 'p_cost', 'p_doc', 'p_doc_wait', 'p_plan']) expect(thanks(id), id).toEqual(['msg_diag_thanks']);
    expect(thanks('p_reserved')).toEqual(['msg_diag_thanks_res']);
    // 答えなかった人（期限切れ）には送らない
    expect(thanks('p_block')).toEqual([]);
    // お礼は最後の回答の直後で、そのあとに質問は出ない
    const lines = chatLines(user(w, 'p_consider'));
    expect(lines[lines.indexOf('ユーザー：平屋') + 1]).toBe('公式：ご回答ありがとうございます！');
    // 1問目だけで止めた人にも送らない
    let d = dispatch(newDemo(p, 'p_consider'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: p.surveys[0]!.questions[0]!.id, choiceIds: [p.surveys[0]!.questions[0]!.choices[0]!.id] });
    d = seek(d, d.now + 2 * DAY);
    expect(me(p, d).deliveries.some((x) => x.messageId === 'msg_diag_thanks')).toBe(false);
  });

  it('架空の人物名で呼びかける（画面は「人物名＠役割」）。役割名は文面に出ない', () => {
    for (const persona of p.personas) {
      expect(persona.personName, persona.id).toMatch(/^\S+ \S+$/);
      const texts = user(w, persona.id).talk.flatMap((t) => (t.side === 'bot' && t.bubble.kind === 'text' ? [t.bubble.text] : []));
      expect(texts.some((t) => t.includes(persona.name)), persona.id).toBe(false);
    }
    const first = user(w, 'p_consider').talk.find((t) => t.side === 'bot')!;
    expect(first.side === 'bot' && first.bubble.kind === 'text' && first.bubble.text.startsWith(`${p.personas[0]!.personName}さん、`)).toBe(true);
  });

  it('代替画像を使うカード・ページには、デモ用の代替画像だと書いてあり、実物と断定しない', () => {
    for (const m of p.messages)
      for (const b of m.bubbles) {
        const text = b.kind === 'card' ? `${b.title}${b.body ?? ''}` : b.kind === 'text' || b.kind === 'buttons' ? b.text : '';
        expect(text, m.id).not.toMatch(/実際に建てた|実際のお住まい|実際の施工事例/);
        if (b.kind === 'card' && b.assetId) expect(b.body, m.id).toContain('代替画像');
      }
    for (const page of p.webPages)
      if (page.blocks.some((b) => b.kind === 'image')) expect(page.blocks.some((b) => b.kind === 'text' && /代替画像|貴社の資料に差し替え/.test(b.text)), page.id).toBe(true);
  });

  it('②のボタンは「資料請求へ／来場予約後へ／来場後へ／契約後へ」。押すとCVとして記録され、配信・タグが動く', () => {
    expect(p.cvs.map((c) => c.demoAction)).toEqual(['資料請求へ', '来場予約後へ', '来場後へ', '契約後へ']);
    let d = dispatch(newDemo(p, 'p_consider'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'cv', cvId: 'cv_reserve' });
    const u = me(p, d);
    expect(u.cvs['cv_reserve']?.count).toBe(1);
    expect(u.flows['flow_step']?.status).toBe('stopped');
    expect(u.flows['flow_reserve']).toBeTruthy();
  });

  it('はじめかたの入口は3つ（Webサイト・資料請求・来場予約）。詳細ページはデータに残り、配信やCTAから開ける', () => {
    expect(p.settings.demoEntries?.map((e) => [e.label, e.pageId])).toEqual([
      ['Webサイトから開始', 'page_lp'],
      ['資料請求から開始', 'page_doc'],
      ['来場予約から開始', 'page_reserve'],
    ]);
    for (const id of ['page_cases', 'page_tour', 'page_loan', 'page_plans']) expect(p.webPages.some((x) => x.id === id)).toBe(true);
  });

  it('7人全員について、Day0 から最後までの出来事が③の時系列に出る（表側・裏側の両方）', () => {
    const d = seek(allScripted(p), 35 * DAY);
    const tl = buildTimeline(p, new Labels(p), worldOf(p, d), d.session.users.map((u) => u.id)).filter((e) => !e.minor);
    expect(new Set(tl.map((e) => e.userId)).size).toBe(7);
    for (const u of d.session.users) {
      const mine = tl.filter((e) => e.userId === u.id);
      expect(mine[0]!.at, u.personaId).toBeLessThan(DAY);
      expect(mine.some((e) => e.back.length > 0)).toBe(true);
    }
    // 予約で止まる・ブロックで届かない、も出来事として出る
    expect(tl.some((e) => e.back.some((x) => x.includes('を停止')))).toBe(true);
    expect(tl.some((e) => e.kind === 'stop' && e.back.some((x) => x.includes('ブロックしている')))).toBe(true);
  });

  it('14日間・全5通は変えず、15日目以降は未予約者だけに関心別のフォロー（16日後・22日後）と29日後の案内', () => {
    expect(flowMsgs(w, 'p_consider', 'flow_step').map(([, d]) => d)).toEqual([0, 1, 3, 6, 12]);
    expect(flowMsgs(w, 'p_consider', 'flow_after14')).toEqual([['msg_f1_info', 16], ['msg_f2_info', 22], ['msg_f3', 29]]);
    expect(flowMsgs(w, 'p_doc_wait', 'flow_after14').slice(0, 2)).toEqual([['msg_f1_doc', 16], ['msg_f2_doc', 22]]);
    expect(flowMsgs(w, 'p_plan', 'flow_after14').slice(0, 2)).toEqual([['msg_f1_plan', 16], ['msg_f2_plan', 22]]);
    // 費用に関心のある人の分岐は、手で操作して確かめる（資料請求なし・費用タグあり）
    let d = dispatch(newDemo(p, 'p_consider'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_main', cellId: 'cell_money' });
    d = seek(d, at(23));
    expect(me(p, d).deliveries.filter((x) => x.flowId === 'flow_after14').map((x) => x.messageId)).toEqual(['msg_f1_cost', 'msg_f2_cost']);
    // 来場予約済み・ブロック中には届かない
    expect(user(w, 'p_doc').deliveries.some((x) => x.flowId === 'flow_after14')).toBe(false);
    expect(w.decisions.filter((x) => x.userId === 'p_block' && x.flowId === 'flow_after14' && x.kind === 'deliver').every((x) => x.reason === 'blocked')).toBe(true);
  });

  it('来場予約で来場促進が止まり、予約確認・3日前・前日 → 来場後のお礼・振り返り・不安別の情報・構造見学会・次回打ち合わせ', () => {
    expect(flowMsgs(w, 'p_cost', 'flow_reserve')).toEqual([['msg_r_confirm', 6], ['msg_r_3days', 10], ['msg_r_daybefore', 12]]);
    expect(flowMsgs(w, 'p_cost', 'flow_visit')).toEqual([['msg_v_thanks', 13], ['msg_v_cost', 16], ['msg_k_after', 19], ['msg_v_next', 23]]);
    const t = user(w, 'p_cost').cvs['cv_reserve']!.at;
    expect(user(w, 'p_cost').deliveries.filter((d) => d.at > t && ['flow_step', 'flow_k_event', 'flow_after14'].includes(d.flowId ?? ''))).toEqual([]);
  });

  it('契約したら、契約前の配信（次回打ち合わせ）は届かない', () => {
    expect(flowMsgs(w, 'p_doc', 'flow_visit').map(([m]) => m)).toEqual(['msg_v_thanks', 'msg_v_general', 'msg_k_after']);
    expect(user(w, 'p_doc').flows['flow_visit']!.status).toBe('stopped');
  });

  it('同じ文面が同じ人に重ねて届かない（ステップ・企画・15日目以降・来場後）', () => {
    expect(repeats(w)).toEqual([]);
  });

  it('過去に戻って別の答えを選ぶと、その先の配信・タグ・トークの履歴が計算し直される', () => {
    let d = newDemo(p, 'p_consider');
    d = dispatch(d, { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_status', choiceIds: ['c_compare'] });
    const before = d.now;
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_type', choiceIds: ['c_flat'] });
    d = seek(d, at(7));
    expect(me(p, d).deliveries.map((x) => x.messageId)).toContain('msg_s4_flat');
    d = seek(d, before);
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_type', choiceIds: ['c_two'] });
    d = seek(d, at(7));
    const now = me(p, d);
    expect(now.deliveries.map((x) => x.messageId)).toContain('msg_s4_two');
    expect(now.deliveries.map((x) => x.messageId)).not.toContain('msg_s4_flat');
    expect('house_flat' in now.tags).toBe(false);
    expect(chatLines(now)).toContain('ユーザー：2階建て');
    expect(chatLines(now)).not.toContain('ユーザー：平屋');
  });

  it('③：資料請求は中間CV、来場予約が GOAL。15日目以降・来場後のルートは②のタグと一致する', () => {
    const rows = (id: string) => journeyOf(p, w, user(w, id));
    expect(rows('p_cost').find((r) => r.isGoal)!.step.id).toBe('j_reserve');
    expect(rows('p_doc_wait').find((r) => r.step.id === 'j_after14')!.traversed).toEqual(['v_f_doc']);
    expect(rows('p_plan').find((r) => r.step.id === 'j_after14')!.traversed).toEqual(['v_f_plan']);
    expect(rows('p_cost').find((r) => r.step.id === 'j_after')!.traversed).toEqual(['v_a_cost']);
    expect(rows('p_cost').at(-1)!.status).toBe('current');
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

  it('全CTA・リッチメニュー・ページのボタンが、押すと反応する', () => {
    expect(deadButtons(p, 'p_now', 'route_popup')).toEqual([]);
  });

  it('ボタンの文言と行き先が合っている（園のこと・1日の流れ・よくある質問・予約・アクセス）', () => {
    expect(mismatched(p, [[/園のことを詳しく|教育について|3つの教育を見る|について見る/, 'page_about'], [/1日の流れ/, 'page_day'], [/よくある質問|費用・見学内容/, 'page_faq'], [/予約/, 'page_tour'], [/アクセス|園の場所/, 'page_place']]).filter((x) => !/(長時間保育|しつけ・生活習慣)について見る → page_day/.test(x))).toEqual([]);
    // Day1 の「園のことを詳しく見る」は、LINE追加の画面ではなく教育の説明画面へ
    const about = p.webPages.find((x) => x.id === 'page_about')!;
    expect(about.blocks.filter((b) => b.kind === 'heading').map((b) => (b as { text: string }).text)).toEqual([`${p.brand.accountName} の教育方針`, '英語教育', '知能教育', '運動指導', '1日の流れ']);
    expect(JSON.stringify(about)).not.toContain('addFriend');
    // Day18 は「予約」と「見学前によくある質問」
    const d18 = p.messages.find((m) => m.id === 'msg_d18')!.bubbles[0] as { buttons: { label: string }[] };
    expect(d18.buttons.map((b) => b.label)).toEqual(['見学・体験を予約する', '見学前によくある質問を見る']);
  });

  it('アンケートは質問と回答が交互にトークへ残る（簡易 → 本アンケート Q1〜Q4）', () => {
    const lines = chatLines(user(w, 'p_later'));
    const i = lines.indexOf('公式：Q1 お子様の年齢・学年を教えてください');
    expect(lines.slice(i, i + 8)).toEqual([
      '公式：Q1 お子様の年齢・学年を教えてください',
      'ユーザー：3歳',
      '公式：Q2 入園の希望時期は？',
      'ユーザー：随時',
      '公式：Q3 いちばん重視することは？',
      'ユーザー：知能',
      '公式：Q4 見学のご希望は？',
      'ユーザー：後日見学したい',
    ]);
    const quick = lines.findIndex((l) => l.startsWith('公式：') && l.endsWith('で、いちばん気になることは？'));
    expect(quick).toBeGreaterThanOrEqual(0);
    expect(quick).toBeLessThan(i);
    // 簡易アンケートに答えない人：時間切れで回答のボタンを閉じ、本アンケートの質問は出さない
    const skip = user(w, 'p_skip');
    expect(chatLines(skip).filter((l) => l.startsWith('公式：Q'))).toEqual([]);
    expect(skip.surface.kind).not.toBe('survey');
  });

  it('ステップ配信は資料の10本の枠（Day0 挨拶・アンケート＋Day1〜30の9通）のまま', () => {
    expect(flowMsgs(w, 'p_info', 'flow_step').map(([, d]) => d)).toEqual([1, 3, 5, 7, 10, 14, 23, 30]); // 情報収集のみは Day18 の見学案内を送らない
    expect(flowMsgs(w, 'p_skip', 'flow_step').length + w.decisions.filter((x) => x.userId === 'p_skip' && x.flowId === 'flow_step' && x.reason === 'blocked').length).toBe(9);
  });

  it('早期未予約フォロー：Q4の回答で届く情報とタイミングが変わる（A 今すぐ／B 後日／C 情報収集）', () => {
    // A：2日後に再案内 → Day18・23 は今すぐ向け → 予約で即停止（Day30 は届かない）
    expect(flowMsgs(w, 'p_now_wait', 'flow_early_a')).toEqual([['msg_ea1', 2]]);
    expect(flowMsgs(w, 'p_now_wait', 'flow_step').slice(-2)).toEqual([['msg_d18_now', 18], ['msg_d23_now', 23]]);
    expect(ids(w, 'p_now_wait')).not.toContain('msg_d30');
    // B：12日後に重視点（知能）の見学メリット → Day18 で改めて見学予約
    expect(flowMsgs(w, 'p_later', 'flow_early_b')).toEqual([['msg_eb_iq', 12]]);
    expect(ids(w, 'p_later')).toContain('msg_d18_later');
    // C：見学を迫らない（Day18 なし）→ 21日後に検討状況 → Day23・30 は情報収集向け
    expect(ids(w, 'p_info')).not.toContain('msg_d18');
    expect(ids(w, 'p_info')).toEqual(expect.arrayContaining(['msg_ec_check', 'msg_d23_info', 'msg_d30_info']));
    // 同じ人に A・B・C が重ならない
    for (const id of ['p_now_wait', 'p_later', 'p_info']) expect(['flow_early_a', 'flow_early_b', 'flow_early_c'].filter((f) => user(w, id).flows[f]).length).toBe(1);
  });

  it('「他園に決定」と答えたら、見学の案内（ステップ・長期休眠）を止める', () => {
    const other = user(w, 'p_other');
    expect(ids(w, 'p_other')).toContain('msg_ec_other');
    const t = other.tags['decided_other']!.at;
    expect(other.deliveries.filter((d) => d.at > t)).toEqual([]);
    expect(other.flows['flow_step']!.status).toBe('stopped');
    expect(other.flows['flow_dormant_b']!.status).toBe('stopped');
  });

  it('早期未予約・長期休眠・見学後未入園は別ルートで、入園・見学予約・ブロックで止まる', () => {
    // 長期休眠：最終案内（Day30）と並ばないよう 37日後 → 60日後 → 90日後
    expect(flowMsgs(w, 'p_info', 'flow_dormant_b')).toEqual([['msg_b1', 37], ['msg_b2', 60], ['msg_b3', 90]]);
    // 見学後・未入園：3日後・2週間後・45日後。入園で止まる
    expect(flowMsgs(w, 'p_later', 'flow_dormant_a').map(([m]) => m)).toEqual(['msg_a1', 'msg_a2_bus', 'msg_a3']);
    expect(flowMsgs(w, 'p_now', 'flow_dormant_a').map(([m]) => m)).toEqual(['msg_a1', 'msg_a2_cost']);
    expect(user(w, 'p_now').flows['flow_dormant_a']!.status).toBe('stopped');
    // 見学予約した人に、早期フォロー・長期休眠は届かない
    for (const id of ['p_now', 'p_later', 'p_now_wait']) {
      const t = user(w, id).cvs['cv_tour']!.at;
      expect(user(w, id).deliveries.filter((d) => d.at > t && ['flow_step', 'flow_early_a', 'flow_early_b', 'flow_early_c', 'flow_dormant_b'].includes(d.flowId ?? ''))).toEqual([]);
    }
    expect(w.decisions.filter((x) => x.userId === 'p_skip' && x.kind === 'deliver' && x.outcome !== 'taken').every((x) => x.reason === 'blocked')).toBe(true);
  });

  it('同じ文面が同じ人に重ねて届かない', () => {
    expect(repeats(w)).toEqual([]);
  });

  it('見学予約で予約済みのメニューになり、見学のあとは元のメニューに戻る', () => {
    expect(user(run(p, at(20)), 'p_later').richMenuId).toBe('rm_reserved');
    expect(user(w, 'p_later').richMenuId).toBe('rm_main');
  });

  it('過去に戻って Q4 を変えると、配信・タグ・トークの質問と回答が計算し直される', () => {
    let d = newDemo(p, 'p_later');
    d = dispatch(d, { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_quick', questionId: 'q_quick', choiceIds: ['c_qk_edu'] });
    for (const [q, c] of [['q_age', 'c_age3'], ['q_enter', 'c_any'], ['q_focus', 'c_iq']]) d = dispatch(d, { type: 'answer', surveyId: 'sv_main', questionId: q!, choiceIds: [c!] });
    const before = d.now;
    d = dispatch(d, { type: 'answer', surveyId: 'sv_main', questionId: 'q_tour', choiceIds: ['c_later'] });
    d = seek(d, at(13));
    expect(me(p, d).deliveries.map((x) => x.messageId)).toEqual(expect.arrayContaining(['msg_f_iq', 'msg_eb_iq']));
    d = seek(d, before);
    d = dispatch(d, { type: 'answer', surveyId: 'sv_main', questionId: 'q_tour', choiceIds: ['c_now'] });
    d = seek(d, at(13));
    const now = me(p, d);
    const got = now.deliveries.map((x) => x.messageId);
    expect(got).toEqual(expect.arrayContaining(['msg_tour_now', 'msg_ea1']));
    expect(got).not.toContain('msg_f_iq');
    expect(got).not.toContain('msg_eb_iq');
    expect('tour_later' in now.tags).toBe(false);
    expect(chatLines(now)).toContain('ユーザー：今すぐ見学したい');
    expect(chatLines(now)).not.toContain('ユーザー：後日見学したい');
  });

  it('③：見学予約が GOAL。早期フォロー・長期休眠・見学後のルートは②のタグと一致する', () => {
    const rows = (id: string) => journeyOf(p, w, user(w, id));
    expect(rows('p_now').find((r) => r.isGoal)!.step.id).toBe('j_tour');
    expect(rows('p_now_wait').find((r) => r.step.id === 'j_early')!.traversed).toEqual(['v_e_now']);
    expect(rows('p_later').find((r) => r.step.id === 'j_early')!.traversed).toEqual(['v_e_later']);
    expect(rows('p_info').find((r) => r.step.id === 'j_early')!.traversed).toEqual(['v_e_info']);
    expect(rows('p_info').find((r) => r.step.id === 'j_longterm')!.status).toBe('current');
    expect(rows('p_later').find((r) => r.step.id === 'j_after_visit')!.traversed).toEqual(['v_av_bus']);
    expect(rows('p_now').at(-1)).toMatchObject({ status: 'current' });
  });

  it('自動再生で全員を最後まで流すと CLI と同じ結果になり、営業用HTMLを書き出せる', async () => {
    const d = playToEnd(p);
    expect(d.now).toBe(95 * DAY);
    const ww = worldOf(p, d);
    for (const u of Object.values(ww.users)) expect(u.deliveries).toEqual(user(w, d.session.users.find((x) => x.id === u.id)!.personaId).deliveries);
    await exportChecks(p);
  }, 30_000);
});

describe.skipIf(!has('client-c-resort'))('企業別デモ③（リゾートバイト・ENTRY／応募）', () => {
  const p = has('client-c-resort') ? loadProject('client-c-resort') : (null as unknown as Project);
  const w = has('client-c-resort') ? run(p, 190 * DAY) : (null as unknown as WorldState);
  const cardsOf = (messageId: string) => p.messages.find((m) => m.id === messageId)!.bubbles.flatMap((b) => (b.kind === 'carousel' ? b.cards.map((c) => c.title) : []));
  const startsOf = (messageId: string) => p.messages.find((m) => m.id === messageId)!.bubbles.flatMap((b) => (b.kind === 'carousel' ? b.cards.map((c) => c.facts!.find((f) => f.label === '開始時期')!.value) : []));
  /** 友だち追加して3問に答えたところ */
  const answered = (start: string, area: string) => {
    let d = dispatch(newDemo(p, 'p_ski'), { type: 'addFriend', routeId: 'route_popup' });
    for (const [q, c] of [['q_entry', 'c_ent_yet'], ['q_start', start], ['q_area', area]] as const) d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: q, choiceIds: [c] });
    return d;
  };

  it('検証で重大エラー・警告が無く、全シナリオに誰かが入る', () => {
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
  });

  it('全CTA・カルーセルのボタン・リッチメニュー（2タブ）・ページのボタンが、押すと反応する', () => {
    expect(deadButtons(p, 'p_ski', 'route_popup')).toEqual([]);
    let d = dispatch(newDemo(p, 'p_ski'), { type: 'addFriend', routeId: 'route_popup' });
    expect(me(p, d).richMenuId).toBe('rm_jobs');
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_jobs', cellId: 'cell_tab_support' });
    expect(me(p, d).richMenuId).toBe('rm_support');
  });

  it('アンケートは今の3問（エントリー状況・開始時期・エリア）で、質問と回答が交互にトークへ残る', () => {
    const lines = chatLines(user(w, 'p_ski'));
    const i = lines.indexOf('公式：エントリーはされましたか？');
    expect(lines.slice(i, i + 6)).toEqual(['公式：エントリーはされましたか？', 'ユーザー：これから', '公式：お仕事の開始時期は？', 'ユーザー：1〜2ヶ月先', '公式：希望のエリアは？', 'ユーザー：山・高原（スキー場など）']);
  });

  it('【修正1】ENTRYしていない人は応募できない（応募ボタンはENTRYへ）。ENTRYのあと、選んだ求人の応募に戻れる', () => {
    let d = dispatch(newDemo(p, 'p_ski'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'visitPage', pageId: 'page_job_w1' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_job_w1', blockIndex: 5 });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_entry' });
    expect('want_w1' in me(p, d).tags).toBe(true);
    // 応募フォームはENTRY前には出ない（どこからも開けない）
    for (const pg of p.webPages) for (const b of pg.blocks) if (b.kind === 'button' && b.action.kind === 'openPage') expect(b.action.pageId.startsWith('page_apply_')).toBe(false);
    d = dispatch(d, { type: 'submitForm', pageId: 'page_entry', values: { f_start: '1〜2ヶ月先' } });
    expect(Object.keys(me(p, d).cvs)).toEqual(['cv_entry']);
    const back = p.webPages.find((x) => x.id === 'page_entry_thanks')!.blocks.findIndex((b) => b.kind === 'button');
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_entry_thanks', blockIndex: back });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_apply_w1' });
    // ENTRY済みなら、応募ボタンからそのまま応募フォームへ
    d = dispatch(d, { type: 'visitPage', pageId: 'page_job_s2' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_job_s2', blockIndex: 5 });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_apply_s2' });
    expect('want_w1' in me(p, d).tags).toBe(false);
  });

  it('【修正1】LINEの友だちでない人の画面で、LINEで連絡すると言い切らない', () => {
    const t = p.webPages.find((x) => x.id === 'page_apply_thanks')!.blocks.map((b) => (b.kind === 'text' ? b.text : '')).join('');
    expect(t).toContain('ご希望の連絡方法');
    expect(t).not.toMatch(/LINEでご連絡します/);
  });

  it('【修正2】カルーセル → 詳細 → 応募フォームで同じ求人。フォームの求人とCVの記録が一致する', () => {
    let d = answered('c_st_1_2', 'c_ar_mt');
    d = dispatch(d, { type: 'submitForm', pageId: 'page_entry', values: {} });
    d = seek(d, d.now + 120);
    d = dispatch(d, { type: 'tapButton', buttonId: 'b_agent_mt_1_2_w3' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_job_w3' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_job_w3', blockIndex: 5 });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_apply_w3' });
    const form = p.webPages.find((x) => x.id === 'page_apply_w3')!.form!;
    expect(form.fields.find((f) => f.id === 'f_job')!.options).toEqual(['スキーレンタルショップ']);
    d = dispatch(d, { type: 'submitForm', pageId: 'page_apply_w3', values: { f_job: 'スキーレンタルショップ', f_name: 'x', f_contact: 'LINE' } });
    const u = me(p, d);
    expect(Object.keys(u.cvs).sort()).toEqual(['cv_apply', 'cv_entry']);
    const back = buildTimeline(p, new Labels(p), worldOf(p, d), [u.id]).flatMap((e) => e.back);
    expect(back).toContain('CV「応募（個別の求人）」を記録（応募する求人：スキーレンタルショップ／ご希望の連絡方法：LINE）');
    // 台本の7人も、応募した求人＝最後に選んだ求人
    for (const id of ['p_ski', 'p_summer', 'p_entered', 'p_later']) {
      const x = user(w, id);
      const sent = x.behaviors.find((b) => b.type === 'submitForm' && (b.detail['pageId'] ?? '').startsWith('page_apply_'))!;
      const job = sent.detail['pageId']!.replace('page_apply_', '');
      expect(`want_${job}` in x.tags).toBe(true);
      expect(sent.detail['values.f_job']).toBe(p.webPages.find((pg) => pg.id === `page_job_${job}`)!.title.replace('求人：', ''));
    }
  });

  it('【修正3】診断後のカルーセルは、エリアと開始時期の両方で選ぶ。合う求人が無いときは求人を出さず案内する', () => {
    const got = (start: string, area: string) => me(p, answered(start, area)).deliveries.map((x) => x.messageId).filter((m) => m.startsWith('msg_car_'));
    expect(got('c_st_1_2', 'c_ar_mt')).toEqual(['msg_car_mt_1_2']);
    expect(got('c_st_3', 'c_ar_mt')).toEqual(['msg_car_mt_3plus']);
    expect(cardsOf('msg_car_mt_1_2')).not.toEqual(cardsOf('msg_car_mt_3plus'));
    expect(startsOf('msg_car_mt_1_2').every((s) => s.startsWith('12月'))).toBe(true);
    expect(startsOf('msg_car_mt_3plus').some((s) => s.startsWith('12月') || s.startsWith('11月'))).toBe(false);
    expect(startsOf('msg_car_sea_asap').every((s) => s.startsWith('11月'))).toBe(true);
    // 山・高原×できるだけ早く：10〜11月に始められる求人が無い → カルーセルは出さず、担当が探す案内
    const none = me(p, answered('c_st_asap', 'c_ar_mt'));
    expect(none.deliveries.map((x) => x.messageId)).toContain('msg_car_none');
    expect(none.talk.some((t) => t.side === 'bot' && t.bubble.kind === 'carousel')).toBe(false);
    const txt = none.talk.flatMap((t) => (t.side === 'bot' && t.bubble.kind === 'buttons' ? [t.bubble.text] : [])).join('');
    expect(txt).toContain('ご希望（山・高原（スキー場など）・できるだけ早く）に合う求人は、いま掲載がありません。');
    // 台本の7人
    expect(ids(w, 'p_ski')).toContain('msg_car_mt_1_2');
    expect(ids(w, 'p_summer')).toContain('msg_car_sea_3plus');
    expect(ids(w, 'p_later')).toContain('msg_car_sea_asap');
    expect(ids(w, 'p_menu')).toContain('msg_car_any_unknown');
    expect(ids(w, 'p_noanswer')).toContain('msg_diag_remind');
  });

  it('【修正3】Day10 の案内・ENTRY後のおすすめ・特集も、エリアと開始時期で選ぶ', () => {
    expect(ids(w, 'p_summer')).toContain('msg_s4_sea_3plus');
    expect(ids(w, 'p_later')).toEqual(expect.arrayContaining(['msg_k_sea_asap', 'msg_agent_sea_asap']));
    expect(ids(w, 'p_entered')).toEqual(expect.arrayContaining(['msg_k_onsen_asap', 'msg_agent_onsen_asap']));
    // 開始が先（3ヶ月以上先）・未定・未回答には特集を送らない
    for (const id of ['p_summer', 'p_noanswer', 'p_menu']) expect(ids(w, id).filter((m) => m.startsWith('msg_k_'))).toEqual([]);
    // 山・高原×できるだけ早く は合う求人が無いので、特集も送らない
    const d = seek(answered('c_st_asap', 'c_ar_mt'), at(8));
    expect(me(p, d).deliveries.map((x) => x.messageId).filter((m) => m.startsWith('msg_k_'))).toEqual([]);
  });

  it('【修正4】文面と処理が合っている（「毎週火曜」と書かない・勤務を確認していない人に「おつかれさま」と書かない）', () => {
    const all = p.messages.flatMap((m) => m.bubbles.flatMap((b) => (b.kind === 'text' ? [b.text] : b.kind === 'buttons' ? [b.text] : [])));
    expect(all.some((t) => t.includes('毎週火曜'))).toBe(false);
    expect(all.some((t) => /おつかれさま/.test(t))).toBe(false);
    expect(p.flows.some((f) => f.label.includes('毎週火曜'))).toBe(false);
    expect(p.journey!.some((j) => (j.purpose ?? '').includes('毎週火曜'))).toBe(false);
    // 「次のお仕事」は応募した人だけに届き、文面は応募を前提にした言い方
    const next = p.messages.find((m) => m.id === 'msg_l_next')!.bubbles[0]!;
    expect(next.kind === 'text' && next.text.startsWith('冬のお仕事に応募いただいた方へ')).toBe(true);
    for (const id of ['p_ski', 'p_entered', 'p_later']) expect(ids(w, id)).toContain('msg_l_next');
    for (const id of ['p_noanswer', 'p_menu']) expect(ids(w, id)).not.toContain('msg_l_next');
    // 特集のカードは、文面に書いた開始時期に合う
    expect(startsOf('msg_k_mt_1_2').every((s) => s.startsWith('12月'))).toBe(true);
  });

  it('Day10 は、ENTRYまだの方にENTRYの案内、ENTRY済みの方に応募の案内を送る（ENTRYの記録で判定）', () => {
    expect(ids(w, 'p_noanswer')).toContain('msg_s4_any_unknown');
    for (const id of ['p_entered', 'p_menu']) {
      expect(ids(w, id)).toContain('msg_s4_apply');
      expect(ids(w, id).filter((m) => /^msg_s4_(mt|sea|onsen|any)_/.test(m))).toEqual([]);
    }
  });

  it('ENTRY（会員登録）と応募は別のCV。応募したら求人の案内（ステップ・特集・ENTRY後のフォロー）は止まる', () => {
    expect(Object.keys(user(w, 'p_ski').cvs).sort()).toEqual(['cv_apply', 'cv_entry']);
    expect(Object.keys(user(w, 'p_menu').cvs)).toEqual(['cv_entry']);
    expect(ids(w, 'p_ski')).not.toContain('msg_s2_money');
    expect(ids(w, 'p_ski')).not.toContain('msg_entry_follow');
    expect(ids(w, 'p_entered')).toContain('msg_welcome_thanks');
    expect(ids(w, 'p_entered')).not.toContain('msg_entry_thanks');
  });

  it('季節で中長期の配信が切り替わる（10月下旬は冬・2月は春・4月は夏の海・離島）', () => {
    expect(flowMsgs(w, 'p_summer', 'flow_long')).toEqual([['msg_l_sea_winter', 18], ['msg_l_spring', 120], ['msg_l_summer_sea', 180]]);
    expect(flowMsgs(w, 'p_noanswer', 'flow_long')).toEqual([['msg_l_ski', 18], ['msg_l_spring', 120], ['msg_l_summer_other', 180]]);
    expect(ids(w, 'p_summer')).toContain('msg_agent_summer');
  });

  it('ブロックした人には、その後の配信が届かない', () => {
    const u = user(w, 'p_block');
    const blockAt = u.behaviors.find((b) => b.type === 'block')!.at;
    expect(u.deliveries.every((d) => d.at < blockAt)).toBe(true);
    expect(w.decisions.some((x) => x.userId === 'p_block' && x.kind === 'deliver' && x.reason === 'blocked')).toBe(true);
  });

  it('過去に戻って開始時期の答えを変えると、カルーセル・タグ・トークが計算し直される', () => {
    let d = dispatch(newDemo(p, 'p_ski'), { type: 'addFriend', routeId: 'route_popup' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_entry', choiceIds: ['c_ent_yet'] });
    const before = d.now;
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_start', choiceIds: ['c_st_1_2'] });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_area', choiceIds: ['c_ar_mt'] });
    d = dispatch(d, { type: 'viewCard', messageId: 'msg_car_mt_1_2', cardId: 'card_car_mt_1_2_w1' });
    d = seek(d, before);
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_start', choiceIds: ['c_st_3'] });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_area', choiceIds: ['c_ar_mt'] });
    const now = me(p, d);
    expect(now.deliveries.map((x) => x.messageId)).toContain('msg_car_mt_3plus');
    expect(now.deliveries.map((x) => x.messageId)).not.toContain('msg_car_mt_1_2');
    expect('start_1_2' in now.tags).toBe(false);
    expect(now.behaviors.some((b) => b.type === 'viewCard')).toBe(false);
  });

  it('③は応募が GOAL。入口・カルーセルのルートは②のタグと一致する', () => {
    expect(p.cvs.map((c) => c.demoAction)).toEqual(['ENTRY後へ', '応募後へ']);
    const rows = (id: string) => journeyOf(p, w, user(w, id));
    expect(rows('p_ski').find((r) => r.isGoal)!.step.id).toBe('j_apply');
    expect(rows('p_ski').find((r) => r.step.id === 'j_match')!.traversed).toEqual(['v_m_mt']);
    expect(rows('p_noanswer').find((r) => r.step.id === 'j_match')!.traversed).toEqual(['v_m_none']);
    expect(rows('p_entered').find((r) => r.step.id === 'j_friend')!.traversed).toEqual(['v_thanks']);
  });

  it('7人全員の出来事が③の時系列に出る。人物名は架空で、役割名は文面に出ない', () => {
    const tl = buildTimeline(p, new Labels(p), w, p.personas.map((x) => x.id));
    for (const x of p.personas) expect(tl.some((e) => e.userId === x.id)).toBe(true);
    expect(tl.filter((e) => !e.minor).length).toBeGreaterThan(120);
    const bot = Object.values(w.users).flatMap((u) => u.talk.flatMap((t) => (t.side === 'bot' && t.bubble.kind === 'text' ? [t.bubble.text] : [])));
    for (const x of p.personas) expect(bot.some((t) => t.includes(x.name))).toBe(false);
  });

  it('自動再生で全員を最後まで流すと CLI と同じ結果になり、営業用HTMLを書き出せる', async () => {
    const d = playToEnd(p);
    expect(d.now).toBe(190 * DAY);
    const ww = worldOf(p, d);
    for (const u of Object.values(ww.users)) expect(u.deliveries).toEqual(user(w, d.session.users.find((x) => x.id === u.id)!.personaId).deliveries);
    await exportChecks(p);
  }, 30_000);
});
