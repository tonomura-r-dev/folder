// 提案資料から作った企業別デモ案件（projects/client-*/）の検証。
// 案件データは顧客の資料にもとづくのでリポジトリに入れていない。手元に無いときは、このテストは飛ばす。
// テストには顧客名を書かず、ID だけで確かめる。

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { InputEventBody, Project, UserState, WorldState } from '../packages/engine/src/index.ts';
import { bubbleButtons, personaCoverage, reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { allScripted, autoStep, dispatch, newDemo, seek, worldOf, type Demo } from '../packages/player/src/controller.ts';
import { journeyOf, journeyPane } from '../packages/player/src/journey.ts';
import { backstage } from '../packages/player/src/backstage.ts';
import { Labels } from '../packages/player/src/describe.ts';
import { buildTimeline, guideOf } from '../packages/player/src/timeline.ts';
import { phone } from '../packages/player/src/phone.ts';
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

describe.skipIf(!has('client-d-event'))('企業別デモ④（イベント運営・新規応募／勤務／再応募の循環）', () => {
  const p = has('client-d-event') ? loadProject('client-d-event') : (null as unknown as Project);
  const w = has('client-d-event') ? run(p, 22 * DAY) : (null as unknown as WorldState);
  const tags = (id: string) => Object.keys(user(w, id).tags);
  const cardTitles = (messageId: string) => p.messages.find((m) => m.id === messageId)!.bubbles.flatMap((b) => (b.kind === 'carousel' ? b.cards.map((c) => c.title) : []));
  const evTitle = (k: string) => p.webPages.find((x) => x.id === `page_ev_${k}`)!.title.replace('（デモ用サンプル）', '');
  const list = () => {
    const b = p.webPages.find((x) => x.id === 'page_jobs')!.blocks[2]!;
    if (b.kind !== 'list') throw new Error('list');
    return b;
  };
  /** 友だち追加して希望条件に答えたところ（はじめての人） */
  const newcomer = (area: string, venue: string, job: string, day: string) => {
    let d = dispatch(newDemo(p, 'p_a'), { type: 'addFriend', routeId: 'route_site' });
    for (const [q, c] of [['q_area', area], ['q_venue', venue], ['q_job', job], ['q_day', day]] as const) d = dispatch(d, { type: 'answer', surveyId: 'sv_pref', questionId: q, choiceIds: [`c_${q.slice(2)}_${c}`] });
    return d;
  };

  it('検証で重大エラー・警告が無く、全シナリオに誰かが入る', () => {
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
  });

  it('全CTA・カルーセルのボタン・リッチメニュー・ページのボタンが、押すと反応する。メニューは資料どおり6マスで「仕事を探す」が左上', () => {
    expect(deadButtons(p, 'p_a', 'route_site')).toEqual([]);
    const m = p.richMenus[0]!;
    expect(m.cells.map((c) => c.label).slice(0, 5)).toEqual(['仕事を探す', '勤務の案内', '希望条件を変える', 'よくある質問', 'お問い合わせ']);
    expect(m.cells[5]!.label).toBe(`${p.brand.companyName.replace(/^株式会社/, '')}とは`);
    expect(m.cells[0]!.area).toEqual({ col: 0, row: 0 });
    expect(m.cells[0]!.emphasis).toBe(true);
  });

  it('LINEへの入口は3つ（サイトのボタン・離脱防止・サンクスLINE誘導）が別々の人の行動として流れる', () => {
    expect(user(w, 'p_a').route).toBe('route_site');
    expect(user(w, 'p_b').route).toBe('route_popup');
    expect(user(w, 'p_c').route).toBe('route_thanks');
    // サンクスLINE誘導：先に応募してから友だちに。応募フォームの内容が入力欄に入り、本人が送ってから面接の案内が届く
    const c = user(w, 'p_c');
    expect(c.cvs['cv_apply']!.at).toBeLessThan(c.friendAddedAt!);
    expect(c.talk.find((t) => t.side === 'user')?.bubble).toEqual({ kind: 'text', text: expect.stringContaining('【応募しました】') });
    // 追加済みの人がサンクスLINE誘導を押すと、友だち追加ではなくトークを開くだけ（2回目の友だち追加にならない）
    let d = dispatch(newDemo(p, 'p_a'), { type: 'addFriend', routeId: 'route_site' });
    d = dispatch(d, { type: 'visitPage', pageId: 'page_apply_thanks_e1' });
    d = dispatch(d, { type: 'tapLineCta', pageId: 'page_apply_thanks_e1' });
    expect(me(p, d).surface).toEqual({ kind: 'talk' });
    expect(me(p, d).route).toBe('route_site');
  });

  it('求人一覧：LINEから開くと希望条件で絞り込んだ状態。エリア・会場・業務・勤務日で絞れ、締切を過ぎた募集は出ない', () => {
    expect(list().facets.map((f) => f.label)).toEqual(['エリア', '会場', '業務', '勤務日']);
    let d = newcomer('yokohama', 'any', 'guide', 'weekend');
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_main', cellId: 'cell_search' });
    const v = { project: p, world: worldOf(p, d), demo: d, ui: { menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' } as never;
    const html = (phone as (v: never, u: UserState) => string)(v, me(p, d));
    for (const k of ['e1', 'e3', 'e6']) expect(html).toContain(evTitle(k));
    expect(html).not.toContain(evTitle('e4'));
    // 締切（前日18:00）を過ぎた e1 は、あとで開くと出ない
    d = seek(d, at(6, 9));
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_main', cellId: 'cell_search' });
    const later = (phone as (v: never, u: UserState) => string)({ ...(v as object), world: worldOf(p, d), demo: d } as never, me(p, d));
    expect(later).not.toContain(evTitle('e1'));
  });

  it('はじめての人は応募フォーム、登録済みの人は勤務希望フォームへ。応募完了・登録完了・勤務希望・勤務確定は別の状態で、送信だけでは確定しない', () => {
    let d = newcomer('yokohama', 'any', 'guide', 'weekend');
    d = dispatch(d, { type: 'visitPage', pageId: 'page_ev_e3' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_ev_e3', blockIndex: 4 });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_apply_e3' });
    d = dispatch(d, { type: 'submitForm', pageId: 'page_apply_e3', values: { f_event: 'x' } });
    let u = me(p, d);
    expect(Object.keys(u.cvs)).toEqual(['cv_apply']);
    expect('ap_e3' in u.tags && !('fx_e3' in u.tags) && !('st_staff' in u.tags)).toBe(true);
    // 登録の前に別のイベントを押すと、応募フォームは出さずに案内だけ
    d = dispatch(d, { type: 'visitPage', pageId: 'page_ev_e4' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_ev_e4', blockIndex: 4 });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_pending' });
    d = dispatch(d, { type: 'cv', cvId: 'st_register' });
    expect('st_staff' in me(p, d).tags && !('fx_e3' in me(p, d).tags)).toBe(true);
    // 登録済みになると、同じボタンから勤務希望フォームへ。送っても勤務は確定しない
    d = dispatch(d, { type: 'visitPage', pageId: 'page_ev_e6' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_ev_e6', blockIndex: 4 });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_wish_e6' });
    d = dispatch(d, { type: 'submitForm', pageId: 'page_wish_e6', values: { f_event: 'x' } });
    u = me(p, d);
    expect(u.cvs['cv_reapply']!.count).toBe(1);
    expect(['wi_e6', 'fx_e3', 'fx_e6'].map((t) => t in u.tags)).toEqual([true, false, false]);
    // 運営側が勤務確定を登録したときだけ確定（開催の早い e3 から）
    d = dispatch(d, { type: 'cv', cvId: 'st_fix' });
    expect(['fx_e3', 'fx_e6'].map((t) => t in me(p, d).tags)).toEqual([true, false]);
    // 確定ずみのイベントをもう一度押しても、二重に応募しない
    d = dispatch(d, { type: 'visitPage', pageId: 'page_ev_e3' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_ev_e3', blockIndex: 4 });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_already_fixed' });
  });

  it('勤務前の案内は、そのイベントの勤務が確定した人にだけ届く', () => {
    const got = (k: string) => Object.values(w.users).filter((u) => u.deliveries.some((d) => d.messageId === `msg_rm_${k}`)).map((u) => u.id).sort();
    expect(got('e3')).toEqual(['p_d', 'p_f']);
    expect(got('e1')).toEqual(['p_a']);
    expect(got('e4')).toEqual(['p_c']);
    for (const u of Object.values(w.users))
      for (const d of u.deliveries.filter((x) => x.messageId.startsWith('msg_rm_') || x.messageId.startsWith('msg_day_'))) {
        const k = d.messageId.replace(/^msg_(rm|day)_/, '');
        expect(`fx_${k}` in u.tags).toBe(true);
      }
  });

  it('「前回のご勤務、ありがとう」は、勤務完了を登録した人にだけ。勤務の案内を送っただけの人（岡田）には送らず、勤務ありにもしない', () => {
    const thanks = new Set(p.messages.filter((m) => m.bubbles.some((b) => (b.kind === 'text' || b.kind === 'buttons') && /ご勤務、ありがとう/.test(b.text))).map((m) => m.id));
    for (const u of Object.values(w.users))
      for (const d of u.deliveries.filter((x) => thanks.has(x.messageId))) expect(u.behaviors.some((b) => b.type === 'cv' && b.detail['cvId'] === 'st_done' && b.at < d.at)).toBe(true);
    expect(tags('p_b')).toContain('fx_e7');
    expect(tags('p_b')).not.toContain('st_worked');
    expect(ids(w, 'p_b').some((m) => thanks.has(m))).toBe(false);
  });

  it('次回の募集は希望条件で出し分け：静岡・平日の人に横浜の土日の募集は届かない／気になる会場を選んだ人には、その会場の募集', () => {
    const e = ids(w, 'p_e').filter((m) => m.startsWith('msg_bc') || m.startsWith('msg_aw'));
    expect(e).toEqual(['msg_bc1_a_shizuoka_weekday_guide', 'msg_aw_e2_a_shizuoka_weekday_guide']);
    expect(cardTitles('msg_bc1_a_shizuoka_weekday_guide')).toEqual([evTitle('e2'), evTitle('e8')]);
    const skipped = w.decisions.filter((d) => d.userId === 'p_e' && d.kind === 'branch' && d.outcome === 'skipped' && d.ref === 'c_bc1_a_yokohama_weekend_guide');
    expect(skipped).toHaveLength(1);
    expect(ids(w, 'p_d')).toContain('msg_bc0_v_nissan_weekend_security');
    // 希望条件を変えた人は、変えた後の希望で届く（湘南 → Kアリーナ横浜）
    expect(ids(w, 'p_g')).toContain('msg_bc1_v_karena_weekend_sales');
    expect(tags('p_g')).not.toContain('area_shonan');
    // 配信日時のあとで友だちになった人には、その回の配信を送らない
    expect(user(w, 'p_e').trail.some((t) => t.nodeId === 'n_bc0_late')).toBe(true);
  });

  it('企画配信・勤務後の先行案内は、エリア（または気になる会場）・曜日に加えて、希望の業務も条件にする', () => {
    const jobs = ['guide', 'sales', 'security', 'setup'];
    const cases = p.flows
      .filter((f) => f.id.startsWith('flow_bc') || f.id === 'flow_done')
      .flatMap((f) => f.nodes.flatMap((n) => (n.type === 'branch' && /^n_(bc\d|aw_e\d)_b[va]$/.test(n.id) ? n.cases : [])));
    expect(cases.length).toBeGreaterThan(0);
    for (const c of cases) {
      const tagsIn = JSON.stringify(c.when);
      expect(jobs.filter((j) => tagsIn.includes(`"job_${j}"`))).toHaveLength(1);
      expect(/"(area|venue)_/.test(tagsIn) && /"day_/.test(tagsIn)).toBe(true);
    }
    // 届いた募集のイベントは、受け取った人の希望の業務を必ず含む（業務を変えていない6人）
    const jobOf = (k: string) => p.webPages.find((x) => x.id === `page_ev_${k}`)!.blocks.flatMap((b) => (b.kind === 'text' ? [b.text] : [])).join('').match(/業務：(.+)/)![1]!;
    const label: Record<string, string> = { guide: '案内・誘導', sales: '販売', security: '警備', setup: '設営・撤去' };
    for (const id of ['p_a', 'p_b', 'p_c', 'p_d', 'p_e', 'p_f']) {
      const job = Object.keys(user(w, id).tags).find((t) => t.startsWith('job_'))!.replace('job_', '');
      for (const m of ids(w, id).filter((x) => /^msg_(bc|aw)/.test(x)))
        for (const b of p.messages.find((x) => x.id === m)!.bubbles)
          if (b.kind === 'carousel') for (const c of b.cards) expect(jobOf(c.id.slice(-2))).toContain(label[job]!);
    }
    // 販売希望の人（前田）に、案内・誘導だけのイベントは届かない
    const c = ids(w, 'p_c').filter((x) => /^msg_(bc|aw)/.test(x));
    expect(c.flatMap((m) => cardTitles(m)).every((t) => t.includes('販売'))).toBe(true);
  });

  it('面接の予約：選んだ日時を記録 → 完了画面の「LINEに戻る」→ 予約完了の確認が届いている。二重に届かず、登録済みにもならない', () => {
    const slot = p.webPages.find((x) => x.id === 'page_interview')!.form!.fields.find((f) => f.id === 'f_slot')!.options![1]!;
    let d = newcomer('yokohama', 'any', 'guide', 'weekend');
    d = dispatch(d, { type: 'visitPage', pageId: 'page_ev_e3' });
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_ev_e3', blockIndex: 4 });
    d = dispatch(d, { type: 'submitForm', pageId: 'page_apply_e3', values: { f_event: 'x' } });
    d = dispatch(d, { type: 'closePage', pageId: 'page_apply_thanks_e3' });
    d = dispatch(d, { type: 'sendText', text: '応募しました' });
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_r_applied_iv' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_interview' });
    d = dispatch(d, { type: 'submitForm', pageId: 'page_interview', values: { f_name: 'デモ', f_slot: slot } });
    let u = me(p, d);
    expect(u.surface).toEqual({ kind: 'page', pageId: 'page_interview_thanks' });
    expect(['iv_booked', 'iv_2', 'iv_1', 'st_staff'].map((t) => t in u.tags)).toEqual([true, true, false, false]);
    expect(Object.keys(u.cvs)).toEqual(['cv_apply']);
    // 完了画面の「LINEに戻る」でトークへ。確認には選んだ日時と、今後の案内（前日のリマインド・登録は面接のあと）
    const back = p.webPages.find((x) => x.id === 'page_interview_thanks')!.blocks.findIndex((b) => b.kind === 'button' && b.label === 'LINEに戻る');
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_interview_thanks', blockIndex: back });
    u = me(p, d);
    expect(u.surface).toEqual({ kind: 'talk' });
    const last = u.talk.at(-1)!;
    expect(last.side === 'bot' && last.bubble.kind === 'text' ? last.bubble.text : '').toMatch(new RegExp(`^【面接の予約が完了しました】\\n日時：${slot.replace(/[()（）]/g, '.')}[\\s\\S]*前日18:00に[\\s\\S]*スタッフ登録[\\s\\S]*予約システムとLINEの連携を想定`));
    // もう一度「面接の日時を選ぶ」を押しても予約フォームは出ず、送り直しても確認は1回だけ
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_r_applied_iv' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_iv_booked' });
    d = dispatch(d, { type: 'submitForm', pageId: 'page_interview', values: { f_name: 'デモ', f_slot: slot } });
    expect(me(p, d).deliveries.filter((x) => x.messageId.startsWith('msg_iv_ok_')).map((x) => x.messageId)).toEqual(['msg_iv_ok_2']);
    // 面接の前日 18:00 にリマインド（日時つき）
    d = seek(d, at(2, 18, 1));
    expect(me(p, d).deliveries.filter((x) => x.messageId.startsWith('msg_iv_')).map((x) => [x.messageId, x.at])).toEqual([
      ['msg_iv_ok_2', me(p, d).deliveries.find((x) => x.messageId === 'msg_iv_ok_2')!.at],
      ['msg_iv_rm_2', at(2, 18)],
    ]);
  });

  it('台本の面接：予約→確認→前日のリマインド→（面接のあと）スタッフ登録の順。確認・リマインドは1人1回ずつ', () => {
    const slots = p.webPages.find((x) => x.id === 'page_interview')!.form!.fields.find((f) => f.id === 'f_slot')!.options!;
    const ivAt = (label: string) => {
      const m = label.match(/11月(\d+)日.*?(\d+):(\d+)/)!;
      return at(Number(m[1]) - 2, Number(m[2]), Number(m[3]));
    };
    for (const id of ['p_a', 'p_b', 'p_c']) {
      const u = user(w, id);
      const sent = u.behaviors.find((b) => b.type === 'submitForm' && b.detail['pageId'] === 'page_interview')!;
      const n = slots.indexOf(sent.detail['values.f_slot']!) + 1;
      expect(n).toBeGreaterThan(0);
      expect(u.tags[`iv_${n}`]).toBeDefined();
      const iv = u.deliveries.filter((x) => x.messageId.startsWith('msg_iv_'));
      expect(iv.map((x) => x.messageId)).toEqual([`msg_iv_ok_${n}`, `msg_iv_rm_${n}`]);
      expect(iv[0]!.at).toBe(sent.at);
      expect(iv[1]!.at).toBe(at(Math.floor(ivAt(slots[n - 1]!) / DAY) - 1, 18));
      // スタッフ登録は面接の時刻より後。LINEに戻るボタンでトークに戻っている
      expect(u.cvs['st_register']!.at).toBeGreaterThan(ivAt(slots[n - 1]!));
      expect(u.behaviors.some((b) => b.type === 'tapPageButton' && b.detail['pageId'] === 'page_interview_thanks' && b.at === sent.at + 1)).toBe(true);
    }
    // ③の時系列にも、面接の予約（日時つき）・予約完了の確認・前日のリマインドが出る
    const tl = buildTimeline(p, new Labels(p), w, ['p_a']);
    const texts = tl.flatMap((e) => [...e.front, ...e.back]);
    expect(texts.some((t) => t.includes('面接の日時を選ぶ') && t.includes(slots[0]!))).toBe(true);
    expect(tl.some((e) => e.at === at(1, 18) && JSON.stringify(e).includes('面接の前日のリマインド'))).toBe(true);
  });

  it('資料用16:9の③：応募→面接の予約→登録→勤務確定→勤務前の案内→勤務→次回募集→再応募は「⋮ ○段」にまとめず、いつも出す', () => {
    const keys = p.journey!.filter((s) => s.key).map((s) => s.label);
    expect(keys).toEqual(['新規応募', '面接の予約', '面接・スタッフ登録', '勤務確定', '勤務前の案内（前日・当日）', '勤務完了（勤務履歴）', '次回イベントの募集（勤務した方へ）', '再応募（次のイベントへ）']);
    for (const [id, when] of [['p_a', at(0, 10, 3)], ['p_b', at(3, 20)], ['p_a', at(21, 23)], ['p_g', at(12)]] as const) {
      let d = seek(allScripted(p), when);
      d = { ...d, activeUserId: d.session.users.find((u) => u.personaId === id)!.id };
      const v = { project: p, world: worldOf(p, d), demo: d, ui: { menuOpen: true, picked: [], mode: 'slide3', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' } as never;
      const html = journeyPane(v, me(p, d), true);
      const from = html.indexOf('<ol class="lsp-strip');
      expect(from).toBeGreaterThan(-1);
      const strip = html.slice(from, html.indexOf('</ol>', from));
      const names = [...strip.matchAll(/<li class="([^"]*)">(.*?)<\/li>/g)].map((m) => (m[1]!.includes('gap') ? '⋮' : m[2]!.replace(/<[^>]+>/g, '')));
      // 主要な工程は全部、並び順のまま、間に「⋮」を挟まずに出る
      const at0 = names.findIndex((n) => n.includes('新規応募'));
      const seq = names.slice(at0, at0 + keys.length).map((n) => keys.find((k) => n.includes(k)) ?? n);
      expect(seq).toEqual(keys);
      expect(strip).toContain('「勤務確定」へ戻る');
    }
  });

  it('ブロックした人には配信しない（企画配信・ステップ配信が止まる）', () => {
    const g = user(w, 'p_g');
    expect(g.friendState).toBe('blocked');
    expect(g.deliveries.every((d) => d.at < at(9, 20))).toBe(true);
    expect(['flow_bc2', 'flow_bc3', 'flow_step'].map((f) => g.flows[f]!.status)).toEqual(['stopped', 'stopped', 'stopped']);
  });

  it('カルーセル → 詳細 → 勤務希望フォーム → CVの記録まで、同じイベント。新規応募と再応募は別のCV', () => {
    const a = user(w, 'p_a');
    const sent = a.behaviors.filter((b) => b.type === 'submitForm' && /^page_(apply|wish)_/.test(b.detail['pageId'] ?? ''));
    expect(sent.map((b) => b.detail['pageId'])).toEqual(['page_apply_e1', 'page_wish_e6']);
    expect(sent[1]!.detail['values.f_event']).toBe(p.webPages.find((x) => x.id === 'page_wish_e6')!.form!.fields[0]!.options![0]);
    expect(a.cvs['cv_apply']!.count).toBe(1);
    expect(a.cvs['cv_reapply']!.count).toBe(1);
    const back = buildTimeline(p, new Labels(p), w, ['p_a']).flatMap((e) => e.back);
    expect(back.some((t) => t.startsWith('CV「再応募（勤務希望）」を記録（勤務希望を出すイベント：11月21日（土） 横浜アリーナ'))).toBe(true);
    // 勤務 → 次回の募集 → 再応募 → 勤務確定 → 勤務前の案内 の循環（2周目）
    expect(ids(w, 'p_a').filter((m) => /^msg_(fix|rm|day|aw)_/.test(m))).toEqual(['msg_fix_e1', 'msg_rm_e1', 'msg_day_e1', 'msg_aw_e1_a_yokohama_weekend_guide', 'msg_fix_e6', 'msg_rm_e6', 'msg_day_e6', 'msg_aw_e6_a_yokohama_weekend_guide']);
  });

  it('③：GOALは再応募で、勤務確定へ戻るくり返しがある', () => {
    const rows = journeyOf(p, w, user(w, 'p_a'));
    const goal = rows.find((r) => r.isGoal)!;
    expect(goal.step.id).toBe('j_reapply');
    expect(goal.step.loopTo).toBe('j_fix');
    expect(rows.filter((r) => r.status === 'done' || r.status === 'current').map((r) => r.step.id)).toEqual(p.journey!.map((s) => s.id));
  });

  it('自動再生で全員を最後まで流すと CLI と同じ結果になり、営業用HTMLを書き出せる', async () => {
    const d = playToEnd(p);
    expect(d.now).toBe(22 * DAY);
    const ww = worldOf(p, d);
    for (const u of Object.values(ww.users)) expect(u.deliveries).toEqual(user(w, d.session.users.find((x) => x.id === u.id)!.personaId).deliveries);
    await exportChecks(p);
  }, 30_000);
});

describe.skipIf(!has('client-e-juku'))('企業別デモ⑤（学習塾・業界汎用：診断→資料請求→14日間→体験予約）', () => {
  const p = has('client-e-juku') ? loadProject('client-e-juku') : (null as unknown as Project);
  const w = has('client-e-juku') ? run(p, 29 * DAY) : (null as unknown as WorldState);
  const tags = (id: string) => Object.keys(user(w, id).tags);
  const slot = (n: number) => p.webPages.find((x) => x.id === 'page_rsv')!.form!.fields.find((f) => f.id === 'f_slot')!.options![n - 1]!;
  const SLOT_TAGS = (p ? p.tags : []).filter((t) => /^rsv_s\d+$/.test(t.id)).map((t) => t.id);
  const bot = (u: UserState) => u.talk.flatMap((t) => (t.side === 'bot' && (t.bubble.kind === 'text' || t.bubble.kind === 'buttons') ? [t.bubble.text] : []));
  const remind = (id: string) => user(w, id).deliveries.filter((d) => d.flowId === 'flow_rsv_remind').map((d) => [d.messageId, d.at] as const);
  /** 手で操作する保護者：ポップアップから友だち追加して、4問に答える（null なら答えない） */
  const parent = (ans: [string, string, string, string] | null) => {
    let d = dispatch(newDemo(p, 'p1'), { type: 'addFriend', routeId: 'route_popup' });
    if (!ans) return d;
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_w_popup_diag' });
    ['q_grade', 'q_worry', 'q_type', 'q_when'].forEach((q, i) => (d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: q, choiceIds: [`c_${q.slice(2)}_${ans[i]}`] })));
    return d;
  };
  const book = (d: Demo, n: number, pageId = 'page_rsv', type = '無料体験授業') => dispatch(d, { type: 'submitForm', pageId, values: { f_type: type, f_grade: '中1・2', f_slot: slot(n), f_people: '1名' } });
  /** 予約の完了画面から、LINEのトークに戻る */
  const backToTalk = (d: Demo) => dispatch(d, { type: 'tapPageButton', pageId: 'page_rsv_done', blockIndex: 2 });

  it('検証で重大エラー・警告が無く、全シナリオに誰かが入る', () => {
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
  });

  it('全ボタン・メニュー・ページのボタンが反応する。リッチメニューは資料P46どおり2タブ×6マスで、右下は予約／日程の確認・変更', () => {
    expect(deadButtons(p, 'p1', 'route_popup')).toEqual([]);
    const [first, reserved] = p.richMenus;
    const body = (m: (typeof p.richMenus)[number]) => m.cells.filter((c) => c.area.row > 0).map((c) => c.label);
    expect(body(first!)).toEqual(['30秒ぴったり診断', 'コース・料金', '教室・講師紹介', '資料を受け取る', '友だちと一緒に体験', '無料体験・教室見学の予約']);
    expect(body(reserved!)).toEqual(['体験当日の流れ', 'アクセス・持ち物', '担当講師の紹介', '保護者の声', 'LINEで質問', '日程の確認・変更']);
    const br = (m: (typeof p.richMenus)[number]) => m.cells.find((c) => c.area.row === 2 && c.area.col === 4)!;
    expect([br(first!).label, br(first!).emphasis, br(reserved!).label, br(reserved!).emphasis]).toEqual(['無料体験・教室見学の予約', true, '日程の確認・変更', true]);
    // 予約すると「体験予約済みの方」、キャンセル・体験の日時が過ぎると「はじめての方」へ
    expect(user(w, 'p4').richMenuId).toBe('rm_first');
    expect(me(p, book(parent(null), 3)).richMenuId).toBe('rm_reserved');
  });

  it('入口：ポップアップは資料P41の文面、あいさつは資料P44（教室長・週1回・営業電話なし・対応時間）。QR・紹介・完了画面からも入れる', () => {
    const pop = p.webPages.find((x) => x.id === 'page_method')!.exitPopup!;
    expect([pop.title, pop.cta.label]).toEqual(['まだ、塾を決めなくて大丈夫です。', 'LINEで資料と診断を受け取る']);
    const hello = bot(me(p, parent(null)))[0]!;
    for (const s of ['教室長、佐藤', '週1回', '営業のお電話は一切しません', '平日14:00〜21:00', 'デモ用の仮設定']) expect(hello).toContain(s);
    expect(['p1', 'p4', 'p5', 'p7'].map((id) => user(w, id).route)).toEqual(['route_popup', 'route_qr', 'route_referral', 'route_thanks']);
  });

  it('30秒診断：資料P45の4問を1問ずつ。答えた瞬間にタグ、4問そろって結果（性格で学び方の候補・断定しない）', () => {
    expect(p.surveys[0]!.questions.map((q) => [q.text, q.choices.map((c) => c.label)])).toEqual([
      ['Q1 お子さまの学年は？', ['小1〜4', '小5・6', '中1・2', '中3', '高1・2', '高3']],
      ['Q2 いちばん近いお悩みは？', ['成績が下がってきた', 'やる気が出ない', '受験対策を始めたい', '授業についていけない']],
      ['Q3 お子さまの性格に近いのは？', ['競争があると燃える', 'マイペースにコツコツ', '質問するのが苦手', '集中が続かない']],
      ['Q4 塾はいつから？', ['今すぐ', '今学期中に', '次の学年から', '情報収集だけ']],
    ]);
    let d = dispatch(parent(null), { type: 'tapButton', buttonId: 'btn_w_popup_diag' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_grade', choiceIds: ['c_grade_m3'] });
    expect(Object.keys(me(p, d).tags)).toEqual(['grade_m3']);
    expect(me(p, d).surface).toEqual({ kind: 'survey', surveyId: 'sv_diag', questionId: 'q_worry' });
    const rec = (ty: string) => Object.keys(me(p, parent(['m12', 'down', ty, 'term'])).tags).filter((t) => t.startsWith('rec_')).sort();
    expect([rec('compete'), rec('shy'), rec('mypace'), rec('focus')]).toEqual([['rec_group'], ['rec_individual'], ['rec_individual', 'rec_video'], ['rec_individual']]);
    const res = bot(me(p, parent(['m3', 'exam', 'shy', 'now']))).find((t) => t.startsWith('【30秒診断の結果】'))!;
    for (const s of ['中3・「受験対策を始めたい」・「質問するのが苦手」', '① お子さまの特徴', '② おすすめの学び方（参考）', '個別指導が候補', '③ 同じような悩み', '学び方を決めるものではありません']) expect(res).toContain(s);
    const next = p.messages.find((m) => m.id === 'msg_res_shy')!.bubbles[1]!;
    expect(next.kind === 'buttons' ? next.buttons.map((b) => b.label) : []).toEqual(['④ コース・料金の資料を受け取る', '⑤ 無料体験を予約する', '診断結果を友だちに送る']);
    // やり直すと候補が入れ替わる
    let again = dispatch(parent(['m12', 'down', 'compete', 'term']), { type: 'tapMenu', richMenuId: 'rm_first', cellId: 'cell_diag' });
    for (const [q, c] of [['q_grade', 'c_grade_m12'], ['q_worry', 'c_worry_down'], ['q_type', 'c_type_shy'], ['q_when', 'c_when_term']] as const) again = dispatch(again, { type: 'answer', surveyId: 'sv_diag', questionId: q, choiceIds: [c] });
    expect(Object.keys(me(p, again).tags).filter((t) => /^(rec_|type_)/.test(t)).sort()).toEqual(['rec_individual', 'type_shy']);
  });

  it('診断に答えていない人：1日後に一度だけフォロー。回答が無いのに診断結果のタグは付けない', () => {
    expect(Object.values(w.users).filter((u) => u.deliveries.some((d) => d.messageId === 'msg_s1_nodiag')).map((u) => u.id)).toEqual(['p4']);
    expect(tags('p4').filter((t) => /^(rec_|diag_done|grade_|worry_|type_|when_)/.test(t))).toEqual([]);
    // 途中でやめた人も同じ
    let d = dispatch(parent(null), { type: 'tapButton', buttonId: 'btn_w_popup_diag' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_diag', questionId: 'q_grade', choiceIds: ['c_grade_m3'] });
    d = seek(d, at(3));
    expect(['diag_done', 'rec_group', 'rec_individual', 'rec_video'].some((t) => t in me(p, d).tags)).toBe(false);
  });

  it('CV①：送信が完了したときだけ記録。資料はLINEに届き、請求ずみの人には同じ資料を請求させない', () => {
    let d = dispatch(parent(['m12', 'down', 'compete', 'term']), { type: 'tapButton', buttonId: 'btn_res_compete_doc' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_doc_req' });
    expect(me(p, d).cvs['cv_doc']).toBeUndefined();
    expect(p.webPages.find((x) => x.id === 'page_doc_req')!.form!.confirm).toBeDefined();
    d = dispatch(d, { type: 'submitForm', pageId: 'page_doc_req', values: { f_doc_grade: '中1・2', f_doc_style: '集団指導' } });
    const u = me(p, d);
    expect([u.cvs['cv_doc']!.count, 'doc_req' in u.tags]).toEqual([1, true]);
    expect(bot(u).at(-1)).toContain('おすすめの指導形式（診断より・参考）：集団指導');
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_first', cellId: 'cell_doc' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_doc_already' });
    expect(me(p, d).cvs['cv_doc']!.count).toBe(1);
    const doc = p.webPages.find((x) => x.id === 'page_doc_view')!.blocks.flatMap((b) => (b.kind === 'text' ? [b.text] : [])).join('\n');
    for (const s of ['おすすめの指導形式', '対象学年・コース', '月謝の目安', 'デモ用の料金例', '講習費を含めた年間', '教室の特徴', '体験授業のご案内']) expect(doc).toContain(s);
    expect(doc).not.toMatch(/[0-9０-９][0-9,０-９]*円/);
  });

  it('14日間の配信：3・5・7・10・14日後。悩み・資料請求・性格で出し分ける。予約中・来塾ずみの人には送らず、「情報収集だけ」の人には予約の案内を入れない', () => {
    // 友だち追加の日から何日後か（暦の日で数える）
    const step = (id: string) => user(w, id).deliveries.filter((d) => d.flowId === 'flow_step').map((d) => [d.messageId, Math.floor(d.at / DAY) - Math.floor(user(w, id).friendAddedAt! / DAY)]);
    expect(step('p1')).toEqual([['msg_s3_down', 3], ['msg_s5_have', 5], ['msg_s7_compete', 7]]);
    expect(step('p2')).toEqual([['msg_s3_motivation', 3], ['msg_s5_ask', 5], ['msg_s7_info', 7], ['msg_s10_info', 10], ['msg_s14_info', 14]]);
    expect(step('p3')).toEqual([['msg_s3_down', 3], ['msg_s5_have', 5], ['msg_s7_mypace', 7], ['msg_s10_faq', 10], ['msg_s14', 14]]);
    // 予約のあとは声がけを減らす：14日間の残りは届かない
    expect([step('p5'), step('p6'), step('p7')]).toEqual([[], [], []]);
    // キャンセルすると、残りの配信に戻る
    expect(step('p4').map((x) => x[0])).toEqual(['msg_s1_nodiag', 'msg_s10_faq', 'msg_s14']);
    // 情報収集だけの人に、配信（ステップ・週1回・企画）で体験予約のボタンを送らない
    const toTrial = new Set(p.messages.filter((m) => m.bubbles.some((b) => bubbleButtons(b).some((x) => x.action.kind === 'startFlow' && x.action.flowId === 'flow_rsv_open'))).map((m) => m.id));
    expect(user(w, 'p2').deliveries.filter((d) => /^flow_(step|weekly|bc_)/.test(d.flowId ?? '') && toTrial.has(d.messageId))).toEqual([]);
    // 送らなかった理由（分岐の判定）が②に残る
    expect(w.decisions.some((x) => x.userId === 'p7' && x.kind === 'branch' && x.outcome === 'taken' && x.ref === 'c_s3_rsv')).toBe(true);
  });

  it('CV②：確認画面のあとの確定で1回だけ。お手続き・日時・人数が、確認の文面・②のタグ・③の時系列まで同じ。当日の流れはLINEの文面に書く', () => {
    const page = p.webPages.find((x) => x.id === 'page_rsv')!;
    expect([page.form!.confirm?.submitLabel, page.form!.fields.map((f) => f.label), page.form!.fields[0]!.options]).toEqual(['この内容で予約する', ['お手続き', '学年', '希望日時（空き枠）', '人数'], ['無料体験授業', '教室見学', '入会説明']]);
    expect(page.blocks[0]).toMatchObject({ kind: 'heading', text: '○○塾 △△教室を予約' });
    expect(p.webPages.some((x) => x.id === 'page_dayflow')).toBe(false);
    const u = user(w, 'p1');
    expect(u.cvs['cv_trial']).toEqual({ at: at(7, 20, 13), count: 1 });
    const ok = bot(u).find((t) => t.startsWith('【ご予約を受け付けました】'))!;
    for (const s of ['お手続き：無料体験授業', `日時：${slot(3)}`, '教室：○○塾 △△教室', '学年：中1・2', '人数：1名', '▼ 当日の流れ\n① 受付で', '④ 保護者さまと']) expect(ok).toContain(s);
    expect(p.tags.find((t) => t.id === 'rsv_s3')!.label).toBe(`予約日時：${slot(3)}`);
    expect(tags('p1')).toContain('rsv_s3');
    expect(JSON.stringify(buildTimeline(p, new Labels(p), w, ['p1']))).toContain(slot(3));
    expect(bot(user(w, 'p5')).find((t) => t.startsWith('【ご予約を受け付けました】'))).toContain('人数：2名（お友だちと一緒）');
    // 資料請求をせずに、診断結果から直接予約もできる
    expect([user(w, 'p7').cvs['cv_doc'], user(w, 'p7').cvs['cv_trial']!.count]).toEqual([undefined, 1]);
    // メニュー「体験当日の流れ」も、ページを開かずトークに文面で届く
    expect(ids(w, 'p5')).toContain('msg_dayflow');
  });

  it('入会説明の予約もCV②で1回。確認は入会説明の当日の流れ。入会は教室の登録だけ。再予約は同じ画面で、裏側でタグを付ける', () => {
    const u = me(p, backToTalk(book(parent(null), 3, 'page_rsv', '入会説明')));
    expect([u.cvs['cv_trial']?.count, 'rsv_t_enroll' in u.tags, 'enrolled' in u.tags, u.deliveries.at(-1)!.messageId]).toEqual([1, true, false, 'msg_rsv_ok_enroll']);
    expect(bot(u).at(-1)).toContain('③ コース・料金・講習費のご説明');
    const [a, b] = [p.webPages.find((x) => x.id === 'page_rsv')!, p.webPages.find((x) => x.id === 'page_rsv_re')!];
    expect([b.title, b.blocks, b.form!.fields]).toEqual([a.title, a.blocks, a.form!.fields]);
  });

  it('予約後の案内：直後・前日18:00・翌日12:00（2日前・当日の案内はない）。予約していない人には送らない', () => {
    expect(remind('p1')).toEqual([['msg_rsv_ok', at(7, 20, 13)], ['msg_rsv_1d', at(11, 18)], ['msg_rsv_thanks', at(13, 12)]]);
    expect(remind('p7')).toEqual([['msg_rsv_ok', at(1, 13, 15)], ['msg_rsv_1d', at(7, 18)], ['msg_rsv_thanks', at(9, 12)]]);
    expect(p.messages.filter((m) => ['msg_rsv_2d', 'msg_rsv_2h'].includes(m.id))).toEqual([]);
    expect(remind('p2')).toEqual([]);
    const t = bot(user(w, 'p1')).find((x) => x.startsWith('【明日'))!;
    expect(t).toContain(`【明日 ${slot(3)}】無料体験授業のご確認です`);
    expect(t).toContain('（キャンセル料はありません）');
  });

  it('日程変更で前の日時の案内は止まり、キャンセルで予約者向けの案内が止まる。CV②は増えない', () => {
    expect(remind('p4')).toEqual([['msg_rsv_ok', at(2, 20, 33)], ['msg_rsv_chg', at(3, 12, 3)], ['msg_rsv_1d', at(11, 18)]]);
    const u = user(w, 'p4');
    expect(bot(u).find((x) => x.startsWith('【日程を変更しました】'))).toContain(`新しい日時：${slot(3)}`);
    expect(ids(w, 'p4')).toContain('msg_cancel_done');
    expect(u.cvs['cv_trial']!.count).toBe(1);
    expect(['rsv_booked', ...SLOT_TAGS].filter((t) => t in u.tags)).toEqual([]);
    expect(['rsv_changed', 'rsv_cancel'].every((t) => t in u.tags)).toBe(true);
  });

  it('当日までが短い予約：過ぎた時刻の案内（前日）はあとから送らない。来塾の登録が無ければ翌日は参加を前提にしない', () => {
    let d = seek(parent(['m12', 'down', 'compete', 'term']), at(4, 19));
    d = book(d, 1);
    const bookedAt = me(p, d).cvs['cv_trial']!.at;
    d = seek(d, at(6, 13));
    expect(me(p, d).deliveries.filter((x) => x.flowId === 'flow_rsv_remind').map((x) => [x.messageId, x.at])).toEqual([['msg_rsv_ok', bookedAt], ['msg_rsv_missed', at(6, 12)]]);
    const got = (m: string) => Object.values(w.users).filter((u) => u.deliveries.some((x) => x.messageId === m)).map((u) => u.id).sort();
    expect(got('msg_rsv_thanks')).toEqual(['p1', 'p5', 'p6', 'p7']);
    expect(got('msg_rsv_missed')).toEqual(['p3']);
    for (const u of Object.values(w.users)) for (const x of u.deliveries.filter((y) => y.messageId === 'msg_rsv_thanks')) expect(u.behaviors.some((b) => b.type === 'cv' && b.detail['cvId'] === 'st_visit' && b.at < x.at)).toBe(true);
  });

  it('日時を過ぎて来塾のタグが無いと、②の先頭に教室の操作（来塾タグ・来塾なし）を出す。押すまで来塾後には進まず、14日間の配信も止めたまま', () => {
    const view = (d: Demo) => ({ project: p, world: worldOf(p, d), demo: d, ui: { menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' }) as never;
    const html = (d: Demo) => backstage(view(d), me(p, d));
    const booked = backToTalk(book(parent(['m12', 'down', 'compete', 'term']), 1));
    expect(html(seek(booked, at(5, 13)))).not.toContain('lsp-ask');
    const past = seek(booked, at(5, 17));
    expect(['lsp-ask', '来塾タグを付ける（教室が登録）', 'data-cv="st_noshow"'].every((x) => html(past).includes(x))).toBe(true);
    expect(guideOf(view(past), me(p, past)).text).toContain('②の「来塾タグを付ける（教室が登録）」');
    const came = dispatch(past, { type: 'cv', cvId: 'st_visit' });
    expect([html(came).includes('lsp-ask'), 'visited' in me(p, came).tags]).toEqual([false, true]);
    const none = dispatch(past, { type: 'cv', cvId: 'st_noshow' });
    expect([html(none).includes('lsp-ask'), 'no_show' in me(p, none).tags, 'visited' in me(p, none).tags, html(none).includes('data-cv="st_noshow"')]).toEqual([false, true, false, false]);
    expect(me(p, seek(none, at(6, 13))).deliveries.at(-1)!.messageId).toBe('msg_rsv_missed');
    expect(tags('p3')).toContain('no_show');
    const rest = (d: Demo) => me(p, seek(d, at(8))).deliveries.filter((x) => x.flowId === 'flow_step').map((x) => x.messageId);
    expect([rest(past), rest(came), rest(none)]).toEqual([[], [], ['msg_s5_ask', 'msg_s7_compete']]);
  });

  it('来塾ずみの人に体験の予約を案内しない。予約中の人に企画配信を送らない。比較をすすめる文面は書かない', () => {
    const toRsv = (m: string) => (p.messages.find((x) => x.id === m)?.bubbles ?? []).flatMap((b) => bubbleButtons(b)).filter((b) => b.action.kind === 'startFlow' && b.action.flowId === 'flow_rsv_open');
    for (const id of ['p1', 'p5', 'p6']) {
      const u = user(w, id);
      const came = u.cvs['st_visit']!.at;
      expect(u.deliveries.filter((d) => d.at > came && d.flowId !== null).flatMap((d) => toRsv(d.messageId))).toEqual([]);
    }
    // 三郎は期末テストの企画（11/26）の時点で予約中
    expect(ids(w, 'p3').filter((m) => m.startsWith('msg_bc_test'))).toEqual([]);
    const all = p.messages.flatMap((m) => m.bubbles.flatMap((b) => (b.kind === 'text' || b.kind === 'buttons' ? [b.text] : [])));
    expect(all.filter((x) => /比べて決め|ほかの塾と比べ|他塾と比べ/.test(x))).toEqual([]);
  });

  it('再予約（キャンセルのあと）でも CV② は重ねない。ブロックした人には何も届かない', () => {
    let d = book(parent(['m12', 'down', 'compete', 'term']), 3);
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_cancel_yes' });
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_first', cellId: 'cell_rsv' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_rsv_re' });
    d = book(d, 4, 'page_rsv_re');
    const u = me(p, d);
    expect([u.cvs['cv_trial']!.count, 'rsv_booked' in u.tags, 'rsv_s4' in u.tags, 'rsv_s3' in u.tags, 'rsv_again' in u.tags, u.richMenuId]).toEqual([1, true, true, false, true, 'rm_reserved']);
    d = dispatch(d, { type: 'block' });
    const blockedAt = d.now;
    d = seek(d, at(28));
    expect(me(p, d).deliveries.filter((x) => x.at > blockedAt)).toEqual([]);
  });

  it('入会は教室が登録したときだけ。入会した人は未入会向けの配信から外れ、通塾中の方向けに切り替わる。GOALは体験予約', () => {
    const u = user(w, 'p7');
    const enrolledAt = u.cvs['st_enroll']!.at;
    expect(u.deliveries.filter((d) => d.at > enrolledAt && /^flow_(step|weekly)$/.test(d.flowId ?? ''))).toEqual([]);
    expect(ids(w, 'p7').filter((m) => m.startsWith('msg_bc_'))).toEqual(['msg_bc_test_member', 'msg_bc_w_member']);
    expect(Object.values(w.users).filter((x) => 'enrolled' in x.tags).map((x) => x.id)).toEqual(['p7']);
    expect(journeyOf(p, w, u).find((r) => r.isGoal)!.step.id).toBe('j_trial');
  });

  it('季節の企画配信：入会ずみ／体験ずみ・未入会／受験学年／情報収集／その他 を分ける。週1回の配信は14日間を終えて予約していない人だけ', () => {
    expect(['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7'].map((id) => ids(w, id).find((m) => m.startsWith('msg_bc_w_')))).toEqual(['msg_bc_w_visited', 'msg_bc_w_info', 'msg_bc_w_exam', 'msg_bc_w_all', 'msg_bc_w_visited', 'msg_bc_w_visited', 'msg_bc_w_member']);
    expect(Object.values(w.users).filter((u) => u.deliveries.some((d) => d.flowId === 'flow_weekly')).map((u) => u.id).sort()).toEqual(['p2', 'p3', 'p4']);
  });

  it('友だち紹介：紹介した人には「送った」記録だけ（相手の追加・予約にはしない）。紹介から来た人は経路が残り、2名で予約', () => {
    expect(tags('p1')).toContain('ref_shared');
    expect(bot(user(w, 'p1')).join('\n')).toContain('紹介しただけでは、お友だちの友だち追加・予約にはなりません');
    expect(Object.keys(w.users)).toHaveLength(7);
    expect([user(w, 'p5').route, tags('p5').includes('rsv_two'), tags('p5').includes('diag_shared')]).toEqual(['route_referral', true, true]);
  });

  it('自動応答と有人対応：料金は自動応答、相談は教室スタッフ（②に「有人対応中」）', () => {
    expect(user(w, 'p2').deliveries.find((d) => d.messageId === 'msg_ar_fee')!.flowId).toBeNull();
    expect(tags('p6')).toContain('human_chat');
    expect(bot(user(w, 'p6')).join('\n')).toContain('（有人対応に切り替え）');
    expect(tags('p2')).not.toContain('human_chat');
  });

  it('戻る：前の時点に戻すと、タグ・CV・予約・届いた文面もその時点に戻る', () => {
    let d = seek(allScripted(p), at(8));
    const id = d.session.users.find((u) => u.personaId === 'p1')!.id;
    expect(worldOf(p, d).users[id]!.cvs['cv_trial']).toBeDefined();
    d = seek(d, at(7, 20, 12));
    const before = worldOf(p, d).users[id]!;
    expect([before.cvs['cv_trial'], 'rsv_booked' in before.tags, before.deliveries.some((x) => x.messageId === 'msg_rsv_ok')]).toEqual([undefined, false, false]);
  });

  it('資料用16:9の③：LINE追加→診断→資料請求→ステップ配信→体験予約→予約後の案内→来塾を、この順に出す', () => {
    const keys = p.journey!.filter((s) => s.key).map((s) => s.label);
    expect(keys).toEqual(['LINE友だち追加', '30秒ぴったり診断（4問）', '資料請求（CV①）', '14日間のステップ配信', '来塾の予約（CV②）', '予約後の案内（直後・前日）', '体験来塾（教室が登録）']);
    for (const [pid, when] of [['p1', at(0, 10, 9)], ['p3', at(16)], ['p7', at(10)]] as const) {
      let d = seek(allScripted(p), when);
      d = { ...d, activeUserId: d.session.users.find((u) => u.personaId === pid)!.id };
      const v = { project: p, world: worldOf(p, d), demo: d, ui: { menuOpen: true, picked: [], mode: 'slide3', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' } as never;
      const html = journeyPane(v, me(p, d), true);
      const from = html.indexOf('<ol class="lsp-strip');
      const strip = html.slice(from, html.indexOf('</ol>', from));
      const names = [...strip.matchAll(/<li class="([^"]*)">(.*?)<\/li>/g)].map((m) => m[2]!.replace(/<[^>]+>/g, ''));
      const found = keys.map((k) => names.findIndex((n) => n.includes(k)));
      expect(found.every((i) => i >= 0)).toBe(true);
      expect([...found].sort((a, b) => a - b)).toEqual(found);
    }
  });

  it('自動再生で全員を最後まで流すと CLI と同じ結果になり、営業用HTMLを書き出せる', async () => {
    const d = playToEnd(p);
    expect(d.now).toBe(29 * DAY);
    const ww = worldOf(p, d);
    for (const u of Object.values(ww.users)) expect(u.deliveries).toEqual(user(w, d.session.users.find((x) => x.id === u.id)!.personaId).deliveries);
    await exportChecks(p);
  }, 30_000);
});

describe.skipIf(!has('client-f-gym'))('企業別デモ⑥（ジム・フィットネス・業界汎用：診断→体験予約→来店→入会→入会後14日間→CV③）', () => {
  const p = has('client-f-gym') ? loadProject('client-f-gym') : (null as unknown as Project);
  const w = has('client-f-gym') ? run(p, 37 * DAY) : (null as unknown as WorldState);
  const tags = (id: string) => Object.keys(user(w, id).tags);
  const slot = (n: number) => p.webPages.find((x) => x.id === 'page_rsv')!.form!.fields.find((f) => f.id === 'f_slot')!.options![n - 1]!;
  /** 枠の表示（例：12月5日（土）10:00〜）から、デモの時刻へ（Day 0 = 12月1日） */
  const startOf = (label: string) => {
    const m = /(\d+)月(\d+)日.*?(\d{2}):(\d{2})/.exec(label)!;
    return at(Number(m[1]) === 12 ? Number(m[2]) - 1 : 30 + Number(m[2]), Number(m[3]), Number(m[4]));
  };
  const bot = (u: UserState) => u.talk.flatMap((t) => (t.side === 'bot' && (t.bubble.kind === 'text' || t.bubble.kind === 'buttons') ? [t.bubble.text] : []));
  const remind = (id: string) => user(w, id).deliveries.filter((d) => d.flowId === 'flow_rsv_remind').map((d) => [d.messageId, d.at] as const);
  /** 友だち追加の日から何日後か（暦の日で数える） */
  const step = (u: UserState) => u.deliveries.filter((d) => d.flowId === 'flow_step').map((d) => [d.messageId, Math.floor(d.at / DAY) - Math.floor(u.friendAddedAt! / DAY)] as const);
  const got = (m: string) => Object.values(w.users).filter((u) => u.deliveries.some((x) => x.messageId === m)).map((u) => u.id).sort();
  /** 手で操作する人：ポップアップから友だち追加して、30秒診断の4問に答える（null なら答えない） */
  const visitor = (ans: [string, string, string, string] | null) => {
    let d = dispatch(newDemo(p, 'p1'), { type: 'addFriend', routeId: 'route_popup' });
    if (!ans) return d;
    d = dispatch(d, { type: 'tapButton', buttonId: 'btn_w_popup_diag' });
    ['q_when', 'q_goal', 'q_exp', 'q_freq'].forEach((q, i) => (d = dispatch(d, { type: 'answer', surveyId: 'sv_body', questionId: q, choiceIds: [`c_${q.slice(2)}_${ans[i]}`] })));
    return d;
  };
  const book = (d: Demo, n: number, type = '無料体験', pageId = 'page_rsv') => dispatch(d, { type: 'submitForm', pageId, values: { f_type: type, f_slot: slot(n), f_name: 'デモ' } });
  const answer = (d: Demo, sv: string, picks: [string, string][]) => picks.reduce((x, [q, c]) => dispatch(x, { type: 'answer', surveyId: sv, questionId: q, choiceIds: [c] }), d);
  const textOf = (id: string) => {
    const b = p.messages.find((m) => m.id === id)!.bubbles.find((x) => x.kind === 'buttons' || x.kind === 'text')!;
    return b.kind === 'buttons' || b.kind === 'text' ? b.text : '';
  };

  it('検証で重大エラー・警告が無く、全シナリオに誰かが入る', () => {
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
  });

  it('リッチメニューは資料P27・P28どおり3タブ×6枠（画像）。全ボタンが反応し、未入会の人は③（会員の方へ）を使えない', () => {
    // ③のメニューは未入会の人には開かない（③のタブ①は、入会した人で下で確かめる）
    expect(deadButtons(p, 'p1', 'route_popup')).toEqual(['menu:cell_t3_tab1']);
    const body = (m: (typeof p.richMenus)[number]) => m.cells.filter((c) => c.area.row > 0).map((c) => c.label);
    expect(p.richMenus.map((m) => [m.label, m.imageAssetId, body(m)])).toEqual([
      ['① 見学・体験', 'img_rm1', ['無料体験を予約', '店舗を探す', '設備・マシン', '料金プラン', '24時間の使い方', 'アクセス・駐車場']],
      ['② 診断・シミュレーター', 'img_rm2', ['30秒体づくり診断', '性格タイプ診断', 'ダイエットタイプ診断', '続けられるか診断', '1分料金診断', 'よくある質問・相談']],
      ['③ 会員の方へ', 'img_rm3', ['会員コード', 'レッスン予約', 'トレーニング記録', '混雑状況', 'お友だち紹介', '休会・退会']],
    ]);
    expect(p.richMenus[0]!.cells.filter((c) => c.emphasis).map((c) => c.label)).toEqual(['無料体験を予約']);
    // 未入会：タブ③も会員の機能も「会員の方専用」の案内だけ（メニューは切り替わらない・ページも開かない）
    let d = dispatch(visitor(null), { type: 'tapMenu', richMenuId: 'rm_t1', cellId: 'cell_t1_tab3' });
    expect([me(p, d).richMenuId, me(p, d).deliveries.at(-1)!.messageId]).toEqual(['rm_t1', 'msg_mo_tab3']);
    for (const cell of ['cell_code', 'cell_lesson', 'cell_log', 'cell_crowd', 'cell_ref', 'cell_quit']) {
      const x = me(p, dispatch(visitor(null), { type: 'tapMenu', richMenuId: 'rm_t3', cellId: cell }));
      expect([x.surface.kind, x.deliveries.at(-1)!.messageId.startsWith('msg_mo_')]).toEqual(['talk', true]);
    }
    // 入会の登録（店舗）で③に自動で切り替わり、会員の機能が使える
    d = dispatch(d, { type: 'cv', cvId: 'st_enroll' });
    expect(me(p, d).richMenuId).toBe('rm_t3');
    expect(me(p, dispatch(d, { type: 'tapMenu', richMenuId: 'rm_t3', cellId: 'cell_code' })).surface).toEqual({ kind: 'page', pageId: 'page_code' });
    // 会員は①②のタブにも行き来できる（③のタブで③に戻る）
    const m1 = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_t3', cellId: 'cell_t3_tab1' });
    expect([me(p, m1).richMenuId, me(p, dispatch(m1, { type: 'tapMenu', richMenuId: 'rm_t1', cellId: 'cell_t1_tab3' })).richMenuId]).toEqual(['rm_t1', 'rm_t3']);
    for (const id of ['p2', 'p3', 'p4', 'p6', 'p7']) expect(user(w, id).richMenuId).not.toBe('rm_t3');
  });

  it('入口：ポップアップ・サイトのLINEボタン・Web予約のあと・LINE広告・店頭QR・紹介。あいさつは資料P29（無料体験・見学OK・営業電話なし・時間）＋期限のバナー', () => {
    const pop = p.webPages.find((x) => x.id === 'page_facility_web')!.exitPopup!;
    expect([pop.title, pop.cta.label]).toEqual(['LINEで30秒体づくり診断', 'LINEで30秒診断を受け取る']);
    const u = me(p, visitor(null));
    const hello = bot(u)[0]!;
    for (const s of ['○○フィットネス △△店', '無料体験を受付中', '見学だけでもOK', '営業のお電話をすることはありません', '10:00〜20:00', 'デモ用の仮設定', '体づくりタイプ診断']) expect(hello).toContain(s);
    expect(u.deliveries.map((d) => d.messageId)).toEqual(['msg_welcome_popup', 'image:img_banner_dec']);
    expect(p.messages.find((m) => m.id === 'msg_welcome_popup')!.bubbles.flatMap((b) => bubbleButtons(b).map((x) => x.label))).toEqual(['▶ 30秒 体づくりタイプ診断', '無料体験・見学を予約する']);
    expect(['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7'].map((id) => user(w, id).route)).toEqual(['route_popup', 'route_popup', 'route_cpf', 'route_web_rsv', 'route_store', 'route_site', 'route_referral']);
  });

  it('30秒診断：資料P29の4問を1問ずつ。答えた瞬間にタグ、4問で回答の組み合わせからタイプ（断定しない）＋料金の参考', () => {
    expect(p.surveys[0]!.questions.map((q) => [q.text, q.choices.map((c) => c.label)])).toEqual([
      ['Q1 いつから始めたいですか？', ['今すぐ', '1ヶ月以内', '2〜3ヶ月先', '検討中']],
      ['Q2 一番の目的は？', ['ダイエット', '筋肉', '健康維持', '姿勢・肩こり']],
      ['Q3 運動の経験は？', ['まったく初めて', '昔やっていた', '運動中', 'ジム経験あり']],
      ['Q4 通えそうな頻度は？', ['週1回', '週2回', '週3回以上', 'わからない']],
    ]);
    let d = dispatch(visitor(null), { type: 'tapButton', buttonId: 'btn_w_popup_diag' });
    d = dispatch(d, { type: 'answer', surveyId: 'sv_body', questionId: 'q_when', choiceIds: ['c_when_m1'] });
    expect(Object.keys(me(p, d).tags)).toEqual(['when_m1']);
    expect(me(p, d).surface).toEqual({ kind: 'survey', surveyId: 'sv_body', questionId: 'q_goal' });
    const type = (a: [string, string, string, string]) => Object.keys(me(p, visitor(a)).tags).filter((t) => t.startsWith('bt_'));
    // 資料P29の例：初心者＋週2回→コツコツ習慣化／筋肉＋経験あり→筋力アップ重視／ダイエット＋検討中→無理なく始める
    expect([type(['m1', 'health', 'none', 'w2']), type(['now', 'muscle', 'gym', 'w3']), type(['undecided', 'diet', 'past', 'w1']), type(['now', 'diet', 'past', 'w3']), type(['m1', 'posture', 'now', 'w1'])]).toEqual([['bt_habit'], ['bt_muscle'], ['bt_gentle'], ['bt_slim'], ['bt_care']]);
    const u = me(p, visitor(['m1', 'health', 'none', 'w2']));
    const res = bot(u).find((t) => t.startsWith('＼あなたは'))!;
    for (const s of ['「コツコツ習慣化タイプ」', '① おすすめの通い方', '② 最初の1ヶ月の取り組み方', '③ おすすめの設備・マシン', '④ 週あたりの通う頻度', 'あなたの回答：週2回', '健康状態や減量の結果を判断するものではありません']) expect(res).toContain(s);
    for (const s of ['料金の参考シミュレーション', '1回あたり 1,000円（月会費8,000円 ÷ 月8回）', '実在のジムの料金ではありません']) expect(bot(u).at(-1)).toContain(s);
    expect(p.messages.find((m) => m.id === 'msg_bref_w2')!.bubbles.flatMap((b) => bubbleButtons(b).map((x) => x.action))).toContainEqual({ kind: 'startFlow', flowId: 'flow_rsv_open' });
    // 答え直すとタイプが入れ替わる（古いタイプのタグは残らない）
    let again = dispatch(visitor(['m1', 'health', 'none', 'w2']), { type: 'tapMenu', richMenuId: 'rm_t2', cellId: 'cell_diag' });
    again = answer(again, 'sv_body', [['q_when', 'c_when_now'], ['q_goal', 'c_goal_muscle'], ['q_exp', 'c_exp_gym'], ['q_freq', 'c_freq_w3']]);
    expect(Object.keys(me(p, again).tags).filter((t) => /^(bt_|goal_|exp_)/.test(t)).sort()).toEqual(['bt_muscle', 'exp_gym', 'goal_muscle']);
    // 途中でやめた人に、結果のタグは付けない
    const quit = seek(dispatch(dispatch(visitor(null), { type: 'tapButton', buttonId: 'btn_w_popup_diag' }), { type: 'answer', surveyId: 'sv_body', questionId: 'q_when', choiceIds: ['c_when_now'] }), at(3));
    expect(Object.keys(me(p, quit).tags).filter((t) => /^(bt_|diag_done)/.test(t))).toEqual([]);
  });

  it('性格タイプ（3問・2問以上そろったタイプ／ばらばらならQ1）・ダイエットタイプ（4問）・続けられるか（3問）。結果はタグになり、断定しない', () => {
    const psy = (a: string, b: string, c: string) =>
      Object.keys(me(p, answer(dispatch(visitor(null), { type: 'tapMenu', richMenuId: 'rm_t2', cellId: 'cell_psy' }), 'sv_psy', [['q_p1', `c_q_p1_${a}`], ['q_p2', `c_q_p2_${b}`], ['q_p3', `c_q_p3_${c}`]])).tags).filter((t) => t.startsWith('ps_') && t !== 'ps_done');
    expect([psy('mood', 'log', 'log'), psy('buddy', 'buddy', 'plan'), psy('plan', 'mood', 'buddy')]).toEqual([['ps_log'], ['ps_buddy'], ['ps_plan']]);
    expect(tags('p2')).toEqual(expect.arrayContaining(['ps_mood', 'kt_bored']));
    expect(tags('p3')).toContain('dt_snack');
    const t = (id: string) => bot(user(w, id)).join('\n');
    for (const s of ['あなたは「気分派」', '① 続けやすい通い方', '② 目標の立て方', '③ 挫折しにくい工夫', '④ おすすめの行動']) expect(t('p2')).toContain(s);
    for (const s of ['あなたは「間食見直し型」', '① ジムで取り組めること', '② 家で試せること（1つだけ）', '減量の結果を約束するものではありません']) expect(t('p3')).toContain(s);
    for (const s of ['あなたは「変化を楽しむ型」', '① 挫折しやすい場面', '② 体験当日の回り方', '「周りの目」', '③ 最初の2週間の予定例', '休日に', '必ず続けられると決めるものではありません']) expect(t('p2')).toContain(s);
    expect(Object.values(w.users).flatMap((u) => bot(u)).filter((x) => /必ず続けられます|必ず痩せ/.test(x))).toEqual([]);
  });

  it('1分料金シミュレーター：月会費÷（週の回数×4週）。回数を変えると結果も②のタグも入れ替わり、計算しただけでは予約（CV①）にならない', () => {
    const u = user(w, 'p6');
    const res = bot(u).filter((x) => x.startsWith('【1分料金シミュレーター】'));
    expect(res).toHaveLength(2);
    for (const s of ['プラン：いつでも通えるプラン', '月会費：8,000円', '月の想定利用回数 8回', '1回あたり：1,000円', '8,000円 ÷ 月8回（週2回×4週）', '実在のジムの料金ではありません']) expect(res[0]).toContain(s);
    for (const s of ['月の想定利用回数 4回', '1回あたり：2,000円']) expect(res[1]).toContain(s);
    expect(tags('p6').filter((x) => x.startsWith('pr_'))).toEqual(['pr_done', 'pr_8000_1']);
    expect(u.cvs['cv_rsv']).toEqual({ at: at(34, 21, 33), count: 1 });
    expect([textOf('msg_pr_10000_3'), textOf('msg_pr_6000_4')]).toEqual([expect.stringContaining('1回あたり：約833円'), expect.stringContaining('1回あたり：375円')]);
  });

  it('入会前14日間（資料P25）：2・3・4・5・6・8・10・11・14日後。5日後はダイエット目的の人だけ。予約中・来店ずみの人には送らない。送らない理由は②に残る', () => {
    expect(step(user(w, 'p2'))).toEqual([['msg_s2_gentle', 2], ['msg_s3', 3], ['msg_s4', 4], ['msg_s6_mood', 6], ['msg_s8_gentle', 8], ['msg_s10', 10], ['msg_s11', 11], ['msg_s14_dec', 14]]);
    expect(step(user(w, 'p3'))).toEqual([['msg_s2_slim', 2], ['msg_s3', 3], ['msg_s4', 4], ['msg_s5', 5], ['msg_s6_default', 6]]);
    expect(got('msg_s5')).toEqual(['p3']);
    // 予約のあとは声がけを減らす：14日間の残りも定期配信も届かない（会員の月1回は別）
    for (const id of ['p1', 'p3', 'p4', 'p6']) {
      const u = user(w, id);
      const bookedAt = u.cvs['cv_rsv']!.at;
      expect(u.deliveries.filter((d) => d.at > bookedAt && (d.flowId === 'flow_step' || (d.flowId === 'flow_weekly' && !d.messageId.endsWith('_member'))))).toEqual([]);
    }
    // 送らなかった理由（分岐の判定）が残る：目的がダイエットではない／予約ずみ
    const why = (id: string, ref: string) => w.decisions.some((x) => x.userId === id && x.kind === 'branch' && x.outcome === 'taken' && x.ref === ref);
    expect([why('p2', 'c_s5_other'), why('p3', 'c_s10_rsv'), why('p4', 'c_s6_rsv'), why('p3', 'c_wk2_q')]).toEqual([true, true, true, true]);
  });

  it('CV①（LINE予約）：確認画面のあとの確定で1回だけ。お手続き・日時が確認の文面・②のタグ・③の時系列まで同じ。当日の流れはLINEの文面に書く。「LINEに戻る」で確認が届いている', () => {
    const page = p.webPages.find((x) => x.id === 'page_rsv')!;
    const form = page.form!;
    expect([form.confirm?.submitLabel, form.fields.map((f) => f.label), form.fields[0]!.options]).toEqual(['この内容で予約する', ['お手続き', '希望日時（空き枠）', 'お名前（デモ用の架空のお名前）'], ['無料体験', '施設見学', '入会説明']]);
    expect(page.blocks[0]).toMatchObject({ kind: 'heading', text: '○○フィットネス △△店を予約' });
    expect(p.webPages.some((x) => x.id === 'page_dayflow')).toBe(false);
    const u = user(w, 'p1');
    expect(u.cvs['cv_rsv']).toEqual({ at: at(0, 12, 23), count: 1 });
    const ok = bot(u).find((t) => t.startsWith('【ご予約を受け付けました】'))!;
    for (const s of ['お手続き：無料体験', `日時：${slot(1)}`, '店舗：○○フィットネス △△店', '持ち物', '▼ 当日の流れ\n① 受付で', '④ ご質問・料金のご説明', '見学だけでも大丈夫', '変更・キャンセル', 'デモ用の通知']) expect(ok).toContain(s);
    expect(ok).not.toContain('希望プラン');
    expect(tags('p1')).toEqual(expect.arrayContaining(['rsv_s1', 'rsv_t_trial', 'rsv_via_line']));
    expect(JSON.stringify(buildTimeline(p, new Labels(p), w, ['p1']))).toContain(slot(1));
    // 予約だけでは来店・入会にならない。LINEに戻ると確認が届いている
    let d = book(visitor(['now', 'muscle', 'past', 'w2']), 5);
    expect([me(p, d).surface, 'visited' in me(p, d).tags, 'enrolled' in me(p, d).tags, me(p, d).cvs['cv_visit']]).toEqual([{ kind: 'page', pageId: 'page_rsv_done' }, false, false, undefined]);
    d = dispatch(d, { type: 'tapPageButton', pageId: 'page_rsv_done', blockIndex: 2 });
    expect([me(p, d).surface, me(p, d).deliveries.at(-1)!.messageId]).toEqual([{ kind: 'talk' }, 'msg_rsv_ok']);
  });

  it('入会説明の予約もCV①で1回。確認は入会説明の当日の流れ。入会は店舗の登録だけ。再予約は同じ画面で、CVを重ねず裏側でタグを付ける', () => {
    const u = me(p, dispatch(book(visitor(null), 5, '入会説明'), { type: 'tapPageButton', pageId: 'page_rsv_done', blockIndex: 2 }));
    expect([u.cvs['cv_rsv']?.count, 'rsv_t_enroll' in u.tags, 'enrolled' in u.tags, u.deliveries.at(-1)!.messageId]).toEqual([1, true, false, 'msg_rsv_ok_enroll']);
    expect(bot(u).at(-1)).toContain('③ 料金・プラン・解約の条件のご説明');
    // キャンセルのあとの予約：見た目は同じ画面。CV①は増えず、「再予約」のタグだけが付く
    const [a, b] = [p.webPages.find((x) => x.id === 'page_rsv')!, p.webPages.find((x) => x.id === 'page_rsv_re')!];
    expect([b.title, b.blocks, b.form!.fields]).toEqual([a.title, a.blocks, a.form!.fields]);
    let d = dispatch(book(visitor(null), 5), { type: 'tapPageButton', pageId: 'page_rsv_done', blockIndex: 2 });
    d = dispatch(dispatch(dispatch(d, { type: 'tapButton', buttonId: 'btn_rsv_ok_manage' }), { type: 'tapButton', buttonId: 'btn_manage_cancel' }), { type: 'tapButton', buttonId: 'btn_cancel_yes' });
    d = dispatch(d, { type: 'tapMenu', richMenuId: 'rm_t1', cellId: 'cell_rsv' });
    expect(me(p, d).surface).toEqual({ kind: 'page', pageId: 'page_rsv_re' });
    const again = me(p, book(d, 6, '無料体験', 'page_rsv_re'));
    expect([again.cvs['cv_rsv']?.count, 'rsv_again' in again.tags, 'rsv_booked' in again.tags]).toEqual([1, true, true]);
  });

  it('Web予約（入口C）：CV①は経路Webで1回。LINEに来た時点で予約ずみとして扱い、予約内容をあいさつで確認・前日の案内も届く', () => {
    const u = user(w, 'p4');
    expect([u.cvs['cv_rsv'], tags('p4').includes('rsv_via_web'), tags('p4').includes('rsv_via_line')]).toEqual([{ at: at(3, 22, 4), count: 1 }, true, false]);
    for (const s of ['Webからのご予約', 'お手続き：施設見学', `日時：${slot(2)}`, '店舗：○○フィットネス △△店', '▼ 当日の流れ']) expect(bot(u)[0]).toContain(s);
    expect(remind('p4')).toEqual([['msg_rsv_1d', at(6, 18)], ['msg_rsv_thanks', at(8, 12)]]);
    expect(u.deliveries.filter((d) => /^msg_(s6_|s14_|s11$)/.test(d.messageId))).toEqual([]);
  });

  it('予約後の案内：直後・前日18:00・翌日12:00（当日の案内はない）。予約していない人には送らない。直前の予約で、過ぎた時刻の案内はあとから送らない', () => {
    expect(remind('p1')).toEqual([['msg_rsv_ok', at(0, 12, 23)], ['msg_rsv_1d', at(3, 18)], ['msg_rsv_thanks', at(5, 12)]]);
    expect(p.messages.some((m) => m.id === 'msg_rsv_2h')).toBe(false);
    expect([remind('p2'), remind('p5')]).toEqual([[], []]);
    expect(bot(user(w, 'p1')).find((x) => x.startsWith('【明日'))).toContain(`【明日 ${slot(1)}】無料体験のご確認です`);
    // 六郎：前日18:00を過ぎてから翌朝の枠を予約 → 前日の案内は送らない
    expect(remind('p6')).toEqual([['msg_rsv_ok', at(34, 21, 33)], ['msg_rsv_missed', at(36, 12)]]);
    let d = seek(visitor(['now', 'muscle', 'past', 'w2']), at(3, 20));
    d = seek(book(d, 1), at(5, 13));
    expect(me(p, d).deliveries.filter((x) => x.flowId === 'flow_rsv_remind').map((x) => x.messageId)).toEqual(['msg_rsv_ok', 'msg_rsv_missed']);
    // 7人のどの案内も、予約した日時の前日18:00ちょうどにだけ届く（後追いで届かない）
    let n = 0;
    for (const u of Object.values(w.users))
      for (const t of u.talk) {
        if (t.side !== 'bot' || !('text' in t.bubble)) continue;
        const day = /^【明日 (.+?〜)】/.exec(t.bubble.text);
        if (day) expect(t.at).toBe(at(Math.floor(startOf(day[1]!) / DAY) - 1, 18));
        if (day) n++;
      }
    expect(n).toBe(4);
  });

  it('日程変更で前の日時の案内は止まり、キャンセルで予約者向けの案内が止まる。CV①は増えない。ブロック後は何も届かない', () => {
    expect(remind('p7')).toEqual([['msg_rsv_ok', at(17, 19, 6)], ['msg_rsv_chg', at(19, 12, 3)], ['msg_rsv_1d', at(25, 18)]]);
    const u = user(w, 'p7');
    expect(bot(u).find((x) => x.startsWith('【日程を変更しました】'))).toContain(`新しい日時：${slot(7)}`);
    expect(u.cvs['cv_rsv']!.count).toBe(1);
    // 予約中は14日間の配信を止める。キャンセルすると残りの配信に戻る
    expect(step(u)).toEqual([['msg_s10', 10]]);
    expect(tags('p7').filter((t) => /^rsv_(booked|s\d)$/.test(t))).toEqual([]);
    expect(['rsv_changed', 'rsv_cancel'].every((t) => tags('p7').includes(t))).toBe(true);
    const blockedAt = u.behaviors.find((b) => b.type === 'block')!.at;
    expect(u.deliveries.filter((x) => x.at > blockedAt)).toEqual([]);
    // ブロックで止まったシナリオ（入会前の配信・1月の企画など）は、②に「停止」として残る
    expect(w.decisions.filter((x) => x.userId === 'p7' && x.kind === 'stopWhen' && x.at === blockedAt).map((x) => x.flowId).sort()).toEqual(['flow_bc_jan', 'flow_cal', 'flow_step']);
  });

  it('CV②（来店）は店舗が確認して登録したときだけ。予約の日時を過ぎても来店にしない。翌日は来店した人にお礼、来ていない人に前提にしない文面', () => {
    expect(['p1', 'p3', 'p4'].map((id) => user(w, id).cvs['cv_visit']!.at)).toEqual([at(4, 11, 30), at(11, 15, 30), at(7, 20, 30)]);
    expect([user(w, 'p6').cvs['cv_visit'], tags('p6').includes('visited'), tags('p6').includes('rsv_past')]).toEqual([undefined, false, true]);
    expect([got('msg_rsv_thanks'), got('msg_rsv_missed'), got('msg_compare')]).toEqual([['p1', 'p3', 'p4'], ['p6'], ['p3', 'p4']]);
    const d = seek(book(visitor(['now', 'muscle', 'past', 'w2']), 1), at(6));
    expect([me(p, d).cvs['cv_visit'], 'visited' in me(p, d).tags]).toEqual([undefined, false]);
  });

  it('日時を過ぎて来店のタグが無いと、②の先頭に店舗の操作（来店タグ・来店なし）を出す。押すまで来店後には進まない', () => {
    const view = (d: Demo) => ({ project: p, world: worldOf(p, d), demo: d, ui: { menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' }) as never;
    const html = (d: Demo) => backstage(view(d), me(p, d));
    const booked = dispatch(book(visitor(['now', 'muscle', 'past', 'w2']), 1), { type: 'tapPageButton', pageId: 'page_rsv_done', blockIndex: 2 });
    expect(html(seek(booked, at(4, 9)))).not.toContain('lsp-ask');
    const past = seek(booked, at(4, 13));
    expect(['lsp-ask', '来店タグを付ける（店舗が登録）', 'data-cv="st_noshow"'].every((x) => html(past).includes(x))).toBe(true);
    expect(guideOf(view(past), me(p, past)).text).toContain('②の「来店タグを付ける（店舗が登録）」');
    // 来店タグ：カードが消えて来店後へ
    const came = dispatch(past, { type: 'cv', cvId: 'cv_visit' });
    expect([html(came).includes('lsp-ask'), 'visited' in me(p, came).tags, me(p, came).cvs['cv_visit']?.count]).toEqual([false, true, 1]);
    // 来店なし：CV②にせず、タグだけ。下のボタン一覧には出さない
    const none = dispatch(past, { type: 'cv', cvId: 'st_noshow' });
    expect([html(none).includes('lsp-ask'), 'no_show' in me(p, none).tags, me(p, none).cvs['cv_visit'], html(none).includes('data-cv="st_noshow"')]).toEqual([false, true, undefined, false]);
    expect(me(p, seek(none, at(5, 13))).deliveries.at(-1)!.messageId).toBe('msg_rsv_missed');
    expect(tags('p6')).toContain('no_show');
    // 14日間の配信：店舗が確認するまでは止めたまま。「来店なし」で残りの配信に戻る
    const rest = (d: Demo) => step(me(p, seek(d, at(7))));
    expect(rest(past)).toEqual([]);
    expect(rest(came)).toEqual([]);
    expect(rest(none).map((x) => x[0])).toEqual(['msg_s4', 'msg_s6_default']);
  });

  it('文面の修正：「比べて決めて」と言わない・持ち物5つは縦に並べる・来店ずみの人に体験／見学の予約を案内しない', () => {
    const all = p.messages.flatMap((m) => m.bubbles.flatMap((b) => (b.kind === 'text' || b.kind === 'buttons' ? [b.text] : [])));
    expect(all.filter((x) => /比べて決め|ほかのジムと比べ/.test(x))).toEqual([]);
    expect(textOf('msg_m1')).toContain('✅ 室内シューズ\n✅ 動きやすいウェア\n✅ タオル\n✅ 飲み物\n✅ ロッカーの鍵');
    const toRsv = (m: string) => (p.messages.find((x) => x.id === m)?.bubbles ?? []).flatMap((b) => (b.kind === 'buttons' ? b.buttons : [])).filter((b) => b.action.kind === 'startFlow' && b.action.flowId === 'flow_rsv_open');
    for (const id of ['p1', 'p3', 'p4']) {
      const u = user(w, id);
      const came = u.cvs['cv_visit']!.at;
      expect(u.deliveries.filter((d) => d.at > came && d.flowId !== null).flatMap((d) => toRsv(d.messageId))).toEqual([]);
      expect(bot(u).filter((x) => x.includes('【1月の予告】'))).toEqual([]);
    }
  });

  it('入会は店舗の登録だけ。入会で入会前の配信が止まり、メニュー③・入会日から1・5・7・14日後の配信に切り替わる', () => {
    expect(Object.values(w.users).filter((x) => 'enrolled' in x.tags).map((x) => x.id)).toEqual(['p1', 'p5']);
    for (const id of ['p1', 'p5']) {
      const u = user(w, id);
      const enrolledAt = u.cvs['st_enroll']!.at;
      expect(u.deliveries.filter((d) => d.at > enrolledAt && /^msg_(s\d|jan$|wk\d$)/.test(d.messageId))).toEqual([]);
      expect(u.deliveries.filter((d) => d.flowId === 'flow_member').map((d) => [d.messageId, Math.floor(d.at / DAY) - Math.floor(enrolledAt / DAY)])).toEqual([['msg_m1', 1], ['msg_m5', 5], ['msg_m7', 7], ['msg_m14', 14]]);
      expect(tags(id)).toContain('member_step_done');
    }
    expect(bot(user(w, 'p1')).find((t) => t.startsWith('【1週目】'))).toContain('個人差があり、効果を約束するものではありません');
    // ボタンや言葉では入会は成立しない（入会の自動応答も「店舗で」）
    const d = dispatch(visitor(null), { type: 'sendText', text: '入会したい' });
    expect([me(p, d).deliveries.at(-1)!.messageId, 'enrolled' in me(p, d).tags]).toEqual(['msg_ar_join', false]);
  });

  it('CV③：PT申込（詳細を見ただけでは記録しない）・物販申込・友だち紹介（送っただけでは記録しない・店舗が確認）。GOALはCV③', () => {
    const u1 = user(w, 'p1');
    expect(u1.cvs['cv3_pt']).toEqual({ at: at(11, 19, 13), count: 1 });
    expect(u1.behaviors.some((b) => b.type === 'visitPage' && b.detail['pageId'] === 'page_pt_info' && b.at < at(11, 19, 13))).toBe(true);
    expect([user(w, 'p5').cvs['cv3_goods']!.at, user(w, 'p5').cvs['cv3_ref']!.at]).toEqual([at(2, 19, 33), at(17, 19, 30)]);
    // 紹介された七郎さんの友だち追加（19:01）のあとに、店舗が紹介の成立を登録
    expect(user(w, 'p7').friendAddedAt).toBeLessThan(user(w, 'p5').cvs['cv3_ref']!.at);
    let d = dispatch(dispatch(visitor(null), { type: 'cv', cvId: 'st_enroll' }), { type: 'tapMenu', richMenuId: 'rm_t3', cellId: 'cell_ref' });
    d = dispatch(dispatch(d, { type: 'tapButton', buttonId: 'btn_ref_send' }), { type: 'visitPage', pageId: 'page_pt_info' });
    expect(['cv3_ref', 'cv3_pt', 'cv3_goods'].map((c) => me(p, d).cvs[c])).toEqual([undefined, undefined, undefined]);
    expect('ref_shared' in me(p, d).tags).toBe(true);
    expect(journeyOf(p, w, u1).find((r) => r.isGoal)!.step.id).toBe('j_cv3');
    expect(journeyOf(p, w, user(w, 'p5')).find((r) => r.step.id === 'j_cv3')!.status).not.toBe('todo');
  });

  it('季節の企画：10〜12月に友だちになった人へ1月1日に新年の無料体験（会員・予約中・来店ずみ・その他で出し分け）。特典の期限は配信した月の月末', () => {
    expect(['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7'].map((id) => user(w, id).deliveries.find((d) => d.flowId === 'flow_bc_jan')?.messageId)).toEqual(['msg_jan_member', 'msg_jan', 'msg_jan_visited', 'msg_jan_visited', 'msg_jan_member', 'msg_jan', undefined]);
    expect([step(user(w, 'p2')).at(-1)![0], step(user(w, 'p6')).at(-1)![0]]).toEqual(['msg_s14_dec', 'msg_s14_jan']);
    for (const u of Object.values(w.users))
      for (const t of u.talk) {
        if (t.side !== 'bot' || !('text' in t.bubble)) continue;
        if (t.bubble.text.includes('12月31日（木）まで')) expect(t.at).toBeLessThan(at(31));
        if (t.bubble.text.includes('1月31日（日）まで')) expect(t.at).toBeGreaterThanOrEqual(at(31));
      }
    // 月4回の定期配信（金曜）。会員は入会後14日間のあとの月1回だけ
    expect(user(w, 'p2').deliveries.filter((d) => d.flowId === 'flow_weekly').map((d) => d.messageId)).toEqual(['msg_wk1_done', 'msg_wk2_done', 'msg_wk3', 'msg_wk4_done']);
    for (const id of ['p1', 'p5']) expect(user(w, id).deliveries.filter((d) => d.flowId === 'flow_weekly').map((d) => [d.messageId, Math.floor(d.at / DAY)])).toEqual([['msg_wk4_member', 24]]);
  });

  it('一度答えた診断・料金シミュレーターは配信で案内しない。定期配信は同じ金曜に診断のない文面へ差し替える（まだ答えていない人には届く）', () => {
    const doneTag: Record<string, string> = { flow_diag: 'diag_done', flow_psy: 'ps_done', flow_diet: 'dt_done', flow_keep: 'kt_done', flow_price: 'pr_done' };
    const starts = (m: string) =>
      (p.messages.find((x) => x.id === m)?.bubbles ?? []).flatMap((b) => (b.kind === 'buttons' ? b.buttons : [])).flatMap((b) => (b.action.kind === 'startFlow' && b.action.flowId in doneTag ? [b.action.flowId] : []));
    /** 答えたあとに届いた案内（その結果についた「計算し直す」と、自分で送った言葉への自動応答は除く） */
    const resent = (u: UserState) => u.deliveries.flatMap((d) => (d.flowId === null ? [] : starts(d.messageId).filter((f) => f !== d.flowId && (u.tags[doneTag[f]!]?.at ?? Infinity) <= d.at).map((f) => `${d.messageId}→${f}`)));
    for (const u of Object.values(w.users)) expect(resent(u)).toEqual([]);

    const fill = (d: Demo, cellId: string, sv: string) =>
      answer(dispatch(d, { type: 'tapMenu', richMenuId: 'rm_t2', cellId }), sv, p.surveys.find((s) => s.id === sv)!.questions.map((q) => [q.id, q.choices[0]!.id] as [string, string]));
    const weekly = (u: UserState) => u.deliveries.filter((x) => x.flowId === 'flow_weekly').map((x) => x.messageId);
    // すべてに答えた人：14日間の診断・料金の配信は止まり、定期配信は4回とも診断のない文面
    const all = me(p, seek(fill(fill(fill(fill(visitor(['now', 'diet', 'none', 'w2']), 'cell_psy', 'sv_psy'), 'cell_diet', 'sv_diet'), 'cell_keep', 'sv_keep'), 'cell_price', 'sv_price'), at(37)));
    expect(all.deliveries.filter((x) => /^msg_(s3|s4|s5|s10|wk\d)$/.test(x.messageId))).toEqual([]);
    expect(weekly(all)).toEqual(['msg_wk1_done', 'msg_wk2_done', 'msg_wk3_done', 'msg_wk4_done']);
    expect(resent(all)).toEqual([]);
    // 30秒診断だけの人：残りの診断は届く。ダイエットタイプ診断は目的がダイエットの人だけ
    const diet = me(p, seek(visitor(['now', 'diet', 'none', 'w2']), at(25)));
    expect(weekly(diet)).toEqual(['msg_wk1_done', 'msg_wk2', 'msg_wk3', 'msg_wk4']);
    expect(step(diet).map((x) => x[0])).toEqual(expect.arrayContaining(['msg_s3', 'msg_s4', 'msg_s5', 'msg_s10']));
    expect(weekly(me(p, seek(visitor(['m1', 'health', 'none', 'w2']), at(25))))).toEqual(['msg_wk1_done', 'msg_wk2_done', 'msg_wk3', 'msg_wk4']);
    // 料金を計算ずみの人の30秒診断の結果には、計算のボタンを出さない
    const body: [string, string][] = [['q_when', 'c_when_now'], ['q_goal', 'c_goal_muscle'], ['q_exp', 'c_exp_past'], ['q_freq', 'c_freq_w2']];
    expect(me(p, answer(dispatch(fill(visitor(null), 'cell_price', 'sv_price'), { type: 'tapMenu', richMenuId: 'rm_t2', cellId: 'cell_diag' }), 'sv_body', body)).deliveries.at(-1)!.messageId).toBe('msg_bref_w2_priced');
    expect(me(p, visitor(['now', 'muscle', 'past', 'w2'])).deliveries.at(-1)!.messageId).toBe('msg_bref_w2');
  });

  it('自動応答（10のキーワード）と有人対応：よくある質問は自動、相談・契約・手続きは店舗スタッフ（②に「有人対応中」）', () => {
    for (const k of ['料金', '無料体験', '見学', '入会', '解約', '持ち物', '設備', '混雑', '営業時間', 'パーソナルトレーニング']) {
      const u = me(p, dispatch(visitor(null), { type: 'sendText', text: k }));
      expect([u.deliveries.at(-1)!.flowId, bot(u).at(-1)!.startsWith('（自動応答）')]).toEqual([null, true]);
    }
    expect(user(w, 'p6').deliveries.find((d) => d.messageId === 'msg_ar_quit')!.flowId).toBeNull();
    expect(['p3', 'p5'].every((id) => tags(id).includes('human_chat'))).toBe(true);
    expect(tags('p6')).not.toContain('human_chat');
    expect(bot(user(w, 'p3')).join('\n')).toContain('（有人対応に切り替え）');
  });

  it('7人の状態は独立し、戻るとタグ・CV・予約・メニュー・届いた文面もその時点に戻る', () => {
    expect(['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7'].map((id) => user(w, id).personName)).toEqual(['佐藤一郎', '佐藤二郎', '佐藤三郎', '佐藤四郎', '佐藤五郎', '佐藤六郎', '佐藤七郎']);
    expect(tags('p2').filter((t) => t.startsWith('rsv_'))).toEqual([]);
    let d = seek(allScripted(p), at(8));
    const id = d.session.users.find((u) => u.personaId === 'p1')!.id;
    expect([worldOf(p, d).users[id]!.cvs['st_enroll'], worldOf(p, d).users[id]!.richMenuId]).toEqual([{ at: at(6, 19), count: 1 }, 'rm_t3']);
    d = seek(d, at(0, 12, 22));
    const before = worldOf(p, d).users[id]!;
    expect([before.cvs['cv_rsv'], 'rsv_booked' in before.tags, 'bt_muscle' in before.tags, before.richMenuId, before.deliveries.some((x) => x.messageId === 'msg_rsv_ok')]).toEqual([undefined, false, true, 'rm_t1', false]);
  });

  it('資料用16:9の③：LINE追加→診断→体験予約→来店→入会→入会後フォロー→CV③を、この順にいつも出す', () => {
    const keys = p.journey!.filter((s) => s.key).map((s) => s.label);
    expect(keys).toEqual(['LINE友だち追加', '30秒 体づくりタイプ診断', '来店の予約（CV①）', '来店・体験（CV②・店舗が確認）', '入会（店舗が登録）', '入会後14日間のステップ配信', '継続・紹介・PT／物販の申込（CV③）']);
    for (const [pid, when] of [['p1', at(0, 12, 12)], ['p1', at(12)], ['p3', at(9)], ['p5', at(20)], ['p6', at(36)], ['p7', at(30)]] as const) {
      let d = seek(allScripted(p), when);
      d = { ...d, activeUserId: d.session.users.find((u) => u.personaId === pid)!.id };
      const v = { project: p, world: worldOf(p, d), demo: d, ui: { menuOpen: true, picked: [], mode: 'slide3', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' } as never;
      const html = journeyPane(v, me(p, d), true);
      const from = html.indexOf('<ol class="lsp-strip');
      const strip = html.slice(from, html.indexOf('</ol>', from));
      const names = [...strip.matchAll(/<li class="([^"]*)">(.*?)<\/li>/g)].map((m) => m[2]!.replace(/<[^>]+>/g, ''));
      const found = keys.map((k) => names.findIndex((x) => x.includes(k)));
      expect(found.every((i) => i >= 0)).toBe(true);
      expect([...found].sort((a, b) => a - b)).toEqual(found);
    }
  });

  it('自動再生で全員を最後まで流すと CLI と同じ結果になり、営業用HTMLを書き出せる', async () => {
    const d = playToEnd(p);
    expect(d.now).toBe(37 * DAY);
    const ww = worldOf(p, d);
    for (const u of Object.values(ww.users)) expect(u.deliveries).toEqual(user(w, d.session.users.find((x) => x.id === u.id)!.personaId).deliveries);
    await exportChecks(p);
  }, 30_000);
});
