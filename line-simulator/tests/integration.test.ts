// Phase 6：AI提案（packages/ai/src/review.ts）と、Phase 1〜6 を通した最終統合テスト。
// 架空の美容サロンのスクショから Claude Code が作った提案（examples/sample-beauty/）を、
// 新しい案件に取り込み → 承認 → 自動再生 → 営業用HTML まで、一時フォルダの中で通す。

import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { afterAll, describe, expect, it } from 'vitest';
import type { Project, ScenarioPatch, WorldState } from '../packages/engine/src/index.ts';
import { approve, personaCoverage, reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { checkShape, countProv, dryRun, listProv, removeAt } from '../packages/ai/src/review.ts';
import { allScripted, autoStep, worldOf } from '../packages/player/src/controller.ts';
import { explainDecision, Labels } from '../packages/player/src/describe.ts';
import { bundlePlayer } from '../tools/bundle.ts';
import { scanEngineSource } from '../tools/engine-scan.ts';
import { buildExport, sanitize } from '../tools/export-lib.ts';
import { ProjectStore } from '../tools/project-store.ts';
import { DAY, sample } from './helpers.ts';

const ROOT = join(import.meta.dirname, '..');
const PATCH = JSON.parse(readFileSync(join(ROOT, 'examples', 'sample-beauty', '20261027-0130_salon_from_screens.json'), 'utf8')) as ScenarioPatch;
const all = new Set(PATCH.ops.map((_, i) => i));

const root = mkdtempSync(join(tmpdir(), 'linesim-final-'));
afterAll(() => rmSync(root, { recursive: true, force: true }));
const store = new ProjectStore(root);

function playToEnd(p: Project) {
  let d = allScripted(p);
  for (let i = 0; i < 10_000; i++) {
    const r = autoStep(p, d);
    d = r.demo;
    if (r.stopped) break;
  }
  return { demo: d, world: worldOf(p, d) };
}
const byPersona = (w: WorldState, demo: ReturnType<typeof allScripted>, personaId: string) => w.users[demo.session.users.find((u) => u.personaId === personaId)!.id]!;

describe('AI提案の確認（review）', () => {
  it('形の確認：AIが「承認済み」を付けた・根拠が無い推定・確認元が無い確認済みを見つける', () => {
    expect(checkShape(PATCH)).toEqual([]);
    const bad = structuredClone(PATCH);
    bad.ops[0]!.prov = { state: 'approved' };
    bad.ops[2]!.prov = { state: 'inferred' };
    bad.ops[1]!.prov = { state: 'observed' };
    expect(checkShape(bad).map((p) => p.index)).toEqual([0, 1, 2]);
    expect(checkShape({ hello: 1 })[0]?.message).toMatch(/schemaVersion/);
  });

  it('写っていない分岐・タグ・配信のつながりは、すべて「AI推定」になっている', () => {
    for (const op of PATCH.ops.filter((o) => /^\/(tags|cvs|routes|personas|narration)\//.test(o.path))) expect(op.prov.state).toBe('inferred');
    const flowNodes = PATCH.ops.filter((o) => o.path.startsWith('/flows')).flatMap((o) => (o.value as { nodes: { type: string; prov?: { state: string } }[] }).nodes);
    for (const n of flowNodes.filter((x) => x.type === 'branch' || x.type === 'wait' || x.type === 'jump')) expect(n.prov?.state ?? 'inferred(親の提案)').not.toBe('observed');
  });

  it('取り込む前の下見：元の案件は変わらず、選んだ分だけ当たる', () => {
    const base = store.create('preview', 'P');
    const r = dryRun(base, PATCH, new Set([0, 1]));
    expect(r.project.brand.companyName).toBe('ヘアサロンサンプル（仮）');
    expect(base.brand.companyName).toBe('P');
    expect(r.stale).toBe(true); // 提案は別の空の案件（ハッシュ）で作ったもの
  });

  it('却下：フローのノードを抜くと、出入りする線も一緒に消える', () => {
    const p = dryRun(store.create('rm', 'R'), PATCH, all).project;
    const fi = p.flows.findIndex((f) => f.id === 'flow_welcome');
    const ni = p.flows[fi]!.nodes.findIndex((n) => n.id === 'n_b_visit');
    const next = removeAt(p, `/flows/${fi}/nodes/${ni}`);
    expect(next.flows[fi]!.nodes.some((n) => n.id === 'n_b_visit')).toBe(false);
    expect(next.flows[fi]!.edges.some((e) => e.from.nodeId === 'n_b_visit' || e.to === 'n_b_visit')).toBe(false);
    expect(removeAt(p, '/')).toEqual(p); // 案件そのものは消さない
  });
});

describe('最終統合：スクショ → 提案 → 取り込み → 承認 → デモ → 営業用HTML（架空の美容サロン）', () => {
  let project: Project;

  it('1. 新しい案件に提案を取り込むと、確認済みとAI推定に分かれ、AI推定はすべて承認待ち（重大エラー）になる', () => {
    const base = store.create('beauty', 'サンプルビューティー（仮）');
    const r = dryRun(base, PATCH, all);
    expect(r.failed).toEqual([]);
    const errors = validate(r.project).filter((i) => i.level === 'error');
    expect(errors.length).toBeGreaterThan(0);
    expect(errors.every((i) => i.code === 'prov.unapproved')).toBe(true);
    expect(countProv(r.project)).toEqual({ observed: 11, inferred: 41, approved: 0 });
    store.save('beauty', r.project);
    project = store.read('beauty');
  });

  it('2. 承認待ちが残っていると、営業用HTMLは書き出せない', async () => {
    const r = await buildExport(project, { readAsset: () => null, confirmFictional: true });
    expect(r.ok).toBe(false);
    expect(r.html).toBeUndefined();
    expect(new Set(r.issues.filter((i) => i.level === 'error').map((i) => i.code))).toEqual(new Set(['prov.unapproved']));
  });

  it('3. 人が承認・却下・修正すると、承認待ちが無くなり、重大エラーも警告も無い', () => {
    // 修正：仮の文面を書き換えてから承認
    let p = structuredClone(project);
    const mi = p.messages.findIndex((m) => m.id === 'msg_coupon_repeat');
    p.messages[mi]!.bubbles = [{ kind: 'text', text: 'いつもご来店ありがとうございます。' }];
    p = approve(p, `/messages/${mi}`);
    // 却下：説明文1件
    const ni = p.narration.findIndex((n) => n.id === 'nar_stop');
    p = removeAt(p, `/narration/${ni}`);
    // 残りは確かめて承認
    expect(listProv(p, 'inferred').length).toBe(39);
    p = approve(p, '');
    expect(countProv(p)).toEqual({ observed: 11, inferred: 0, approved: 40 });
    expect(validate(p)).toEqual([]);
    expect(personaCoverage(p)).toEqual([]);
    store.save('beauty', p, store.read('beauty').meta.updatedAt);
    project = store.read('beauty');
  });

  it('4. 台本つきの全員を自動再生すると、業種が違っても同じエンジンで出し分け・停止・自動応答が動く', () => {
    const { demo, world } = playToEnd(project);
    const first = byPersona(world, demo, 'p_first_cut');
    const repeat = byPersona(world, demo, 'p_repeat_color');
    const perm = byPersona(world, demo, 'p_first_perm');
    expect(first.deliveries.map((d) => d.messageId)).toEqual(['msg_welcome', 'msg_coupon_first', 'msg_coupon_first', 'msg_style', 'msg_book_thanks']);
    expect(repeat.deliveries.map((d) => d.messageId)).toEqual(['msg_welcome', 'msg_coupon_repeat', 'msg_reminder']);
    expect(repeat.deliveries.at(-1)!.at).toBe(28 * DAY + 10 * 60);
    expect(perm.flows['flow_first']!.status).toBe('stopped'); // 予約で初来店向けが止まる
    // 自動応答（頼まれた返信）は、1時間前に同じクーポンを配信していても届く（Phase 6 で直した不具合）
    expect(first.deliveries.filter((d) => d.messageId === 'msg_coupon_first').map((d) => d.count)).toEqual([1, 0]);
    // 「なぜ」は営業先に分かる日本語
    const why = world.decisions.find((d) => d.userId === first.id && d.nodeId === 'n_b_visit' && d.outcome === 'skipped')!;
    expect(explainDecision(why, new Labels(project), project, world)).toEqual({
      tone: 'ng',
      title: '「来店回数」で「2回目以上」には進みませんでした',
      reason: 'タグ「来店：2回目以上」が付いていないため',
    });
  });

  it('5. 営業用HTMLを書き出せ、出所・根拠・伏せた固有名詞は入らず、HTMLの中のエンジンが同じ結果を出す', async () => {
    const r = await buildExport(project, { readAsset: () => null, confirmFictional: true });
    expect(r.ok).toBe(true);
    for (const s of ['"prov"', 'rationale', 'screenshot:', 'LUMI', 'lumiere', '渋谷']) expect(r.html).not.toContain(s);
    const ctx: { LineSim?: { simulate: (p: Project, s: unknown, until: number) => WorldState } } = {};
    runInNewContext(`${await bundlePlayer()};this.LineSim=LineSim;`, ctx);
    const session = sessionFromPersonas(project);
    expect(JSON.stringify(ctx.LineSim!.simulate(sanitize(project), session, 30 * DAY))).toBe(JSON.stringify(reduce(project, session, { until: 30 * DAY })));
  }, 30_000);

  it('6. 最初の業種（人材の検証案件）も、同じ版のエンジン・書き出しで壊れていない', async () => {
    const staff = sample();
    expect(validate(staff)).toEqual([]);
    expect(personaCoverage(staff)).toEqual([]);
    expect((await buildExport(staff, { readAsset: () => null, confirmFictional: true })).ok).toBe(true);
  });

  it('7. エンジンには業種の言葉が入っていない（日本語・禁止API・外部 import が0件）', () => {
    const dir = join(ROOT, 'packages', 'engine', 'src');
    const problems = readdirSync(dir).filter((f) => f.endsWith('.ts')).flatMap((f) => scanEngineSource(readFileSync(join(dir, f), 'utf8'), f));
    expect(problems).toEqual([]);
  });
});
