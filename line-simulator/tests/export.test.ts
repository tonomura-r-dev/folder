// Phase 5：営業用HTMLの書き出し（tools/export-lib.ts）と、書き出したHTMLの中のエンジンが編集ツールと同じ結果を出すこと。

import sharp from 'sharp';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import type { Project, WorldState } from '../packages/engine/src/index.ts';
import { reduce, sessionFromPersonas } from '../packages/engine/src/index.ts';
import { initialDemo } from '../packages/player/src/controller.ts';
import type { ExportData } from '../packages/player/src/standalone.ts';
import { bundlePlayer } from '../tools/bundle.ts';
import { buildExport, exportFileName, MAX_BYTES, safeJson, sanitize, scanExport, uniqueName, WARN_BYTES } from '../tools/export-lib.ts';
import { areaBranch, DAY, sample } from './helpers.ts';

const noAssets = { readAsset: () => null, confirmFictional: true };
const dataOf = (html: string) => JSON.parse(html.match(/<script type="application\/json" id="linesim-data">([\s\S]*?)<\/script>/)![1]!) as ExportData;
const withAsset = (p: Project, path: string): Project => ({ ...p, assets: [{ id: 'img1', kind: 'image', path, width: 1, height: 1, alt: 'テスト画像' }] });

describe('書き出し：検証', () => {
  it('検証案件は重大エラー・警告なしで書き出せ、小さい（5MB未満）', async () => {
    const r = await buildExport(sample(), noAssets);
    expect(r.ok).toBe(true);
    expect(r.issues).toEqual([]);
    expect(r.bytes).toBeLessThan(WARN_BYTES);
    expect(r.fileName).toMatch(/^サンプルスタッフ株式会社（仮）_デモ_\d{8}\.html$/);
  });

  it('架空データの確認が無ければ書き出さない', async () => {
    const r = await buildExport(sample(), { ...noAssets, confirmFictional: false });
    expect(r.ok).toBe(false);
    expect(r.html).toBeUndefined();
    expect(r.issues.map((i) => i.code)).toContain('export.notConfirmed');
  });

  it('重大エラー（存在しないノードへの線・不正な条件・無限ループ）があれば書き出さない', async () => {
    const p = sample();
    p.flows[0]!.edges[0]!.to = 'n_nowhere';
    areaBranch(p).node.cases[0]!.when = { op: 'and', children: [] };
    const t = p.flows[1]!;
    t.nodes.push({ id: 'x1', type: 'tag', add: ['applied'], position: { x: 0, y: 0 } }, { id: 'x2', type: 'tag', add: ['applied'], position: { x: 0, y: 0 } });
    t.edges.push({ id: 'ex1', from: { nodeId: 'x1', port: 'out' }, to: 'x2' }, { id: 'ex2', from: { nodeId: 'x2', port: 'out' }, to: 'x1' });
    const r = await buildExport(p, noAssets);
    expect(r.ok).toBe(false);
    expect(r.html).toBeUndefined();
    expect(r.issues.filter((i) => i.level === 'error').map((i) => i.code)).toEqual(expect.arrayContaining(['edge.danglingTarget', 'expr.invalid', 'flow.zeroTimeCycle']));
  });

  it('警告だけなら書き出せる（使われていないタグ）', async () => {
    const p = sample();
    p.tags.push({ id: 'tag_unused', label: '未使用' });
    const r = await buildExport(p, noAssets);
    expect(r.ok).toBe(true);
    expect(r.issues.map((i) => i.code)).toEqual(['tag.unused']);
  });

  it('画像ファイルが無ければ書き出さない', async () => {
    const r = await buildExport(withAsset(sample(), 'assets/none.png'), noAssets);
    expect(r.ok).toBe(false);
    expect(r.issues.map((i) => i.code)).toContain('asset.fileMissing');
  });
});

describe('書き出し：中身', () => {
  it('社内メモ・AIの出所メモは、どの深さにあっても入らない', async () => {
    const p = sample();
    p._internal = { memo: '社外秘：この案件の粗利は３割' };
    p.flows[0]!._internal = { memo: '担当メモ：先方は東京だけ見ている' };
    p.flows[0]!.nodes[1]!.prov = { state: 'approved', rationale: 'AIが推定した根拠の文章' };
    const r = await buildExport(p, noAssets);
    expect(r.ok).toBe(true);
    for (const s of ['社外秘', '担当メモ', 'AIが推定した根拠', '_internal', '"prov"']) expect(r.html).not.toContain(s);
    expect(JSON.stringify(sanitize(p))).not.toMatch(/_internal|"prov"/);
  });

  it('文面に </script> や <!-- が入っていても、HTMLが壊れず元の文字のまま表示される', async () => {
    const p = sample();
    const evil = '</script><script>alert(1)</script><!-- & 〈テスト〉';
    p.messages[0]!.bubbles = [{ kind: 'text', text: evil }];
    const r = await buildExport(p, noAssets);
    expect(r.ok).toBe(true);
    expect(r.html!.match(/<script\b/g)).toHaveLength(2); // 案件データと、束ねたJSの2つだけ
    const back = dataOf(r.html!).project.messages[0]!.bubbles[0]!;
    expect(back).toEqual({ kind: 'text', text: evil });
    expect(JSON.parse(`"${safeJson('</script>').slice(1, -1)}"`)).toBe('</script>');
  });

  it('外から何も読まない：通信を禁止する設定が入り、外部URL・編集ツールのコードが無い', async () => {
    const r = await buildExport(sample(), noAssets);
    expect(r.html).toContain(`content="default-src 'none'; img-src data:;`);
    expect(scanExport(r.html!, sample())).toEqual([]);
    expect(r.html).not.toMatch(/react-dom|@xyflow|\/api\/projects|<link\b/);
  });

  it('検査は、外部の読み込み・APIキー・社内メモの混入を見つける', () => {
    const p = { ...sample(), _internal: { memo: '社外秘の数字' } };
    const base = '<!doctype html><script type="application/json" id="linesim-data">{"assets":{}}</script>';
    expect(scanExport(`${base}<img src="https://example.com/x.png">`, p)[0]?.code).toBe('export.external');
    expect(scanExport(`${base}<script>const k="sk-abcdefghijklmnopqrstuvwxyz123"</script>`, p)[0]?.code).toBe('export.secret');
    expect(scanExport(`${base}<p>社外秘の数字</p>`, p)[0]?.code).toBe('export.leak');
    expect(scanExport(`${base}<script>ReactFlow()</script>`, p)[0]?.code).toBe('export.editorCode');
  });

  it('大きい画像は横800pxのWebPに縮めて埋め込み、小さい画像はそのまま埋め込む', async () => {
    const big = await sharp({ create: { width: 1600, height: 1000, channels: 3, background: '#3467b2' } }).png().toBuffer();
    const small = await sharp({ create: { width: 200, height: 100, channels: 3, background: '#ffffff' } }).png().toBuffer();
    const r1 = await buildExport(withAsset(sample(), 'assets/big.png'), { readAsset: () => big, confirmFictional: true });
    const url1 = dataOf(r1.html!).assets['img1']!;
    expect(url1.startsWith('data:image/webp;base64,')).toBe(true);
    expect((await sharp(Buffer.from(url1.split(',')[1]!, 'base64')).metadata()).width).toBe(800);
    expect(r1.issues.map((i) => i.code)).toContain('export.imageResized');
    const r2 = await buildExport(withAsset(sample(), 'assets/small.png'), { readAsset: () => small, confirmFictional: true });
    expect(dataOf(r2.html!).assets['img1']!.startsWith('data:image/png;base64,')).toBe(true);
  });

  it('5MBを超えると警告、15MBを超えると書き出さない', async () => {
    const gif = (mb: number) => Buffer.alloc(Math.round(mb * 1024 * 1024), 7); // GIF は縮めずに埋め込む
    const warn = await buildExport(withAsset(sample(), 'assets/a.gif'), { readAsset: () => gif(4.5), confirmFictional: true });
    expect(warn.ok).toBe(true);
    expect(warn.bytes).toBeGreaterThan(WARN_BYTES);
    expect(warn.issues.map((i) => i.code)).toContain('export.large');
    const over = await buildExport(withAsset(sample(), 'assets/a.gif'), { readAsset: () => gif(12), confirmFictional: true });
    expect(over.ok).toBe(false);
    expect(over.bytes).toBeGreaterThan(MAX_BYTES);
    expect(over.html).toBeUndefined();
    expect(over.issues.map((i) => i.code)).toContain('export.tooLarge');
  }, 30_000);

  it('ファイル名に Windows で使えない文字を入れない', () => {
    const p = sample();
    p.meta.name = 'A/B:C*D?"E"<F>|G';
    expect(exportFileName(p, new Date(2026, 9, 27))).toBe('A_B_C_D__E__F__G_デモ_20261027.html');
  });

  it('同じ名前のファイルがあっても上書きせず、_2・_3 を付ける', () => {
    const have = new Set(['x_デモ_20261027.html', 'x_デモ_20261027_2.html']);
    expect(uniqueName('x_デモ_20261027.html', (n) => have.has(n))).toBe('x_デモ_20261027_3.html');
    expect(uniqueName('y.html', () => false)).toBe('y.html');
  });
});

describe('書き出したHTMLのエンジンと、編集ツールのエンジンが同じ結果を出す', () => {
  it('束ねたJSを単体で動かし、4人をDay 30まで流した結果（配信・タグ・履歴・除外理由）が一致する', async () => {
    const code = await bundlePlayer();
    const ctx: { LineSim?: { simulate: (p: Project, s: unknown, until: number) => WorldState } } = {};
    runInNewContext(`${code};this.LineSim=LineSim;`, ctx);
    const project = sanitize(sample());
    const session = sessionFromPersonas(project);
    const fromHtml = ctx.LineSim!.simulate(project, session, 30 * DAY);
    const fromEditor = reduce(sample(), session, { until: 30 * DAY });
    expect(JSON.stringify(fromHtml)).toBe(JSON.stringify(fromEditor));
    expect(fromHtml.decisions.some((d) => d.outcome !== 'taken')).toBe(true); // 除外理由も比べている
  }, 30_000);
});

describe('デモの始まり方（案件設定）', () => {
  it('既定は1人を手で操作。台本つきを選ぶと、台本のあるペルソナ全員', () => {
    const p = sample();
    expect(initialDemo(p).session.users).toHaveLength(1);
    expect(initialDemo(p).scripted).toEqual([]);
    p.settings.demoStart = { mode: 'single', personaId: 'p_osaka' };
    expect(initialDemo(p).session.users[0]!.personaId).toBe('p_osaka');
    p.settings.demoStart = { mode: 'scripted' };
    expect(initialDemo(p).session.users.map((u) => u.personaId)).toEqual(['p_tokyo', 'p_osaka', 'p_applied', 'p_notyet', 'p_late']);
  });
});
