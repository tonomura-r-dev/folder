// 保存まわり（tools/project-store.ts）。一時フォルダの中だけで動かす。

import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { validate } from '../packages/engine/src/index.ts';
import { HISTORY_KEEP, ProjectStore, StoreError } from '../tools/project-store.ts';
import { sample } from './helpers.ts';

let root: string;
let store: ProjectStore;
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'linesim-'));
  store = new ProjectStore(root);
});
afterEach(() => rmSync(root, { recursive: true, force: true }));

const fail = (fn: () => unknown) => {
  try {
    fn();
  } catch (e) {
    return e as StoreError;
  }
  throw new Error('失敗するはずが成功した');
};

describe('案件の保存', () => {
  it('新規作成した空の案件は、検証で重大エラーが出ない', () => {
    const p = store.create('new-1', '新しい案件');
    expect(p.meta.name).toBe('新しい案件');
    expect(validate(p).filter((i) => i.level === 'error')).toEqual([]);
  });

  it('保存のたびに履歴が1つ増え、過去の版に戻せる（戻したことも履歴に残る）', () => {
    let p = store.create('a', 'A');
    p = { ...p, brand: { ...p.brand, accountName: '版2' } };
    store.save('a', p, store.read('a').meta.updatedAt);
    const v2 = store.read('a');
    store.save('a', { ...v2, brand: { ...v2.brand, accountName: '版3' } }, v2.meta.updatedAt);
    const hist = store.history('a');
    expect(hist).toHaveLength(2); // 作成時は退避する前の版が無い
    expect(store.read('a').brand.accountName).toBe('版3');
    store.restore('a', hist[0]!.id); // 直前（版2）に戻す
    expect(store.read('a').brand.accountName).toBe('版2');
    expect(store.history('a')).toHaveLength(3);
  });

  it('履歴の日時は、その版を保存した時刻（置き換えられた時刻ではない）', () => {
    const p = store.create('a', 'A');
    const savedAt = p.meta.updatedAt;
    store.save('a', p, savedAt);
    expect(store.history('a')[0]!.savedAt).toBe(savedAt);
  });

  it('読み込んだ後に別の場所で更新されていたら、保存を断る', () => {
    const p = store.create('a', 'A');
    const opened = p.meta.updatedAt;
    store.save('a', { ...p, brand: { ...p.brand, accountName: '別タブ' } }, opened);
    const e = fail(() => store.save('a', { ...p, brand: { ...p.brand, accountName: '古い画面' } }, opened));
    expect(e.status).toBe(409);
    expect(store.read('a').brand.accountName).toBe('別タブ');
  });

  it('書き込み途中で落ちても（.tmp だけ残っても）、project.json は前の版のまま読める', () => {
    store.create('a', 'A');
    const file = join(root, 'projects', 'a', 'project.json');
    const before = readFileSync(file, 'utf8');
    writeFileSync(`${file}.tmp`, '{ "schemaVersion": 1, "meta": { "na'); // 書きかけ
    expect(readFileSync(file, 'utf8')).toBe(before);
    expect(store.read('a').meta.name).toBe('A');
    // 次の保存で .tmp は上書きされ、正常に保存できる
    store.save('a', store.read('a'), store.read('a').meta.updatedAt);
    expect(existsSync(`${file}.tmp`)).toBe(false);
  });

  it(`名前の無い履歴は ${HISTORY_KEEP} 件まで。名前を付けた版は消えない`, () => {
    store.create('a', 'A');
    for (let i = 0; i < 3; i++) store.save('a', store.read('a'));
    const keep = store.history('a').at(-1)!.id; // いちばん古い版
    store.label('a', keep, '提案前の版');
    for (let i = 0; i < HISTORY_KEEP + 5; i++) store.save('a', store.read('a'));
    const hist = store.history('a');
    expect(hist.filter((h) => !h.label)).toHaveLength(HISTORY_KEEP);
    expect(hist.find((h) => h.id === keep)?.label).toBe('提案前の版');
  });
});

describe('複製・読み込み・削除', () => {
  it('複製した案件を編集しても、元の案件は変わらない', () => {
    store.import('src', sample());
    const copy = store.duplicate('src', 'copy', '複製');
    store.save('copy', { ...copy, tags: [] }, copy.meta.updatedAt);
    expect(store.read('src').tags.length).toBeGreaterThan(0);
    expect(store.read('copy').tags).toEqual([]);
    expect(store.read('copy').meta).toMatchObject({ id: 'copy', slug: 'copy', name: '複製' });
    expect(store.history('copy')).toHaveLength(0 + 1); // 複製元の履歴は引き継がない（自分の保存1回分だけ）
  });

  it('JSONエクスポート → 別名でインポート → 中身が同じ', () => {
    store.import('src', sample());
    const exported = JSON.parse(JSON.stringify(store.read('src')));
    const imported = store.import('other', exported);
    const strip = (p: typeof imported) => ({ ...p, meta: { ...p.meta, id: '', slug: '', updatedAt: '' } });
    expect(strip(imported)).toEqual(strip(store.read('src')));
  });

  it('同じフォルダ名へのインポート・形式違いの JSON は断る', () => {
    store.import('src', sample());
    expect(fail(() => store.import('src', sample())).status).toBe(409);
    expect(fail(() => store.import('x', { hello: 1 } as never)).status).toBe(422);
  });

  it('削除はゴミ箱への移動。中身は残る', () => {
    store.create('a', 'A');
    const dest = store.trash('a');
    expect(store.exists('a')).toBe(false);
    expect(existsSync(join(dest, 'project.json'))).toBe(true);
    expect(readdirSync(join(root, '.trash'))).toHaveLength(1);
  });

  it('フォルダ名でパスの外に出られない', () => {
    for (const bad of ['../x', 'a/b', '..', 'A', '']) expect(fail(() => store.read(bad)).code).toBe('slug.invalid');
  });

  it('壊れた project.json の案件は、一覧から消さずに「読み込めません」と出す', () => {
    store.create('a', 'A');
    writeFileSync(join(root, 'projects', 'a', 'project.json'), '{broken');
    expect(store.list()[0]!.name).toContain('読み込めません');
    expect(fail(() => store.read('a')).status).toBe(422);
  });
});
