// 案件フォルダの読み込み。エンジンはファイルを読まないので、ディスクに触るのはツール側だけ。

import { existsSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Project } from '../packages/engine/src/index.ts';

export const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const projectDir = (slug: string) => join(ROOT, 'projects', slug);

export function loadProject(slug: string): Project {
  const file = join(projectDir(slug), 'project.json');
  if (!existsSync(file)) throw new Error(`案件が見つかりません: ${file}`);
  return JSON.parse(readFileSync(file, 'utf8')) as Project;
}

/** 画像の存在確認（案件フォルダからの相対パス） */
export const assetExists = (slug: string) => (path: string) => existsSync(join(projectDir(slug), path));
