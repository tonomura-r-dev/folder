// エンジンの汎用性検査（NFR-07）。1件でも見つかれば終了コード1。中身は engine-scan.ts。

import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { scanEngineSource } from './engine-scan.ts';
import { ROOT } from './project-io.ts';

const ENGINE = join(ROOT, 'packages', 'engine', 'src');
const files = readdirSync(ENGINE, { recursive: true, encoding: 'utf8' })
  .filter((f) => f.endsWith('.ts'))
  .map((f) => join(ENGINE, f));
const problems = files.flatMap((f) => scanEngineSource(readFileSync(f, 'utf8'), relative(ROOT, f)));

if (problems.length) {
  console.error(`エンジンの検査：${problems.length}件\n${problems.map((p) => `  ${p}`).join('\n')}`);
  process.exit(1);
}
console.log(`エンジンの検査：OK（${files.length}ファイル・日本語0件・禁止API 0件・外部 import 0件）`);
