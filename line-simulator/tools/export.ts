// CLI：営業用HTMLを書き出す。
//   npm run export -- sample-staff --confirm-fictional
//   npm run export -- sample-staff --confirm-fictional --out <フォルダ>
// --confirm-fictional は「架空または匿名化したデータだけであることを確認した」という意味（無いと書き出さない）。
// 既定の保存先は Downloads 直下。

import { existsSync, mkdirSync, readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { describeIssue } from '../packages/player/src/issues.ts';
import { writeAtomic } from './project-store.ts';
import { buildExport, uniqueName } from './export-lib.ts';
import { loadProject, projectDir } from './project-io.ts';

export const defaultOutDir = () => join(homedir(), 'Downloads');

const args = process.argv.slice(2);
const slug = args.find((a) => !a.startsWith('--')) ?? 'sample-staff';
const outIdx = args.indexOf('--out');
const outDir = outIdx >= 0 ? (args[outIdx + 1] ?? defaultOutDir()) : defaultOutDir();

const project = loadProject(slug);
const result = await buildExport(project, {
  readAsset: (p) => {
    const f = join(projectDir(slug), p);
    return existsSync(f) ? readFileSync(f) : null;
  },
  confirmFictional: args.includes('--confirm-fictional'),
});

for (const i of result.issues) console.log(`  [${i.level === 'error' ? '重大' : '警告'}] ${describeIssue(i)}　${i.path}`);
if (!result.ok || !result.html) {
  console.error(`\n書き出しませんでした（重大エラー ${result.issues.filter((i) => i.level === 'error').length}件）。`);
  if (!args.includes('--confirm-fictional')) console.error('架空・匿名化データだけであることを確認したら --confirm-fictional を付けてください。');
  process.exit(1);
}
mkdirSync(outDir, { recursive: true });
const file = join(outDir, uniqueName(result.fileName, (n) => existsSync(join(outDir, n))));
writeAtomic(file, result.html);
console.log(`\n書き出しました：${file}（${(result.bytes / 1024).toFixed(0)}KB・警告 ${result.issues.filter((i) => i.level === 'warning').length}件）`);
