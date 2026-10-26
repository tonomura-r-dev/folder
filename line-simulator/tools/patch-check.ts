// CLI：AIの提案（シナリオ差分）を、案件に当てる前に確かめる。Claude Code が提案を書いたら必ず流す。
//   npm run patch:check -- <案件> <提案ファイル名 or パス>
// 形の誤り・当たらない場所・（承認待ち以外の）重大エラーがあれば終了コード1。
// 「AI推定が承認されていない」は正常（人が承認するまで書き出せない状態）なので、件数だけ出す。

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { ScenarioPatch } from '../packages/engine/src/index.ts';
import { checkShape, countProv, dryRun, opLabel, summarizeValue } from '../packages/ai/src/review.ts';
import { describeIssue } from '../packages/player/src/issues.ts';
import { loadProject, projectDir } from './project-io.ts';

const [slug = '', file = ''] = process.argv.slice(2).filter((a) => !a.startsWith('--'));
if (!slug || !file) {
  console.error('使い方：npm run patch:check -- <案件> <提案ファイル名 or パス>');
  process.exit(2);
}
const path = existsSync(file) ? file : join(projectDir(slug), 'patches', file);
const data = JSON.parse(readFileSync(path, 'utf8')) as ScenarioPatch;

const shape = checkShape(data);
if (shape.length) {
  console.error(`提案ファイルの形に問題があります（${shape.length}件）`);
  for (const s of shape) console.error(`  ${s.index === null ? '全体' : `提案${s.index + 1}`}：${s.message}`);
  process.exit(1);
}

const project = loadProject(slug);
const all = new Set(data.ops.map((_, i) => i));
const r = dryRun(project, data, all);
const errors = r.issues.filter((i) => i.level === 'error' && i.code !== 'prov.unapproved');
const pending = r.issues.filter((i) => i.code === 'prov.unapproved').length;
const c = countProv(r.project);

console.log(`\n${data.summary}`);
console.log(`提案 ${data.ops.length}件（確認済み ${data.ops.filter((o) => o.prov.state === 'observed').length}・AI推定 ${data.ops.filter((o) => o.prov.state === 'inferred').length}）${r.stale ? '　※提案を作った後に案件が変わっています' : ''}`);
data.ops.forEach((op, i) => console.log(`  ${String(i + 1).padStart(2)}. [${op.prov.state === 'observed' ? '確認済み' : 'AI推定'}] ${opLabel(op.op)} ${op.path}　${summarizeValue(op.value)}`));
console.log(`\n全部当てたとき：確認済み ${c.observed}・AI推定（承認待ち）${c.inferred}　/　承認待ちの重大エラー ${pending}件（正常）`);
for (const f of r.failed) console.log(`  [当たらない] 提案${f.index + 1}：${f.message}`);
for (const i of errors) console.log(`  [重大] ${describeIssue(i)}　${i.path}`);
for (const i of r.issues.filter((x) => x.level === 'warning')) console.log(`  [警告] ${describeIssue(i)}　${i.path}`);
if (r.failed.length || errors.length) {
  console.error(`\n直してください：当たらない提案 ${r.failed.length}件・重大エラー ${errors.length}件`);
  process.exit(1);
}
console.log('\nOK：編集ツールの「AI提案」で読み込めます。');
