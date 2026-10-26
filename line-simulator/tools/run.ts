// CLI：案件をペルソナ全員で Day N まで流し、配信・分岐・停止の結果を表で出す。
//   npm run sim -- sample-staff --until-day 30
//   npm run sim -- sample-staff --json          （比較用の素のJSON）
//   npm run sim -- sample-staff --user p_osaka   （1人だけ）

import { hasErrors, MINUTES_PER_DAY, reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { assetExists, loadProject } from './project-io.ts';
import { summarize } from './summary.ts';

const args = process.argv.slice(2);
const slug = args.find((a) => !a.startsWith('--')) ?? 'sample-staff';
const opt = (name: string) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : undefined;
};
const untilDay = Number(opt('until-day') ?? 30);
const onlyUser = opt('user');
const asJson = args.includes('--json');

const project = loadProject(slug);
const issues = validate(project, { assetExists: assetExists(slug) });
const errors = issues.filter((i) => i.level === 'error');
const warnings = issues.filter((i) => i.level === 'warning');

if (hasErrors(issues)) {
  console.error(`重大エラー ${errors.length}件のため実行しません。`);
  for (const i of errors) console.error(`  [error] ${i.code}  ${i.path}  ${i.params ? JSON.stringify(i.params) : ''}`);
  process.exit(1);
}

const started = performance.now();
const world = reduce(project, sessionFromPersonas(project), { until: untilDay * MINUTES_PER_DAY });
const elapsedMs = performance.now() - started;
const summary = summarize(project, world);
const users = Object.entries(summary).filter(([id]) => !onlyUser || id === onlyUser);

if (asJson) {
  console.log(JSON.stringify(Object.fromEntries(users), null, 2));
  process.exit(0);
}

const line = '─'.repeat(72);
console.log(`\n${project.meta.name}　Day ${untilDay} まで（${users.length}人・${elapsedMs.toFixed(0)}ms）`);
console.log(`検証：重大エラー ${errors.length}件／警告 ${warnings.length}件`);
for (const w of warnings) console.log(`  [warning] ${w.code}  ${w.path}  ${w.params ? JSON.stringify(w.params) : ''}`);
for (const w of world.warnings) console.log(`  [runtime] ${w.code}  ${w.userId}  ${w.ref}`);

for (const [, u] of users) {
  console.log(`\n${line}\n■ ${u.name}`);
  console.log(`  友だち状態：${u.friend}　経路：${u.route}　通数：${u.messageCount}`);
  console.log(`  タグ　　　：${u.tags.join('、') || '—'}`);
  console.log(`  回答　　　：${Object.entries(u.answers).map(([q, a]) => `${q}=${a}`).join('、') || '—'}`);
  console.log(`  CV　　　　：${Object.entries(u.cvs).map(([c, n]) => `${c}×${n}`).join('、') || '—'}`);
  console.log('  シナリオ　：');
  for (const [f, s] of Object.entries(u.flows)) console.log(`    ${s.padEnd(8, '　')} ${f}`);
  console.log('  出来事　　：');
  for (const l of u.log) console.log(`    ${l}`);
}
console.log(`\n${line}`);
