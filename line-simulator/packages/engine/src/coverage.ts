// ペルソナを台本どおりに流したときの確認（検証の警告の一部）。
// 静的な検証（validate）では分からない「どのペルソナもそのシナリオに入らない」を、実際に流して見つける。

import type { Issue, Project } from './schema.ts';
import { MINUTES_PER_DAY } from './schema.ts';
import { reduce, sessionFromPersonas } from './reduce.ts';

/** 台本のあるペルソナを untilDay 日まで流し、一度も始まらなかったシナリオを警告にする */
export function personaCoverage(project: Project, untilDay = 30): Issue[] {
  const scripted = project.personas.filter((p) => (p.script ?? []).length > 0);
  if (scripted.length === 0) return [];
  const world = reduce(project, sessionFromPersonas(project, scripted.map((p) => p.id)), { until: untilDay * MINUTES_PER_DAY });
  const issues: Issue[] = [];
  project.flows.forEach((flow, fi) => {
    // 何も配信しない補助（タグを付けてページを開く等）は、デモで見せるシナリオではないので数えない
    const helper = !flow.nodes.some((n) => n.type === 'message' || n.type === 'image' || n.type === 'choice' || n.type === 'survey' || n.type === 'jump');
    if (helper) return;
    const started = Object.values(world.users).some((u) => u.flows[flow.id]);
    if (!started) issues.push({ level: 'warning', code: 'persona.noneMatch', path: `/flows/${fi}`, params: { id: flow.id, days: untilDay } });
  });
  return issues;
}
