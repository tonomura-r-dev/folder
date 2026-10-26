// シミュレーション結果を、比較しやすい素の形にまとめる。CLI の --json と、通しテストのスナップショットで使う。

import type { Project, WorldState } from '../packages/engine/src/index.ts';
import { describeDecision, fmtTime, friendText, Labels } from '../packages/player/src/describe.ts';

export type UserSummary = {
  name: string;
  friend: string;
  route: string;
  tags: string[];
  answers: Record<string, string>;
  cvs: Record<string, number>;
  messageCount: number;
  flows: Record<string, string>;
  /** 時刻順の出来事（届いた・届かない・分岐・停止） */
  log: string[];
};

const FLOW_STATUS = { running: '実行中', waiting: '待機中', ended: '終了', stopped: '停止' } as const;

export function summarize(project: Project, world: WorldState): Record<string, UserSummary> {
  const L = new Labels(project);
  const out: Record<string, UserSummary> = {};
  // 並びはペルソナの定義順（案件で決めた順番で見せる）
  const order = project.personas.map((p) => p.id);
  const users = Object.values(world.users).sort((a, b) => order.indexOf(a.personaId) - order.indexOf(b.personaId));
  for (const user of users) {
    const answers: Record<string, string> = {};
    for (const [sv, qs] of Object.entries(user.answers))
      for (const [q, ids] of Object.entries(qs)) answers[L.questionText(sv, q)] = L.choices(sv, q, ids);
    const flows: Record<string, string> = {};
    for (const f of project.flows) {
      const c = user.flows[f.id];
      flows[f.label] = c ? FLOW_STATUS[c.status] + (c.runs > 1 ? `（${c.runs}回目）` : '') : '未開始';
    }
    const log = world.decisions
      .filter((d) => d.userId === user.id && d.kind !== 'trigger' && !(d.kind === 'autoReply' && d.outcome === 'taken'))
      .map((d) => `${fmtTime(d.at)}　${describeDecision(d, L, project)}`);
    out[user.id] = {
      name: user.name,
      friend: friendText(user.friendState),
      route: L.route(user.route),
      tags: Object.keys(user.tags).map(L.tag),
      answers,
      cvs: Object.fromEntries(Object.entries(user.cvs).map(([id, v]) => [L.cv(id), v.count])),
      messageCount: user.messageCount,
      flows,
      log,
    };
  }
  return out;
}
