// テスト用の小さな案件を組み立てる道具。

import type { Edge, Flow, FlowNode, MessageDef, Project } from '../packages/engine/src/index.ts';
import { loadProject } from '../tools/project-io.ts';

export const sample = () => loadProject('sample-staff');

/** 検証案件の「希望エリア」分岐ノードと、その場所（どのシナリオにあっても探す） */
export function areaBranch(p: Project) {
  for (const [fi, flow] of p.flows.entries()) {
    const ni = flow.nodes.findIndex((n) => n.id === 'n_w_area');
    const node = flow.nodes[ni];
    if (node?.type === 'branch') return { flow, node, path: `/flows/${fi}/nodes/${ni}` };
  }
  throw new Error('n_w_area がありません');
}

export const DAY = 1440;

export function baseProject(over: Partial<Project> = {}): Project {
  return {
    schemaVersion: 1,
    meta: { id: 't', name: 't', slug: 't', createdAt: '', updatedAt: '' },
    brand: { companyName: 'c', accountName: 'a', colors: { primary: '#000', accent: '#0f0' } },
    settings: { dedupeWindowMinutes: 0 },
    assets: [],
    tags: [{ id: 'tag_a', label: 'A' }],
    cvs: [{ id: 'cv1', label: 'cv' }],
    routes: [{ id: 'r1', label: 'r' }],
    messages: [msg('m1'), msg('m2')],
    richMenus: [],
    autoReplies: [],
    surveys: [],
    webPages: [],
    flows: [],
    personas: [{ id: 'u1', name: 'u1', attrs: {} }],
    narration: [],
    ...over,
  };
}

export const msg = (id: string): MessageDef => ({ id, bubbles: [{ kind: 'text', text: id }] });

let seq = 0;
const pos = { x: 0, y: 0 };
export const node = (n: Omit<FlowNode, 'position'> & Record<string, unknown>): FlowNode => ({ position: pos, ...n }) as FlowNode;
export const edge = (from: string, to: string, port = 'out'): Edge => ({ id: `e${seq++}`, from: { nodeId: from, port }, to });

/** entry → nodes を一直線につないだフローを作る */
export function chain(id: string, nodes: FlowNode[], over: Partial<Flow> = {}): Flow {
  const entry = node({ id: `${id}_entry`, type: 'entry' });
  const all = [entry, ...nodes];
  const edges = all.slice(0, -1).map((n, i) => edge(n.id, all[i + 1]!.id));
  return { id, label: id, trigger: { kind: 'friendAdd' }, priority: 10, concurrency: 'exclusive', reentry: 'never', nodes: all, edges, ...over };
}

export const codes = (issues: { code: string; level: string }[], level: 'error' | 'warning') =>
  issues.filter((i) => i.level === level).map((i) => i.code);
