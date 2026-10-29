// 案件JSONを ID で引けるようにする索引。ノードID・ボタンIDは案件全体で一意（validate が保証）。

import type { Bubble, Button, Flow, FlowNode, Id, MessageDef, Project, RichMenu, Survey, WebPage, Edge } from './schema.ts';

/** 吹き出しにあるボタン（カルーセルは全カードの分） */
export function bubbleButtons(b: Bubble): Button[] {
  if (b.kind === 'buttons' || b.kind === 'card') return b.buttons;
  if (b.kind === 'carousel') return b.cards.flatMap((c) => c.buttons);
  return [];
}

export type ProjectIndex = {
  flows: Map<Id, Flow>;
  nodes: Map<Id, { flow: Flow; node: FlowNode }>;
  edges: Map<string, Edge>;
  messages: Map<Id, MessageDef>;
  surveys: Map<Id, Survey>;
  pages: Map<Id, WebPage>;
  richMenus: Map<Id, RichMenu>;
  buttons: Map<Id, Button>;
  /** 優先度 → ID の順に並べたフロー（トリガー判定の順番） */
  flowOrder: Flow[];
};

export const edgeKey = (flowId: Id, nodeId: Id, port: string) => `${flowId}\u0000${nodeId}\u0000${port}`;

export const byId = (a: string, b: string) => (a < b ? -1 : a > b ? 1 : 0);

export function buildIndex(project: Project): ProjectIndex {
  const flows = new Map(project.flows.map((f) => [f.id, f]));
  const nodes = new Map<Id, { flow: Flow; node: FlowNode }>();
  const edges = new Map<string, Edge>();
  for (const flow of project.flows) {
    for (const node of flow.nodes) nodes.set(node.id, { flow, node });
    for (const edge of flow.edges) edges.set(edgeKey(flow.id, edge.from.nodeId, edge.from.port), edge);
  }
  const buttons = new Map<Id, Button>();
  for (const m of project.messages) for (const b of m.bubbles) for (const btn of bubbleButtons(b)) buttons.set(btn.id, btn);
  const flowOrder = [...project.flows].sort((a, b) => a.priority - b.priority || byId(a.id, b.id));
  return {
    flows,
    nodes,
    edges,
    messages: new Map(project.messages.map((m) => [m.id, m])),
    surveys: new Map(project.surveys.map((s) => [s.id, s])),
    pages: new Map(project.webPages.map((p) => [p.id, p])),
    richMenus: new Map(project.richMenus.map((r) => [r.id, r])),
    buttons,
    flowOrder,
  };
}
