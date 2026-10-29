// 案件のフロー ⇄ React Flow のノード・線。React に依存しない純粋な変換だけを置く。

import type { Edge as RFEdge, Node as RFNode } from '@xyflow/react';
import { MarkerType } from '@xyflow/react';
import type { Flow, FlowNode, Id, Issue, NodeType, Project, Trigger } from '../../../engine/src/index.ts';
import { acceptsInput, bubbleButtons, portsOf } from '../../../engine/src/index.ts';
import { describeExpr, describeTrigger, fmtDuration, Labels } from '../../../player/src/describe.ts';

export { describeTrigger };

export const TYPE_LABEL: Record<NodeType, string> = {
  entry: '入口',
  message: '文面',
  image: '画像',
  richMenu: 'メニュー切替',
  choice: 'ボタン選択',
  survey: 'アンケート',
  tag: 'タグ',
  branch: '分岐',
  wait: '待機',
  page: 'ページ遷移',
  jump: '別シナリオへ',
  end: '終了',
  stopAll: '全停止',
};

/** パレットに並べる順（入口は1フローに1つなので出さない） */
export const PALETTE: NodeType[] = ['message', 'choice', 'survey', 'tag', 'branch', 'wait', 'image', 'richMenu', 'page', 'jump', 'end', 'stopAll'];

export type Port = { id: string; label: string; connected: boolean };
export type CardData = {
  node: FlowNode;
  type: NodeType;
  title: string;
  body: string;
  missing: boolean;
  ports: Port[];
  hasInput: boolean;
  errors: number;
  warnings: number;
};
export type CardNode = RFNode<CardData, 'card'>;

export function portLabel(node: FlowNode, port: string, project: Project): string {
  if (port === 'out') return '次へ';
  if (port === 'done') return '回答した';
  if (port === 'timeout') return '時間切れ';
  if (port === 'else') return 'それ以外';
  if (port.startsWith('case:') && node.type === 'branch') return node.cases.find((c) => `case:${c.id}` === port)?.label || port.slice(5);
  if (port.startsWith('btn:')) {
    const id = port.slice(4);
    for (const m of project.messages)
      for (const b of m.bubbles) for (const btn of bubbleButtons(b)) if (btn.id === id) return btn.label;
    return id;
  }
  return port;
}

/** ノードの見出しと中身（カードに出す要約） */
export function summarize(node: FlowNode, flow: Flow, project: Project): { title: string; body: string; missing: boolean } {
  const L = new Labels(project);
  const has = <T extends { id: Id }>(list: T[], id: Id) => list.some((x) => x.id === id);
  const title = node.label ?? '';
  switch (node.type) {
    case 'entry':
      return { title: title || 'はじまり', body: describeTrigger(flow.trigger, project), missing: false };
    case 'message':
    case 'choice':
      return has(project.messages, node.messageId)
        ? { title, body: L.message(node.messageId), missing: false }
        : { title, body: '文面が未選択です', missing: true };
    case 'image': {
      const a = project.assets.find((x) => x.id === node.assetId);
      return a ? { title, body: a.alt || a.path, missing: false } : { title, body: '画像が未選択です', missing: true };
    }
    case 'richMenu': {
      const m = project.richMenus.find((x) => x.id === node.richMenuId);
      return m ? { title, body: `「${m.label}」に切り替える`, missing: false } : { title, body: 'メニューが未選択です', missing: true };
    }
    case 'survey': {
      const s = project.surveys.find((x) => x.id === node.surveyId);
      return s
        ? { title, body: `${s.label}（${s.questions.length}問）${node.timeoutMinutes ? `\n${fmtDuration(node.timeoutMinutes)}で回答待ちを終える` : ''}`, missing: false }
        : { title, body: 'アンケートが未選択です', missing: true };
    }
    case 'tag': {
      const parts = [...(node.add ?? []).map((t) => `＋${L.tag(t)}`), ...(node.remove ?? []).map((t) => `－${L.tag(t)}`)];
      return parts.length ? { title, body: parts.join('\n'), missing: false } : { title, body: '付ける／外すタグが未設定です', missing: true };
    }
    case 'branch':
      return node.cases.length
        ? { title, body: node.cases.map((c) => `${c.label}：${safeExpr(c.when, L)}`).join('\n'), missing: false }
        : { title, body: '条件がありません', missing: true };
    case 'wait': {
      const w = node.wait;
      const body =
        w.kind === 'duration'
          ? `${fmtDuration(w.minutes)}待つ`
          : w.kind === 'dayAt'
            ? `${w.from === 'flowStart' ? 'シナリオ開始' : '友だち追加'}から ${w.day}日目 ${w.time}`
            : `操作を待つ（最大${fmtDuration(w.timeoutMinutes)}）`;
      return { title, body, missing: false };
    }
    case 'page':
      return has(project.webPages, node.pageId) ? { title, body: `「${L.page(node.pageId)}」を開く`, missing: false } : { title, body: 'ページが未選択です', missing: true };
    case 'jump':
      return has(project.flows, node.flowId) ? { title, body: `→ ${L.flow(node.flowId)}`, missing: false } : { title, body: '行き先のシナリオが未選択です', missing: true };
    case 'end':
      return { title, body: node.reason ?? 'このシナリオを終える', missing: false };
    case 'stopAll':
      return { title, body: node.reason ?? 'この人のシナリオをすべて止める', missing: false };
  }
}

function safeExpr(e: unknown, L: Labels): string {
  try {
    return describeExpr(e as never, L);
  } catch {
    return '（条件式が正しくありません）';
  }
}

export function toCards(flow: Flow, flowIndex: number, project: Project, issues: Issue[], selected?: Id): CardNode[] {
  const used = new Set(flow.edges.map((e) => `${e.from.nodeId}\u0000${e.from.port}`));
  return flow.nodes.map((node, ni) => {
    const path = `/flows/${flowIndex}/nodes/${ni}`;
    const mine = issues.filter((i) => i.path === path || i.path.startsWith(`${path}/`));
    const s = summarize(node, flow, project);
    return {
      id: node.id,
      type: 'card',
      position: node.position,
      selected: node.id === selected,
      deletable: node.type !== 'entry',
      data: {
        node,
        type: node.type,
        ...s,
        ports: portsOf(node, project).map((id) => ({ id, label: portLabel(node, id, project), connected: used.has(`${node.id}\u0000${id}`) })),
        hasInput: acceptsInput(node),
        errors: mine.filter((i) => i.level === 'error').length,
        warnings: mine.filter((i) => i.level === 'warning').length,
      },
    };
  });
}

export function toLines(flow: Flow, project: Project, selected?: Id): RFEdge[] {
  const ids = new Set(flow.nodes.map((n) => n.id));
  return flow.edges
    .filter((e) => ids.has(e.from.nodeId) && ids.has(e.to)) // 行き先の無い線は描けない（検証一覧に出る）
    .map((e) => {
      const from = flow.nodes.find((n) => n.id === e.from.nodeId)!;
      const label = e.from.port === 'out' ? undefined : portLabel(from, e.from.port, project);
      const on = e.id === selected;
      return {
        id: e.id,
        source: e.from.nodeId,
        sourceHandle: e.from.port,
        target: e.to,
        type: 'smoothstep',
        selected: on,
        reconnectable: true,
        markerEnd: { type: MarkerType.ArrowClosed, color: on ? '#3467b2' : '#8a93b0' },
        style: { stroke: on ? '#3467b2' : '#8a93b0', strokeWidth: on ? 2.5 : 1.6 },
        ...(label ? { label, labelBgPadding: [4, 2] as [number, number], labelStyle: { fontSize: 11, fill: '#1f285a' } } : {}),
      };
    });
}

/** 新しいノードの初期値（参照は空にしておき、検証の赤で「選んでください」と知らせる） */
export function defaultNode(type: NodeType, id: Id, position: { x: number; y: number }): FlowNode {
  const base = { id, position };
  switch (type) {
    case 'entry':
      return { ...base, type };
    case 'message':
      return { ...base, type, messageId: '' };
    case 'image':
      return { ...base, type, assetId: '' };
    case 'richMenu':
      return { ...base, type, richMenuId: '' };
    case 'choice':
      return { ...base, type, messageId: '' };
    case 'survey':
      return { ...base, type, surveyId: '', timeoutMinutes: 1440 };
    case 'tag':
      return { ...base, type, add: [] };
    case 'branch':
      return { ...base, type, label: '分岐', cases: [{ id: 'c_1', label: '条件1', when: { op: 'always' } }] };
    case 'wait':
      return { ...base, type, wait: { kind: 'dayAt', day: 1, time: '09:00' } };
    case 'page':
      return { ...base, type, pageId: '' };
    case 'jump':
      return { ...base, type, flowId: '' };
    case 'end':
      return { ...base, type };
    case 'stopAll':
      return { ...base, type };
  }
}

/** 案件全体のノードID（ノードIDは案件全体で一意にする） */
export const allNodeIds = (project: Project) => project.flows.flatMap((f) => f.nodes.map((n) => n.id));
export const allEdgeIds = (project: Project) => project.flows.flatMap((f) => f.edges.map((e) => e.id));
