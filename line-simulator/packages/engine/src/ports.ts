// ノード種別ごとに許される出口（ポート）。検証と編集画面の両方がこれを使う。

import { bubbleButtons } from './indexing.ts';
import type { FlowNode, Id, Project } from './schema.ts';

/** ボタン選択ノードが参照するメッセージのボタンID */
export function choiceButtons(node: Extract<FlowNode, { type: 'choice' }>, project: Project): Id[] {
  const m = project.messages.find((x) => x.id === node.messageId);
  return (m?.bubbles ?? []).flatMap((b) => bubbleButtons(b).map((x) => x.id));
}

export function portsOf(node: FlowNode, project: Project): string[] {
  switch (node.type) {
    case 'choice':
      return [...choiceButtons(node, project).map((id) => `btn:${id}`), ...(node.timeoutMinutes !== undefined ? ['timeout'] : [])];
    case 'survey':
      return ['done', ...(node.timeoutMinutes !== undefined ? ['timeout'] : [])];
    case 'wait':
      return node.wait.kind === 'untilEvent' ? ['out', 'timeout'] : ['out'];
    case 'branch':
      return [...node.cases.map((c) => `case:${c.id}`), 'else'];
    case 'jump':
    case 'end':
    case 'stopAll':
      return [];
    default:
      return ['out'];
  }
}

/** 入ってくる線を受けられるか（入口ノードは受けない） */
export const acceptsInput = (node: FlowNode) => node.type !== 'entry';
