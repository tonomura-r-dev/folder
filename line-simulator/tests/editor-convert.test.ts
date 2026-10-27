// 編集画面のフロー変換（packages/editor/src/flow/convert.ts）。画面を開かずに確かめられる部分。

import { describe, expect, it } from 'vitest';
import { validate } from '../packages/engine/src/index.ts';
import { describeIssue } from '../packages/player/src/issues.ts';
import { defaultNode, PALETTE, toCards, toLines } from '../packages/editor/src/flow/convert.ts';
import { locate } from '../packages/editor/src/route.ts';
import { areaBranch, sample } from './helpers.ts';

describe('フロー → カード・線', () => {
  const p = sample();
  const fi = p.flows.findIndex((f) => f.id === 'flow_welcome');
  const flow = p.flows[fi]!;

  it('分岐ノードには条件ごと＋「それ以外」の出口ができ、名前は案件のラベル', () => {
    const area = areaBranch(p).flow;
    const card = toCards(area, p.flows.indexOf(area), p, validate(p)).find((c) => c.id === 'n_w_area')!;
    expect(card.data.ports.map((x) => x.label)).toEqual(['東京', '大阪', 'それ以外']);
    expect(card.data.ports.every((x) => x.connected)).toBe(true);
  });

  it('ボタン選択ノードの出口は、文面のボタンの文字になる', () => {
    const t = p.flows.find((f) => f.id === 'flow_step_tokyo')!;
    const card = toCards(t, 1, p, []).find((c) => c.id === 'n_t_d7')!;
    expect(card.data.ports.map((x) => x.label)).toEqual(['応募する', '今回は見送る', '時間切れ']);
  });

  it('入口ノードは消せず、線を受け付けない', () => {
    const entry = toCards(flow, fi, p, []).find((c) => c.data.type === 'entry')!;
    expect(entry.deletable).toBe(false);
    expect(entry.data.hasInput).toBe(false);
  });

  it('行き先の無い線は描かない（検証一覧には出る）', () => {
    const broken = structuredClone(flow);
    broken.edges[0]!.to = 'n_none';
    expect(toLines(broken, p).map((e) => e.id)).not.toContain(broken.edges[0]!.id);
  });

  it('検証結果のあるノードには件数が付く', () => {
    const q = structuredClone(p);
    const n = q.flows[fi]!.nodes.find((x) => x.id === 'n_w_msg')!;
    if (n.type === 'message') n.messageId = '';
    const card = toCards(q.flows[fi]!, fi, q, validate(q)).find((c) => c.id === 'n_w_msg')!;
    expect(card.data.errors).toBe(1);
    expect(card.data.missing).toBe(true);
  });

  it('パレットのどのノードも、置いた直後の初期値で画面が落ちない（未選択は検証の赤で知らせる）', () => {
    const q = structuredClone(p);
    PALETTE.forEach((t, i) => q.flows[fi]!.nodes.push(defaultNode(t, `n_new_${i}`, { x: 0, y: 0 })));
    const cards = toCards(q.flows[fi]!, fi, q, validate(q));
    expect(cards).toHaveLength(flow.nodes.length + PALETTE.length);
    const msgs = validate(q).filter((i) => i.code === 'ref.missing' && i.params?.['id'] === '').map(describeIssue);
    expect(msgs).toContain('文面が選ばれていません');
  });
});

describe('検証結果 → 画面の場所', () => {
  it('ノードの問題はそのフローとノード、線の問題は出発点のノードを開く', () => {
    const p = sample();
    expect(locate(p, 's', '/flows/1/nodes/4/cases/0/when')).toMatchObject({ section: 'flows', item: 'flow_step_tokyo', node: 'n_t_job' });
    expect(locate(p, 's', '/flows/1/edges/4')).toMatchObject({ item: 'flow_step_tokyo', node: 'n_t_job', edge: 'e_t5' });
    expect(locate(p, 's', '/messages/2/bubbles/0')).toMatchObject({ section: 'messages', item: 'msg_t_d1' });
    expect(locate(p, 's', '/tags/0')).toMatchObject({ section: 'labels' });
  });
});
