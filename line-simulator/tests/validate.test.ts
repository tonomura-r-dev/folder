// 壊した案件が、該当する重大エラー・警告になるか。

import { describe, expect, it } from 'vitest';
import type { Project } from '../packages/engine/src/index.ts';
import { validate } from '../packages/engine/src/index.ts';
import { areaBranch, baseProject, chain, codes, edge, node, sample } from './helpers.ts';

const broken = (mutate: (p: Project) => void) => {
  const p = sample();
  mutate(p);
  return validate(p);
};
const flow = (p: Project, id: string) => p.flows.find((f) => f.id === id)!;

describe('重大エラー', () => {
  it('存在しないノードへの接続', () => {
    const issues = broken((p) => (flow(p, 'flow_welcome').edges[0]!.to = 'n_nowhere'));
    expect(codes(issues, 'error')).toContain('edge.danglingTarget');
  });

  it('そのノードに無い出口から出ている', () => {
    const issues = broken((p) => (flow(p, 'flow_welcome').edges[0]!.from.port = 'case:x'));
    expect(codes(issues, 'error')).toContain('edge.invalidPort');
  });

  it('同じ出口から2本出ている', () => {
    const issues = broken((p) => {
      const f = flow(p, 'flow_welcome');
      f.edges.push({ ...f.edges[0]!, id: 'e_dup', to: 'n_w_msg' });
    });
    expect(codes(issues, 'error')).toContain('edge.invalidPort');
  });

  it('定義していないタグを条件で使う', () => {
    const issues = broken((p) => {
      areaBranch(p).node.cases[0]!.when = { op: 'hasTag', tagId: 'area_nagoya' };
    });
    expect(issues).toContainEqual(expect.objectContaining({ level: 'error', code: 'ref.missing', params: { kind: 'tag', id: 'area_nagoya' } }));
  });

  it('壊れた条件式（中身のない AND・未知の演算子）', () => {
    const issues = broken((p) => {
      flow(p, 'flow_step_tokyo').stopWhen = { op: 'and', children: [] };
      flow(p, 'flow_step_osaka').stopWhen = { op: 'xor' } as never;
    });
    expect(codes(issues, 'error').filter((c) => c === 'expr.invalid')).toHaveLength(2);
  });

  it('分岐に「それ以外」の行き先が無い', () => {
    const issues = broken((p) => {
      const f = areaBranch(p).flow;
      f.edges = f.edges.filter((e) => !(e.from.nodeId === 'n_w_area' && e.from.port === 'else'));
    });
    expect(codes(issues, 'error')).toContain('branch.noExit');
  });

  it('時間が進まないループ（タグ付与どうしの往復）', () => {
    const p = baseProject({
      flows: [chain('f', [node({ id: 'a', type: 'tag', add: ['tag_a'] }), node({ id: 'b', type: 'tag', add: ['tag_a'] })])],
    });
    p.flows[0]!.edges.push(edge('b', 'a'));
    expect(codes(validate(p), 'error')).toContain('flow.zeroTimeCycle');
  });

  it('時間が進まないループ（フローをまたぐ jump の往復）', () => {
    const p = baseProject({
      flows: [
        chain('f1', [node({ id: 'j1', type: 'jump', flowId: 'f2' })]),
        chain('f2', [node({ id: 'j2', type: 'jump', flowId: 'f1' })], { trigger: { kind: 'manual' } }),
      ],
    });
    expect(codes(validate(p), 'error')).toContain('flow.zeroTimeCycle');
  });

  it('待機を挟むループは正当なので、エラーにしない', () => {
    const p = baseProject({
      flows: [
        chain('f', [node({ id: 'w', type: 'wait', wait: { kind: 'duration', minutes: 1440 } }), node({ id: 'm', type: 'message', messageId: 'm1' })], {
          stopWhen: { op: 'cv', cvId: 'cv1', done: true },
        }),
      ],
    });
    p.flows[0]!.edges.push(edge('m', 'w'));
    expect(codes(validate(p), 'error')).toEqual([]);
  });

  it('入口が無いフロー', () => {
    const issues = broken((p) => {
      const f = flow(p, 'flow_step_osaka');
      f.nodes = f.nodes.filter((n) => n.type !== 'entry');
      f.edges = f.edges.filter((e) => e.from.nodeId !== 'n_o_entry');
    });
    expect(codes(issues, 'error')).toContain('flow.noEntry');
  });

  it('IDの重複', () => {
    const issues = broken((p) => p.tags.push({ id: 'area_tokyo', label: 'dup' }));
    expect(codes(issues, 'error')).toContain('id.duplicate');
  });

  it('必須設定の欠落（吹き出しが0個・待機の時刻が不正）', () => {
    const issues = broken((p) => {
      p.messages[0]!.bubbles = [];
      const w = flow(p, 'flow_step_tokyo').nodes.find((n) => n.id === 'n_t_wait1')!;
      if (w.type === 'wait') w.wait = { kind: 'dayAt', day: 1, time: '25:00' };
    });
    expect(codes(issues, 'error').filter((c) => c === 'node.requiredMissing')).toHaveLength(2);
  });

  it('AIの推定が承認されずに残っている', () => {
    const issues = broken((p) => (flow(p, 'flow_step_tokyo').nodes[1]!.prov = { state: 'inferred', rationale: 'test' }));
    expect(issues).toContainEqual(expect.objectContaining({ level: 'error', code: 'prov.unapproved', path: '/flows/1/nodes/1' }));
  });

  it('画像ファイルが無い', () => {
    const p = sample();
    p.assets.push({ id: 'img1', kind: 'image', path: 'assets/none.png', width: 1, height: 1, alt: 'x' });
    const issues = validate(p, { assetExists: () => false });
    expect(codes(issues, 'error')).toContain('asset.fileMissing');
  });
});

describe('警告（書き出しは可）', () => {
  it('使われていないタグ', () => {
    const issues = broken((p) => p.tags.push({ id: 'tag_orphan', label: '孤児' }));
    expect(codes(issues, 'warning')).toContain('tag.unused');
    expect(codes(issues, 'error')).toEqual([]);
  });

  it('到達できないノード', () => {
    const issues = broken((p) => flow(p, 'flow_step_osaka').nodes.push(node({ id: 'n_orphan', type: 'end' })));
    expect(codes(issues, 'warning')).toContain('node.unreachable');
  });

  it('停止条件の無いステップ配信', () => {
    const issues = broken((p) => delete flow(p, 'flow_step_osaka').stopWhen);
    expect(codes(issues, 'warning')).toContain('flow.noStop');
  });

  it('同じメッセージに別のフローから届く経路があり、重複防止が無効', () => {
    const issues = broken((p) => {
      p.settings.dedupeWindowMinutes = 0;
      const n = flow(p, 'flow_step_osaka').nodes.find((x) => x.id === 'n_o_d1')!;
      if (n.type === 'message') n.messageId = 'msg_t_d1';
    });
    expect(codes(issues, 'warning')).toContain('delivery.possibleDuplicate');
  });

  it('差し込みのキーが不明', () => {
    const issues = broken((p) => (p.messages[0]!.bubbles = [{ kind: 'text', text: '{{user.nickname}}さん' }]));
    expect(codes(issues, 'warning')).toContain('text.unknownVar');
  });

  it('LINEの上限超え（ステップ365日超）', () => {
    const issues = broken((p) => {
      const w = flow(p, 'flow_step_tokyo').nodes.find((n) => n.id === 'n_t_wait7')!;
      if (w.type === 'wait') w.wait = { kind: 'dayAt', day: 400, time: '12:00' };
    });
    expect(codes(issues, 'warning')).toContain('line.limit');
  });

  it('友だち追加前に始まるフローが、待機なしで配信しようとしている', () => {
    const p = baseProject({
      webPages: [{ id: 'pg', kind: 'lp', title: 'lp', blocks: [] }],
      flows: [chain('f', [node({ id: 'm', type: 'message', messageId: 'm1' })], { trigger: { kind: 'pageExit', pageId: 'pg' } })],
    });
    expect(codes(validate(p), 'warning')).toContain('flow.friendRequired');
  });
});
