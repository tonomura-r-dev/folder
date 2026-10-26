// Phase 4：営業先に見せる「なぜ」の日本語と、条件式ビルダーの初期値。

import { describe, expect, it } from 'vitest';
import type { Expr } from '../packages/engine/src/index.ts';
import { reduce, sessionFromPersonas, validate } from '../packages/engine/src/index.ts';
import { defaultLeaf } from '../packages/editor/src/ui/ExprBuilder.tsx';
import { describeTraceNode, explainDecision, Labels } from '../packages/player/src/describe.ts';
import { baseProject, chain, DAY, node, sample } from './helpers.ts';

const project = sample();
const L = new Labels(project);
const world = reduce(project, sessionFromPersonas(project), { until: 30 * DAY });
const find = (userId: string, pred: (d: (typeof world.decisions)[number]) => boolean) => world.decisions.find((d) => d.userId === userId && pred(d))!;

describe('「なぜ」の説明（営業先向け）', () => {
  it('大阪希望さんが東京ルートに進まなかった理由は「東京のタグが無いため」', () => {
    const x = explainDecision(find('p_osaka', (d) => d.nodeId === 'n_w_area' && d.ref === 'tokyo'), L, project, world);
    expect(x).toEqual({ tone: 'ng', title: '「希望エリア」で「東京」には進みませんでした', reason: 'タグ「希望エリア：東京」が付いていないため' });
  });

  it('応募済みさんの東京ステップは「応募済みに当てはまったため」止まる', () => {
    const x = explainDecision(find('p_applied', (d) => d.kind === 'stopWhen' && d.flowId === 'flow_step_tokyo'), L, project, world);
    expect(x).toEqual({ tone: 'stop', title: '「東京希望のステップ配信」を止めました', reason: '「応募済み」に当てはまったため' });
  });

  it('未応募さんにブロック後の配信が届かない理由は「ブロックしているため」', () => {
    const x = explainDecision(find('p_notyet', (d) => d.kind === 'deliver' && d.ref === 'msg_c_general'), L, project, world);
    expect(x.tone).toBe('ng');
    expect(x.reason).toBe('この人がブロックしているため');
  });

  it('優先度で見送ったときは、先に届いたシナリオの名前を出す', () => {
    const p = baseProject({
      flows: [
        chain('hi', [node({ id: 'hi_m', type: 'message', messageId: 'm1' })], { priority: 1, label: '大事なお知らせ' }),
        chain('lo', [node({ id: 'lo_m', type: 'message', messageId: 'm2' })], { priority: 2, label: 'ついでの案内' }),
      ],
    });
    const w = reduce(p, { schemaVersion: 1, projectId: 't', users: [{ id: 'u1', personaId: 'u1' }], events: [{ at: 0, userId: 'u1', type: 'addFriend', routeId: 'r1' }] });
    const x = explainDecision(w.decisions.find((d) => d.ref === 'm2' && d.kind === 'deliver')!, new Labels(p), p, w);
    expect(x.reason).toBe('同じ時刻に、優先度の高い「大事なお知らせ」の配信があったため');
  });

  it('判定の木の各段に、実際の値を添える', () => {
    const d = find('p_osaka', (x) => x.nodeId === 'n_w_area' && x.ref === 'tokyo');
    expect(describeTraceNode(d.trace!, L)).toBe('タグ「希望エリア：東京」が付いている（付いているタグ：希望エリア：大阪、希望職種：軽作業）');
    const stop = find('p_applied', (x) => x.kind === 'stopWhen');
    expect(describeTraceNode(stop.trace!, L)).toMatch(/^応募済み（応募：Day 0 10:10）$/);
  });
});

describe('条件式ビルダー', () => {
  const kinds = ['hasTag', 'answer', 'answered', 'cv', 'friend', 'route', 'attr', 'did', 'elapsed', 'delivered', 'always'];

  it('どの種類の条件も、選んだ直後の初期値が正しい条件式になる（検証で重大エラーが出ない）', () => {
    for (const k of kinds) {
      const p = structuredClone(project);
      const n = p.flows[0]!.nodes.find((x) => x.id === 'n_w_area')!;
      if (n.type === 'branch') n.cases[0]!.when = defaultLeaf(k, p);
      const errors = validate(p).filter((i) => i.level === 'error' && i.path.startsWith('/flows/0/nodes/5/cases/0/when'));
      // 属性（attr）の値は空のまま置くので型の問題は出ない。どれも重大エラー0件
      expect({ k, errors }).toEqual({ k, errors: [] });
    }
  });

  it('組み立てた AND／OR／NOT の入れ子がそのまま評価できる', () => {
    const e: Expr = {
      op: 'and',
      children: [
        { op: 'or', children: [defaultLeaf('hasTag', project), { op: 'hasTag', tagId: 'area_osaka' }] },
        { op: 'not', child: { op: 'cv', cvId: 'cv_apply', done: true } },
      ],
    };
    const p = structuredClone(project);
    const n = p.flows[0]!.nodes.find((x) => x.id === 'n_w_area')!;
    if (n.type === 'branch') n.cases[0]!.when = e;
    expect(validate(p).filter((i) => i.level === 'error')).toEqual([]);
  });
});
