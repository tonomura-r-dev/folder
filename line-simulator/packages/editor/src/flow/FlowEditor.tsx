// シナリオ（フロー）の編集画面：上の帯でシナリオを選び、キャンバスでノードを置いて繋ぐ。右は設定パネル。

import {
  applyNodeChanges,
  Background,
  Controls,
  MiniMap,
  ReactFlow,
  ReactFlowProvider,
  useReactFlow,
  type Connection,
  type Edge as RFEdge,
  type IsValidConnection,
  type NodeChange,
  type OnDelete,
} from '@xyflow/react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { Flow, NodeType, Project } from '../../../engine/src/index.ts';
import { useEditor } from '../ctx.ts';
import { replace } from '../route.ts';
import { newId } from '../ui/form.tsx';
import { allEdgeIds, allNodeIds, defaultNode, PALETTE, toCards, toLines, TYPE_LABEL, type CardNode } from './convert.ts';
import { Inspector } from './Inspector.tsx';
import { NodeCard } from './NodeCard.tsx';

const nodeTypes = { card: NodeCard };

export function FlowEditor() {
  const { project, route, update } = useEditor();
  const flow = project.flows.find((f) => f.id === route.item) ?? project.flows[0];

  const addFlow = () => {
    const id = newId('flow', project.flows.map((f) => f.id));
    const entryId = newId('n', allNodeIds(project));
    update((d) => {
      d.flows.push({
        id,
        label: '新しいシナリオ',
        trigger: { kind: 'manual' },
        priority: 20,
        concurrency: 'exclusive',
        reentry: 'never',
        nodes: [{ id: entryId, type: 'entry', position: { x: 0, y: 0 } }],
        edges: [],
      });
    });
    replace({ ...route, item: id, node: undefined, edge: undefined } as typeof route);
  };

  if (!flow)
    return (
      <div className="detail">
        <div className="empty">
          シナリオがまだありません。
          <div style={{ marginTop: 10 }}>
            <button className="btn primary" onClick={addFlow}>
              ＋ シナリオを作る
            </button>
          </div>
        </div>
      </div>
    );

  return (
    <div className="flow-wrap">
      <div className="canvas">
        <ReactFlowProvider key={flow.id}>
          <Canvas flow={flow} onAddFlow={addFlow} />
        </ReactFlowProvider>
      </div>
      <Inspector flow={flow} />
    </div>
  );
}

function Canvas({ flow, onAddFlow }: { flow: Flow; onAddFlow: () => void }) {
  const { project, issues, update, route } = useEditor();
  const fi = project.flows.indexOf(flow);
  const rf = useReactFlow();
  const box = useRef<HTMLDivElement>(null);

  const select = useCallback(
    (node?: string, edge?: string) => replace({ ...route, item: flow.id, node, edge } as typeof route),
    [route, flow.id],
  );

  // 案件から作ったカード。React Flow が測った大きさ（measured）は手元の状態に残す
  const cards = useMemo(() => toCards(flow, fi, project, issues, route.node), [flow, fi, project, issues, route.node]);
  const [nodes, setNodes] = useState<CardNode[]>(cards);
  useEffect(() => {
    setNodes((prev) => cards.map((c) => {
      const old = prev.find((o) => o.id === c.id);
      return old?.measured ? { ...c, measured: old.measured } : c;
    }));
  }, [cards]);
  const lines = useMemo(() => toLines(flow, project, route.edge), [flow, project, route.edge]);

  // 検証一覧などから特定のノードを指定されたら、そこへ寄せる
  const lastFocus = useRef<string | undefined>(undefined);
  useEffect(() => {
    if (!route.node || route.node === lastFocus.current) return;
    const n = rf.getNode(route.node);
    // 大きさが測られる前に寄せると位置がずれるので、測られてから寄せる（nodes が変わるたびに再試行）
    if (!n?.measured?.width || !n.measured.height) return;
    lastFocus.current = route.node;
    const { width, height } = n.measured;
    requestAnimationFrame(() => void rf.setCenter(n.position.x + width / 2, n.position.y + height / 2, { zoom: Math.max(rf.getZoom(), 0.9), duration: 300 }));
  }, [route.node, nodes, rf]);

  const commit = (fn: (f: Flow, d: Project) => void) =>
    update((d) => {
      const f = d.flows.find((x) => x.id === flow.id);
      if (f) fn(f, d);
    });

  const onNodesChange = useCallback((changes: NodeChange<CardNode>[]) => {
    // 削除は onDelete でまとめて扱う。それ以外（大きさ・ドラッグ中の位置・選択）は手元にだけ反映
    setNodes((nds) => applyNodeChanges(changes.filter((c) => c.type !== 'remove'), nds));
  }, []);

  const onNodeDragStop = useCallback(
    (_: unknown, __: CardNode, dragged: CardNode[]) =>
      commit((f) => {
        for (const n of dragged) {
          const target = f.nodes.find((x) => x.id === n.id);
          if (target) target.position = { x: Math.round(n.position.x), y: Math.round(n.position.y) };
        }
      }),
    [flow.id],
  );

  const onDelete: OnDelete<CardNode, RFEdge> = useCallback(
    ({ nodes: dn, edges: de }) => {
      const nodeIds = new Set(dn.filter((n) => n.data.type !== 'entry').map((n) => n.id));
      const edgeIds = new Set(de.map((e) => e.id));
      if (!nodeIds.size && !edgeIds.size) return;
      commit((f) => {
        f.nodes = f.nodes.filter((n) => !nodeIds.has(n.id));
        // 消したノードに出入りする線も一緒に消す
        f.edges = f.edges.filter((e) => !edgeIds.has(e.id) && !nodeIds.has(e.from.nodeId) && !nodeIds.has(e.to));
      });
      select(undefined, undefined);
    },
    [flow.id, select],
  );

  /** 1つの出口からは1本だけ。同じ出口に繋ぎ直したら古い線は置き換える */
  const connect = (c: Connection, replaceId?: string) =>
    commit((f, d) => {
      if (!c.source || !c.target || !c.sourceHandle) return;
      f.edges = f.edges.filter((e) => e.id === replaceId || !(e.from.nodeId === c.source && e.from.port === c.sourceHandle));
      const existing = replaceId ? f.edges.find((e) => e.id === replaceId) : undefined;
      if (existing) {
        existing.from = { nodeId: c.source, port: c.sourceHandle };
        existing.to = c.target;
      } else {
        f.edges.push({ id: newId('e', allEdgeIds(d)), from: { nodeId: c.source, port: c.sourceHandle }, to: c.target });
      }
    });

  const isValidConnection: IsValidConnection = useCallback(
    (c) => {
      const target = flow.nodes.find((n) => n.id === c.target);
      return !!target && target.type !== 'entry' && !!c.sourceHandle;
    },
    [flow.nodes],
  );

  const addNode = (type: NodeType) => {
    const rect = box.current?.getBoundingClientRect();
    const center = rf.screenToFlowPosition({ x: (rect?.left ?? 0) + (rect?.width ?? 600) / 2, y: (rect?.top ?? 0) + (rect?.height ?? 400) / 2 });
    const id = newId('n', allNodeIds(project));
    // 画面の真ん中から右へずらしていき、既存のノードと重ならない場所に置く
    const W = 230;
    const H = 120;
    const boxes = nodes.map((n) => ({ x: n.position.x, y: n.position.y, w: n.measured?.width ?? 210, h: n.measured?.height ?? 100 }));
    const hits = (x: number, y: number) => boxes.some((b) => x < b.x + b.w + 20 && x + W > b.x - 20 && y < b.y + b.h + 20 && y + H > b.y - 20);
    let pos = { x: Math.round(center.x - W / 2), y: Math.round(center.y - H / 2) };
    for (let i = 0; i < 12 && hits(pos.x, pos.y); i++) pos = { x: pos.x + 250, y: pos.y };
    commit((f) => f.nodes.push(defaultNode(type, id, pos)));
    select(id, undefined);
    // 置いた場所が画面の外なら、そこへ寄せる
    window.setTimeout(() => void rf.setCenter(pos.x + W / 2, pos.y + H / 2, { zoom: rf.getZoom(), duration: 250 }), 0);
  };

  return (
    <div className="canvas-col">
      <div className="palette">
        <select
          className="btn"
          value={flow.id}
          onChange={(e) => replace({ ...route, item: e.target.value, node: undefined, edge: undefined } as typeof route)}
          aria-label="編集するシナリオ"
        >
          {project.flows.map((f) => (
            <option key={f.id} value={f.id}>
              {f.label}
            </option>
          ))}
        </select>
        <button className="btn" onClick={onAddFlow}>
          ＋ シナリオ
        </button>
        <span style={{ width: 10 }} />
        {PALETTE.map((t) => (
          <button key={t} className="btn small" onClick={() => addNode(t)} title={`${TYPE_LABEL[t]}ノードを追加`}>
            ＋{TYPE_LABEL[t]}
          </button>
        ))}
      </div>
      <div ref={box} className="canvas-area">
        <ReactFlow<CardNode, RFEdge>
          nodes={nodes}
          edges={lines}
          nodeTypes={nodeTypes}
          onNodesChange={onNodesChange}
          onNodeDragStop={onNodeDragStop}
          onDelete={onDelete}
          onConnect={(c) => connect(c)}
          onReconnect={(old, c) => connect(c, old.id)}
          isValidConnection={isValidConnection}
          onNodeClick={(_, n) => select(n.id, undefined)}
          onEdgeClick={(_, e) => select(undefined, e.id)}
          onPaneClick={() => select(undefined, undefined)}
          deleteKeyCode={['Delete', 'Backspace']}
          // 特定のノードを開いたときは全体表示にせず、そのノードに寄せる（上の useEffect）
          fitView={!route.node}
          fitViewOptions={{ padding: 0.25, maxZoom: 1 }}
          minZoom={0.2}
          maxZoom={1.6}
          proOptions={{ hideAttribution: false }}
        >
          <Background gap={20} color="#dfe3ec" />
          <Controls showInteractive={false} />
          <MiniMap pannable zoomable nodeStrokeWidth={2} style={{ width: 160, height: 110 }} />
        </ReactFlow>
      </div>
    </div>
  );
}
