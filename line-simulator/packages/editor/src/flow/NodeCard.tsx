// フローチャート上のノード1枚。出口（ポート）ごとに接続口を出す。

import { Handle, Position, type NodeProps } from '@xyflow/react';
import { memo } from 'react';
import { ProvBadge } from '../ai/ProvBadge.tsx';
import { TYPE_LABEL, type CardNode } from './convert.ts';

export const NodeCard = memo(function NodeCard({ data, selected }: NodeProps<CardNode>) {
  const state = data.errors ? ' has-err' : data.warnings ? ' has-warn' : '';
  return (
    <div className={`fnode t-${data.type}${selected ? ' selected' : ''}${state}`}>
      {data.hasInput && <Handle type="target" position={Position.Top} />}
      <div className="fnode-head">
        <span className="fnode-type">{TYPE_LABEL[data.type]}</span>
        {data.title && <span className="fnode-title">{data.title}</span>}
        <ProvBadge prov={data.node.prov} />
      </div>
      {data.body && <div className={`fnode-body${data.missing ? ' missing' : ''}`}>{data.body}</div>}
      {data.errors > 0 && <div className="fnode-err">✖ 重大エラー {data.errors}</div>}
      {data.ports.length > 0 && (
        <div className="fnode-ports">
          {data.ports.map((p) => (
            <div key={p.id} className={`fnode-port${p.connected ? ' connected' : ''}`} title={p.id}>
              {p.label}
              <Handle type="source" position={Position.Bottom} id={p.id} />
            </div>
          ))}
        </div>
      )}
    </div>
  );
});
