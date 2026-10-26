// 条件式を画面で組み立てる部品。AND（すべて）／OR（どれか）のグループに条件を入れ、
// どの条件・グループにも「ではない（NOT）」を付けられる。保存される形は DATA_MODEL §5 の JSON のまま。

import type { Expr, InputEventType, Project } from '../../../engine/src/index.ts';
import { useEditor } from '../ctx.ts';
import { Checks, NumberInput, Select, Text, type Option } from './form.tsx';
import { DurationInput, EVENT_OPTIONS, options } from './pickers.tsx';

type Leaf = Exclude<Expr, { op: 'and' | 'or' | 'not' }>;
type Group = Extract<Expr, { op: 'and' | 'or' }>;

const LEAF_KINDS: Option[] = [
  { value: 'hasTag', label: 'タグが付いている' },
  { value: 'answer', label: 'アンケートの回答' },
  { value: 'answered', label: 'アンケートに回答済み' },
  { value: 'cv', label: 'CV（応募・予約など）' },
  { value: 'friend', label: '友だちの状態' },
  { value: 'route', label: '友だち追加の経路' },
  { value: 'attr', label: '属性（年齢・地域など）' },
  { value: 'did', label: '行動（操作したか）' },
  { value: 'elapsed', label: '経過時間' },
  { value: 'delivered', label: '配信を受け取ったか' },
  { value: 'always', label: '常に当てはまる' },
];

export function defaultLeaf(op: string, p: Project): Leaf {
  const sv = p.surveys[0];
  const q = sv?.questions[0];
  switch (op) {
    case 'answer':
      return { op, surveyId: sv?.id ?? '', questionId: q?.id ?? '', match: 'any', choiceIds: q?.choices[0] ? [q.choices[0].id] : [] };
    case 'answered':
      return { op, surveyId: sv?.id ?? '' };
    case 'cv':
      return { op, cvId: p.cvs[0]?.id ?? '', done: true };
    case 'friend':
      return { op, state: 'friend' };
    case 'route':
      return { op, routeIds: p.routes[0] ? [p.routes[0].id] : [] };
    case 'attr':
      return { op, key: 'area', cmp: 'eq', value: '' };
    case 'did':
      return { op, event: 'tapButton', cmp: 'exists' };
    case 'elapsed':
      return { op, since: 'friendAdd', cmp: 'gte', minutes: 1440 };
    case 'delivered':
      return { op, target: { messageId: p.messages[0]?.id ?? '' }, done: true };
    case 'always':
      return { op };
    default:
      return { op: 'hasTag', tagId: p.tags[0]?.id ?? '' };
  }
}

const isGroup = (e: Expr): e is Group => e.op === 'and' || e.op === 'or';

export function ExprBuilder({ value, onChange }: { value: Expr; onChange: (e: Expr) => void }) {
  const { project } = useEditor();
  const root = value && typeof value === 'object' ? value : defaultLeaf('hasTag', project);
  const inner = root.op === 'not' ? root.child : root;
  return (
    <div className="xb">
      <Item e={root} onChange={onChange} depth={0} />
      {!isGroup(inner) && (
        <div className="row" style={{ marginTop: 6 }}>
          <button type="button" className="btn small" onClick={() => onChange({ op: 'and', children: [root, defaultLeaf('hasTag', project)] })}>
            ＋ 条件を足す（すべて満たす）
          </button>
          <button type="button" className="btn small" onClick={() => onChange({ op: 'or', children: [root, defaultLeaf('hasTag', project)] })}>
            ＋ 条件を足す（どれかを満たす）
          </button>
        </div>
      )}
    </div>
  );
}

/** 1つの条件またはグループ。「ではない」の切り替えと削除をここで持つ */
function Item({ e, onChange, onRemove, depth }: { e: Expr; onChange: (e: Expr) => void; onRemove?: () => void; depth: number }) {
  const negated = e.op === 'not';
  const inner: Expr = negated ? e.child : e;
  const setInner = (x: Expr) => onChange(negated ? { op: 'not', child: x } : x);
  return (
    <div className={`xb-item${negated ? ' neg' : ''}`}>
      <div className="xb-head">
        <label className="xb-not" title="この条件を反対にする（NOT）">
          <input type="checkbox" checked={negated} onChange={() => onChange(negated ? inner : { op: 'not', child: inner })} />
          ではない
        </label>
        <span className="spacer" />
        {onRemove && (
          <button type="button" className="btn small ghost danger" onClick={onRemove} title="この条件を外す">
            ×
          </button>
        )}
      </div>
      {isGroup(inner) ? <GroupBox g={inner} onChange={setInner} depth={depth} /> : <LeafBox e={inner as Leaf} onChange={setInner} />}
    </div>
  );
}

function GroupBox({ g, onChange, depth }: { g: Group; onChange: (e: Expr) => void; depth: number }) {
  const { project } = useEditor();
  const set = (i: number, x: Expr) => onChange({ ...g, children: g.children.map((c, k) => (k === i ? x : c)) });
  const remove = (i: number) => {
    const rest = g.children.filter((_, k) => k !== i);
    // 1つだけ残ったらグループを外して、その条件だけにする
    onChange(rest.length === 1 ? rest[0]! : { ...g, children: rest });
  };
  return (
    <div className={`xb-group d${depth % 3}`}>
      <div className="xb-group-head">
        <select value={g.op} onChange={(ev) => onChange({ ...g, op: ev.target.value as 'and' | 'or' })} aria-label="グループの種類">
          <option value="and">次のすべてを満たす（AND）</option>
          <option value="or">次のどれかを満たす（OR）</option>
        </select>
      </div>
      {g.children.map((c, i) => (
        <div key={i} className="xb-child">
          {i > 0 && <div className="xb-joint">{g.op === 'and' ? 'かつ' : 'または'}</div>}
          <Item e={c} depth={depth + 1} onChange={(x) => set(i, x)} onRemove={() => remove(i)} />
        </div>
      ))}
      <div className="row" style={{ marginTop: 6 }}>
        <button type="button" className="btn small" onClick={() => onChange({ ...g, children: [...g.children, defaultLeaf('hasTag', project)] })}>
          ＋ 条件
        </button>
        <button
          type="button"
          className="btn small"
          onClick={() => onChange({ ...g, children: [...g.children, { op: g.op === 'and' ? 'or' : 'and', children: [defaultLeaf('hasTag', project), defaultLeaf('hasTag', project)] }] })}
        >
          ＋ グループ（{g.op === 'and' ? 'どれか' : 'すべて'}）
        </button>
      </div>
    </div>
  );
}

const CMP: Option[] = [
  { value: 'eq', label: 'と同じ' },
  { value: 'neq', label: 'と違う' },
  { value: 'gte', label: '以上' },
  { value: 'lte', label: '以下' },
  { value: 'in', label: 'のどれか（、区切り）' },
];

function OptionalWithin({ value, onChange, label }: { value: number | undefined; onChange: (v: number | undefined) => void; label: string }) {
  return (
    <span className="row" style={{ flexWrap: 'nowrap', gap: 4 }}>
      <label className="row small" style={{ flexWrap: 'nowrap', gap: 4 }}>
        <input type="checkbox" checked={value !== undefined} onChange={(ev) => onChange(ev.target.checked ? 1440 * 7 : undefined)} />
        {label}
      </label>
      {value !== undefined && <DurationInput value={value} onChange={onChange} />}
    </span>
  );
}

function LeafBox({ e, onChange }: { e: Leaf; onChange: (e: Expr) => void }) {
  const { project: p } = useEditor();
  const kind = (
    <Select value={e.op} options={LEAF_KINDS} onChange={(op) => onChange(defaultLeaf(op, p))} />
  );
  const body = (() => {
    switch (e.op) {
      case 'hasTag':
        return <Select value={e.tagId} empty="タグを選ぶ" options={options.tags(p)} onChange={(tagId) => onChange({ ...e, tagId })} />;
      case 'answer': {
        const sv = p.surveys.find((s) => s.id === e.surveyId);
        const q = sv?.questions.find((x) => x.id === e.questionId);
        return (
          <>
            <Select value={e.surveyId} empty="アンケート" options={options.surveys(p)} onChange={(surveyId) => onChange({ ...e, surveyId, questionId: '', choiceIds: [] })} />
            <Select value={e.questionId} empty="質問" options={(sv?.questions ?? []).map((x) => ({ value: x.id, label: x.text }))} onChange={(questionId) => onChange({ ...e, questionId, choiceIds: [] })} />
            <Select
              value={e.match}
              options={[
                { value: 'any', label: 'で次のどれかを選んだ' },
                { value: 'all', label: 'で次をすべて選んだ' },
                { value: 'none', label: 'で次をどれも選んでいない' },
              ]}
              onChange={(match) => onChange({ ...e, match: match as typeof e.match })}
            />
            <Checks value={e.choiceIds} options={(q?.choices ?? []).map((c) => ({ value: c.id, label: c.label }))} onChange={(choiceIds) => onChange({ ...e, choiceIds })} />
          </>
        );
      }
      case 'answered': {
        const sv = p.surveys.find((s) => s.id === e.surveyId);
        return (
          <>
            <Select value={e.surveyId} empty="アンケート" options={options.surveys(p)} onChange={(surveyId) => onChange({ op: 'answered', surveyId })} />
            <Select
              value={e.questionId ?? ''}
              empty="すべての質問"
              options={(sv?.questions ?? []).map((x) => ({ value: x.id, label: x.text }))}
              onChange={(questionId) => onChange(questionId ? { ...e, questionId } : { op: 'answered', surveyId: e.surveyId })}
            />
          </>
        );
      }
      case 'cv':
        return (
          <>
            <Select value={e.cvId} empty="CVを選ぶ" options={options.cvs(p)} onChange={(cvId) => onChange({ ...e, cvId })} />
            <Select
              value={e.done ? 'yes' : 'no'}
              options={[
                { value: 'yes', label: 'をした' },
                { value: 'no', label: 'をしていない' },
              ]}
              onChange={(v) => onChange({ ...e, done: v === 'yes' })}
            />
            <OptionalWithin label="期間を区切る" value={e.withinMinutes} onChange={(w) => onChange(w === undefined ? { op: 'cv', cvId: e.cvId, done: e.done } : { ...e, withinMinutes: w })} />
          </>
        );
      case 'friend':
        return (
          <Select
            value={e.state}
            options={[
              { value: 'friend', label: '友だち' },
              { value: 'not_friend', label: 'まだ友だちではない' },
              { value: 'blocked', label: 'ブロック中' },
            ]}
            onChange={(state) => onChange({ ...e, state: state as typeof e.state })}
          />
        );
      case 'route':
        return <Checks value={e.routeIds} options={options.routes(p)} onChange={(routeIds) => onChange({ ...e, routeIds })} />;
      case 'attr':
        return (
          <>
            <Text value={e.key} onChange={(key) => onChange({ ...e, key })} placeholder="属性の名前（例：age）" />
            <Text
              value={Array.isArray(e.value) ? e.value.join('、') : String(e.value ?? '')}
              onChange={(raw) => {
                const toValue = (s: string) => (s.trim() !== '' && !Number.isNaN(Number(s)) ? Number(s) : s.trim());
                onChange({ ...e, value: e.cmp === 'in' ? raw.split(/[、,]/).map(toValue) : toValue(raw) });
              }}
              placeholder="値"
            />
            <Select value={e.cmp} options={CMP} onChange={(cmp) => onChange({ ...e, cmp: cmp as typeof e.cmp, value: cmp === 'in' ? [] : '' })} />
          </>
        );
      case 'did':
        return (
          <>
            <Select value={e.event} options={EVENT_OPTIONS} onChange={(event) => onChange({ ...e, event: event as InputEventType })} />
            <Select
              value={e.cmp}
              options={[
                { value: 'exists', label: 'ことがある' },
                { value: 'countGte', label: '回数が○回以上' },
              ]}
              onChange={(cmp) => onChange(cmp === 'countGte' ? { ...e, cmp: 'countGte', value: e.value ?? 2 } : { op: 'did', event: e.event, cmp: 'exists', ...(e.withinMinutes ? { withinMinutes: e.withinMinutes } : {}) })}
            />
            {e.cmp === 'countGte' && <NumberInput value={e.value} min={1} onChange={(v) => onChange({ ...e, value: v ?? 1 })} />}
            <OptionalWithin label="期間を区切る" value={e.withinMinutes} onChange={(w) => onChange(w === undefined ? (({ withinMinutes: _, ...rest }) => rest)(e) : { ...e, withinMinutes: w })} />
          </>
        );
      case 'elapsed':
        return (
          <>
            <Select
              value={e.since}
              options={[
                { value: 'friendAdd', label: '友だち追加から' },
                { value: 'flowStart', label: 'このシナリオの開始から' },
                { value: 'lastDelivery', label: '最後の配信から' },
              ]}
              onChange={(since) => onChange({ ...e, since: since as typeof e.since })}
            />
            <DurationInput value={e.minutes} onChange={(minutes) => onChange({ ...e, minutes })} />
            <Select
              value={e.cmp}
              options={[
                { value: 'gte', label: '以上たった' },
                { value: 'lt', label: 'たっていない' },
              ]}
              onChange={(cmp) => onChange({ ...e, cmp: cmp as typeof e.cmp })}
            />
          </>
        );
      case 'delivered':
        return (
          <>
            <Select
              value={e.target.messageId ?? ''}
              empty="文面を選ぶ"
              options={options.messages(p)}
              onChange={(messageId) => onChange({ ...e, target: { messageId } })}
            />
            <Select
              value={e.done ? 'yes' : 'no'}
              options={[
                { value: 'yes', label: 'を受け取った' },
                { value: 'no', label: 'をまだ受け取っていない' },
              ]}
              onChange={(v) => onChange({ ...e, done: v === 'yes' })}
            />
          </>
        );
      case 'always':
        return null;
    }
  })();
  return (
    <div className="xb-leaf">
      {kind}
      {body}
    </div>
  );
}
