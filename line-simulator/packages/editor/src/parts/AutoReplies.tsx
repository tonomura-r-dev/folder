// 自動応答：キーワード → 条件ごとの返信（上から順に判定）→ どれにも当てはまらないときの返信。

import type { AutoReply } from '../../../engine/src/index.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { replace } from '../route.ts';
import { Field, InlineIssues, ItemTools, move, newId, NumberInput, ReadOnly, Select, Text } from '../ui/form.tsx';
import { ExprField, options } from '../ui/pickers.tsx';
import { SplitList } from '../ui/split.tsx';

export function AutoReplies() {
  const { project, update, slug } = useEditor();
  const sorted = [...project.autoReplies].sort((a, b) => a.priority - b.priority);
  return (
    <SplitList<AutoReply>
      section="autoReplies"
      pathBase="/autoReplies"
      title="自動応答"
      items={project.autoReplies}
      row={(a) => ({ t: a.match.keywords.join('・') || '（キーワード未設定）', s: `${a.match.mode === 'exact' ? '完全一致' : '含む'}・判定順 ${sorted.indexOf(a) + 1}` })}
      onAdd={() => {
        const id = newId('ar', project.autoReplies.map((a) => a.id));
        update((d) => d.autoReplies.push({ id, priority: (Math.max(0, ...d.autoReplies.map((a) => a.priority)) || 0) + 10, match: { mode: 'exact', keywords: [''] }, cases: [] }));
        return id;
      }}
      detail={(a, ai, path) => <ReplyDetail a={a} ai={ai} path={path} onDeleted={() => replace({ page: 'project', slug, section: 'autoReplies' })} />}
    />
  );
}

function ReplyDetail({ a, ai, path, onDeleted }: { a: AutoReply; ai: number; path: string; onDeleted: () => void }) {
  const { project, update, issues } = useEditor();
  const set = (fn: (x: AutoReply) => void) => update((d) => fn(d.autoReplies[ai]!));
  return (
    <>
      <h2>自動応答</h2>
      <ReadOnly>{a.id}</ReadOnly>
      <InlineIssues issues={issues.filter((i) => i.path === path)} />
      <div className="section">
        <div className="fields">
          <Field label="キーワード（読点「、」区切り）" issues={issuesAt(issues, `${path}/match/keywords`)}>
            <Text value={a.match.keywords.join('、')} onChange={(v) => set((x) => (x.match.keywords = v.split(/[、,]/).map((s) => s.trim())))} />
          </Field>
          <Field label="一致のしかた">
            <Select
              value={a.match.mode}
              options={[
                { value: 'exact', label: 'キーワードと完全に同じ' },
                { value: 'contains', label: 'キーワードを含む' },
              ]}
              onChange={(v) => set((x) => (x.match.mode = v as typeof a.match.mode))}
            />
          </Field>
          <Field label="判定の順番（小さいほど先）" hint="複数の自動応答に当てはまるときは、先の1つだけが返信します">
            <NumberInput value={a.priority} onChange={(v) => set((x) => (x.priority = v ?? 0))} />
          </Field>
        </div>
      </div>

      <div className="section">
        <h3>条件ごとの返信（上から順に、最初に当てはまったものを返す）</h3>
        {a.cases.length === 0 && <div className="muted small">条件はありません。下の「どれにも当てはまらないとき」の返信だけを返します。</div>}
        {a.cases.map((c, ci) => (
          <div className="sub" key={ci}>
            <div className="sub-head">
              <span className="t">返信 {ci + 1}</span>
              <span className="spacer" />
              <ItemTools index={ci} length={a.cases.length} onMove={(to) => set((x) => (x.cases = move(x.cases, ci, to)))} onRemove={() => set((x) => x.cases.splice(ci, 1))} />
            </div>
            <div className="fields">
              <ExprField label="この条件のとき" path={`${path}/cases/${ci}/when`} value={c.when} onChange={(e) => set((x) => (x.cases[ci]!.when = e))} />
              <Field label="返す文面" issues={issuesAt(issues, `${path}/cases/${ci}/messageId`)}>
                <Select value={c.messageId} empty="選んでください" options={options.messages(project)} onChange={(v) => set((x) => (x.cases[ci]!.messageId = v))} />
              </Field>
            </div>
          </div>
        ))}
        <button
          className="btn small"
          style={{ marginTop: 8 }}
          onClick={() => set((x) => x.cases.push({ when: project.tags[0] ? { op: 'hasTag', tagId: project.tags[0].id } : { op: 'always' }, messageId: '' }))}
        >
          ＋ 条件ごとの返信
        </button>
      </div>

      <div className="section">
        <h3>どれにも当てはまらないとき</h3>
        <Field label="返す文面" issues={issuesAt(issues, `${path}/fallbackMessageId`).concat(issuesAt(issues, `${path}/cases`).filter((i) => i.path === `${path}/cases`))}>
          <Select
            value={a.fallbackMessageId ?? ''}
            empty="返信しない"
            options={options.messages(project)}
            onChange={(v) => set((x) => (v ? (x.fallbackMessageId = v) : delete x.fallbackMessageId))}
          />
        </Field>
      </div>

      <div className="row" style={{ marginTop: 12 }}>
        <span className="spacer" />
        <button
          className="btn danger"
          onClick={() => {
            if (!window.confirm('この自動応答を削除しますか？')) return;
            update((d) => d.autoReplies.splice(ai, 1));
            onDeleted();
          }}
        >
          この自動応答を削除
        </button>
      </div>
    </>
  );
}
