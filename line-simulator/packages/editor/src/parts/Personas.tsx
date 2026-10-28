// ペルソナ（架空ユーザーの雛形）：名前・属性・初期状態・台本（いつ何をするか）。

import type { Persona, ScriptStep } from '../../../engine/src/index.ts';
import { fmtTime } from '../../../player/src/describe.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { replace } from '../route.ts';
import { Checks, Field, InlineIssues, JsonField, newId, ReadOnly, Select, Text } from '../ui/form.tsx';
import { options } from '../ui/pickers.tsx';
import { SplitList } from '../ui/split.tsx';

export function Personas() {
  const { project, update, slug } = useEditor();
  return (
    <SplitList<Persona>
      section="personas"
      pathBase="/personas"
      title="ペルソナ"
      items={project.personas}
      row={(p) => ({ t: p.name, s: `台本 ${p.script?.length ?? 0}件` })}
      onAdd={() => {
        const id = newId('p', project.personas.map((p) => p.id));
        update((d) => d.personas.push({ id, name: '新しいユーザー', attrs: {}, script: [] }));
        return id;
      }}
      detail={(p, pi, path) => <PersonaDetail p={p} pi={pi} path={path} onDeleted={() => replace({ page: 'project', slug, section: 'personas' })} />}
    />
  );
}

function PersonaDetail({ p, pi, path, onDeleted }: { p: Persona; pi: number; path: string; onDeleted: () => void }) {
  const { project, update, issues } = useEditor();
  const set = (fn: (x: Persona) => void) => update((d) => fn(d.personas[pi]!));
  const attrs = Object.entries(p.attrs);
  return (
    <>
      <h2>ペルソナ</h2>
      <ReadOnly>{p.id}</ReadOnly>
      <div className="section">
        <div className="fields">
          <Field label="役割の名前（デモ画面でユーザーを見分ける名前。例：検討中さん）">
            <Text value={p.name} onChange={(v) => set((x) => (x.name = v))} />
          </Field>
          <Field label="人物名（架空・任意）" hint="入れると文面の {{user.name}} はこの名前になり、画面は「人物名＠役割」で出ます。実在の人の名前は使わないでください">
            <Text value={p.personName ?? ''} onChange={(v) => set((x) => (v ? (x.personName = v) : delete x.personName))} />
          </Field>
          <Field label="最初の状態">
            <Select
              value={p.initial?.friendState ?? 'not_friend'}
              options={[
                { value: 'not_friend', label: 'まだ友だちではない' },
                { value: 'friend', label: 'すでに友だち' },
              ]}
              onChange={(v) => set((x) => (x.initial = { ...x.initial, friendState: v as 'friend' | 'not_friend' }))}
            />
          </Field>
          <Field label="最初から付いているタグ" wide issues={issuesAt(issues, `${path}/initial/tags`)}>
            <Checks value={p.initial?.tags ?? []} options={options.tags(project)} onChange={(v) => set((x) => (x.initial = { ...x.initial, tags: v }))} />
          </Field>
        </div>
      </div>

      <div className="section">
        <h3>
          属性
          <span className="spacer" />
          <button className="btn small" onClick={() => set((x) => (x.attrs[newId('attr', Object.keys(x.attrs))] = ''))}>
            ＋ 属性
          </button>
        </h3>
        {attrs.length === 0 && <div className="muted small">年齢・地域など、条件式の「attr」で使う値です。</div>}
        {attrs.map(([k, v]) => (
          <div className="fields" key={k} style={{ marginTop: 6 }}>
            <Field label="名前（英数字）">
              <Text
                value={k}
                onChange={(nk) =>
                  set((x) => {
                    if (!nk || nk in x.attrs) return;
                    x.attrs = Object.fromEntries(Object.entries(x.attrs).map(([kk, vv]) => (kk === k ? [nk, vv] : [kk, vv])));
                  })
                }
              />
            </Field>
            <Field label="値（数字は数として扱います）">
              <div className="row" style={{ flexWrap: 'nowrap' }}>
                <Text value={String(v)} onChange={(nv) => set((x) => (x.attrs[k] = nv !== '' && !Number.isNaN(Number(nv)) ? Number(nv) : nv))} />
                <button className="btn small ghost danger" onClick={() => set((x) => delete x.attrs[k])}>
                  削除
                </button>
              </div>
            </Field>
          </div>
        ))}
      </div>

      <div className="section">
        <h3>台本（自動で流すときに、この人が取る行動）</h3>
        <div className="muted small" style={{ marginBottom: 6 }}>
          atMinutes はシミュレーション開始からの分（Day 1 の 9:00 ＝ 1440 + 540 ＝ 1980）。画面で操作して記録する方法は Phase 4 で付けます。
        </div>
        {(p.script ?? []).length > 0 && (
          <ol className="small" style={{ margin: '0 0 8px', paddingLeft: 20 }}>
            {(p.script ?? []).map((s, i) => (
              <li key={i}>
                {fmtTime(s.atMinutes)}　{s.event.type}
              </li>
            ))}
          </ol>
        )}
        <Field label="台本（JSON）" wide issues={issuesAt(issues, `${path}/script`)}>
          <JsonField<ScriptStep[]> rows={10} value={p.script ?? []} onChange={(v) => set((x) => (x.script = Array.isArray(v) ? v : []))} />
        </Field>
      </div>

      <InlineIssues issues={issues.filter((i) => i.path === path)} />
      <div className="row" style={{ marginTop: 12 }}>
        <span className="spacer" />
        <button
          className="btn danger"
          onClick={() => {
            if (!window.confirm('このペルソナを削除しますか？')) return;
            update((d) => d.personas.splice(pi, 1));
            onDeleted();
          }}
        >
          このペルソナを削除
        </button>
      </div>
    </>
  );
}
