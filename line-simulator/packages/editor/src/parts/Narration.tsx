// 説明文：デモ画面の右に出す「いま何が起きたか」のカード（八十吉デモの説明カード相当）。

import type { InputEventType, NarrationDef } from '../../../engine/src/index.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { TYPE_LABEL } from '../flow/convert.ts';
import { replace } from '../route.ts';
import { Field, newId, Select, Text, TextArea } from '../ui/form.tsx';
import { EVENT_OPTIONS } from '../ui/pickers.tsx';
import { SplitList } from '../ui/split.tsx';

const EVENTS = [
  ...EVENT_OPTIONS,
  { value: 'exitIntent', label: '離脱しようとする' },
  { value: 'addFriend', label: '友だち追加する' },
  { value: 'block', label: 'ブロックする' },
];

export function Narration() {
  const { project, update, slug, issues } = useEditor();
  const nodeOptions = project.flows.flatMap((f) => f.nodes.map((n) => ({ value: n.id, label: `${f.label}／${TYPE_LABEL[n.type]} ${n.label ?? n.id}` })));
  return (
    <SplitList<NarrationDef>
      section="narration"
      pathBase="/narration"
      title="説明文"
      items={project.narration}
      row={(n) => ({ t: n.title || '（見出しなし）', ...(n.eyebrow ? { s: n.eyebrow } : {}) })}
      onAdd={() => {
        const id = newId('nar', project.narration.map((n) => n.id));
        update((d) => d.narration.push({ id, on: { kind: 'event', type: 'addFriend' }, title: '説明の見出し', body: '' }));
        return id;
      }}
      empty="デモ画面で、場面ごとに出す説明です。"
      detail={(n, i, path) => {
        const set = (fn: (x: NarrationDef) => void) => update((d) => fn(d.narration[i]!));
        return (
          <>
            <h2>説明文</h2>
            <div className="section">
              <div className="fields">
                <Field label="いつ出すか">
                  <Select
                    value={n.on.kind}
                    options={[
                      { value: 'event', label: 'ユーザーが操作したとき' },
                      { value: 'node', label: 'シナリオのノードを通ったとき' },
                    ]}
                    onChange={(k) => set((x) => (x.on = k === 'node' ? { kind: 'node', nodeId: nodeOptions[0]?.value ?? '' } : { kind: 'event', type: 'addFriend' }))}
                  />
                </Field>
                {n.on.kind === 'node' ? (
                  <Field label="ノード" issues={issuesAt(issues, `${path}/on/nodeId`)}>
                    <Select value={n.on.nodeId} empty="選んでください" options={nodeOptions} onChange={(v) => set((x) => (x.on = { kind: 'node', nodeId: v }))} />
                  </Field>
                ) : (
                  <Field label="操作">
                    <Select value={n.on.type} options={EVENTS} onChange={(v) => set((x) => (x.on = { kind: 'event', type: v as InputEventType }))} />
                  </Field>
                )}
                <Field label="小見出し（任意）">
                  <Text value={n.eyebrow ?? ''} onChange={(v) => set((x) => (v ? (x.eyebrow = v) : delete x.eyebrow))} />
                </Field>
                <Field label="見出し" wide issues={issuesAt(issues, `${path}/title`)}>
                  <Text value={n.title} onChange={(v) => set((x) => (x.title = v))} />
                </Field>
                <Field label="本文" wide issues={issuesAt(issues, `${path}/body`)}>
                  <TextArea value={n.body} onChange={(v) => set((x) => (x.body = v))} />
                </Field>
              </div>
            </div>
            <div className="row" style={{ marginTop: 12 }}>
              <span className="spacer" />
              <button
                className="btn danger"
                onClick={() => {
                  update((d) => d.narration.splice(i, 1));
                  replace({ page: 'project', slug, section: 'narration' });
                }}
              >
                この説明文を削除
              </button>
            </div>
          </>
        );
      }}
    />
  );
}
