// アンケート：質問・選択肢・選んだときに付ける／外すタグ。

import type { Survey } from '../../../engine/src/index.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { replace } from '../route.ts';
import { Checks, Field, InlineIssues, ItemTools, move, newId, ReadOnly, Select, Text } from '../ui/form.tsx';
import { options } from '../ui/pickers.tsx';
import { SplitList, UsedBy } from '../ui/split.tsx';

export function Surveys() {
  const { project, update, slug } = useEditor();
  return (
    <SplitList<Survey>
      section="surveys"
      pathBase="/surveys"
      title="アンケート"
      items={project.surveys}
      row={(s) => ({ t: s.label, s: `${s.questions.length}問` })}
      onAdd={() => {
        const id = newId('sv', project.surveys.map((s) => s.id));
        update((d) =>
          d.surveys.push({ id, label: '新しいアンケート', questions: [{ id: 'q_1', text: '質問', mode: 'single', choices: [{ id: 'c_1', label: '選択肢1' }] }] }),
        );
        return id;
      }}
      detail={(s, si, path) => <SurveyDetail s={s} si={si} path={path} onDeleted={() => replace({ page: 'project', slug, section: 'surveys' })} />}
    />
  );
}

function SurveyDetail({ s, si, path, onDeleted }: { s: Survey; si: number; path: string; onDeleted: () => void }) {
  const { project, update, issues } = useEditor();
  const set = (fn: (x: Survey) => void) => update((d) => fn(d.surveys[si]!));
  const used = project.flows.flatMap((f) => f.nodes.filter((n) => n.type === 'survey' && n.surveyId === s.id).map((n) => `${f.label}／${n.label ?? n.id}`));
  return (
    <>
      <h2>アンケート</h2>
      <div className="row">
        <ReadOnly>{s.id}</ReadOnly>
        <UsedBy list={used} />
      </div>
      <InlineIssues issues={issues.filter((i) => i.path === path || i.path === `${path}/questions`)} />
      <div className="section">
        <div className="fields">
          <Field label="アンケート名">
            <Text value={s.label} onChange={(v) => set((x) => (x.label = v))} />
          </Field>
        </div>
      </div>
      {s.questions.map((q, qi) => {
        const qp = `${path}/questions/${qi}`;
        const setQ = (fn: (x: Survey['questions'][number]) => void) => set((x) => fn(x.questions[qi]!));
        return (
          <div className="section" key={q.id}>
            <h3>
              質問 {qi + 1}　<span className="muted small mono">{q.id}</span>
              <span className="spacer" />
              <ItemTools index={qi} length={s.questions.length} onMove={(to) => set((x) => (x.questions = move(x.questions, qi, to)))} onRemove={() => set((x) => x.questions.splice(qi, 1))} />
            </h3>
            <div className="fields">
              <Field label="質問文" issues={issuesAt(issues, `${qp}/text`)}>
                <Text value={q.text} onChange={(v) => setQ((x) => (x.text = v))} />
              </Field>
              <Field label="答え方">
                <Select
                  value={q.mode}
                  options={[
                    { value: 'single', label: '1つだけ選ぶ' },
                    { value: 'multiple', label: '複数選べる' },
                  ]}
                  onChange={(v) => setQ((x) => (x.mode = v as typeof q.mode))}
                />
              </Field>
            </div>
            <InlineIssues issues={issues.filter((i) => i.path === `${qp}/choices`)} />
            {q.choices.map((c, ci) => (
              <div className="sub" key={c.id}>
                <div className="sub-head">
                  <span className="t">選択肢 {ci + 1}</span>
                  <span className="muted small mono">{c.id}</span>
                  <span className="spacer" />
                  <ItemTools index={ci} length={q.choices.length} onMove={(to) => setQ((x) => (x.choices = move(x.choices, ci, to)))} onRemove={() => setQ((x) => x.choices.splice(ci, 1))} />
                </div>
                <div className="fields">
                  <Field label="選択肢の文字">
                    <Text value={c.label} onChange={(v) => setQ((x) => (x.choices[ci]!.label = v))} />
                  </Field>
                  <Field label="選んだら付けるタグ" issues={issuesAt(issues, `${qp}/choices/${ci}/addTags`)}>
                    <Checks value={c.addTags ?? []} options={options.tags(project)} onChange={(v) => setQ((x) => (v.length ? (x.choices[ci]!.addTags = v) : delete x.choices[ci]!.addTags))} />
                  </Field>
                  <Field label="選んだら外すタグ" hint="同じ質問の別の選択肢のタグを外すと、答え直したときに入れ替わります">
                    <Checks
                      value={c.removeTags ?? []}
                      options={options.tags(project)}
                      onChange={(v) => setQ((x) => (v.length ? (x.choices[ci]!.removeTags = v) : delete x.choices[ci]!.removeTags))}
                    />
                  </Field>
                </div>
              </div>
            ))}
            <button className="btn small" style={{ marginTop: 8 }} onClick={() => setQ((x) => x.choices.push({ id: newId('c', s.questions.flatMap((qq) => qq.choices.map((cc) => cc.id))), label: '選択肢' }))}>
              ＋ 選択肢
            </button>
          </div>
        );
      })}
      <div className="row" style={{ marginTop: 12 }}>
        <button className="btn" onClick={() => set((x) => x.questions.push({ id: newId('q', x.questions.map((qq) => qq.id)), text: '質問', mode: 'single', choices: [] }))}>
          ＋ 質問
        </button>
        <span className="spacer" />
        <button
          className="btn danger"
          onClick={() => {
            if (!window.confirm(used.length ? `${used.length} か所で使われています。削除しますか？` : 'このアンケートを削除しますか？')) return;
            update((d) => d.surveys.splice(si, 1));
            onDeleted();
          }}
        >
          このアンケートを削除
        </button>
      </div>
    </>
  );
}
