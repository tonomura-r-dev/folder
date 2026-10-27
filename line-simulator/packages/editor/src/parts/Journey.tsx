// CVまでの道筋：デモ画面の③に出す、営業説明用の段（Webサイト訪問 → 友だち追加 → … → CV）。
// シナリオの動きには関係しない。段ごとに「来たとみなす条件」と「人によって変わるルート」を決める。

import type { JourneyStep } from '../../../engine/src/index.ts';
import { journeySteps } from '../../../player/src/journey.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { TYPE_LABEL } from '../flow/convert.ts';
import { replace } from '../route.ts';
import { ExprField } from '../ui/ExprField.tsx';
import { Checks, Field, ItemTools, move, newId, Select, Text, TextArea } from '../ui/form.tsx';
import { options } from '../ui/pickers.tsx';
import { SplitList } from '../ui/split.tsx';

export function Journey() {
  const { project, update, slug, issues } = useEditor();
  const steps = project.journey ?? [];
  const defaultWhen = () => (project.tags[0] ? { op: 'hasTag' as const, tagId: project.tags[0].id } : { op: 'always' as const });

  const draft = () => {
    // シナリオの並びとCVから下書きを作る（デモの③の簡易表示と同じもの）
    const auto = journeySteps({ ...project, journey: [] }).steps.map((s, i) => ({ ...s, id: `j_${i + 1}` }));
    update((d) => (d.journey = auto));
    replace({ page: 'project', slug, section: 'journey', item: auto[0]?.id ?? '' });
  };

  return (
    <SplitList<JourneyStep>
      section="journey"
      pathBase="/journey"
      title="CVまでの道筋"
      items={steps}
      row={(s) => ({ t: s.label || '（名前なし）', ...(s.purpose ? { s: s.purpose } : {}) })}
      onAdd={() => {
        const id = newId('j', steps.map((s) => s.id));
        update((d) => (d.journey ??= []).push({ id, label: '新しい段', purpose: '' }));
        return id;
      }}
      empty={
        <div>
          <p>デモ画面の③に、友だち追加からCVまでの流れを、営業先に説明する言葉で並べます。</p>
          <p className="muted small">未設定のあいだは、シナリオの一覧から自動で並べた簡易表示になります。</p>
          <button className="btn small" onClick={draft}>
            シナリオから下書きを作る
          </button>
        </div>
      }
      detail={(s, i, path) => {
        const set = (fn: (x: JourneyStep) => void) => update((d) => fn(d.journey![i]!));
        const variants = s.variants ?? [];
        return (
          <>
            <h2>道筋の段 {i + 1}</h2>
            <div className="row" style={{ marginBottom: 8 }}>
              <span className="muted small">並び順＝デモで上から表示する順。CV地点（ゴール）は最後の方に置きます。</span>
              <span className="spacer" />
              <ItemTools
                index={i}
                length={steps.length}
                onMove={(to) => {
                  update((d) => (d.journey = move(d.journey!, i, to)));
                }}
                onRemove={() => {
                  update((d) => d.journey!.splice(i, 1));
                  replace({ page: 'project', slug, section: 'journey' });
                }}
              />
            </div>
            <div className="section">
              <div className="fields">
                <Field label="段の名前" issues={issuesAt(issues, `${path}/label`)}>
                  <Text value={s.label} onChange={(v) => set((x) => (x.label = v))} />
                </Field>
                <Field label="CV地点にする（任意）" hint="このCVをしたら、この段に来たことにします。ゴールとして★で表示します" issues={issuesAt(issues, `${path}/cvId`)}>
                  <Select value={s.cvId ?? ''} empty="CV地点ではない" options={options.cvs(project)} onChange={(v) => set((x) => (v ? (x.cvId = v) : delete x.cvId))} />
                </Field>
                <Field label="この施策の目的（営業先への説明）" wide>
                  <TextArea rows={2} value={s.purpose ?? ''} onChange={(v) => set((x) => (v ? (x.purpose = v) : delete x.purpose))} />
                </Field>
              </div>
            </div>

            <div className="section">
              <h3>この段に来たとみなす条件（どれか1つで「済」）</h3>
              <Field label="シナリオのこのノードを通った" wide issues={issuesAt(issues, path).filter((x) => x.path === path).concat(issuesAt(issues, `${path}/nodeIds`))}>
                {project.flows.map((f) => (
                  <div key={f.id} style={{ marginBottom: 6 }}>
                    <div className="muted small">{f.label}</div>
                    <Checks
                      value={s.nodeIds ?? []}
                      options={f.nodes.filter((n) => n.type !== 'entry').map((n) => ({ value: n.id, label: `${TYPE_LABEL[n.type]} ${n.label ?? n.id}` }))}
                      onChange={(v) => {
                        const others = (s.nodeIds ?? []).filter((id) => !f.nodes.some((n) => n.id === id));
                        const next = [...others, ...v];
                        set((x) => (next.length ? (x.nodeIds = next) : delete x.nodeIds));
                      }}
                    />
                  </div>
                ))}
              </Field>
              <label className="row" style={{ gap: 6, margin: '6px 0' }}>
                <input type="checkbox" checked={!!s.reachedWhen} onChange={(e) => set((x) => (e.target.checked ? (x.reachedWhen = defaultWhen()) : delete x.reachedWhen))} />
                条件式でも判定する（例：ページを見たことがある、タグが付いている）
              </label>
              {s.reachedWhen && <ExprField label="この条件に当てはまったら来たことにする" path={`${path}/reachedWhen`} value={s.reachedWhen} onChange={(e) => set((x) => (x.reachedWhen = e))} />}
            </div>

            <div className="section">
              <h3>人によって変わるルート（任意）</h3>
              <div className="muted small" style={{ marginBottom: 6 }}>
                上から順に、最初に当てはまったものを「この人のルート」として強調します。例：東京／大阪／未選択
              </div>
              {variants.map((x, vi) => (
                <div className="sub" key={x.id}>
                  <div className="sub-head">
                    <span className="t">ルート {vi + 1}</span>
                    <span className="spacer" />
                    <ItemTools
                      index={vi}
                      length={variants.length}
                      onMove={(to) => set((y) => (y.variants = move(y.variants!, vi, to)))}
                      onRemove={() => set((y) => (y.variants!.splice(vi, 1), y.variants!.length || delete y.variants))}
                    />
                  </div>
                  <div className="fields">
                    <Field label="ルートの名前" issues={issuesAt(issues, `${path}/variants/${vi}/label`)}>
                      <Text value={x.label} onChange={(v) => set((y) => (y.variants![vi]!.label = v))} />
                    </Field>
                    <Field label="このルートでやること" wide>
                      <TextArea rows={2} value={x.purpose ?? ''} onChange={(v) => set((y) => (v ? (y.variants![vi]!.purpose = v) : delete y.variants![vi]!.purpose))} />
                    </Field>
                    <ExprField label="この条件の人" path={`${path}/variants/${vi}/when`} value={x.when} onChange={(e) => set((y) => (y.variants![vi]!.when = e))} />
                  </div>
                </div>
              ))}
              <button
                className="btn small"
                style={{ marginTop: 8 }}
                onClick={() =>
                  set((y) => (y.variants ??= []).push({ id: newId('v', steps.flatMap((t) => (t.variants ?? []).map((z) => z.id))), label: '新しいルート', when: defaultWhen() }))
                }
              >
                ＋ ルート
              </button>
            </div>
          </>
        );
      }}
    />
  );
}
