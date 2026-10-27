// CVまでの道筋：デモ画面の③に出す、営業説明用のフローチャート（START → 施策 → 分岐 → … → GOAL）。
// シナリオの動きには関係しない。段ごとに「来たとみなす条件」「分岐の行き先」「紐づく文面（シナリオを参照）」
// 「Lステップでの実現性」「追加候補」を決める。最初の段が START、最後のCV地点が GOAL になる。

import type { Feasibility, Id, JourneyOption, JourneyStep, Project } from '../../../engine/src/index.ts';
import { FEASIBILITY, journeySteps } from '../../../player/src/journey.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { TYPE_LABEL } from '../flow/convert.ts';
import { replace } from '../route.ts';
import { ExprField } from '../ui/ExprField.tsx';
import { Checks, Field, ItemTools, move, newId, Select, Text, TextArea } from '../ui/form.tsx';
import { options } from '../ui/pickers.tsx';
import { SplitList } from '../ui/split.tsx';

const FEAS_OPTIONS = (Object.keys(FEASIBILITY) as Feasibility[]).map((k) => ({ value: k, label: `${FEASIBILITY[k].mark} ${FEASIBILITY[k].text}` }));

/** シナリオのノードを選ぶ（シナリオごとに並べる）。文面・アンケートのノードを選ぶと、その文面が③に出る */
function NodePicker({ project, value, onChange }: { project: Project; value: Id[]; onChange: (v: Id[]) => void }) {
  return (
    <>
      {project.flows.map((f) => (
        <div key={f.id} style={{ marginBottom: 6 }}>
          <div className="muted small">{f.label}</div>
          <Checks
            value={value}
            options={f.nodes.filter((n) => n.type !== 'entry').map((n) => ({ value: n.id, label: `${TYPE_LABEL[n.type]} ${n.label ?? n.id}` }))}
            onChange={(v) => onChange([...value.filter((id) => !f.nodes.some((n) => n.id === id)), ...v])}
          />
        </div>
      ))}
    </>
  );
}

/** シナリオに無い文面（自動応答で返す文面など）を選ぶ */
function ReplyPicker({ project, value, onChange }: { project: Project; value: Id[]; onChange: (v: Id[]) => void }) {
  const replies = new Set(project.autoReplies.flatMap((a) => [...a.cases.map((c) => c.messageId), ...(a.fallbackMessageId ? [a.fallbackMessageId] : [])]));
  const list = options.messages(project).filter((o) => replies.has(o.value) || value.includes(o.value));
  return <Checks value={value} options={list} onChange={onChange} />;
}

type Target = { feasibility?: Feasibility; feasibilityNote?: string; options?: JourneyOption[] };

/** Lステップでの実現性と、追加候補（表示だけ・配信されない）の入力。段と分岐の行き先で共通 */
function Proposal({ t, path, set, used }: { t: Target; path: string; set: (fn: (x: Target) => void) => void; used: () => Id[] }) {
  const { issues } = useEditor();
  const opts = t.options ?? [];
  return (
    <>
      <div className="fields">
        <Field label="Lステップでの実現性" hint="確認していない機能は「未確認」にしてください（可と断定しない）" issues={issuesAt(issues, `${path}/feasibility`)}>
          <Select value={t.feasibility ?? ''} empty="指定しない" options={FEAS_OPTIONS} onChange={(v) => set((x) => (v ? (x.feasibility = v as Feasibility) : delete x.feasibility))} />
        </Field>
        <Field label="実現性の補足（任意）" wide>
          <TextArea rows={2} value={t.feasibilityNote ?? ''} onChange={(v) => set((x) => (v ? (x.feasibilityNote = v) : delete x.feasibilityNote))} />
        </Field>
      </div>
      <div className="muted small" style={{ margin: '10px 0 4px' }}>
        追加候補：今回は採用していないが、提案できる施策。③に点線で表示するだけで、シナリオには入らず配信もされません。
      </div>
      {opts.map((o, oi) => (
        <div className="sub" key={o.id}>
          <div className="sub-head">
            <span className="t">追加候補 {oi + 1}</span>
            <span className="spacer" />
            <ItemTools
              index={oi}
              length={opts.length}
              onMove={(to) => set((x) => (x.options = move(x.options!, oi, to)))}
              onRemove={() => set((x) => (x.options!.splice(oi, 1), x.options!.length || delete x.options))}
            />
          </div>
          <div className="fields">
            <Field label="施策名" issues={issuesAt(issues, `${path}/options/${oi}/label`)}>
              <Text value={o.label} onChange={(v) => set((x) => (x.options![oi]!.label = v))} />
            </Field>
            <Field label="Lステップでの実現性" issues={issuesAt(issues, `${path}/options/${oi}/feasibility`)}>
              <Select value={o.feasibility ?? ''} empty="指定しない" options={FEAS_OPTIONS} onChange={(v) => set((x) => (v ? (x.options![oi]!.feasibility = v as Feasibility) : delete x.options![oi]!.feasibility))} />
            </Field>
            <Field label="概要" wide>
              <TextArea rows={2} value={o.summary ?? ''} onChange={(v) => set((x) => (v ? (x.options![oi]!.summary = v) : delete x.options![oi]!.summary))} />
            </Field>
            <Field label="文面の例（任意）" wide>
              <TextArea rows={3} value={o.sample ?? ''} onChange={(v) => set((x) => (v ? (x.options![oi]!.sample = v) : delete x.options![oi]!.sample))} />
            </Field>
            <Field label="補足（任意）" wide>
              <Text value={o.note ?? ''} onChange={(v) => set((x) => (v ? (x.options![oi]!.note = v) : delete x.options![oi]!.note))} />
            </Field>
          </div>
        </div>
      ))}
      <button className="btn small" style={{ marginTop: 8 }} onClick={() => set((x) => (x.options ??= []).push({ id: newId('o', used()), label: '新しい候補' }))}>
        ＋ 追加候補
      </button>
    </>
  );
}

export function Journey() {
  const { project, update, slug, issues } = useEditor();
  const steps = project.journey ?? [];
  const defaultWhen = () => (project.tags[0] ? { op: 'hasTag' as const, tagId: project.tags[0].id } : { op: 'always' as const });
  const optionIds = () => steps.flatMap((s) => [...(s.options ?? []), ...(s.variants ?? []).flatMap((x) => x.options ?? [])].map((o) => o.id));

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
      row={(s) => ({ t: `${s.variants?.length ? '◇ ' : ''}${s.label || '（名前なし）'}`, ...(s.purpose ? { s: s.purpose } : {}) })}
      onAdd={() => {
        const id = newId('j', steps.map((s) => s.id));
        update((d) => (d.journey ??= []).push({ id, label: '新しい施策', purpose: '' }));
        return id;
      }}
      empty={
        <div>
          <p>デモ画面の③に、友だち追加からCVまでの流れをフローチャートで出します（営業先への説明用）。</p>
          <p className="muted small">未設定のあいだは、シナリオの一覧から自動で並べた簡易表示になります。</p>
          <button className="btn small" onClick={draft}>
            シナリオから下書きを作る
          </button>
        </div>
      }
      detail={(s, i, path) => {
        const set = (fn: (x: JourneyStep) => void) => update((d) => fn(d.journey![i]!));
        const variants = s.variants ?? [];
        const isGoal = s.cvId !== undefined && steps.reduce((g, x, k) => (x.cvId !== undefined ? k : g), -1) === i;
        return (
          <>
            <h2>
              {i === 0 ? 'START：' : isGoal ? 'GOAL：' : ''}
              {s.label || '（名前なし）'}
            </h2>
            <div className="row" style={{ marginBottom: 8 }}>
              <span className="muted small">いちばん上の段が START、いちばん下のCV地点が GOAL になります。並び順＝③で上から表示する順。</span>
              <span className="spacer" />
              <ItemTools
                index={i}
                length={steps.length}
                onMove={(to) => update((d) => (d.journey = move(d.journey!, i, to)))}
                onRemove={() => {
                  update((d) => d.journey!.splice(i, 1));
                  replace({ page: 'project', slug, section: 'journey' });
                }}
              />
            </div>
            <div className="section">
              <div className="fields">
                <Field label="施策名（箱に出す名前）" issues={issuesAt(issues, `${path}/label`)}>
                  <Text value={s.label} onChange={(v) => set((x) => (x.label = v))} />
                </Field>
                <Field label="CV地点にする（任意）" hint="このCVをしたら、この段に来たことにします" issues={issuesAt(issues, `${path}/cvId`)}>
                  <Select value={s.cvId ?? ''} empty="CV地点ではない" options={options.cvs(project)} onChange={(v) => set((x) => (v ? (x.cvId = v) : delete x.cvId))} />
                </Field>
                <Field label="この施策の目的（営業先への説明）" wide>
                  <TextArea rows={2} value={s.purpose ?? ''} onChange={(v) => set((x) => (v ? (x.purpose = v) : delete x.purpose))} />
                </Field>
              </div>
            </div>

            <div className="section">
              <h3>紐づく文面・来たとみなす条件</h3>
              <div className="muted small" style={{ marginBottom: 6 }}>
                選んだノードを通ったら「済」。文面・アンケートのノードは、その中身が③の箱と詳細に出ます（文面はシナリオのものをそのまま使います）。
              </div>
              <Field label="シナリオのノード" wide issues={issuesAt(issues, path).filter((x) => x.path === path).concat(issuesAt(issues, `${path}/nodeIds`))}>
                <NodePicker project={project} value={s.nodeIds ?? []} onChange={(v) => set((x) => (v.length ? (x.nodeIds = v) : delete x.nodeIds))} />
              </Field>
              <Field label="自動応答の文面（任意）" wide issues={issuesAt(issues, `${path}/messageIds`)}>
                <ReplyPicker project={project} value={s.messageIds ?? []} onChange={(v) => set((x) => (v.length ? (x.messageIds = v) : delete x.messageIds))} />
              </Field>
              <label className="row" style={{ gap: 6, margin: '6px 0' }}>
                <input type="checkbox" checked={!!s.reachedWhen} onChange={(e) => set((x) => (e.target.checked ? (x.reachedWhen = defaultWhen()) : delete x.reachedWhen))} />
                条件式でも判定する（例：友だち追加したことがある）
              </label>
              {s.reachedWhen && <ExprField label="この条件に当てはまったら来たことにする" path={`${path}/reachedWhen`} value={s.reachedWhen} onChange={(e) => set((x) => (x.reachedWhen = e))} />}
            </div>

            <div className="section">
              <h3>分岐（任意）</h3>
              <div className="muted small" style={{ marginBottom: 6 }}>
                行き先を足すと、③でひし形の分岐になります。上から順に、最初に条件に当てはまった行き先が「この人のルート」です（②の判定と同じ条件式）。
              </div>
              {variants.length > 0 && (
                <div className="fields">
                  <Field label="ひし形に書く問い" hint="例：地域選択済み？" wide>
                    <Text value={s.decision ?? ''} onChange={(v) => set((x) => (v ? (x.decision = v) : delete x.decision))} />
                  </Field>
                </div>
              )}
              {variants.map((x, vi) => {
                const vp = `${path}/variants/${vi}`;
                const setV = (fn: (y: NonNullable<JourneyStep['variants']>[number]) => void) => set((y) => fn(y.variants![vi]!));
                return (
                  <div className="sub" key={x.id}>
                    <div className="sub-head">
                      <span className="t">行き先 {vi + 1}</span>
                      <span className="spacer" />
                      <ItemTools
                        index={vi}
                        length={variants.length}
                        onMove={(to) => set((y) => (y.variants = move(y.variants!, vi, to)))}
                        onRemove={() => set((y) => (y.variants!.splice(vi, 1), y.variants!.length || delete y.variants))}
                      />
                    </div>
                    <div className="fields">
                      <Field label="矢印に付ける名前" hint="例：東京" issues={issuesAt(issues, `${vp}/label`)}>
                        <Text value={x.label} onChange={(v) => setV((y) => (y.label = v))} />
                      </Field>
                      <Field label="行き先の施策名" hint="例：東京向け求人配信（空なら矢印の名前）">
                        <Text value={x.title ?? ''} onChange={(v) => setV((y) => (v ? (y.title = v) : delete y.title))} />
                      </Field>
                      <Field label="この行き先でやること" wide>
                        <TextArea rows={2} value={x.purpose ?? ''} onChange={(v) => setV((y) => (v ? (y.purpose = v) : delete y.purpose))} />
                      </Field>
                      <ExprField label="この条件の人" path={`${vp}/when`} value={x.when} onChange={(e) => setV((y) => (y.when = e))} />
                      <Field label="この行き先の施策＝シナリオのノード（通ったら「通った経路」）" wide issues={issuesAt(issues, `${vp}/nodeIds`)}>
                        <NodePicker project={project} value={x.nodeIds ?? []} onChange={(v) => setV((y) => (v.length ? (y.nodeIds = v) : delete y.nodeIds))} />
                      </Field>
                      <Field label="自動応答の文面（任意）" wide issues={issuesAt(issues, `${vp}/messageIds`)}>
                        <ReplyPicker project={project} value={x.messageIds ?? []} onChange={(v) => setV((y) => (v.length ? (y.messageIds = v) : delete y.messageIds))} />
                      </Field>
                    </div>
                    <Proposal t={x} path={vp} set={(fn) => setV((y) => fn(y))} used={optionIds} />
                  </div>
                );
              })}
              <button
                className="btn small"
                style={{ marginTop: 8 }}
                onClick={() => set((y) => (y.variants ??= []).push({ id: newId('v', steps.flatMap((t) => (t.variants ?? []).map((z) => z.id))), label: '新しい行き先', when: defaultWhen() }))}
              >
                ＋ 分岐の行き先
              </button>
            </div>

            <div className="section">
              <h3>Lステップでの実現性・追加候補（この段）</h3>
              <Proposal t={s} path={path} set={(fn) => set((y) => fn(y))} used={optionIds} />
            </div>
          </>
        );
      }}
    />
  );
}
