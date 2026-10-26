// フロー編集の右側：選んだノード／線の設定。何も選んでいなければシナリオ全体の設定。

import type { Expr, Flow, FlowNode } from '../../../engine/src/index.ts';
import { describeExpr, Labels } from '../../../player/src/describe.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { go, replace } from '../route.ts';
import { Checks, Field, InlineIssues, ItemTools, move, newId, NumberInput, ReadOnly, Select, Text } from '../ui/form.tsx';
import { DurationInput, EVENT_OPTIONS, options, TriggerEditor } from '../ui/pickers.tsx';
import { ExprField } from '../ui/ExprField.tsx';
import { ProvNote } from '../ai/ProvBadge.tsx';
import { portLabel, TYPE_LABEL } from './convert.ts';

export function Inspector({ flow }: { flow: Flow }) {
  const { route } = useEditor();
  const node = route.node ? flow.nodes.find((n) => n.id === route.node) : undefined;
  const edge = !node && route.edge ? flow.edges.find((e) => e.id === route.edge) : undefined;
  return (
    <aside className="inspector" aria-label="設定">
      {node ? <NodeInspector flow={flow} node={node} /> : edge ? <EdgeInspector flow={flow} edgeId={edge.id} /> : <FlowSettings flow={flow} />}
    </aside>
  );
}

// ---------- シナリオ全体 ----------

function FlowSettings({ flow }: { flow: Flow }) {
  const { project, update, issues, route } = useEditor();
  const fi = project.flows.indexOf(flow);
  const path = `/flows/${fi}`;
  const set = (fn: (f: Flow) => void) => update((d) => fn(d.flows[fi]!));
  const L = new Labels(project);
  return (
    <>
      <h2>シナリオの設定</h2>
      <div className="muted small">ノードを選ぶと、そのノードの設定に切り替わります。</div>
      <ProvNote prov={flow.prov} path={path} />
      <div className="section">
        <div className="fields">
          <Field label="シナリオ名" issues={issuesAt(issues, `${path}/label`)}>
            <Text value={flow.label} onChange={(v) => set((f) => (f.label = v))} />
          </Field>
          <Field label="ID">
            <ReadOnly>{flow.id}</ReadOnly>
          </Field>
        </div>
      </div>
      <div className="section">
        <h3>入口</h3>
        <TriggerEditor value={flow.trigger} path={`${path}/trigger`} onChange={(t) => set((f) => (f.trigger = t))} />
      </div>
      <div className="section">
        <h3>配信の制御</h3>
        <div className="fields">
          <Field label="優先度（小さいほど優先）" hint="同じ時刻に別のシナリオと配信が重なったとき、優先度の高い方だけが送ります">
            <NumberInput value={flow.priority} onChange={(v) => set((f) => (f.priority = v ?? 0))} />
          </Field>
          <Field label="重なったとき">
            <Select
              value={flow.concurrency}
              options={[
                { value: 'exclusive', label: '優先度で譲る（標準）' },
                { value: 'parallel', label: '並行して送る' },
              ]}
              onChange={(v) => set((f) => (f.concurrency = v as Flow['concurrency']))}
            />
          </Field>
          <Field label="同じ人が再び入れるか">
            <Select
              value={flow.reentry}
              options={[
                { value: 'never', label: '1回だけ' },
                { value: 'afterEnd', label: '終わった後なら再び入れる' },
                { value: 'always', label: 'いつでも入り直す' },
              ]}
              onChange={(v) => set((f) => (f.reentry = v as Flow['reentry']))}
            />
          </Field>
        </div>
      </div>
      <div className="section">
        <h3>
          停止条件
          <span className="spacer" />
          {flow.stopWhen ? (
            <button className="btn small ghost danger" onClick={() => set((f) => delete f.stopWhen)}>
              外す
            </button>
          ) : (
            <button className="btn small" onClick={() => set((f) => (f.stopWhen = project.cvs[0] ? { op: 'cv', cvId: project.cvs[0].id, done: true } : { op: 'always' }))}>
              ＋ 停止条件を付ける
            </button>
          )}
        </h3>
        {flow.stopWhen ? (
          <ExprField label="この条件に当てはまったら、このシナリオを止める" path={`${path}/stopWhen`} value={flow.stopWhen} onChange={(e) => set((f) => (f.stopWhen = e))} />
        ) : (
          <div className="muted small">停止条件はありません。{flow.nodes.some((n) => n.type === 'wait') && '待機のあるシナリオは、CV済みの人にも送り続けないよう停止条件を付けるのがおすすめです。'}</div>
        )}
      </div>
      <div className="section">
        <h3>このシナリオを</h3>
        <div className="row">
          <button
            className="btn danger"
            onClick={() => {
              const callers = project.flows.filter((f) => f.nodes.some((n) => n.type === 'jump' && n.flowId === flow.id)).map((f) => f.label);
              const msg = `「${flow.label}」を削除しますか？${callers.length ? `\n\n次のシナリオから呼ばれています：${callers.join('、')}` : ''}\n（保存するまでは元に戻せます：保存せずに一覧へ戻る）`;
              if (!window.confirm(msg)) return;
              update((d) => (d.flows = d.flows.filter((f) => f.id !== flow.id)));
              replace({ ...route, item: undefined, node: undefined, edge: undefined } as typeof route);
            }}
          >
            削除
          </button>
        </div>
        <div className="muted small" style={{ marginTop: 6 }}>
          入口：{L.flow(flow.id)} ／ ノード {flow.nodes.length}個・線 {flow.edges.length}本
        </div>
      </div>
    </>
  );
}

// ---------- 線 ----------

function EdgeInspector({ flow, edgeId }: { flow: Flow; edgeId: string }) {
  const { project, update, route } = useEditor();
  const fi = project.flows.indexOf(flow);
  const e = flow.edges.find((x) => x.id === edgeId)!;
  const from = flow.nodes.find((n) => n.id === e.from.nodeId);
  const to = flow.nodes.find((n) => n.id === e.to);
  return (
    <>
      <h2>線</h2>
      <div className="section">
        <div className="fields">
          <Field label="出発">
            <ReadOnly>
              {from ? `${TYPE_LABEL[from.type]} ${from.label ?? from.id}` : e.from.nodeId}　の「{from ? portLabel(from, e.from.port, project) : e.from.port}」
            </ReadOnly>
          </Field>
          <Field label="行き先">
            <ReadOnly>{to ? `${TYPE_LABEL[to.type]} ${to.label ?? to.id}` : `（存在しない：${e.to}）`}</ReadOnly>
          </Field>
        </div>
        <div className="muted small" style={{ marginTop: 8 }}>
          線の端をドラッグすると繋ぎ替えられます。Delete キーで削除できます。
        </div>
        <div className="row" style={{ marginTop: 8 }}>
          <button
            className="btn danger"
            onClick={() => {
              update((d) => (d.flows[fi]!.edges = d.flows[fi]!.edges.filter((x) => x.id !== edgeId)));
              replace({ ...route, edge: undefined } as typeof route);
            }}
          >
            この線を削除
          </button>
        </div>
      </div>
    </>
  );
}

// ---------- ノード ----------

function NodeInspector({ flow, node }: { flow: Flow; node: FlowNode }) {
  const { project, update, issues, route, slug } = useEditor();
  const fi = project.flows.indexOf(flow);
  const ni = flow.nodes.indexOf(node);
  const path = `/flows/${fi}/nodes/${ni}`;
  const at = (k: string) => issuesAt(issues, `${path}/${k}`);
  /** このノードを書き換える（型は呼び出し側で合わせる） */
  const set = <T extends FlowNode>(fn: (n: T) => void) => update((d) => fn(d.flows[fi]!.nodes[ni] as T));
  const own = issues.filter((i) => i.path === path);
  const editMessage = (id: string) => go({ page: 'project', slug, section: 'messages', item: id });

  return (
    <>
      <h2>
        {TYPE_LABEL[node.type]}　<span className="muted small mono">{node.id}</span>
      </h2>
      {own.length > 0 && <InlineIssues issues={own} />}
      <ProvNote prov={node.prov} path={path} />
      <div className="section">
        <div className="fields">
          <Field label="見出し（任意）" hint="フローチャート上の名前。営業先に見せる画面には出ません">
            <Text value={node.label ?? ''} onChange={(v) => set((n) => (v ? (n.label = v) : delete n.label))} />
          </Field>
        </div>
      </div>

      <div className="section">
        {node.type === 'entry' && <div className="muted small">はじまりのノードです。いつ始まるかは、何も選んでいないときの「シナリオの設定」で変えます。</div>}

        {(node.type === 'message' || node.type === 'choice') && (
          <div className="fields">
            <Field label="送る文面" issues={at('messageId')} hint={node.type === 'choice' ? 'ボタン付きの文面だけ選べます。ボタンごとに出口ができます' : undefined}>
              <Select
                value={node.messageId}
                empty="選んでください"
                options={options.messages(project, node.type === 'choice')}
                onChange={(v) => set<typeof node>((n) => (n.messageId = v))}
              />
            </Field>
            <div className="row">
              {node.messageId && (
                <button className="btn small" onClick={() => editMessage(node.messageId)}>
                  この文面を編集
                </button>
              )}
              <button
                className="btn small"
                onClick={() => {
                  const id = newId('msg', project.messages.map((m) => m.id));
                  update((d) => {
                    d.messages.push(
                      node.type === 'choice'
                        ? { id, bubbles: [{ kind: 'buttons', text: '選んでください', buttons: [{ id: newId('btn', allButtonIds(d)), label: 'はい', action: { kind: 'none' } }] }] }
                        : { id, bubbles: [{ kind: 'text', text: '新しい文面' }] },
                    );
                    (d.flows[fi]!.nodes[ni] as typeof node).messageId = id;
                  });
                }}
              >
                ＋ 新しい文面を作って使う
              </button>
            </div>
            {node.type === 'choice' && (
              <Field label="回答を待つ時間の上限（任意）" hint="過ぎると「時間切れ」の出口へ進みます">
                <OptionalDuration value={node.timeoutMinutes} onChange={(v) => set<typeof node>((n) => (v === undefined ? delete n.timeoutMinutes : (n.timeoutMinutes = v)))} />
              </Field>
            )}
          </div>
        )}

        {node.type === 'survey' && (
          <div className="fields">
            <Field label="アンケート" issues={at('surveyId')}>
              <Select value={node.surveyId} empty="選んでください" options={options.surveys(project)} onChange={(v) => set<typeof node>((n) => (n.surveyId = v))} />
            </Field>
            <Field label="最初に送る文面（任意）" issues={at('introMessageId')}>
              <Select
                value={node.introMessageId ?? ''}
                empty="なし"
                options={options.messages(project)}
                onChange={(v) => set<typeof node>((n) => (v ? (n.introMessageId = v) : delete n.introMessageId))}
              />
            </Field>
            <Field label="回答を待つ時間の上限（任意）" hint="過ぎると「時間切れ」の出口へ進みます。全問に答えると「回答した」へ">
              <OptionalDuration value={node.timeoutMinutes} onChange={(v) => set<typeof node>((n) => (v === undefined ? delete n.timeoutMinutes : (n.timeoutMinutes = v)))} />
            </Field>
          </div>
        )}

        {node.type === 'tag' && (
          <div className="fields">
            <Field label="付けるタグ" issues={at('add')}>
              <Checks value={node.add ?? []} options={options.tags(project)} onChange={(v) => set<typeof node>((n) => (n.add = v))} />
            </Field>
            <Field label="外すタグ" issues={at('remove')}>
              <Checks
                value={node.remove ?? []}
                options={options.tags(project)}
                onChange={(v) => set<typeof node>((n) => (v.length ? (n.remove = v) : delete n.remove))}
              />
            </Field>
          </div>
        )}

        {node.type === 'branch' && <BranchCases flow={flow} node={node} path={path} />}

        {node.type === 'wait' && <WaitEditor node={node} path={path} set={set} />}

        {node.type === 'image' && (
          <Field label="送る画像" issues={at('assetId')} hint="画像は「案件設定」で登録します（画像の取り込みは Phase 3）">
            <Select value={node.assetId} empty="選んでください" options={options.assets(project)} onChange={(v) => set<typeof node>((n) => (n.assetId = v))} />
          </Field>
        )}

        {node.type === 'richMenu' && (
          <Field label="切り替え先のメニュー" issues={at('richMenuId')}>
            <Select value={node.richMenuId} empty="選んでください" options={options.richMenus(project)} onChange={(v) => set<typeof node>((n) => (n.richMenuId = v))} />
          </Field>
        )}

        {node.type === 'page' && (
          <Field label="開くページ" issues={at('pageId')}>
            <Select value={node.pageId} empty="選んでください" options={options.pages(project)} onChange={(v) => set<typeof node>((n) => (n.pageId = v))} />
          </Field>
        )}

        {node.type === 'jump' && (
          <>
            <Field label="次に始めるシナリオ" issues={at('flowId')} hint="このシナリオは終わり、選んだシナリオが始まります">
              <Select
                value={node.flowId}
                empty="選んでください"
                options={options.flows(project).filter((o) => o.value !== flow.id)}
                onChange={(v) => set<typeof node>((n) => (n.flowId = v))}
              />
            </Field>
            {node.flowId && project.flows.some((f) => f.id === node.flowId) && (
              <button className="btn small" style={{ marginTop: 8 }} onClick={() => replace({ ...route, item: node.flowId, node: undefined, edge: undefined } as typeof route)}>
                行き先のシナリオを開く
              </button>
            )}
          </>
        )}

        {(node.type === 'end' || node.type === 'stopAll') && (
          <Field label="メモ（任意）" hint={node.type === 'stopAll' ? 'この人のほかのシナリオもすべて止めます（CVしたとき等）' : 'このシナリオを終えます'}>
            <Text value={node.reason ?? ''} onChange={(v) => set<typeof node>((n) => (v ? (n.reason = v) : delete n.reason))} />
          </Field>
        )}
      </div>

      {node.type !== 'entry' && (
        <div className="row" style={{ marginTop: 12 }}>
          <button
            className="btn danger"
            onClick={() => {
              update((d) => {
                const f = d.flows[fi]!;
                f.nodes = f.nodes.filter((n) => n.id !== node.id);
                f.edges = f.edges.filter((e) => e.from.nodeId !== node.id && e.to !== node.id);
              });
              replace({ ...route, node: undefined } as typeof route);
            }}
          >
            このノードを削除
          </button>
          <span className="muted small">選んで Delete キーでも削除できます</span>
        </div>
      )}
    </>
  );
}

function OptionalDuration({ value, onChange }: { value: number | undefined; onChange: (v: number | undefined) => void }) {
  return (
    <div className="row" style={{ flexWrap: 'nowrap' }}>
      <label className="row small" style={{ flexWrap: 'nowrap' }}>
        <input type="checkbox" checked={value !== undefined} onChange={(e) => onChange(e.target.checked ? 1440 : undefined)} />
        上限あり
      </label>
      {value !== undefined && <DurationInput value={value} onChange={onChange} />}
    </div>
  );
}

function WaitEditor({ node, path, set }: { node: Extract<FlowNode, { type: 'wait' }>; path: string; set: <T extends FlowNode>(fn: (n: T) => void) => void }) {
  const { issues } = useEditor();
  const w = node.wait;
  const at = (k: string) => issuesAt(issues, `${path}/wait/${k}`);
  const setWait = (fn: (n: typeof node) => void) => set<typeof node>(fn);
  return (
    <div className="fields">
      <Field label="待ち方">
        <Select
          value={w.kind}
          options={[
            { value: 'dayAt', label: '◯日目の◯時に送る（ステップ配信）' },
            { value: 'duration', label: '一定の時間を待つ' },
            { value: 'untilEvent', label: 'ユーザーの操作を待つ' },
          ]}
          onChange={(k) =>
            setWait((n) => {
              n.wait = k === 'duration' ? { kind: 'duration', minutes: 60 } : k === 'untilEvent' ? { kind: 'untilEvent', event: 'tapButton', timeoutMinutes: 1440 } : { kind: 'dayAt', day: 1, time: '09:00' };
            })
          }
        />
      </Field>
      {w.kind === 'dayAt' && (
        <>
          <Field label="起点" hint="友だち追加前に始まるシナリオでは、自動的にシナリオ開始が起点になります">
            <Select
              value={w.from ?? 'friendAdd'}
              options={[
                { value: 'friendAdd', label: '友だち追加した日' },
                { value: 'flowStart', label: 'このシナリオが始まった日' },
              ]}
              onChange={(v) => setWait((n) => n.wait.kind === 'dayAt' && (v === 'friendAdd' ? delete n.wait.from : (n.wait.from = 'flowStart')))}
            />
          </Field>
          <Field label="何日目（0＝当日）" issues={at('day')}>
            <NumberInput value={w.day} min={0} onChange={(v) => setWait((n) => n.wait.kind === 'dayAt' && (n.wait.day = v ?? 0))} />
          </Field>
          <Field label="時刻（HH:MM）" issues={at('time')}>
            <Text value={w.time} onChange={(v) => setWait((n) => n.wait.kind === 'dayAt' && (n.wait.time = v))} placeholder="09:00" />
          </Field>
        </>
      )}
      {w.kind === 'duration' && (
        <Field label="待つ時間" issues={at('minutes')}>
          <DurationInput value={w.minutes} onChange={(v) => setWait((n) => n.wait.kind === 'duration' && (n.wait.minutes = v))} />
        </Field>
      )}
      {w.kind === 'untilEvent' && (
        <>
          <Field label="待つ操作" issues={at('event')}>
            <Select value={w.event} options={EVENT_OPTIONS} onChange={(v) => setWait((n) => n.wait.kind === 'untilEvent' && (n.wait.event = v as typeof w.event))} />
          </Field>
          <Field label="待つ時間の上限" issues={at('timeoutMinutes')} hint="過ぎると「時間切れ」の出口へ進みます">
            <DurationInput value={w.timeoutMinutes} onChange={(v) => setWait((n) => n.wait.kind === 'untilEvent' && (n.wait.timeoutMinutes = v))} />
          </Field>
        </>
      )}
    </div>
  );
}

function BranchCases({ flow, node, path }: { flow: Flow; node: Extract<FlowNode, { type: 'branch' }>; path: string }) {
  const { project, update } = useEditor();
  const fi = project.flows.indexOf(flow);
  const ni = flow.nodes.indexOf(node);
  const setNode = (fn: (n: typeof node) => void) => update((d) => fn(d.flows[fi]!.nodes[ni] as typeof node));
  const L = new Labels(project);
  return (
    <>
      <div className="muted small">上から順に判定し、最初に当てはまった条件の出口へ進みます。どれにも当てはまらなければ「それ以外」へ。</div>
      {node.cases.map((c, i) => (
        <div className="sub" key={c.id}>
          <div className="sub-head">
            <span className="t">条件 {i + 1}</span>
            <span className="spacer" />
            <ItemTools
              index={i}
              length={node.cases.length}
              onMove={(to) => setNode((n) => (n.cases = move(n.cases, i, to)))}
              onRemove={() =>
                update((d) => {
                  const f = d.flows[fi]!;
                  const n = f.nodes[ni] as typeof node;
                  n.cases = n.cases.filter((x) => x.id !== c.id);
                  f.edges = f.edges.filter((e) => !(e.from.nodeId === node.id && e.from.port === `case:${c.id}`));
                })
              }
            />
          </div>
          <div className="fields">
            <Field label="出口の名前">
              <Text value={c.label} onChange={(v) => setNode((n) => (n.cases[i]!.label = v))} />
            </Field>
            <ExprField label="条件" path={`${path}/cases/${i}/when`} value={c.when} onChange={(e: Expr) => setNode((n) => (n.cases[i]!.when = e))} />
          </div>
        </div>
      ))}
      <div className="row" style={{ marginTop: 10 }}>
        <button
          className="btn small"
          onClick={() =>
            setNode((n) =>
              n.cases.push({
                id: newId('c', n.cases.map((x) => x.id)),
                label: `条件${n.cases.length + 1}`,
                when: project.tags[0] ? { op: 'hasTag', tagId: project.tags[0].id } : { op: 'always' },
              }),
            )
          }
        >
          ＋ 条件を追加
        </button>
      </div>
      <div className="muted small" style={{ marginTop: 8 }}>
        いまの判定順：{node.cases.map((c) => `${c.label}（${safeDescribe(c.when, L)}）`).join(' → ')} → それ以外
      </div>
    </>
  );
}

const safeDescribe = (e: Expr, L: Labels) => {
  try {
    return describeExpr(e, L);
  } catch {
    return '読めない条件';
  }
};

const allButtonIds = (p: { messages: { bubbles: { kind: string; buttons?: { id: string }[] }[] }[] }) =>
  p.messages.flatMap((m) => m.bubbles.flatMap((b) => b.buttons?.map((x) => x.id) ?? []));
