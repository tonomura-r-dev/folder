// 文面の編集：吹き出し（文字・画像・ボタン・カード・カルーセル）を並べる。

import type { Bubble, Button, CarouselCard, MessageDef, Project } from '../../../engine/src/index.ts';
import { bubbleButtons } from '../../../engine/src/index.ts';
import { Labels } from '../../../player/src/describe.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { replace } from '../route.ts';
import { Field, InlineIssues, ItemTools, move, newId, ReadOnly, Select, Text, TextArea } from '../ui/form.tsx';
import { ActionEditor, options } from '../ui/pickers.tsx';
import { SplitList, UsedBy } from '../ui/split.tsx';

const KIND = [
  { value: 'text', label: '文字' },
  { value: 'buttons', label: 'ボタン付き' },
  { value: 'card', label: 'カード（見出し＋ボタン）' },
  { value: 'image', label: '画像' },
  { value: 'carousel', label: 'カルーセル（横に並ぶカード）' },
];

export const buttonIds = (p: Project) =>
  p.messages.flatMap((m) => m.bubbles.flatMap((b) => bubbleButtons(b).map((x) => x.id)));

const cardIds = (p: Project) => p.messages.flatMap((m) => m.bubbles.flatMap((b) => (b.kind === 'carousel' ? b.cards.map((c) => c.id) : [])));
const newCard = (p: Project): CarouselCard => ({
  id: newId('card', cardIds(p)),
  title: '',
  buttons: [{ id: newId('btn', buttonIds(p)), label: '詳しく見る', action: { kind: 'none' } }],
});
/** 「ラベル：値」の行 ⇔ カードの項目 */
const factsText = (c: CarouselCard) => (c.facts ?? []).map((f) => `${f.label}：${f.value}`).join('\n');
const parseFacts = (t: string) =>
  t
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => {
      const i = l.search(/[：:]/);
      return i < 0 ? { label: l, value: '' } : { label: l.slice(0, i).trim(), value: l.slice(i + 1).trim() };
    });

export function usedBy(p: Project, messageId: string): string[] {
  const out: string[] = [];
  for (const f of p.flows)
    for (const n of f.nodes)
      if ((n.type === 'message' || n.type === 'choice') && n.messageId === messageId) out.push(`${f.label}／${n.label ?? n.id}`);
      else if (n.type === 'survey' && n.introMessageId === messageId) out.push(`${f.label}／${n.label ?? n.id}`);
  for (const a of p.autoReplies)
    if (a.fallbackMessageId === messageId || a.cases.some((c) => c.messageId === messageId)) out.push(`自動応答「${a.match.keywords.join('・')}」`);
  return out;
}

function emptyBubble(kind: string, p: Project): Bubble {
  const btn = (): Button => ({ id: newId('btn', buttonIds(p)), label: 'ボタン', action: { kind: 'none' } });
  switch (kind) {
    case 'buttons':
      return { kind: 'buttons', text: '', buttons: [btn()] };
    case 'card':
      return { kind: 'card', title: '', buttons: [btn()] };
    case 'image':
      return { kind: 'image', assetId: '' };
    case 'carousel':
      return { kind: 'carousel', cards: [newCard(p)] };
    default:
      return { kind: 'text', text: '' };
  }
}

export function Messages() {
  const { project, update, slug } = useEditor();
  const L = new Labels(project);
  return (
    <SplitList<MessageDef>
      section="messages"
      pathBase="/messages"
      title="文面"
      items={project.messages}
      row={(m) => ({ t: L.message(m.id), s: `${m.id}　吹き出し${m.bubbles.length}個` })}
      onAdd={() => {
        const id = newId('msg', project.messages.map((m) => m.id));
        update((d) => d.messages.push({ id, bubbles: [{ kind: 'text', text: '新しい文面' }] }));
        return id;
      }}
      detail={(m, mi, path) => <MessageDetail m={m} mi={mi} path={path} onDeleted={() => replace({ page: 'project', slug, section: 'messages' })} />}
    />
  );
}

function MessageDetail({ m, mi, path, onDeleted }: { m: MessageDef; mi: number; path: string; onDeleted: () => void }) {
  const { project, update, issues } = useEditor();
  const set = (fn: (x: MessageDef) => void) => update((d) => fn(d.messages[mi]!));
  const users = usedBy(project, m.id);
  const own = issues.filter((i) => i.path === path || i.path === `${path}/bubbles`);
  return (
    <>
      <h2>文面</h2>
      <div className="row">
        <ReadOnly>{m.id}</ReadOnly>
        <UsedBy list={users} />
      </div>
      {own.length > 0 && <InlineIssues issues={own} />}
      <div className="muted small" style={{ marginTop: 6 }}>
        差し込み：<span className="mono">{'{{user.name}}'}</span>（名前）／<span className="mono">{'{{brand.accountName}}'}</span>（アカウント名）／
        <span className="mono">{'{{answer.アンケートID.質問ID}}'}</span>（回答）
      </div>

      {m.bubbles.map((b, bi) => {
        const bp = `${path}/bubbles/${bi}`;
        const setB = (fn: (x: Bubble) => void) => set((x) => fn(x.bubbles[bi]!));
        return (
          <div className="section" key={bi}>
            <h3>
              吹き出し {bi + 1}
              <span className="spacer" />
              <ItemTools index={bi} length={m.bubbles.length} onMove={(to) => set((x) => (x.bubbles = move(x.bubbles, bi, to)))} onRemove={() => set((x) => x.bubbles.splice(bi, 1))} />
            </h3>
            <div className="fields">
              <Field label="種類">
                <Select value={b.kind} options={KIND} onChange={(k) => set((x) => (x.bubbles[bi] = emptyBubble(k, project)))} />
              </Field>
            </div>
            {b.kind === 'text' && (
              <Field label="本文" wide issues={issuesAt(issues, `${bp}/text`)}>
                <TextArea rows={5} value={b.text} onChange={(v) => setB((x) => x.kind === 'text' && (x.text = v))} />
              </Field>
            )}
            {b.kind === 'image' && (
              <Field label="画像" issues={issuesAt(issues, `${bp}/assetId`)} hint="画像は「案件設定」で登録します（取り込みは Phase 3）">
                <Select value={b.assetId} empty="選んでください" options={options.assets(project)} onChange={(v) => setB((x) => x.kind === 'image' && (x.assetId = v))} />
              </Field>
            )}
            {b.kind === 'buttons' && (
              <Field label="本文" wide issues={issuesAt(issues, `${bp}/text`)}>
                <TextArea rows={3} value={b.text} onChange={(v) => setB((x) => x.kind === 'buttons' && (x.text = v))} />
              </Field>
            )}
            {b.kind === 'card' && (
              <div className="fields">
                <Field label="見出し">
                  <Text value={b.title} onChange={(v) => setB((x) => x.kind === 'card' && (x.title = v))} />
                </Field>
                <Field label="説明（任意）">
                  <Text value={b.body ?? ''} onChange={(v) => setB((x) => x.kind === 'card' && (v ? (x.body = v) : delete x.body))} />
                </Field>
              </div>
            )}
            {b.kind === 'carousel' && (
              <>
                {b.cards.map((c, ci) => {
                  const cp = `${bp}/cards/${ci}`;
                  const setC = (fn: (x: CarouselCard) => void) => setB((x) => x.kind === 'carousel' && fn(x.cards[ci]!));
                  return (
                    <div className="sub" key={c.id}>
                      <div className="sub-head">
                        <span className="t">カード {ci + 1}</span>
                        <span className="muted small mono">{c.id}</span>
                        <span className="spacer" />
                        <ItemTools
                          index={ci}
                          length={b.cards.length}
                          onMove={(to) => setB((x) => x.kind === 'carousel' && (x.cards = move(x.cards, ci, to)))}
                          onRemove={() => setB((x) => x.kind === 'carousel' && x.cards.splice(ci, 1))}
                        />
                      </div>
                      <div className="fields">
                        <Field label="見出し" issues={issuesAt(issues, `${cp}/title`)}>
                          <Text value={c.title} onChange={(v) => setC((x) => (x.title = v))} />
                        </Field>
                        <Field label="画像（任意）" issues={issuesAt(issues, `${cp}/assetId`)}>
                          <Select value={c.assetId ?? ''} empty="なし" options={options.assets(project)} onChange={(v) => setC((x) => (v ? (x.assetId = v) : delete x.assetId))} />
                        </Field>
                        <Field label="画像の札（任意）" hint="例：デモ用サンプル">
                          <Text value={c.badge ?? ''} onChange={(v) => setC((x) => (v ? (x.badge = v) : delete x.badge))} />
                        </Field>
                        <Field label="説明（任意）">
                          <Text value={c.body ?? ''} onChange={(v) => setC((x) => (v ? (x.body = v) : delete x.body))} />
                        </Field>
                      </div>
                      <Field label="項目（1行に「ラベル：値」）" wide hint="例：勤務地：長野県（スキー場）">
                        <TextArea rows={4} value={factsText(c)} onChange={(v) => setC((x) => (v.trim() ? (x.facts = parseFacts(v)) : delete x.facts))} />
                      </Field>
                      {c.buttons.map((btn, k) => (
                        <div className="sub" key={btn.id}>
                          <div className="sub-head">
                            <span className="t">ボタン {k + 1}</span>
                            <span className="muted small mono">{btn.id}</span>
                            <span className="spacer" />
                            <ItemTools index={k} length={c.buttons.length} onMove={(to) => setC((x) => (x.buttons = move(x.buttons, k, to)))} onRemove={() => setC((x) => x.buttons.splice(k, 1))} />
                          </div>
                          <div className="fields">
                            <Field label="ボタンの文字">
                              <Text value={btn.label} onChange={(v) => setC((x) => (x.buttons[k]!.label = v))} />
                            </Field>
                          </div>
                          <ActionEditor value={btn.action} path={`${cp}/buttons/${k}/action`} onChange={(a) => setC((x) => (x.buttons[k]!.action = a))} />
                        </div>
                      ))}
                      <button
                        className="btn small"
                        style={{ marginTop: 8 }}
                        onClick={() =>
                          update((d) => {
                            const x = d.messages[mi]!.bubbles[bi]!;
                            if (x.kind === 'carousel') x.cards[ci]!.buttons.push({ id: newId('btn', buttonIds(d)), label: 'ボタン', action: { kind: 'none' } });
                          })
                        }
                      >
                        ＋ ボタン
                      </button>
                    </div>
                  );
                })}
                <button
                  className="btn small"
                  style={{ marginTop: 8 }}
                  onClick={() =>
                    update((d) => {
                      const x = d.messages[mi]!.bubbles[bi]!;
                      if (x.kind === 'carousel') x.cards.push(newCard(d));
                    })
                  }
                >
                  ＋ カード
                </button>
              </>
            )}
            {(b.kind === 'buttons' || b.kind === 'card') && (
              <>
                {b.buttons.map((btn, k) => (
                  <div className="sub" key={btn.id}>
                    <div className="sub-head">
                      <span className="t">ボタン {k + 1}</span>
                      <span className="muted small mono">{btn.id}</span>
                      <span className="spacer" />
                      <ItemTools
                        index={k}
                        length={b.buttons.length}
                        onMove={(to) => setB((x) => (x.kind === 'buttons' || x.kind === 'card') && (x.buttons = move(x.buttons, k, to)))}
                        onRemove={() => setB((x) => (x.kind === 'buttons' || x.kind === 'card') && x.buttons.splice(k, 1))}
                      />
                    </div>
                    <div className="fields">
                      <Field label="ボタンの文字">
                        <Text value={btn.label} onChange={(v) => setB((x) => (x.kind === 'buttons' || x.kind === 'card') && (x.buttons[k]!.label = v))} />
                      </Field>
                    </div>
                    <ActionEditor
                      value={btn.action}
                      path={`${bp}/buttons/${k}/action`}
                      onChange={(a) => setB((x) => (x.kind === 'buttons' || x.kind === 'card') && (x.buttons[k]!.action = a))}
                    />
                  </div>
                ))}
                <button
                  className="btn small"
                  style={{ marginTop: 8 }}
                  onClick={() =>
                    update((d) => {
                      const x = d.messages[mi]!.bubbles[bi]!;
                      if (x.kind === 'buttons' || x.kind === 'card') x.buttons.push({ id: newId('btn', buttonIds(d)), label: 'ボタン', action: { kind: 'none' } });
                    })
                  }
                >
                  ＋ ボタン
                </button>
              </>
            )}
          </div>
        );
      })}

      <div className="row" style={{ marginTop: 12 }}>
        <button className="btn" onClick={() => set((x) => x.bubbles.push({ kind: 'text', text: '' }))}>
          ＋ 吹き出し
        </button>
        <span className="spacer" />
        <button
          className="btn danger"
          onClick={() => {
            if (!window.confirm(users.length ? `この文面は ${users.length} か所で使われています。削除すると、そこが「文面が見つからない」になります。削除しますか？` : 'この文面を削除しますか？')) return;
            update((d) => d.messages.splice(mi, 1));
            onDeleted();
          }}
        >
          この文面を削除
        </button>
      </div>
    </>
  );
}
