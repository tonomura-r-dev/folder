// Webページ：LP・応募フォーム・サンクス。離脱防止ポップアップとLINE誘導もここで設定する。

import type { WebBlock, WebPage } from '../../../engine/src/index.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { replace } from '../route.ts';
import { Field, InlineIssues, ItemTools, move, newId, ReadOnly, Select, Text, TextArea } from '../ui/form.tsx';
import { ActionEditor, options } from '../ui/pickers.tsx';
import { SplitList } from '../ui/split.tsx';

const KIND = [
  { value: 'lp', label: 'LP・サイトのページ' },
  { value: 'form', label: 'フォーム（応募・予約など）' },
  { value: 'thanks', label: 'サンクス（完了）' },
];
const KIND_LABEL: Record<string, string> = { lp: 'LP', form: 'フォーム', thanks: 'サンクス' };

function emptyBlock(kind: string): WebBlock {
  if (kind === 'image') return { kind: 'image', assetId: '' };
  if (kind === 'button') return { kind: 'button', label: 'ボタン', action: { kind: 'none' } };
  return { kind: kind as 'heading' | 'text', text: '' };
}

export function WebPages() {
  const { project, update, slug } = useEditor();
  return (
    <SplitList<WebPage>
      section="webPages"
      pathBase="/webPages"
      title="Webページ"
      items={project.webPages}
      row={(w) => ({ t: w.title, s: `${KIND_LABEL[w.kind]}${w.exitPopup ? '・離脱ポップアップあり' : ''}${w.lineCta ? '・LINE誘導あり' : ''}` })}
      onAdd={() => {
        const id = newId('page', project.webPages.map((w) => w.id));
        update((d) => d.webPages.push({ id, kind: 'lp', title: '新しいページ', blocks: [{ kind: 'heading', text: '見出し' }] }));
        return id;
      }}
      detail={(w, wi, path) => <PageDetail w={w} wi={wi} path={path} onDeleted={() => replace({ page: 'project', slug, section: 'webPages' })} />}
    />
  );
}

function PageDetail({ w, wi, path, onDeleted }: { w: WebPage; wi: number; path: string; onDeleted: () => void }) {
  const { project, update, issues } = useEditor();
  const set = (fn: (x: WebPage) => void) => update((d) => fn(d.webPages[wi]!));
  const at = (k: string) => issuesAt(issues, `${path}/${k}`);
  return (
    <>
      <h2>Webページ</h2>
      <ReadOnly>{w.id}</ReadOnly>
      <InlineIssues issues={issues.filter((i) => i.path === path)} />

      <div className="section">
        <div className="fields">
          <Field label="ページ名">
            <Text value={w.title} onChange={(v) => set((x) => (x.title = v))} />
          </Field>
          <Field label="種類">
            <Select value={w.kind} options={KIND} onChange={(v) => set((x) => (x.kind = v as WebPage['kind']))} />
          </Field>
          <Field label="表示用のURL（任意）" hint="画面の飾りとして出すだけで、開きません。example.com などの架空のものにしてください">
            <Text value={w.url ?? ''} onChange={(v) => set((x) => (v ? (x.url = v) : delete x.url))} />
          </Field>
        </div>
      </div>

      <div className="section">
        <h3>ページの中身</h3>
        {w.blocks.map((b, bi) => (
          <div className="sub" key={bi}>
            <div className="sub-head">
              <select value={b.kind} onChange={(e) => set((x) => (x.blocks[bi] = emptyBlock(e.target.value)))}>
                <option value="heading">見出し</option>
                <option value="text">文章</option>
                <option value="image">画像</option>
                <option value="button">ボタン</option>
                {b.kind === 'list' && <option value="list">一覧（絞り込みつき）</option>}
              </select>
              <span className="spacer" />
              <ItemTools index={bi} length={w.blocks.length} onMove={(to) => set((x) => (x.blocks = move(x.blocks, bi, to)))} onRemove={() => set((x) => x.blocks.splice(bi, 1))} />
            </div>
            {(b.kind === 'heading' || b.kind === 'text') && (
              <Field label={b.kind === 'heading' ? '見出し' : '文章'} wide issues={at(`blocks/${bi}`)}>
                {b.kind === 'heading' ? (
                  <Text value={b.text} onChange={(v) => set((x) => ((x.blocks[bi] as typeof b).text = v))} />
                ) : (
                  <TextArea value={b.text} onChange={(v) => set((x) => ((x.blocks[bi] as typeof b).text = v))} />
                )}
              </Field>
            )}
            {b.kind === 'image' && (
              <Field label="画像" issues={at(`blocks/${bi}/assetId`)}>
                <Select value={b.assetId} empty="選んでください" options={options.assets(project)} onChange={(v) => set((x) => ((x.blocks[bi] as typeof b).assetId = v))} />
              </Field>
            )}
            {b.kind === 'button' && (
              <>
                <div className="fields">
                  <Field label="ボタンの文字">
                    <Text value={b.label} onChange={(v) => set((x) => ((x.blocks[bi] as typeof b).label = v))} />
                  </Field>
                </div>
                <ActionEditor value={b.action} path={`${path}/blocks/${bi}/action`} onChange={(a) => set((x) => ((x.blocks[bi] as typeof b).action = a))} />
              </>
            )}
            {b.kind === 'list' && (
              <div className="hint">
                絞り込み {b.facets.map((f) => f.label).join('・')}／項目 {b.items.length}件（{b.items.map((it) => it.title).join('、')}）。一覧の中身は案件データ（project.json）で編集します。
              </div>
            )}
          </div>
        ))}
        <button className="btn small" style={{ marginTop: 8 }} onClick={() => set((x) => x.blocks.push({ kind: 'text', text: '' }))}>
          ＋ 中身を追加
        </button>
      </div>

      <div className="section">
        <h3>
          フォーム
          <span className="spacer" />
          <Toggle
            on={!!w.form}
            onLabel="フォームを外す"
            offLabel="＋ フォームを付ける"
            onChange={(on) => set((x) => (on ? (x.form = { fields: [{ id: 'f_name', label: 'お名前', type: 'text' }], onSubmit: {} }) : delete x.form))}
          />
        </h3>
        {w.form && (
          <>
            {w.form.fields.map((f, fi) => (
              <div className="sub" key={fi}>
                <div className="sub-head">
                  <span className="t">項目 {fi + 1}</span>
                  <span className="spacer" />
                  <ItemTools
                    index={fi}
                    length={w.form!.fields.length}
                    onMove={(to) => set((x) => (x.form!.fields = move(x.form!.fields, fi, to)))}
                    onRemove={() => set((x) => x.form!.fields.splice(fi, 1))}
                  />
                </div>
                <div className="fields">
                  <Field label="項目名">
                    <Text value={f.label} onChange={(v) => set((x) => (x.form!.fields[fi]!.label = v))} />
                  </Field>
                  <Field label="入力の種類">
                    <Select
                      value={f.type}
                      options={[
                        { value: 'text', label: '文字' },
                        { value: 'tel', label: '電話番号' },
                        { value: 'email', label: 'メール' },
                        { value: 'select', label: '選択' },
                      ]}
                      onChange={(v) => set((x) => (x.form!.fields[fi]!.type = v as typeof f.type))}
                    />
                  </Field>
                  {f.type === 'select' && (
                    <Field label="選択肢（、区切り）">
                      <Text value={(f.options ?? []).join('、')} onChange={(v) => set((x) => (x.form!.fields[fi]!.options = v.split(/[、,]/).map((s) => s.trim())))} />
                    </Field>
                  )}
                </div>
              </div>
            ))}
            <button
              className="btn small"
              style={{ marginTop: 8 }}
              onClick={() => set((x) => x.form!.fields.push({ id: newId('f', x.form!.fields.map((f) => f.id)), label: '項目', type: 'text' }))}
            >
              ＋ 項目
            </button>
            <div className="fields" style={{ marginTop: 12 }}>
              <Field label="送信したら記録するCV" issues={at('form/onSubmit/cvId')}>
                <Select
                  value={w.form.onSubmit.cvId ?? ''}
                  empty="記録しない"
                  options={options.cvs(project)}
                  onChange={(v) => set((x) => (v ? (x.form!.onSubmit.cvId = v) : delete x.form!.onSubmit.cvId))}
                />
              </Field>
              <Field label="送信後に開くページ" issues={at('form/onSubmit/nextPageId')}>
                <Select
                  value={w.form.onSubmit.nextPageId ?? ''}
                  empty="移動しない"
                  options={options.pages(project).filter((o) => o.value !== w.id)}
                  onChange={(v) => set((x) => (v ? (x.form!.onSubmit.nextPageId = v) : delete x.form!.onSubmit.nextPageId))}
                />
              </Field>
            </div>
          </>
        )}
      </div>

      <div className="section">
        <h3>
          離脱防止ポップアップ
          <span className="spacer" />
          <Toggle
            on={!!w.exitPopup}
            onLabel="ポップアップを外す"
            offLabel="＋ ポップアップを付ける"
            onChange={(on) =>
              set((x) =>
                on
                  ? (x.exitPopup = {
                      title: 'LINEで受け取りませんか？',
                      body: '',
                      cta: { label: 'LINEで受け取る', action: project.routes[0] ? { kind: 'addFriend', routeId: project.routes[0].id } : { kind: 'none' } },
                      dismissLabel: '閉じる',
                    })
                  : delete x.exitPopup,
              )
            }
          />
        </h3>
        {w.exitPopup ? (
          <>
            <div className="fields">
              <Field label="見出し" issues={at('exitPopup/title')}>
                <Text value={w.exitPopup.title} onChange={(v) => set((x) => (x.exitPopup!.title = v))} />
              </Field>
              <Field label="閉じるボタンの文字">
                <Text value={w.exitPopup.dismissLabel} onChange={(v) => set((x) => (x.exitPopup!.dismissLabel = v))} />
              </Field>
              <Field label="本文" wide issues={at('exitPopup/body')}>
                <TextArea rows={2} value={w.exitPopup.body} onChange={(v) => set((x) => (x.exitPopup!.body = v))} />
              </Field>
              <Field label="誘導ボタンの文字">
                <Text value={w.exitPopup.cta.label} onChange={(v) => set((x) => (x.exitPopup!.cta.label = v))} />
              </Field>
            </div>
            <ActionEditor value={w.exitPopup.cta.action} path={`${path}/exitPopup/cta/action`} onChange={(a) => set((x) => (x.exitPopup!.cta.action = a))} />
          </>
        ) : (
          <div className="muted small">このページから離脱しようとしたときに出すポップアップです。</div>
        )}
      </div>

      <div className="section">
        <h3>
          LINE誘導（サンクスページ）
          <span className="spacer" />
          <Toggle
            on={!!w.lineCta}
            onLabel="外す"
            offLabel="＋ LINE誘導を付ける"
            onChange={(on) => set((x) => (on ? (x.lineCta = { label: 'LINEで受け取る', routeId: project.routes[0]?.id ?? '' }) : delete x.lineCta))}
          />
        </h3>
        {w.lineCta ? (
          <div className="fields">
            <Field label="ボタンの文字">
              <Text value={w.lineCta.label} onChange={(v) => set((x) => (x.lineCta!.label = v))} />
            </Field>
            <Field label="友だち追加の経路" issues={at('lineCta/routeId')}>
              <Select value={w.lineCta.routeId} empty="選んでください" options={options.routes(project)} onChange={(v) => set((x) => (x.lineCta!.routeId = v))} />
            </Field>
          </div>
        ) : (
          <div className="muted small">完了ページに出す「LINEで受け取る」ボタンです。</div>
        )}
      </div>

      <div className="row" style={{ marginTop: 12 }}>
        <span className="spacer" />
        <button
          className="btn danger"
          onClick={() => {
            if (!window.confirm('このページを削除しますか？')) return;
            update((d) => d.webPages.splice(wi, 1));
            onDeleted();
          }}
        >
          このページを削除
        </button>
      </div>
    </>
  );
}

function Toggle({ on, onLabel, offLabel, onChange }: { on: boolean; onLabel: string; offLabel: string; onChange: (on: boolean) => void }) {
  return (
    <button className={`btn small${on ? ' ghost danger' : ''}`} onClick={() => onChange(!on)}>
      {on ? onLabel : offLabel}
    </button>
  );
}
