// 案件一覧：新規作成・開く・複製・削除（ゴミ箱）・JSONから読み込み。

import { useEffect, useRef, useState } from 'react';
import type { Project } from '../../engine/src/index.ts';
import { api, ApiError, fmtDateTime, type ProjectSummary } from './api.ts';
import { go } from './route.ts';
import { Field, Modal, Text } from './ui/form.tsx';

const SLUG = /^[a-z0-9][a-z0-9_-]{0,63}$/;
const defaultSlug = (base = 'project') => {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, '0');
  return `${base}-${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
};

type Dialog =
  | { kind: 'new' }
  | { kind: 'duplicate'; from: ProjectSummary }
  | { kind: 'import'; project: Project; fileName: string }
  | { kind: 'trash'; target: ProjectSummary };

export function ProjectList({ toast }: { toast: (m: string, k?: 'ok' | 'err') => void }) {
  const [items, setItems] = useState<ProjectSummary[] | null>(null);
  const [dialog, setDialog] = useState<Dialog | null>(null);
  const file = useRef<HTMLInputElement>(null);

  const reload = () =>
    api
      .list()
      .then(setItems)
      .catch((e: Error) => toast(`一覧を読めません：${e.message}`, 'err'));
  useEffect(() => void reload(), []);

  const onFile = async (f: File | undefined) => {
    if (!f) return;
    try {
      const project = JSON.parse(await f.text()) as Project;
      setDialog({ kind: 'import', project, fileName: f.name });
    } catch {
      toast('JSONとして読めないファイルです', 'err');
    }
    if (file.current) file.current.value = '';
  };

  return (
    <div className="page">
      <div className="page-head">
        <h1>LINE施策シミュレーター　案件一覧</h1>
        <span className="spacer" />
        <button className="btn" onClick={() => file.current?.click()}>
          JSONから読み込み
        </button>
        <input ref={file} type="file" accept=".json,application/json" hidden onChange={(e) => void onFile(e.target.files?.[0])} />
        <button className="btn primary" onClick={() => setDialog({ kind: 'new' })}>
          ＋ 新しい案件
        </button>
      </div>

      {items === null ? (
        <div className="empty">読み込み中…</div>
      ) : items.length === 0 ? (
        <div className="empty">案件がまだありません。「＋ 新しい案件」から作ってください。</div>
      ) : (
        <div className="cards">
          {items.map((it) => (
            <div className="card" key={it.slug}>
              <h2>{it.name}</h2>
              <div className="muted small">
                {it.industry ? `${it.industry}　` : ''}フォルダ：<span className="mono">{it.slug}</span>
              </div>
              <div className="muted small">最終保存：{fmtDateTime(it.updatedAt)}</div>
              <div className="row">
                <button className="btn primary" onClick={() => go({ page: 'project', slug: it.slug, section: 'flows' })}>
                  開く
                </button>
                <button className="btn" onClick={() => setDialog({ kind: 'duplicate', from: it })}>
                  複製
                </button>
                <span className="spacer" />
                <button className="btn ghost danger" onClick={() => setDialog({ kind: 'trash', target: it })}>
                  削除
                </button>
              </div>
            </div>
          ))}
        </div>
      )}

      {dialog && (
        <ListDialog
          dialog={dialog}
          existing={new Set(items?.map((i) => i.slug))}
          onClose={() => setDialog(null)}
          onDone={(msg, slug) => {
            setDialog(null);
            toast(msg);
            if (slug) go({ page: 'project', slug, section: 'flows' });
            else void reload();
          }}
          onError={(e) => toast(e, 'err')}
        />
      )}
    </div>
  );
}

function ListDialog({
  dialog,
  existing,
  onClose,
  onDone,
  onError,
}: {
  dialog: Dialog;
  existing: Set<string>;
  onClose: () => void;
  onDone: (msg: string, slug?: string) => void;
  onError: (msg: string) => void;
}) {
  const initialName =
    dialog.kind === 'duplicate' ? `${dialog.from.name}（コピー）` : dialog.kind === 'import' ? dialog.project.meta?.name ?? '' : '';
  const initialSlug =
    dialog.kind === 'duplicate'
      ? `${dialog.from.slug}-copy`
      : dialog.kind === 'import'
        ? existing.has(dialog.project.meta?.slug ?? '') || !SLUG.test(dialog.project.meta?.slug ?? '')
          ? defaultSlug('import')
          : dialog.project.meta.slug
        : defaultSlug();
  const [name, setName] = useState(initialName);
  const [slug, setSlug] = useState(initialSlug);
  const [busy, setBusy] = useState(false);

  if (dialog.kind === 'trash') {
    return (
      <Modal title="案件を削除" onClose={onClose}>
        <p>
          「{dialog.target.name}」をゴミ箱（<span className="mono">.trash</span> フォルダ）へ移します。完全には消えないので、フォルダを戻せば復元できます。
        </p>
        <div className="actions">
          <button className="btn" onClick={onClose}>
            やめる
          </button>
          <button
            className="btn primary"
            disabled={busy}
            onClick={async () => {
              setBusy(true);
              try {
                await api.trash(dialog.target.slug);
                onDone('ゴミ箱へ移しました');
              } catch (e) {
                onError((e as Error).message);
                setBusy(false);
              }
            }}
          >
            ゴミ箱へ移す
          </button>
        </div>
      </Modal>
    );
  }

  const slugBad = !SLUG.test(slug) ? '英小文字・数字・-・_ で付けてください' : existing.has(slug) ? '同じフォルダ名の案件があります' : null;
  const title = dialog.kind === 'new' ? '新しい案件' : dialog.kind === 'duplicate' ? '案件を複製' : `JSONから読み込み（${dialog.fileName}）`;

  const submit = async () => {
    setBusy(true);
    try {
      if (dialog.kind === 'new') await api.create(slug, name.trim());
      else if (dialog.kind === 'duplicate') await api.duplicate(dialog.from.slug, slug, name.trim());
      else await api.import(slug, { ...dialog.project, meta: { ...dialog.project.meta, name: name.trim(), slug } });
      onDone(dialog.kind === 'new' ? '作成しました' : dialog.kind === 'duplicate' ? '複製しました' : '読み込みました', slug);
    } catch (e) {
      onError(e instanceof ApiError ? e.message : String(e));
      setBusy(false);
    }
  };

  return (
    <Modal title={title} onClose={onClose}>
      <div className="fields">
        <Field label="案件名">
          <Text value={name} onChange={setName} placeholder="例：〇〇株式会社（仮）" />
        </Field>
        <Field label="フォルダ名（英数字）" hint={slugBad ?? `projects/${slug}/ に保存されます`}>
          <Text value={slug} onChange={(v) => setSlug(v.toLowerCase())} />
        </Field>
      </div>
      <div className="actions">
        <button className="btn" onClick={onClose}>
          やめる
        </button>
        <button className="btn primary" disabled={busy || !!slugBad || !name.trim()} onClick={() => void submit()}>
          {dialog.kind === 'new' ? '作成' : dialog.kind === 'duplicate' ? '複製' : '読み込む'}
        </button>
      </div>
    </Modal>
  );
}
