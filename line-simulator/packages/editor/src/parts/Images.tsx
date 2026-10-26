// 画像の取り込みと一覧。ファイルは案件フォルダの assets/ に保存し、案件にはそのパスを登録する。

import { useRef, useState } from 'react';
import { api, assetUrl, readImage } from '../api.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { Field, newId, Text } from '../ui/form.tsx';

export function Images() {
  const { project: p, update, issues, slug, toast } = useEditor();
  const file = useRef<HTMLInputElement>(null);
  const [busy, setBusy] = useState(false);

  const upload = async (files: FileList | null) => {
    if (!files?.length) return;
    setBusy(true);
    try {
      for (const f of Array.from(files)) {
        const { base64, width, height } = await readImage(f);
        const { path } = await api.uploadAsset(slug, f.name, base64);
        update((d) =>
          d.assets.push({ id: newId('img', d.assets.map((a) => a.id)), kind: 'image', path, width, height, alt: f.name.replace(/\.[^.]+$/, '') }),
        );
      }
      toast('画像を取り込みました（案件への登録は保存で確定します）');
    } catch (e) {
      toast(`取り込めませんでした：${(e as Error).message}`, 'err');
    } finally {
      setBusy(false);
      if (file.current) file.current.value = '';
    }
  };

  const usedBy = (id: string) => JSON.stringify({ ...p, assets: [] }).split(`"${id}"`).length - 1;

  return (
    <div className="section">
      <h3>
        画像
        <span className="spacer" />
        <button className="btn small" disabled={busy} onClick={() => file.current?.click()}>
          {busy ? '取り込み中…' : '＋ 画像を取り込む'}
        </button>
        <input ref={file} type="file" accept="image/png,image/jpeg,image/gif,image/webp" multiple hidden onChange={(e) => void upload(e.target.files)} />
      </h3>
      <div className="muted small">
        PNG・JPEG・GIF・WebP（1枚10MBまで）。案件フォルダの assets/ に保存されます。実在の顧客のロゴや写真は使わないでください。
      </div>
      {p.assets.length === 0 && <div className="muted small" style={{ marginTop: 8 }}>まだ画像はありません。</div>}
      {p.assets.map((a, i) => (
        <div className="sub" key={a.id} style={{ display: 'grid', gridTemplateColumns: '120px 1fr', gap: 12 }}>
          <img
            src={assetUrl(slug, a.path)}
            alt={a.alt}
            style={{ width: 120, height: 90, objectFit: 'contain', background: '#eef1f7', borderRadius: 8 }}
          />
          <div>
            <div className="sub-head">
              <span className="t mono">{a.id}</span>
              <span className="muted small">
                {a.width}×{a.height}　使われている数：{usedBy(a.id)}
              </span>
              <span className="spacer" />
              <button
                className="btn small ghost danger"
                onClick={() => {
                  const n = usedBy(a.id);
                  if (n > 0 && !window.confirm(`${n} か所で使われています。登録を外しますか？（ファイル自体は assets/ に残ります）`)) return;
                  update((d) => d.assets.splice(i, 1));
                }}
              >
                登録を外す
              </button>
            </div>
            <div className="fields">
              <Field label="説明（代替テキスト）" issues={issuesAt(issues, `/assets/${i}/alt`)}>
                <Text value={a.alt} onChange={(v) => update((d) => (d.assets[i]!.alt = v))} />
              </Field>
              <Field label="ファイル" issues={issuesAt(issues, `/assets/${i}/path`)}>
                <div className="readonly">{a.path}</div>
              </Field>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
