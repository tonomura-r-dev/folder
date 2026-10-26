// リッチメニュー：区画の並び（列×行）と、押したときの動作。

import { useState } from 'react';
import type { RichMenu } from '../../../engine/src/index.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { replace } from '../route.ts';
import { Field, InlineIssues, newId, NumberInput, ReadOnly, Select, Text } from '../ui/form.tsx';
import { ActionEditor, options } from '../ui/pickers.tsx';
import { SplitList } from '../ui/split.tsx';

export function RichMenus() {
  const { project, update, slug } = useEditor();
  return (
    <SplitList<RichMenu>
      section="richMenus"
      pathBase="/richMenus"
      title="リッチメニュー"
      items={project.richMenus}
      row={(r) => ({ t: r.label, s: `${r.layout.cols}列×${r.layout.rows}行・${r.cells.length}区画${project.settings.defaultRichMenuId === r.id ? '・最初に表示' : ''}` })}
      onAdd={() => {
        const id = newId('rm', project.richMenus.map((r) => r.id));
        update((d) =>
          d.richMenus.push({
            id,
            label: '新しいメニュー',
            layout: { cols: 3, rows: 2 },
            cells: Array.from({ length: 6 }, (_, i) => ({ id: `cell_${i + 1}`, label: `ボタン${i + 1}`, area: { col: i % 3, row: Math.floor(i / 3) }, action: { kind: 'none' as const } })),
          }),
        );
        return id;
      }}
      detail={(r, ri, path) => <MenuDetail r={r} ri={ri} path={path} onDeleted={() => replace({ page: 'project', slug, section: 'richMenus' })} />}
    />
  );
}

function MenuDetail({ r, ri, path, onDeleted }: { r: RichMenu; ri: number; path: string; onDeleted: () => void }) {
  const { project, update, issues } = useEditor();
  const [cellIdx, setCellIdx] = useState(0);
  const set = (fn: (x: RichMenu) => void) => update((d) => fn(d.richMenus[ri]!));
  const cell = r.cells[Math.min(cellIdx, r.cells.length - 1)];
  const ci = cell ? r.cells.indexOf(cell) : -1;
  const isDefault = project.settings.defaultRichMenuId === r.id;
  return (
    <>
      <h2>リッチメニュー</h2>
      <ReadOnly>{r.id}</ReadOnly>
      <InlineIssues issues={issues.filter((i) => i.path === path || i.path === `${path}/cells`)} />
      <div className="section">
        <div className="fields">
          <Field label="メニュー名">
            <Text value={r.label} onChange={(v) => set((x) => (x.label = v))} />
          </Field>
          <Field label="列の数">
            <NumberInput value={r.layout.cols} min={1} onChange={(v) => set((x) => (x.layout.cols = Math.max(1, v ?? 1)))} />
          </Field>
          <Field label="行の数">
            <NumberInput value={r.layout.rows} min={1} onChange={(v) => set((x) => (x.layout.rows = Math.max(1, v ?? 1)))} />
          </Field>
          <Field label="画像（任意）" hint="無ければ区画の名前で表示します" issues={issuesAt(issues, `${path}/imageAssetId`)}>
            <Select value={r.imageAssetId ?? ''} empty="なし" options={options.assets(project)} onChange={(v) => set((x) => (v ? (x.imageAssetId = v) : delete x.imageAssetId))} />
          </Field>
        </div>
        <label className="row small" style={{ marginTop: 10 }}>
          <input
            type="checkbox"
            checked={isDefault}
            onChange={(e) => update((d) => (e.target.checked ? (d.settings.defaultRichMenuId = r.id) : delete d.settings.defaultRichMenuId))}
          />
          友だち追加したときに、このメニューを表示する
        </label>
      </div>

      <div className="section">
        <h3>区画（クリックで選択）</h3>
        <div className="menu-grid" style={{ gridTemplateColumns: `repeat(${r.layout.cols}, 1fr)`, gridTemplateRows: `repeat(${r.layout.rows}, auto)` }}>
          {r.cells.map((c, i) => (
            <div
              key={c.id}
              className={i === ci ? 'on' : ''}
              style={{ gridColumn: `${c.area.col + 1} / span ${c.area.colSpan ?? 1}`, gridRow: `${c.area.row + 1} / span ${c.area.rowSpan ?? 1}` }}
              onClick={() => setCellIdx(i)}
            >
              {c.label}
            </div>
          ))}
        </div>
        <div className="row" style={{ marginTop: 8 }}>
          <button
            className="btn small"
            onClick={() =>
              set((x) => {
                x.cells.push({ id: newId('cell', x.cells.map((c) => c.id)), label: '新しい区画', area: { col: 0, row: x.layout.rows }, action: { kind: 'none' } });
                x.layout.rows += 1;
              })
            }
          >
            ＋ 区画
          </button>
        </div>
      </div>

      {cell && (
        <div className="section">
          <h3>
            区画「{cell.label}」　<span className="muted small mono">{cell.id}</span>
            <span className="spacer" />
            <button className="btn small ghost danger" onClick={() => set((x) => x.cells.splice(ci, 1))}>
              この区画を削除
            </button>
          </h3>
          <div className="fields">
            <Field label="表示名">
              <Text value={cell.label} onChange={(v) => set((x) => (x.cells[ci]!.label = v))} />
            </Field>
            <Field label="列（0から）" issues={issuesAt(issues, `${path}/cells/${ci}/area`)}>
              <NumberInput value={cell.area.col} min={0} onChange={(v) => set((x) => (x.cells[ci]!.area.col = v ?? 0))} />
            </Field>
            <Field label="行（0から）">
              <NumberInput value={cell.area.row} min={0} onChange={(v) => set((x) => (x.cells[ci]!.area.row = v ?? 0))} />
            </Field>
            <Field label="横に何マス使うか">
              <NumberInput value={cell.area.colSpan ?? 1} min={1} onChange={(v) => set((x) => ((v ?? 1) > 1 ? (x.cells[ci]!.area.colSpan = v) : delete x.cells[ci]!.area.colSpan))} />
            </Field>
            <Field label="縦に何マス使うか">
              <NumberInput value={cell.area.rowSpan ?? 1} min={1} onChange={(v) => set((x) => ((v ?? 1) > 1 ? (x.cells[ci]!.area.rowSpan = v) : delete x.cells[ci]!.area.rowSpan))} />
            </Field>
          </div>
          <ActionEditor value={cell.action} path={`${path}/cells/${ci}/action`} onChange={(a) => set((x) => (x.cells[ci]!.action = a))} />
        </div>
      )}

      <div className="row" style={{ marginTop: 12 }}>
        <span className="spacer" />
        <button
          className="btn danger"
          onClick={() => {
            if (!window.confirm('このメニューを削除しますか？')) return;
            update((d) => {
              d.richMenus.splice(ri, 1);
              if (d.settings.defaultRichMenuId === r.id) delete d.settings.defaultRichMenuId;
            });
            onDeleted();
          }}
        >
          このメニューを削除
        </button>
      </div>
    </>
  );
}
