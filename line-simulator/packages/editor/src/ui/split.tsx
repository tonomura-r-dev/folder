// 「左に一覧・右に詳細」の部品編集画面の共通の型。

import type { ReactNode } from 'react';
import { issuesAt, useEditor } from '../ctx.ts';
import { replace, type Section } from '../route.ts';

type Props<T extends { id: string }> = {
  section: Section;
  /** 検証結果の path の頭（例：'/messages'） */
  pathBase: string;
  title: string;
  items: T[];
  row: (item: T) => { t: string; s?: string };
  onAdd?: () => string;
  addLabel?: string;
  detail: (item: T, index: number, path: string) => ReactNode;
  empty?: ReactNode;
};

export function SplitList<T extends { id: string }>({ section, pathBase, title, items, row, onAdd, addLabel = '＋ 追加', detail, empty }: Props<T>) {
  const { route, issues, slug } = useEditor();
  const index = Math.max(0, items.findIndex((x) => x.id === route.item));
  const current = items[index];
  const open = (id: string) => replace({ page: 'project', slug, section, item: id });
  return (
    <div className="split">
      <div className="list">
        <div className="list-head">
          <h3>{title}</h3>
          <span className="spacer" />
          {onAdd && (
            <button className="btn small" onClick={() => open(onAdd())}>
              {addLabel}
            </button>
          )}
        </div>
        {items.map((it, i) => {
          const r = row(it);
          const mine = issuesAt(issues, `${pathBase}/${i}`);
          const err = mine.some((x) => x.level === 'error');
          return (
            <div key={it.id} className={`list-item${it === current ? ' on' : ''}`} onClick={() => open(it.id)}>
              <div className="t">
                {err && <span style={{ color: 'var(--err)' }}>✖ </span>}
                {r.t}
              </div>
              {r.s && <div className="s">{r.s}</div>}
            </div>
          );
        })}
        {items.length === 0 && <div className="list-item muted">{empty ?? 'まだありません'}</div>}
      </div>
      <div className="detail">{current ? detail(current, index, `${pathBase}/${index}`) : <div className="empty">{empty ?? '左の「追加」から作ってください'}</div>}</div>
    </div>
  );
}

/** どこで使われているか（削除前の確認と、詳細画面の案内に使う） */
export function UsedBy({ list }: { list: string[] }) {
  if (!list.length) return <div className="muted small">どこからも使われていません。</div>;
  return <div className="muted small">使われている場所：{list.join('、')}</div>;
}
