// タグ・CV・友だち追加経路：名前を付けるだけの部品なので、表でまとめて編集する。

import type { Project } from '../../../engine/src/index.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { InlineIssues, newId, ReadOnly, Text } from '../ui/form.tsx';

type Key = 'tags' | 'cvs' | 'routes';

const TITLE: Record<Key, { title: string; prefix: string; hint: string }> = {
  tags: { title: 'タグ', prefix: 'tag', hint: 'アンケートの回答やシナリオで付け外しし、分岐・配信の条件に使います。グループは見やすさのためだけの分類です。' },
  cvs: { title: 'CV（応募・予約など）', prefix: 'cv', hint: 'フォーム送信などで「達成した」と記録するもの。停止条件や分岐に使います。' },
  routes: { title: '友だち追加の経路', prefix: 'route', hint: 'どこから友だちになったか（離脱ポップアップ・サンクスページ等）。' },
};

export function LabelsEditor() {
  return (
    <div className="detail" style={{ width: '100%' }}>
      <h2>タグ・CV・経路</h2>
      <Table k="tags" />
      <Table k="cvs" />
      <Table k="routes" />
    </div>
  );
}

function Table({ k }: { k: Key }) {
  const { project, update, issues } = useEditor();
  const items = project[k] as { id: string; label: string; group?: string }[];
  const meta = TITLE[k];
  const set = (i: number, fn: (x: { id: string; label: string; group?: string }) => void) => update((d) => fn((d[k] as typeof items)[i]!));
  return (
    <div className="section">
      <h3>
        {meta.title}
        <span className="spacer" />
        <button
          className="btn small"
          onClick={() =>
            update((d) => (d[k] as typeof items).push({ id: newId(meta.prefix, (d[k] as typeof items).map((x) => x.id)), label: `新しい${meta.title.replace(/（.*）/, '')}` }))
          }
        >
          ＋ 追加
        </button>
      </h3>
      <div className="muted small" style={{ marginBottom: 8 }}>
        {meta.hint}
      </div>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr className="muted small" style={{ textAlign: 'left' }}>
            <th style={{ width: 170 }}>ID</th>
            <th>名前</th>
            {k === 'tags' && <th style={{ width: 160 }}>グループ</th>}
            <th style={{ width: 210 }}>使われている数</th>
            <th style={{ width: 60 }} />
          </tr>
        </thead>
        <tbody>
          {items.map((it, i) => {
            const n = countUses(project, k, it.id);
            const mine = issuesAt(issues, `/${k}/${i}`);
            return (
              <tr key={it.id} style={{ borderTop: '1px solid var(--soft)' }}>
                <td>
                  <ReadOnly>{it.id}</ReadOnly>
                </td>
                <td className="field">
                  <Text value={it.label} onChange={(v) => set(i, (x) => (x.label = v))} />
                  {mine.length > 0 && <InlineIssues issues={mine} />}
                </td>
                {k === 'tags' && (
                  <td className="field">
                    <Text value={it.group ?? ''} onChange={(v) => set(i, (x) => (v ? (x.group = v) : delete x.group))} />
                  </td>
                )}
                <td className="muted small">{n === 0 ? '未使用' : `${n} か所`}</td>
                <td>
                  <button
                    className="btn small ghost danger"
                    onClick={() => {
                      if (n > 0 && !window.confirm(`${n} か所で使われています。削除すると、そこが「見つからない」になります。削除しますか？`)) return;
                      update((d) => (d[k] as typeof items).splice(i, 1));
                    }}
                  >
                    削除
                  </button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

/** 案件JSONの中で、その ID を指している値の数（タグ・CV・経路の ID は種類ごとに接頭辞が違うので衝突しない前提） */
function countUses(p: Project, k: Key, id: string): number {
  const keys = k === 'tags' ? ['tagId', 'tags', 'add', 'remove', 'addTags', 'removeTags'] : k === 'cvs' ? ['cvId', 'cvs'] : ['routeId', 'routeIds'];
  let n = 0;
  const walk = (v: unknown, key?: string) => {
    if (Array.isArray(v)) {
      for (const x of v) {
        if (key && keys.includes(key) && x === id) n++;
        else walk(x);
      }
      return;
    }
    if (v && typeof v === 'object') {
      for (const [kk, vv] of Object.entries(v)) {
        if (keys.includes(kk) && vv === id) n++;
        else walk(vv, kk);
      }
    }
  };
  const { [k]: _self, ...rest } = p;
  walk(rest);
  return n;
}
