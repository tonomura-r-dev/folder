// 出所の印：確認済み（スクショで見えた）／AI推定（承認待ち）／承認済み。手で作ったもの（manual・無し）は印を出さない。

import type { Provenance, Project } from '../../../engine/src/index.ts';
import { approve } from '../../../engine/src/index.ts';
import { useEditor } from '../ctx.ts';

const LABEL = { observed: '確認済み', inferred: 'AI推定', approved: '承認済み' } as const;
const TITLE = {
  observed: 'スクリーンショットで確認できた内容',
  inferred: 'AIが推定・提案した内容（承認するまで営業用HTMLに書き出せません）',
  approved: 'AIの推定を人が承認した内容',
} as const;

export function ProvBadge({ prov }: { prov?: Provenance | undefined }) {
  if (!prov || prov.state === 'manual') return null;
  return (
    <span className={`prov prov-${prov.state}`} title={TITLE[prov.state]}>
      {LABEL[prov.state]}
    </span>
  );
}

/** 下書きを丸ごと差し替える（approve などが新しい案件を返すため） */
export function replaceDraft(draft: Project, next: Project) {
  for (const k of Object.keys(draft)) delete (draft as unknown as Record<string, unknown>)[k];
  Object.assign(draft, structuredClone(next));
}

/** 設定画面の上に出す「AI推定です／根拠／承認する」の帯 */
export function ProvNote({ prov, path }: { prov?: Provenance | undefined; path: string }) {
  const { update } = useEditor();
  if (!prov || (prov.state !== 'inferred' && prov.state !== 'observed')) return null;
  return (
    <div className={`prov-note prov-note-${prov.state}`}>
      <ProvBadge prov={prov} />
      <span className="small">
        {prov.state === 'inferred' ? `根拠：${prov.rationale ?? '（記載なし）'}` : `確認元：${prov.source ?? '（記載なし）'}`}
      </span>
      {prov.state === 'inferred' && (
        <button type="button" className="btn small" onClick={() => update((d) => replaceDraft(d, approve(d, path)))}>
          内容を確かめて承認する
        </button>
      )}
    </div>
  );
}
