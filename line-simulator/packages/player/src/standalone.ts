// 営業用HTMLの入口。HTMLに埋め込まれた案件データを読み、デモ画面を出す。
// 編集ツールと同じ player・エンジンをそのまま束ねる（tools/bundle.ts）。編集機能・保存APIは含まない。

import type { Project, Session, WorldState } from '../../engine/src/index.ts';
import { reduce } from '../../engine/src/index.ts';
import { worldOf } from './controller.ts';
import { createPlayer } from './player.ts';

/** HTMLの <script type="application/json" id="linesim-data"> に入れる中身 */
export type ExportData = {
  /** 社内情報を取り除いた案件 */
  project: Project;
  /** 画像ID → data: URL（外部ファイルを読まない） */
  assets: Record<string, string>;
  exportedAt: string;
};

export const DATA_ID = 'linesim-data';
export const APP_ID = 'linesim-app';

/** エンジンの計算をそのまま呼ぶ（書き出したHTMLと編集ツールで結果が同じかを確かめるため） */
export const simulate = (project: Project, session: Session, until: number): WorldState => reduce(project, session, { until });

export function mount(doc: Document): void {
  const dataEl = doc.getElementById(DATA_ID);
  const root = doc.getElementById(APP_ID);
  if (!dataEl || !root) return;
  const data = JSON.parse(dataEl.textContent ?? '{}') as ExportData;
  const handle = createPlayer(root, { project: data.project, assetUrl: (a) => data.assets[a.id] ?? '' });
  // 動作確認用の読み取り口（書き換えはできない）
  (doc.defaultView as unknown as { lineSim: unknown }).lineSim = Object.freeze({
    demo: () => structuredClone(handle.getDemo()),
    world: () => structuredClone(worldOf(data.project, handle.getDemo())),
  });
}

if (typeof document !== 'undefined') mount(document);
