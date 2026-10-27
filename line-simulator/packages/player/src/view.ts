// 画面を描く関数が共通で受け取るもの。描画は「この状態から HTML 文字列を作る」だけで、操作は player.ts が受ける。

import type { Asset, Id, Project, WorldState } from '../../engine/src/index.ts';
import { MINUTES_PER_DAY } from '../../engine/src/index.ts';
import type { Demo } from './controller.ts';
import type { Labels } from './describe.ts';

/** three = ①画面・②裏側・③CVまでの道筋／full = 画面＋裏側／phone = 画面のみ／slide・slide3 = 資料用16:9（2画面・3画面） */
export type Mode = 'three' | 'full' | 'phone' | 'slide' | 'slide3';
export type Tab = 'state' | 'why' | 'log' | 'matrix';

/** 画面だけが持つ状態（イベントには入らない） */
export type Ui = {
  menuOpen: boolean;
  /** 複数選択アンケートで選び中の選択肢 */
  picked: Id[];
  mode: Mode;
  tab: Tab;
  /** 「なぜ」で開いている判定の番号 */
  open: Set<number>;
  /** 「なぜ」で届かなかった・進まなかったものだけ出す */
  onlyNg: boolean;
  playing: boolean;
  speed: 1 | 2 | 5;
  /** 再生の状態の一言（回答待ち・最後まで再生した 等） */
  status: string;
};

export type View = {
  project: Project;
  world: WorldState;
  demo: Demo;
  ui: Ui;
  L: Labels;
  assetUrl: (a: Asset) => string;
};

const ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
export const esc = (v: unknown) => String(v ?? '').replace(/[&<>"']/g, (c) => ESC[c]!);

export const hhmm = (min: number) => {
  const m = ((min % MINUTES_PER_DAY) + MINUTES_PER_DAY) % MINUTES_PER_DAY;
  return `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
};
export const dayOf = (min: number) => Math.floor(min / MINUTES_PER_DAY);

/** 3画面表示の1つの枠（①②③の見出しつき） */
export const pane = (no: 1 | 2 | 3, title: string, body: string) =>
  `<section class="lsp-pane p${no}"><h3 class="lsp-pane-head"><span class="no">${'①②③'[no - 1]}</span>${esc(title)}</h3>${body}</section>`;
