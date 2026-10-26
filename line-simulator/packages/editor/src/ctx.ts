// 編集中の案件を、各編集画面から読む・書くための共有口。

import { createContext, useContext } from 'react';
import type { Issue, Project } from '../../engine/src/index.ts';
import type { Route } from './route.ts';

export type EditorCtx = {
  slug: string;
  project: Project;
  /** 下書きをその場で書き換える関数を渡す。保存は別（保存ボタン／Ctrl+S） */
  update: (fn: (draft: Project) => void) => void;
  issues: Issue[];
  route: Extract<Route, { page: 'project' }>;
  toast: (msg: string, kind?: 'ok' | 'err') => void;
};

export const Ctx = createContext<EditorCtx | null>(null);

export function useEditor(): EditorCtx {
  const c = useContext(Ctx);
  if (!c) throw new Error('useEditor outside editor');
  return c;
}

/** path そのもの、またはその配下の検証結果 */
export const issuesAt = (issues: Issue[], path: string) => issues.filter((i) => i.path === path || i.path.startsWith(`${path}/`));
