// 画面の場所はURLのハッシュに持つ（再読み込みしても同じ場所に戻るため）。
//   #/                                一覧
//   #/p/<slug>/<section>/<item>?node=<nodeId>&edge=<edgeId>

import { useEffect, useState } from 'react';
import type { Project } from '../../engine/src/index.ts';

export type Section =
  | 'demo'
  | 'flows'
  | 'messages'
  | 'labels'
  | 'surveys'
  | 'richMenus'
  | 'autoReplies'
  | 'webPages'
  | 'personas'
  | 'narration'
  | 'ai'
  | 'settings';

export type Route =
  | { page: 'list' }
  | { page: 'project'; slug: string; section: Section; item?: string; node?: string; edge?: string };

export function parse(hash: string): Route {
  const [path = '', query = ''] = hash.replace(/^#/, '').split('?');
  const parts = path.split('/').filter(Boolean).map(decodeURIComponent);
  if (parts[0] !== 'p' || !parts[1]) return { page: 'list' };
  const q = new URLSearchParams(query);
  const r: Route = { page: 'project', slug: parts[1], section: (parts[2] as Section) ?? 'flows' };
  if (parts[3]) r.item = parts[3];
  if (q.get('node')) r.node = q.get('node')!;
  if (q.get('edge')) r.edge = q.get('edge')!;
  return r;
}

export function href(r: Route): string {
  if (r.page === 'list') return '#/';
  const q = new URLSearchParams();
  if (r.node) q.set('node', r.node);
  if (r.edge) q.set('edge', r.edge);
  const qs = q.toString();
  return `#/p/${encodeURIComponent(r.slug)}/${r.section}${r.item ? `/${encodeURIComponent(r.item)}` : ''}${qs ? `?${qs}` : ''}`;
}

export const go = (r: Route) => {
  const h = href(r);
  if (location.hash !== h) location.hash = h;
};

/** 選択の切り替えなど、戻るボタンの履歴に積まない移動 */
export const replace = (r: Route) => {
  const h = href(r);
  if (location.hash === h) return;
  history.replaceState(null, '', h);
  window.dispatchEvent(new HashChangeEvent('hashchange'));
};

export function useRoute(): Route {
  const [route, setRoute] = useState(() => parse(location.hash));
  useEffect(() => {
    const on = () => setRoute(parse(location.hash));
    window.addEventListener('hashchange', on);
    return () => window.removeEventListener('hashchange', on);
  }, []);
  return route;
}

const COLLECTION: Record<string, Section> = {
  flows: 'flows',
  messages: 'messages',
  tags: 'labels',
  cvs: 'labels',
  routes: 'labels',
  assets: 'settings',
  surveys: 'surveys',
  richMenus: 'richMenus',
  autoReplies: 'autoReplies',
  webPages: 'webPages',
  personas: 'personas',
  narration: 'narration',
  settings: 'settings',
  brand: 'settings',
  meta: 'settings',
};

/** 検証結果の path（JSON Pointer）から、開くべき画面を決める */
export function locate(project: Project, slug: string, path: string): Route {
  const k = path.split('/').filter(Boolean);
  const section = COLLECTION[k[0] ?? ''] ?? 'settings';
  const idx = Number(k[1]);
  const r: Route = { page: 'project', slug, section };
  if (k[0] === 'flows' && Number.isInteger(idx)) {
    const flow = project.flows[idx];
    if (!flow) return r;
    r.item = flow.id;
    const j = Number(k[3]);
    if (k[2] === 'nodes' && Number.isInteger(j) && flow.nodes[j]) r.node = flow.nodes[j]!.id;
    if (k[2] === 'edges' && Number.isInteger(j) && flow.edges[j]) {
      const e = flow.edges[j]!;
      r.edge = e.id;
      // 行き先が無い線は描けないので、出発点のノードを開く
      if (flow.nodes.some((n) => n.id === e.from.nodeId)) r.node = e.from.nodeId;
    }
    return r;
  }
  const list = (project as unknown as Record<string, { id: string }[] | undefined>)[k[0] ?? ''];
  if (Array.isArray(list) && Number.isInteger(idx) && list[idx]) r.item = list[idx]!.id;
  return r;
}

/** 画面に出す「どこの問題か」 */
export function where(project: Project, path: string): string {
  const k = path.split('/').filter(Boolean);
  const idx = Number(k[1]);
  const names: Record<string, string> = {
    flows: 'シナリオ', messages: '文面', tags: 'タグ', cvs: 'CV', routes: '経路', surveys: 'アンケート', richMenus: 'リッチメニュー',
    autoReplies: '自動応答', webPages: 'Webページ', personas: 'ペルソナ', narration: '説明文', assets: '画像', settings: '案件設定', brand: '案件設定',
  };
  const base = names[k[0] ?? ''] ?? k[0] ?? '';
  if (k[0] === 'flows') {
    const f = project.flows[idx];
    if (!f) return base;
    const j = Number(k[3]);
    const n = k[2] === 'nodes' ? f.nodes[j] : undefined;
    return `${f.label}${n ? ` ／ ${n.label ?? n.id}` : k[2] === 'edges' ? ' ／ 線' : ''}`;
  }
  const item = (project as unknown as Record<string, { id: string; label?: string; name?: string; title?: string }[] | undefined>)[k[0] ?? '']?.[idx];
  return item ? `${base} ／ ${item.label ?? item.name ?? item.title ?? item.id}` : base;
}
