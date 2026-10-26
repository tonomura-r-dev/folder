// 保存API（tools/dev-api.ts）の呼び出し。

import type { Project } from '../../engine/src/index.ts';

export type ProjectSummary = { slug: string; name: string; industry?: string; updatedAt: string };
export type HistoryEntry = { id: string; savedAt: string; label?: string; bytes: number };

export class ApiError extends Error {
  readonly status: number;
  readonly code: string;
  constructor(status: number, code: string, message: string) {
    super(message);
    this.status = status;
    this.code = code;
  }
}

async function call<T>(method: string, path: string, body?: unknown): Promise<T> {
  const res = await fetch(path, {
    method,
    headers: body === undefined ? {} : { 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const data = (await res.json().catch(() => ({}))) as { code?: string; message?: string };
  if (!res.ok) throw new ApiError(res.status, data.code ?? 'error', data.message ?? `${res.status}`);
  return data as T;
}

const p = (slug: string) => `/api/projects/${encodeURIComponent(slug)}`;

export const api = {
  list: () => call<ProjectSummary[]>('GET', '/api/projects'),
  read: (slug: string) => call<Project>('GET', p(slug)),
  create: (slug: string, name: string) => call<Project>('POST', '/api/projects', { slug, name }),
  duplicate: (from: string, slug: string, name: string) => call<Project>('POST', '/api/projects', { from, slug, name }),
  import: (slug: string, project: Project) => call<Project>('POST', '/api/import', { slug, project }),
  save: (slug: string, project: Project, expectedUpdatedAt?: string) =>
    call<{ updatedAt: string; historyId: string | null }>('PUT', p(slug), { project, expectedUpdatedAt }),
  trash: (slug: string) => call<{ trashed: string }>('DELETE', p(slug)),
  history: (slug: string) => call<HistoryEntry[]>('GET', `${p(slug)}/history`),
  restore: (slug: string, id: string) => call<{ updatedAt: string }>('POST', `${p(slug)}/history/${id}/restore`),
  label: (slug: string, id: string, label: string) => call<{ ok: true }>('POST', `${p(slug)}/history/${id}/label`, { label }),
  uploadAsset: (slug: string, name: string, dataBase64: string) => call<{ path: string }>('POST', `${p(slug)}/assets`, { name, dataBase64 }),
};

/** 案件フォルダの画像を表示する URL */
export const assetUrl = (slug: string, path: string) => `${p(slug)}/files/${path.split('/').map(encodeURIComponent).join('/')}`;

/** 画像ファイルを読み、Base64 と縦横の大きさを返す */
export async function readImage(file: File): Promise<{ base64: string; width: number; height: number }> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  const url = URL.createObjectURL(file);
  try {
    const size = await new Promise<{ width: number; height: number }>((ok, ng) => {
      const img = new Image();
      img.onload = () => ok({ width: img.naturalWidth, height: img.naturalHeight });
      img.onerror = () => ng(new Error('画像として読めません'));
      img.src = url;
    });
    return { base64: btoa(bin), ...size };
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** いまの内容を JSON ファイルとしてダウンロードさせる */
export function downloadJson(filename: string, data: unknown) {
  const blob = new Blob([`${JSON.stringify(data, null, 2)}\n`], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

export const fmtDateTime = (iso: string) => {
  if (!iso) return '—';
  const d = new Date(iso);
  const p2 = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${p2(d.getMonth() + 1)}/${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`;
};
