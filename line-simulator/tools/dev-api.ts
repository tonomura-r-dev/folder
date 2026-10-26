// 編集画面の保存API。Vite の開発サーバーに差し込む（vite.config.ts）。社内PCでだけ動く。
//
//   GET    /api/projects                         一覧
//   POST   /api/projects        {slug,name,from?} 新規（from があれば複製）
//   POST   /api/import          {slug,project}    JSONから読み込み
//   GET    /api/projects/:slug                   読み込み
//   PUT    /api/projects/:slug  {project,expectedUpdatedAt}  保存
//   DELETE /api/projects/:slug                   ゴミ箱へ
//   GET    /api/projects/:slug/history           履歴一覧
//   GET    /api/projects/:slug/history/:id       履歴の中身
//   POST   /api/projects/:slug/history/:id/restore
//   POST   /api/projects/:slug/history/:id/label {label}
//   GET    /api/projects/:slug/files/assets/<name>  画像（assets/ の中だけ）
//   POST   /api/projects/:slug/assets {name,dataBase64}  画像の取り込み
//   POST   /api/projects/:slug/export {confirmFictional}  営業用HTMLを Downloads に書き出す
//   GET    /api/projects/:slug/patches              AIの提案（シナリオ差分）の一覧
//   GET    /api/projects/:slug/patches/:name        提案の中身
//   POST   /api/projects/:slug/patches {name,patch} 画面から読み込んだ提案を保存

import { existsSync, mkdirSync } from 'node:fs';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { homedir } from 'node:os';
import { join } from 'node:path';
import type { Project } from '../packages/engine/src/index.ts';
import { buildExport, uniqueName } from './export-lib.ts';
import { ProjectStore, StoreError, writeAtomic } from './project-store.ts';

/** 営業用HTMLの保存先（社内の決まりで Downloads 直下） */
const exportDir = () => join(homedir(), 'Downloads');

const MAX_BODY = 20 * 1024 * 1024;

function body(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let size = 0;
    req.on('data', (c: Buffer) => {
      size += c.length;
      if (size > MAX_BODY) reject(new StoreError(413, 'body.tooLarge', '送信データが大きすぎます'));
      else chunks.push(c);
    });
    req.on('end', () => {
      if (!chunks.length) return resolve({});
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch {
        reject(new StoreError(400, 'body.invalid', 'JSONとして読めません'));
      }
    });
    req.on('error', reject);
  });
}

function send(res: ServerResponse, status: number, data: unknown) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.end(JSON.stringify(data));
}

export function createApiHandler(store: ProjectStore) {
  return async (req: IncomingMessage, res: ServerResponse, next: () => void) => {
    const url = new URL(req.url ?? '/', 'http://localhost');
    if (!url.pathname.startsWith('/api/')) return next();
    const parts = url.pathname.split('/').filter(Boolean).map(decodeURIComponent); // ['api','projects',slug,...]
    const method = req.method ?? 'GET';
    try {
      if (parts[1] === 'projects' && parts.length === 2) {
        if (method === 'GET') return send(res, 200, store.list());
        if (method === 'POST') {
          const b = (await body(req)) as { slug: string; name: string; from?: string };
          const p = b.from ? store.duplicate(b.from, b.slug, b.name) : store.create(b.slug, b.name);
          return send(res, 201, p);
        }
      }
      if (parts[1] === 'import' && parts.length === 2 && method === 'POST') {
        const b = (await body(req)) as { slug: string; project: Project };
        return send(res, 201, store.import(b.slug, b.project));
      }
      if (parts[1] === 'projects' && parts.length === 3) {
        const slug = parts[2]!;
        if (method === 'GET') return send(res, 200, store.read(slug));
        if (method === 'PUT') {
          const b = (await body(req)) as { project: Project; expectedUpdatedAt?: string };
          return send(res, 200, store.save(slug, b.project, b.expectedUpdatedAt));
        }
        if (method === 'DELETE') return send(res, 200, { trashed: store.trash(slug) });
      }
      if (parts[1] === 'projects' && parts[3] === 'files' && method === 'GET') {
        const { data, type } = store.readAsset(parts[2]!, parts.slice(4).join('/'));
        res.statusCode = 200;
        res.setHeader('Content-Type', type);
        res.setHeader('Cache-Control', 'no-store');
        return void res.end(data);
      }
      if (parts[1] === 'projects' && parts[3] === 'assets' && parts.length === 4 && method === 'POST') {
        const b = (await body(req)) as { name: string; dataBase64: string };
        return send(res, 201, { path: store.saveAsset(parts[2]!, b.name ?? '', Buffer.from(b.dataBase64 ?? '', 'base64')) });
      }
      if (parts[1] === 'projects' && parts[3] === 'export' && parts.length === 4 && method === 'POST') {
        // 保存済みの案件を書き出す（画面の未保存の変更は含めない）
        const slug = parts[2]!;
        const b = (await body(req)) as { confirmFictional?: boolean };
        const result = await buildExport(store.read(slug), {
          readAsset: (p) => {
            try {
              return store.readAsset(slug, p).data;
            } catch {
              return null;
            }
          },
          confirmFictional: b.confirmFictional === true,
        });
        if (!result.ok || !result.html) return send(res, 422, { ok: false, bytes: result.bytes, issues: result.issues });
        const dir = exportDir();
        mkdirSync(dir, { recursive: true });
        const path = join(dir, uniqueName(result.fileName, (n) => existsSync(join(dir, n))));
        writeAtomic(path, result.html);
        return send(res, 201, { ok: true, path, bytes: result.bytes, issues: result.issues });
      }
      if (parts[1] === 'projects' && parts[3] === 'patches') {
        const slug = parts[2]!;
        if (parts.length === 4 && method === 'GET') return send(res, 200, store.listPatches(slug));
        if (parts.length === 4 && method === 'POST') {
          const b = (await body(req)) as { name?: string; patch?: unknown };
          return send(res, 201, { name: store.savePatch(slug, b.name ?? 'patch.json', b.patch) });
        }
        if (parts.length === 5 && method === 'GET') return send(res, 200, store.readPatch(slug, parts[4]!));
      }
      if (parts[1] === 'projects' && parts[3] === 'history') {
        const slug = parts[2]!;
        const id = parts[4];
        if (parts.length === 4 && method === 'GET') return send(res, 200, store.history(slug));
        if (parts.length === 5 && method === 'GET') return send(res, 200, store.readHistory(slug, id!));
        if (parts.length === 6 && parts[5] === 'restore' && method === 'POST') return send(res, 200, store.restore(slug, id!));
        if (parts.length === 6 && parts[5] === 'label' && method === 'POST') {
          const b = (await body(req)) as { label: string };
          store.label(slug, id!, b.label ?? '');
          return send(res, 200, { ok: true });
        }
      }
      send(res, 404, { code: 'api.notFound', message: `${method} ${url.pathname}` });
    } catch (e) {
      if (e instanceof StoreError) return send(res, e.status, { code: e.code, message: e.message });
      send(res, 500, { code: 'server.error', message: e instanceof Error ? e.message : String(e) });
    }
  };
}
