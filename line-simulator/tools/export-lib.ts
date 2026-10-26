// 営業用HTMLの書き出し（中身）。CLI（tools/export.ts）と編集ツールの保存API（dev-api.ts）の両方から使う。
//
//   1. 検証：重大エラーがあれば書き出さない（警告は一覧にして書き出す）
//   2. 社内情報を取り除く：_internal（社内メモ等）・prov（AIの出所メモ）
//   3. 画像を data: URL にして埋め込む（大きい画像は縮小）
//   4. 1つのHTMLに組み立てる：CSS・案件データ・束ねたJS（外部ファイル・外部URLを読まない）
//   5. 大きさを確かめる：5MB超で警告、15MB超で書き出さない
//   6. 出来上がりを検査：社内メモ・APIキーらしき文字列・外部の読み込み・編集ツールのコードが無いこと

import type { Asset, Issue, Project } from '../packages/engine/src/index.ts';
import { personaCoverage, validate } from '../packages/engine/src/index.ts';
import { APP_ID, DATA_ID, type ExportData } from '../packages/player/src/standalone.ts';
import { bundlePlayer, playerCss } from './bundle.ts';

export const WARN_BYTES = 5 * 1024 * 1024;
export const MAX_BYTES = 15 * 1024 * 1024;
/** 画像はこの幅より大きければ縮める（スマホ枠は360px。2倍の密度でも足りる） */
export const IMAGE_MAX_WIDTH = 800;

export type ExportOptions = {
  /** 案件フォルダからの相対パスで画像を読む。無ければ null */
  readAsset: (path: string) => Buffer | null;
  /** 架空・匿名化したデータだけであることを、書き出す人が確認したか */
  confirmFictional: boolean;
  now?: Date;
};

export type ExportResult = { ok: boolean; html?: string; bytes: number; fileName: string; issues: Issue[] };

const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' };

/** _internal と prov を、どの深さにあっても取り除いた複製 */
export function sanitize(project: Project): Project {
  const walk = (v: unknown): unknown => {
    if (Array.isArray(v)) return v.map(walk);
    if (!v || typeof v !== 'object') return v;
    const out: Record<string, unknown> = {};
    for (const [k, x] of Object.entries(v)) if (k !== '_internal' && k !== 'prov') out[k] = walk(x);
    return out;
  };
  return walk(project) as Project;
}

/** 社内メモなど、HTMLに出てはいけない文字列（_internal の中の文字すべて） */
export function internalStrings(project: Project): string[] {
  const out: string[] = [];
  const collect = (v: unknown) => {
    if (typeof v === 'string') {
      if (v.trim().length >= 4) out.push(v.trim());
    } else if (Array.isArray(v)) v.forEach(collect);
    else if (v && typeof v === 'object') Object.values(v).forEach(collect);
  };
  const walk = (v: unknown) => {
    if (Array.isArray(v)) return v.forEach(walk);
    if (!v || typeof v !== 'object') return;
    for (const [k, x] of Object.entries(v)) (k === '_internal' ? collect : walk)(x);
  };
  walk(project);
  return out;
}

/** <script> の中に安全に入れられる JSON（</script> や <!-- で抜け出せないように） */
export function safeJson(data: unknown): string {
  return JSON.stringify(data).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}

const ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
const esc = (s: string) => s.replace(/[&<>"']/g, (c) => ESC[c]!);

/** Windows で使えない文字を外したファイル名 */
export function exportFileName(project: Project, now: Date): string {
  const p = (n: number) => String(n).padStart(2, '0');
  const base = (project.meta.name || project.meta.slug).replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().slice(0, 80) || 'demo';
  return `${base}_デモ_${now.getFullYear()}${p(now.getMonth() + 1)}${p(now.getDate())}.html`;
}

/** 同じ名前のファイルがあれば _2, _3 … を付けて、上書きしない */
export function uniqueName(fileName: string, exists: (name: string) => boolean): string {
  if (!exists(fileName)) return fileName;
  const dot = fileName.lastIndexOf('.');
  for (let i = 2; ; i++) {
    const next = `${fileName.slice(0, dot)}_${i}${fileName.slice(dot)}`;
    if (!exists(next)) return next;
  }
}

async function imageDataUrl(asset: Asset, buf: Buffer): Promise<{ url: string; resized: boolean }> {
  const ext = (asset.path.split('.').pop() ?? '').toLowerCase();
  const mime = MIME[ext] ?? 'application/octet-stream';
  if (ext !== 'gif') {
    try {
      const { default: sharp } = await import('sharp');
      const meta = await sharp(buf).metadata();
      if ((meta.width ?? 0) > IMAGE_MAX_WIDTH || buf.length > 300 * 1024) {
        const out = await sharp(buf).resize({ width: Math.min(meta.width ?? IMAGE_MAX_WIDTH, IMAGE_MAX_WIDTH), withoutEnlargement: true }).webp({ quality: 80 }).toBuffer();
        if (out.length < buf.length) return { url: `data:image/webp;base64,${out.toString('base64')}`, resized: true };
      }
    } catch {
      // 縮小できなければ、そのまま埋め込む（大きさは後で確かめる）
    }
  }
  return { url: `data:${mime};base64,${buf.toString('base64')}`, resized: false };
}

const PAGE_CSS = `
:root{color-scheme:light}
*{box-sizing:border-box}
html,body{margin:0}
body{background:#e9ecf3;font-family:Meiryo,"Yu Gothic UI","Hiragino Sans",system-ui,sans-serif}
.ls-head{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap;padding:12px 16px;background:#1f285a;color:#fff}
.ls-head b{font-size:16px}
.ls-head span{font-size:12px;opacity:.85}
.ls-foot{padding:14px 16px 24px;text-align:center;font-size:11.5px;color:#6b7280}
`;

/** 外から何も読まない・どこにも送らない（ブラウザに守らせる） */
const CSP = "default-src 'none'; img-src data:; style-src 'unsafe-inline'; script-src 'unsafe-inline'; base-uri 'none'; form-action 'none'";

function page(project: Project, data: ExportData, js: string, css: string): string {
  const name = esc(project.brand.accountName || project.meta.name);
  // JS の中に </script が現れたら、HTMLとして閉じてしまわないようにずらす
  const script = js.replace(/<\/(script)/gi, '<\\/$1').replace(/<!--/g, '<\\!--');
  return `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="${CSP}">
<meta name="robots" content="noindex, nofollow">
<meta name="referrer" content="no-referrer">
<title>${name} LINE施策デモ</title>
<style>${PAGE_CSS}${css}</style>
</head>
<body>
<header class="ls-head"><b>${name}</b><span>LINE公式アカウント施策のデモ（架空のデータによるデモです）</span></header>
<main id="${APP_ID}"></main>
<noscript><p style="padding:20px">このデモを見るには、JavaScript を有効にしてください。</p></noscript>
<footer class="ls-foot">このファイルはインターネットに接続しなくても動きます。表示されている会社名・人物・求人などはすべて架空です。</footer>
<script type="application/json" id="${DATA_ID}">${safeJson(data)}</script>
<script>${script}</script>
</body>
</html>
`;
}

/** 出来上がったHTMLの検査。問題があれば重大エラーとして返す */
export function scanExport(html: string, original: Project): Issue[] {
  const issues: Issue[] = [];
  const err = (code: string, params: Record<string, unknown>) => issues.push({ level: 'error', code, path: '/', params });
  const json = html.match(new RegExp(`<script type="application/json" id="${DATA_ID}">([\\s\\S]*?)</script>`))?.[1] ?? '';
  const outsideData = html.replace(json, '');

  // 社内メモ・社内専用の項目
  if (/"_internal"\s*:|"prov"\s*:/.test(json)) err('export.leak', { what: '社内専用の項目（_internal・prov）' });
  const decoded = (() => {
    try {
      return JSON.stringify(JSON.parse(json));
    } catch {
      return json;
    }
  })();
  for (const s of internalStrings(original)) if (decoded.includes(JSON.stringify(s).slice(1, -1)) || html.includes(s)) err('export.leak', { what: '社内メモの文字', sample: s.slice(0, 20) });

  // APIキー・秘密鍵らしき文字列
  const secret = /\bsk-[A-Za-z0-9_-]{20,}|AIza[0-9A-Za-z_-]{30,}|-----BEGIN [A-Z ]*PRIVATE KEY-----|"private_key"\s*:|\bxox[abprs]-[A-Za-z0-9-]{10,}|\bghp_[A-Za-z0-9]{30,}/;
  const hit = html.match(secret);
  if (hit) err('export.secret', { sample: hit[0].slice(0, 12) });

  // 外部からの読み込み（案件データの中の「表示用のURL」は文字として出すだけなので対象外）
  const external = outsideData.match(/(?:src|href)\s*=\s*["']?\s*(?:https?:)?\/\/|url\(\s*["']?\s*(?:https?:)?\/\/|@import|<link\b|<iframe\b|<object\b|<embed\b|\bfetch\(|XMLHttpRequest|new WebSocket|navigator\.sendBeacon/i);
  if (external) err('export.external', { sample: external[0] });
  for (const url of Object.values((JSON.parse(json || '{}') as Partial<ExportData>).assets ?? {})) if (!url.startsWith('data:')) err('export.external', { sample: url.slice(0, 40) });

  // 編集ツールのコード
  const editor = outsideData.match(/react-dom|ReactFlow|@xyflow|\/api\/projects|createRoot/);
  if (editor) err('export.editorCode', { sample: editor[0] });
  return issues;
}

export async function buildExport(project: Project, opts: ExportOptions): Promise<ExportResult> {
  const now = opts.now ?? new Date();
  const fileName = exportFileName(project, now);
  const issues: Issue[] = validate(project, { assetExists: (p) => opts.readAsset(p) !== null });
  const fail = (): ExportResult => ({ ok: false, bytes: 0, fileName, issues });
  if (!opts.confirmFictional) {
    issues.push({ level: 'error', code: 'export.notConfirmed', path: '/' });
    return fail();
  }
  if (issues.some((i) => i.level === 'error')) return fail();
  issues.push(...personaCoverage(project));

  const clean = sanitize(project);
  const assets: Record<string, string> = {};
  for (const a of clean.assets) {
    const buf = opts.readAsset(a.path);
    if (!buf) continue; // validate で重大エラー済み
    const { url, resized } = await imageDataUrl(a, buf);
    assets[a.id] = url;
    if (resized) issues.push({ level: 'warning', code: 'export.imageResized', path: `/assets/${clean.assets.indexOf(a)}`, params: { id: a.id, width: IMAGE_MAX_WIDTH } });
  }

  const data: ExportData = { project: clean, assets, exportedAt: now.toISOString() };
  const html = page(clean, data, await bundlePlayer(), playerCss());
  const bytes = Buffer.byteLength(html, 'utf8');
  if (bytes > MAX_BYTES) {
    issues.push({ level: 'error', code: 'export.tooLarge', path: '/', params: { bytes, limit: MAX_BYTES } });
    return { ok: false, bytes, fileName, issues };
  }
  if (bytes > WARN_BYTES) issues.push({ level: 'warning', code: 'export.large', path: '/', params: { bytes, limit: WARN_BYTES } });

  const leaks = scanExport(html, project);
  if (leaks.length) return { ok: false, bytes, fileName, issues: [...issues, ...leaks] };
  return { ok: true, html, bytes, fileName, issues };
}
