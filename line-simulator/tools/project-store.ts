// 案件フォルダの保存・読み込み・複製・ゴミ箱・履歴。
// 編集画面の保存API（dev-api.ts）とテストから使う。エンジンはディスクに触らないので、ここが唯一の書き手。
//
// 守っていること
//   - 書き込みは「.tmp に書いてから置き換え」。途中で落ちても前の project.json が残る
//   - 上書きの前に、いまの版を history/ に退避する（保存のたびに1版）
//   - 読み込んだ時点の版（updatedAt）と、ディスクの版が違えば保存を断る（別の場所での変更を潰さない）
//   - 削除はゴミ箱（.trash/）への移動だけ。完全削除はしない
//   - フォルダ名（slug）は英数字・-・_ だけ。パスの外に出られない

import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { basename, extname, join, resolve, sep } from 'node:path';
import type { Project } from '../packages/engine/src/index.ts';

export const HISTORY_KEEP = 30;
export const MAX_ASSET_BYTES = 10 * 1024 * 1024;
/** 扱う画像の種類。SVG は中にスクリプトを書けるので受け付けない */
const IMAGE_TYPES: Record<string, string> = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp' };
const SLUG = /^[a-z0-9][a-z0-9_-]{0,63}$/;

export class StoreError extends Error {
  readonly status: number;
  readonly code: string;
  constructor(status: number, code: string, message: string) {
    super(message);
    this.status = status;
    this.code = code;
  }
}

export type ProjectSummary = { slug: string; name: string; industry?: string; updatedAt: string };
/** savedAt は「その版が保存された時刻」（版の meta.updatedAt）。読めない版はファイルの時刻で代用する */
export type HistoryEntry = { id: string; savedAt: string; label?: string; bytes: number };

/** history の ID。時刻順に並ぶ文字列（同じミリ秒でも重ならないよう連番を付ける） */
let lastStamp = '';
let stampSeq = 0;
function stamp(now = new Date()): string {
  const p = (n: number, w = 2) => String(n).padStart(w, '0');
  const s = `${now.getFullYear()}${p(now.getMonth() + 1)}${p(now.getDate())}-${p(now.getHours())}${p(now.getMinutes())}${p(now.getSeconds())}-${p(now.getMilliseconds(), 3)}`;
  stampSeq = s === lastStamp ? stampSeq + 1 : 0;
  lastStamp = s;
  return stampSeq ? `${s}-${p(stampSeq, 2)}` : s;
}

/** 書き込み中に落ちても壊れない書き方（同じフォルダの .tmp に書いて置き換える） */
export function writeAtomic(file: string, text: string) {
  const tmp = `${file}.tmp`;
  writeFileSync(tmp, text, 'utf8');
  renameSync(tmp, file);
}

const toJson = (p: Project) => `${JSON.stringify(p, null, 2)}\n`;

export function blankProject(slug: string, name: string): Project {
  const now = new Date().toISOString();
  return {
    schemaVersion: 1,
    meta: { id: slug, name, slug, createdAt: now, updatedAt: now },
    brand: { companyName: name, accountName: name, colors: { primary: '#1F285A', accent: '#06C755' } },
    settings: { dedupeWindowMinutes: 1440 },
    assets: [],
    tags: [],
    cvs: [],
    routes: [{ id: 'route_default', label: '友だち追加' }],
    messages: [{ id: 'msg_welcome', bubbles: [{ kind: 'text', text: '友だち追加ありがとうございます。' }] }],
    richMenus: [],
    autoReplies: [],
    surveys: [],
    webPages: [],
    flows: [
      {
        id: 'flow_welcome',
        label: '友だち追加 → 挨拶',
        trigger: { kind: 'friendAdd' },
        priority: 10,
        concurrency: 'exclusive',
        reentry: 'never',
        nodes: [
          { id: 'n_entry', type: 'entry', position: { x: 0, y: 0 } },
          { id: 'n_welcome', type: 'message', messageId: 'msg_welcome', position: { x: 0, y: 140 } },
        ],
        edges: [{ id: 'e_1', from: { nodeId: 'n_entry', port: 'out' }, to: 'n_welcome' }],
      },
    ],
    personas: [
      { id: 'p_1', name: 'ユーザーA', attrs: {}, script: [{ atMinutes: 600, event: { type: 'addFriend', routeId: 'route_default' } }] },
    ],
    narration: [],
  };
}

export class ProjectStore {
  readonly root: string;
  constructor(root: string) {
    this.root = root;
  }

  private get projectsDir() {
    return join(this.root, 'projects');
  }
  private get trashDir() {
    return join(this.root, '.trash');
  }
  private dir(slug: string) {
    if (!SLUG.test(slug)) throw new StoreError(400, 'slug.invalid', `フォルダ名は英小文字・数字・-・_ で付けてください: ${slug}`);
    return join(this.projectsDir, slug);
  }
  private file(slug: string) {
    return join(this.dir(slug), 'project.json');
  }
  private historyDir(slug: string) {
    return join(this.dir(slug), 'history');
  }
  private labelsFile(slug: string) {
    return join(this.historyDir(slug), '_labels.json');
  }

  exists(slug: string) {
    return existsSync(this.file(slug));
  }

  list(): ProjectSummary[] {
    if (!existsSync(this.projectsDir)) return [];
    const out: ProjectSummary[] = [];
    for (const slug of readdirSync(this.projectsDir)) {
      if (!SLUG.test(slug) || !existsSync(join(this.projectsDir, slug, 'project.json'))) continue;
      try {
        const p = this.read(slug);
        out.push({ slug, name: p.meta.name, updatedAt: p.meta.updatedAt, ...(p.meta.industry ? { industry: p.meta.industry } : {}) });
      } catch {
        // 壊れた JSON の案件は一覧に「読めない」として出す（消さない）
        out.push({ slug, name: `${slug}（読み込めません）`, updatedAt: '' });
      }
    }
    return out.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : a.updatedAt > b.updatedAt ? -1 : 0));
  }

  read(slug: string): Project {
    const file = this.file(slug);
    if (!existsSync(file)) throw new StoreError(404, 'project.notFound', `案件が見つかりません: ${slug}`);
    try {
      return JSON.parse(readFileSync(file, 'utf8')) as Project;
    } catch {
      throw new StoreError(422, 'project.broken', `project.json が読めません（履歴から戻せます）: ${slug}`);
    }
  }

  /**
   * 保存。expectedUpdatedAt は「画面がこの案件を読み込んだときの updatedAt」。
   * ディスクの版と違えば、別の場所で更新されたとみなして断る（409）。
   */
  save(slug: string, project: Project, expectedUpdatedAt?: string): { updatedAt: string; historyId: string | null } {
    const file = this.file(slug);
    let historyId: string | null = null;
    if (existsSync(file)) {
      const current = readFileSync(file, 'utf8');
      if (expectedUpdatedAt !== undefined) {
        let onDisk = '';
        try {
          onDisk = (JSON.parse(current) as Project).meta.updatedAt;
        } catch {
          onDisk = '';
        }
        if (onDisk !== expectedUpdatedAt)
          throw new StoreError(409, 'project.conflict', '読み込んだ後に、別の場所でこの案件が更新されています。');
      }
      historyId = this.snapshot(slug, current);
    } else {
      mkdirSync(this.dir(slug), { recursive: true });
    }
    const updatedAt = new Date().toISOString();
    const next: Project = { ...project, meta: { ...project.meta, slug, updatedAt } };
    writeAtomic(file, toJson(next));
    this.prune(slug);
    return { updatedAt, historyId };
  }

  create(slug: string, name: string): Project {
    if (this.exists(slug)) throw new StoreError(409, 'project.exists', `同じフォルダ名の案件があります: ${slug}`);
    const p = blankProject(slug, name);
    this.save(slug, p);
    return this.read(slug);
  }

  /** 複製。中身はすべてコピーし、ID・名前・作成日だけ振り直す。履歴は引き継がない */
  duplicate(from: string, to: string, name: string): Project {
    const src = this.read(from);
    if (this.exists(to)) throw new StoreError(409, 'project.exists', `同じフォルダ名の案件があります: ${to}`);
    mkdirSync(this.dir(to), { recursive: true });
    for (const entry of readdirSync(this.dir(from))) {
      if (entry === 'history' || entry === 'project.json' || entry.endsWith('.tmp')) continue;
      cpSync(join(this.dir(from), entry), join(this.dir(to), entry), { recursive: true });
    }
    const now = new Date().toISOString();
    const copy: Project = { ...structuredClone(src), meta: { ...src.meta, id: to, slug: to, name, createdAt: now, updatedAt: now } };
    this.save(to, copy);
    return this.read(to);
  }

  /** JSON から読み込む。既存のフォルダ名とは重ねない */
  import(slug: string, project: Project): Project {
    if (!project || project.schemaVersion !== 1 || !project.meta || !Array.isArray(project.flows))
      throw new StoreError(422, 'import.invalid', 'この案件JSONの形式には対応していません。');
    if (this.exists(slug)) throw new StoreError(409, 'project.exists', `同じフォルダ名の案件があります: ${slug}`);
    this.save(slug, { ...project, meta: { ...project.meta, id: slug } });
    return this.read(slug);
  }

  /** ゴミ箱へ移す（完全には消さない） */
  trash(slug: string): string {
    if (!this.exists(slug)) throw new StoreError(404, 'project.notFound', `案件が見つかりません: ${slug}`);
    mkdirSync(this.trashDir, { recursive: true });
    const dest = join(this.trashDir, `${slug}__${stamp()}`);
    renameSync(this.dir(slug), dest);
    return dest;
  }

  // ---------- 画像 ----------

  /** 案件フォルダの assets/ の中のファイルだけを読む（パスの外・他の場所は読まない） */
  readAsset(slug: string, rel: string): { data: Buffer; type: string } {
    const type = IMAGE_TYPES[extname(rel).toLowerCase()];
    const base = resolve(this.dir(slug), 'assets');
    const file = resolve(this.dir(slug), rel);
    if (!type || !file.startsWith(base + sep) || !existsSync(file)) throw new StoreError(404, 'asset.notFound', `画像が見つかりません: ${rel}`);
    return { data: readFileSync(file), type };
  }

  /** 画像を assets/ に保存し、案件から参照するパスを返す。同じ名前があれば番号を付ける */
  saveAsset(slug: string, name: string, data: Buffer): string {
    const ext = extname(name).toLowerCase();
    if (!IMAGE_TYPES[ext]) throw new StoreError(415, 'asset.type', '画像は PNG・JPEG・GIF・WebP にしてください');
    if (data.length > MAX_ASSET_BYTES) throw new StoreError(413, 'asset.tooLarge', `画像は ${MAX_ASSET_BYTES / 1024 / 1024}MB までにしてください`);
    if (!this.exists(slug)) throw new StoreError(404, 'project.notFound', `案件が見つかりません: ${slug}`);
    const dir = join(this.dir(slug), 'assets');
    mkdirSync(dir, { recursive: true });
    const stem = basename(name, extname(name)).replace(/[^\w.-]+/g, '_').replace(/^[._]+/, '').slice(0, 60) || 'image';
    let file = `${stem}${ext}`;
    for (let i = 2; existsSync(join(dir, file)); i++) file = `${stem}-${i}${ext}`;
    writeFileSync(join(dir, file), data);
    return `assets/${file}`;
  }

  // ---------- AIの提案（シナリオ差分） ----------
  // projects/<案件>/patches/*.json に置く。Claude Code が書き、編集ツールの「AI提案」で読む。

  private patchesDir(slug: string) {
    return join(this.dir(slug), 'patches');
  }

  private patchFile(slug: string, name: string) {
    if (!/^[\w.-]{1,120}\.json$/.test(name)) throw new StoreError(400, 'patch.invalidName', `提案ファイルの名前が不正です: ${name}`);
    return join(this.patchesDir(slug), name);
  }

  listPatches(slug: string): { name: string; savedAt: string; bytes: number }[] {
    const dir = this.patchesDir(slug);
    if (!existsSync(dir)) return [];
    return readdirSync(dir)
      .filter((f) => /^[\w.-]{1,120}\.json$/.test(f))
      .map((name) => {
        const st = statSync(join(dir, name));
        return { name, savedAt: st.mtime.toISOString(), bytes: st.size };
      })
      .sort((a, b) => (a.name < b.name ? 1 : -1));
  }

  readPatch(slug: string, name: string): unknown {
    const f = this.patchFile(slug, name);
    if (!existsSync(f)) throw new StoreError(404, 'patch.notFound', `提案ファイルが見つかりません: ${name}`);
    try {
      return JSON.parse(readFileSync(f, 'utf8'));
    } catch {
      throw new StoreError(422, 'patch.broken', `提案ファイルがJSONとして読めません: ${name}`);
    }
  }

  /** 画面から読み込んだ提案ファイルを保存する。名前は英数字に寄せ、同名は上書きしない */
  savePatch(slug: string, name: string, data: unknown): string {
    if (!this.exists(slug)) throw new StoreError(404, 'project.notFound', `案件が見つかりません: ${slug}`);
    const stem = basename(name, extname(name)).replace(/[^\w.-]+/g, '_').replace(/^[._]+/, '').slice(0, 80) || 'patch';
    mkdirSync(this.patchesDir(slug), { recursive: true });
    let file = `${stem}.json`;
    for (let i = 2; existsSync(join(this.patchesDir(slug), file)); i++) file = `${stem}-${i}.json`;
    writeAtomic(this.patchFile(slug, file), `${JSON.stringify(data, null, 2)}\n`);
    return file;
  }

  // ---------- 履歴 ----------

  private snapshot(slug: string, text: string): string {
    mkdirSync(this.historyDir(slug), { recursive: true });
    const id = stamp();
    writeAtomic(join(this.historyDir(slug), `${id}.json`), text);
    return id;
  }

  private labels(slug: string): Record<string, string> {
    const f = this.labelsFile(slug);
    if (!existsSync(f)) return {};
    try {
      return JSON.parse(readFileSync(f, 'utf8')) as Record<string, string>;
    } catch {
      return {};
    }
  }

  /** 履歴のID（新しい順）。中身は読まない */
  private historyIds(slug: string): string[] {
    const dir = this.historyDir(slug);
    if (!existsSync(dir)) return [];
    return readdirSync(dir)
      .filter((f) => f.endsWith('.json') && f !== '_labels.json')
      .map((f) => f.slice(0, -5))
      .sort()
      .reverse();
  }

  history(slug: string): HistoryEntry[] {
    const dir = this.historyDir(slug);
    const labels = this.labels(slug);
    return this.historyIds(slug)
      .map((id) => {
        const file = join(dir, `${id}.json`);
        const st = statSync(file);
        let savedAt = st.mtime.toISOString();
        try {
          savedAt = (JSON.parse(readFileSync(file, 'utf8')) as Project).meta.updatedAt || savedAt;
        } catch {
          // 壊れた版でも一覧には出す
        }
        const entry: HistoryEntry = { id, savedAt, bytes: st.size };
        if (labels[id]) entry.label = labels[id];
        return entry;
      });
  }

  private historyFile(slug: string, id: string) {
    if (!/^[\d-]+$/.test(id)) throw new StoreError(400, 'history.invalid', `履歴のIDが不正です: ${id}`);
    const f = join(this.historyDir(slug), `${id}.json`);
    if (!existsSync(f)) throw new StoreError(404, 'history.notFound', `履歴が見つかりません: ${id}`);
    return f;
  }

  readHistory(slug: string, id: string): Project {
    return JSON.parse(readFileSync(this.historyFile(slug, id), 'utf8')) as Project;
  }

  /** 過去の版に戻す。いまの版は履歴に退避されるので、戻したこと自体も取り消せる */
  restore(slug: string, id: string): { updatedAt: string; historyId: string | null } {
    const old = this.readHistory(slug, id);
    return this.save(slug, old);
  }

  /** 履歴に名前を付ける。名前付きの版は、件数の上限で消さない */
  label(slug: string, id: string, label: string) {
    this.historyFile(slug, id);
    const labels = this.labels(slug);
    if (label.trim()) labels[id] = label.trim();
    else delete labels[id];
    writeAtomic(this.labelsFile(slug), `${JSON.stringify(labels, null, 2)}\n`);
  }

  /** 名前の無い版は新しい順に HISTORY_KEEP 件だけ残す */
  private prune(slug: string) {
    const labels = this.labels(slug);
    const unnamed = this.historyIds(slug).filter((id) => !labels[id]);
    for (const id of unnamed.slice(HISTORY_KEEP)) rmSync(join(this.historyDir(slug), `${id}.json`));
  }
}
