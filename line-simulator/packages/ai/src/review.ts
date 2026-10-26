// AIの提案（シナリオ差分）を人が確かめるための部品。画面を持たない（編集ツール・CLI・将来のツール内AIで共用）。
//
//   1. 形の確認：差分JSONとして読めるか（checkShape）
//   2. 下見：選んだ提案だけを当てたら、どうなるか・重大エラーが出るか（dryRun）
//   3. 承認待ちの一覧：案件の中の「AI推定」のもの（listInferred）
//   4. 却下：取り込み済みの AI 推定を取り除く（removeAt）
// 当てる・承認するのはエンジンの applyPatch／approve をそのまま使う（エンジンは変えない）。

import type { Issue, PatchOp, Project, ProvenanceState, ScenarioPatch } from '../../engine/src/index.ts';
import { applyPatch, projectHash, validate } from '../../engine/src/index.ts';

export type ShapeProblem = { index: number | null; message: string };

const STATES: ProvenanceState[] = ['observed', 'inferred', 'approved', 'manual'];

/** 差分JSONの形を確かめる。中身の正しさ（参照先があるか等）は dryRun の検証で見る */
export function checkShape(data: unknown): ShapeProblem[] {
  const out: ShapeProblem[] = [];
  const p = data as Partial<ScenarioPatch> | null;
  if (!p || typeof p !== 'object') return [{ index: null, message: 'JSONのオブジェクトではありません' }];
  if (p.schemaVersion !== 1) out.push({ index: null, message: 'schemaVersion が 1 ではありません' });
  if (typeof p.summary !== 'string' || !p.summary.trim()) out.push({ index: null, message: 'summary（何をしたかの一文）がありません' });
  if (!Array.isArray(p.ops) || p.ops.length === 0) return [...out, { index: null, message: 'ops（提案の一覧）がありません' }];
  p.ops.forEach((op, i) => {
    if (!op || typeof op !== 'object') return void out.push({ index: i, message: '提案がオブジェクトではありません' });
    if (!['add', 'replace', 'remove'].includes(op.op)) out.push({ index: i, message: `op が add／replace／remove ではありません（${String(op.op)}）` });
    if (typeof op.path !== 'string' || !op.path.startsWith('/')) out.push({ index: i, message: 'path が「/」で始まる場所の指定ではありません' });
    if (op.op !== 'remove' && op.value === undefined) out.push({ index: i, message: '追加・置き換えなのに value がありません' });
    if (!op.prov || !STATES.includes(op.prov.state)) out.push({ index: i, message: 'prov.state（出所）がありません' });
    else if (op.prov.state === 'approved') out.push({ index: i, message: '「承認済み」はAIが付けてはいけません（人が承認します）' });
    else if (op.prov.state === 'inferred' && !op.prov.rationale) out.push({ index: i, message: 'AI推定なのに rationale（根拠）がありません' });
    if (op.prov?.state === 'observed' && !op.prov.source) out.push({ index: i, message: '確認済みなのに source（どのスクショか）がありません' });
  });
  return out;
}

export type DryRun = { project: Project; issues: Issue[]; failed: { index: number; message: string }[]; stale: boolean };

/** 選んだ提案だけを当てた結果（元の案件は変えない）。stale は「提案を作った後に案件が変わっている」 */
export function dryRun(project: Project, patch: ScenarioPatch, selected: Set<number>): DryRun {
  const r = applyPatch(project, patch, (_op, i) => selected.has(i));
  const stale = !!patch.basedOn && patch.basedOn !== projectHash(project);
  return { project: r.project, issues: r.issues, failed: r.failed, stale };
}

// ---------- 表示用の要約 ----------

const OP_LABEL = { add: '追加', replace: '置き換え', remove: '削除' } as const;
export const opLabel = (op: PatchOp['op']) => OP_LABEL[op];

/** 値の中身を1行で（見出し・名前・本文の頭） */
export function summarizeValue(v: unknown): string {
  if (v === undefined) return '';
  if (v === null || typeof v !== 'object') return String(v).slice(0, 60);
  const o = v as Record<string, unknown>;
  for (const k of ['label', 'title', 'name', 'text', 'companyName']) if (typeof o[k] === 'string' && o[k]) return (o[k] as string).split('\n')[0]!.slice(0, 60);
  if (Array.isArray(o['bubbles'])) {
    const b = (o['bubbles'] as { text?: string; title?: string }[])[0];
    return (b?.text ?? b?.title ?? '').split('\n')[0]!.slice(0, 60);
  }
  if (Array.isArray(o['keywords'])) return `キーワード：${(o['keywords'] as string[]).join('・')}`;
  if (o['match'] && typeof o['match'] === 'object') return `キーワード：${((o['match'] as { keywords?: string[] }).keywords ?? []).join('・')}`;
  if (typeof o['type'] === 'string') return String(o['type']);
  if (typeof o['id'] === 'string') return o['id'] as string;
  return Array.isArray(v) ? `${v.length}件` : '';
}

// ---------- 案件の中の出所 ----------

export type ProvItem = { path: string; state: ProvenanceState; summary: string; source?: string; rationale?: string };

/** 案件の中で出所（prov）が付いているもの。state を指定すればそれだけ */
export function listProv(project: Project, state?: ProvenanceState): ProvItem[] {
  const out: ProvItem[] = [];
  const walk = (v: unknown, path: string) => {
    if (Array.isArray(v)) return v.forEach((x, i) => walk(x, `${path}/${i}`));
    if (!v || typeof v !== 'object') return;
    const o = v as Record<string, unknown>;
    const prov = o['prov'] as { state?: ProvenanceState; source?: string; rationale?: string } | undefined;
    if (prov?.state && (!state || prov.state === state))
      out.push({ path: path || '/', state: prov.state, summary: summarizeValue(o), ...(prov.source ? { source: prov.source } : {}), ...(prov.rationale ? { rationale: prov.rationale } : {}) });
    for (const [k, x] of Object.entries(o)) if (k !== 'prov') walk(x, `${path}/${k}`);
  };
  walk(project, '');
  return out;
}

export const countProv = (project: Project) => {
  const all = listProv(project);
  return { observed: all.filter((x) => x.state === 'observed').length, inferred: all.filter((x) => x.state === 'inferred').length, approved: all.filter((x) => x.state === 'approved').length };
};

/**
 * 取り込み済みの提案を却下して取り除く。配列の要素なら抜き、オブジェクトの項目なら消す。
 * フローのノードを抜くときは、そのノードに出入りする線も一緒に消す（行き先の無い線を残さない）。
 */
export function removeAt(project: Project, pointer: string): Project {
  const next = structuredClone(project);
  const keys = pointer.split('/').slice(1);
  const last = keys.pop();
  if (last === undefined || (keys.length === 0 && last === '')) return next; // 案件そのものは消さない
  let parent: unknown = next;
  for (const k of keys) parent = Array.isArray(parent) ? parent[Number(k)] : (parent as Record<string, unknown>)?.[k];
  if (Array.isArray(parent)) {
    const [removed] = parent.splice(Number(last), 1) as { id?: string }[];
    // /flows/<i>/nodes/<j> のときは線も消す
    if (keys.length === 3 && keys[0] === 'flows' && keys[2] === 'nodes' && removed?.id) {
      const flow = next.flows[Number(keys[1])];
      if (flow) flow.edges = flow.edges.filter((e) => e.from.nodeId !== removed.id && e.to !== removed.id);
    }
  } else if (parent && typeof parent === 'object') delete (parent as Record<string, unknown>)[last];
  return next;
}

/** 下見の重大エラー（validate と同じ基準） */
export const errorsOf = (issues: Issue[]) => issues.filter((i) => i.level === 'error');
export const recheck = (project: Project) => validate(project);
