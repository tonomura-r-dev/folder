// シナリオ差分（AI連携の境界）。AIは案件JSONを丸ごと書き換えず、この差分だけを出す。
// 初期版は Claude Code が出し、将来はツール内のAIが同じ形を出す。

import type { Issue, PatchOp, Project, ScenarioPatch } from './schema.ts';
import { validate } from './validate.ts';

const decode = (s: string) => s.replace(/~1/g, '/').replace(/~0/g, '~');

function split(pointer: string): string[] {
  if (pointer === '' || pointer === '/') return [];
  if (!pointer.startsWith('/')) throw new Error(`bad pointer: ${pointer}`);
  return pointer.slice(1).split('/').map(decode);
}

/** 差分の値に出所を付ける（オブジェクトのときだけ） */
function withProv(value: unknown, op: PatchOp): unknown {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return value;
  return { ...(value as Record<string, unknown>), prov: op.prov };
}

function applyOp(root: unknown, op: PatchOp) {
  const keys = split(op.path);
  const last = keys.pop();
  if (last === undefined) throw new Error('cannot patch the root');
  let parent: unknown = root;
  for (const k of keys) {
    parent = Array.isArray(parent) ? parent[Number(k)] : (parent as Record<string, unknown> | undefined)?.[k];
    if (parent === undefined || parent === null) throw new Error(`path not found: ${op.path}`);
  }
  if (Array.isArray(parent)) {
    const i = last === '-' ? parent.length : Number(last);
    if (!Number.isInteger(i) || i < 0 || i > parent.length) throw new Error(`bad index: ${op.path}`);
    if (op.op === 'add') parent.splice(i, 0, withProv(op.value, op));
    else if (i >= parent.length) throw new Error(`path not found: ${op.path}`);
    else if (op.op === 'replace') parent[i] = withProv(op.value, op);
    else parent.splice(i, 1);
    return;
  }
  const obj = parent as Record<string, unknown>;
  if (op.op !== 'add' && !(last in obj)) throw new Error(`path not found: ${op.path}`);
  if (op.op === 'remove') delete obj[last];
  else obj[last] = withProv(op.value, op);
}

export type PatchResult = { project: Project; issues: Issue[]; failed: { index: number; message: string }[] };

/** 差分を当てた新しい案件を返す（元の案件は変えない）。当たらなかった差分は failed に入る */
export function applyPatch(project: Project, patch: ScenarioPatch, accept?: (op: PatchOp, index: number) => boolean): PatchResult {
  const next = structuredClone(project);
  const failed: PatchResult['failed'] = [];
  const issues: Issue[] = [];
  patch.ops.forEach((op, index) => {
    // 「人が承認した」はAIが付けてはいけない
    if (op.prov.state === 'approved') {
      issues.push({ level: 'error', code: 'patch.approvedByAi', path: op.path, params: { index } });
      return;
    }
    if (accept && !accept(op, index)) return;
    try {
      applyOp(next, op);
    } catch (e) {
      failed.push({ index, message: e instanceof Error ? e.message : String(e) });
    }
  });
  return { project: next, issues: [...issues, ...validate(next)], failed };
}

/** 推定（inferred）を承認済み（approved）に書き換える。path 配下のすべてが対象 */
export function approve(project: Project, pointer: string): Project {
  const next = structuredClone(project);
  let target: unknown = next;
  for (const k of split(pointer)) target = Array.isArray(target) ? target[Number(k)] : (target as Record<string, unknown>)?.[k];
  const visit = (v: unknown) => {
    if (Array.isArray(v)) return v.forEach(visit);
    if (!v || typeof v !== 'object') return;
    const obj = v as Record<string, unknown>;
    const prov = obj['prov'] as { state?: string } | undefined;
    if (prov?.state === 'inferred') obj['prov'] = { ...prov, state: 'approved' };
    Object.values(obj).forEach(visit);
  };
  visit(target);
  return next;
}
