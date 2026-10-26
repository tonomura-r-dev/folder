// シナリオエンジンの公開API。編集ツール・営業用HTML・CLI はここからだけ読み込む。

export * from './schema.ts';
export { evaluate, failingLeaves, type EvalContext } from './evaluate.ts';
export { reduce, sessionFromPersonas, dayAt, STEP_BUDGET, type ReduceOptions } from './reduce.ts';
export { validate, hasErrors, LINE_LIMITS, type ValidateOptions } from './validate.ts';
export { applyPatch, approve, type PatchResult } from './patch.ts';
export { renderText, isKnownVar } from './text.ts';
export { projectHash } from './hash.ts';
export { portsOf, choiceButtons, acceptsInput } from './ports.ts';
export { personaCoverage } from './coverage.ts';
