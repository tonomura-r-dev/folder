// 条件式の入力欄。ふだんは画面で組み立て（ExprBuilder）、必要なら JSON でも書ける。
// どちらで書いても、下に日本語の意味と検証結果を出す。

import { useState } from 'react';
import type { Expr } from '../../../engine/src/index.ts';
import { describeExpr, Labels } from '../../../player/src/describe.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { ExprBuilder } from './ExprBuilder.tsx';
import { InlineIssues, JsonField } from './form.tsx';

export function ExprField({ label, value, onChange, path, hint }: { label: string; value: Expr; onChange: (e: Expr) => void; path: string; hint?: string }) {
  const { project, issues } = useEditor();
  const [json, setJson] = useState(false);
  const L = new Labels(project);
  const describe = (e: Expr) => {
    try {
      return `→ ${describeExpr(e, L)}`;
    } catch {
      return '→ （読み取れない条件です）';
    }
  };
  const mine = issuesAt(issues, path);
  return (
    <div className={`field wide${mine.some((i) => i.level === 'error') ? ' has-err' : ''}`}>
      <div className="row" style={{ justifyContent: 'space-between' }}>
        <span>{label}</span>
        <span className="xb-switch" role="tablist" aria-label="入力のしかた">
          <button type="button" role="tab" aria-selected={!json} className={json ? '' : 'on'} onClick={() => setJson(false)}>
            組み立て
          </button>
          <button type="button" role="tab" aria-selected={json} className={json ? 'on' : ''} onClick={() => setJson(true)}>
            JSON
          </button>
        </span>
      </div>
      {json ? <JsonField<Expr> value={value} onChange={onChange} describe={describe} /> : <ExprBuilder value={value} onChange={onChange} />}
      {!json && <div className="preview">{describe(value)}</div>}
      {hint && <div className="hint">{hint}</div>}
      {mine.length > 0 && <InlineIssues issues={mine} />}
    </div>
  );
}
