// 入力部品。どの編集画面もこれだけで組む（見た目と挙動を揃えるため）。

import { useEffect, useRef, useState, type ReactNode } from 'react';
import type { Issue } from '../../../engine/src/index.ts';
import { describeIssue } from '../../../player/src/issues.ts';

type FieldProps = { label: string; hint?: ReactNode; wide?: boolean; issues?: Issue[]; children: ReactNode };

export function Field({ label, hint, wide, issues, children }: FieldProps) {
  const hasErr = issues?.some((i) => i.level === 'error');
  return (
    <label className={`field${wide ? ' wide' : ''}${hasErr ? ' has-err' : ''}`}>
      <span>{label}</span>
      {children}
      {hint && <div className="hint">{hint}</div>}
      {issues && issues.length > 0 && <InlineIssues issues={issues} />}
    </label>
  );
}

export function InlineIssues({ issues }: { issues: Issue[] }) {
  return (
    <div className="inline-issues">
      {issues.map((i, k) => (
        <div key={k} className={i.level}>
          {i.level === 'error' ? '✖ ' : '▲ '}
          {describeIssue(i)}
        </div>
      ))}
    </div>
  );
}

export function Text({ value, onChange, placeholder }: { value: string; onChange: (v: string) => void; placeholder?: string }) {
  return <input type="text" value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} />;
}

export function TextArea({ value, onChange, rows = 4 }: { value: string; onChange: (v: string) => void; rows?: number }) {
  return <textarea value={value} rows={rows} onChange={(e) => onChange(e.target.value)} />;
}

export function NumberInput({ value, onChange, min, step = 1 }: { value: number | undefined; onChange: (v: number | undefined) => void; min?: number; step?: number }) {
  return (
    <input
      type="number"
      value={value ?? ''}
      min={min}
      step={step}
      onChange={(e) => onChange(e.target.value === '' ? undefined : Number(e.target.value))}
    />
  );
}

export type Option = { value: string; label: string };

export function Select({ value, onChange, options, empty }: { value: string; onChange: (v: string) => void; options: Option[]; empty?: string }) {
  const known = options.some((o) => o.value === value);
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)}>
      {empty !== undefined && <option value="">{empty}</option>}
      {!known && value !== '' && <option value={value}>（存在しない：{value}）</option>}
      {options.map((o) => (
        <option key={o.value} value={o.value}>
          {o.label}
        </option>
      ))}
    </select>
  );
}

export function Checks({ value, onChange, options }: { value: string[]; onChange: (v: string[]) => void; options: Option[] }) {
  if (options.length === 0) return <div className="muted small">選べるものがまだありません</div>;
  return (
    <div className="checks">
      {options.map((o) => (
        <label key={o.value}>
          <input
            type="checkbox"
            checked={value.includes(o.value)}
            onChange={(e) => onChange(e.target.checked ? [...value, o.value] : value.filter((v) => v !== o.value))}
          />
          {o.label}
        </label>
      ))}
    </div>
  );
}

export function ReadOnly({ children }: { children: ReactNode }) {
  return <div className="readonly">{children}</div>;
}

/**
 * JSON を直接書く欄。書いている途中の壊れた JSON は手元に持ち、正しく読めたときだけ反映する。
 * describe を渡すと、読めた値を日本語で下に出す。
 */
export function JsonField<T>({ value, onChange, describe, rows = 6 }: { value: T; onChange: (v: T) => void; describe?: (v: T) => string; rows?: number }) {
  const external = JSON.stringify(value, null, 2) ?? '';
  const [text, setText] = useState(external);
  const [bad, setBad] = useState<string | null>(null);
  const focused = useRef(false);
  useEffect(() => {
    if (!focused.current) {
      setText(external);
      setBad(null);
    }
  }, [external]);
  return (
    <>
      <textarea
        className="code"
        rows={rows}
        value={text}
        spellCheck={false}
        onFocus={() => (focused.current = true)}
        onBlur={() => {
          focused.current = false;
          if (!bad) setText(external);
        }}
        onChange={(e) => {
          setText(e.target.value);
          try {
            const v = JSON.parse(e.target.value) as T;
            setBad(null);
            onChange(v);
          } catch (err) {
            setBad(err instanceof Error ? err.message : String(err));
          }
        }}
      />
      {bad && <div className="bad">JSONとして読めません（反映していません）：{bad}</div>}
      {!bad && describe && <div className="preview">{describe(value)}</div>}
    </>
  );
}

/** 並び替え・削除ボタンの組 */
export function ItemTools({ index, length, onMove, onRemove }: { index: number; length: number; onMove: (to: number) => void; onRemove: () => void }) {
  return (
    <>
      <button type="button" className="btn small ghost" disabled={index === 0} onClick={() => onMove(index - 1)} title="上へ">
        ↑
      </button>
      <button type="button" className="btn small ghost" disabled={index === length - 1} onClick={() => onMove(index + 1)} title="下へ">
        ↓
      </button>
      <button type="button" className="btn small ghost danger" onClick={onRemove} title="削除">
        削除
      </button>
    </>
  );
}

export function move<T>(list: T[], from: number, to: number): T[] {
  const next = [...list];
  const [x] = next.splice(from, 1);
  next.splice(to, 0, x!);
  return next;
}

/** prefix_1, prefix_2 … のうち、まだ使われていない ID を返す */
export function newId(prefix: string, used: Iterable<string>): string {
  const set = new Set(used);
  for (let i = 1; ; i++) if (!set.has(`${prefix}_${i}`)) return `${prefix}_${i}`;
}

export function Modal({ title, children, onClose }: { title: string; children: ReactNode; onClose: () => void }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);
  return (
    <div className="overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="dialog" role="dialog" aria-modal="true" aria-label={title}>
        <h2>{title}</h2>
        {children}
      </div>
    </div>
  );
}
