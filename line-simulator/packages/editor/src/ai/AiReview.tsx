// AI提案の画面：Claude Code が作った提案（シナリオ差分）を確かめ、選んで取り込み、AI推定を承認・却下・修正する。
// 手順は docs/AI_WORKFLOW.md。AIの推定は承認するまで営業用HTMLに書き出せない（検証の重大エラー prov.unapproved）。

import { useEffect, useMemo, useRef, useState } from 'react';
import type { ScenarioPatch } from '../../../engine/src/index.ts';
import { approve } from '../../../engine/src/index.ts';
import { checkShape, countProv, dryRun, listProv, opLabel, removeAt, summarizeValue, type ShapeProblem } from '../../../ai/src/review.ts';
import { describeIssue } from '../../../player/src/issues.ts';
import { api, fmtDateTime } from '../api.ts';
import { useEditor } from '../ctx.ts';
import { go, locate, where } from '../route.ts';
import { InlineIssues } from '../ui/form.tsx';
import { ProvBadge, replaceDraft } from './ProvBadge.tsx';

type PatchEntry = { name: string; savedAt: string; bytes: number };

export function AiReview() {
  const { project, update, slug, toast } = useEditor();
  const [files, setFiles] = useState<PatchEntry[] | null>(null);
  const [name, setName] = useState<string | null>(null);
  const [patch, setPatch] = useState<ScenarioPatch | null>(null);
  const [shape, setShape] = useState<ShapeProblem[]>([]);
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const input = useRef<HTMLInputElement>(null);

  const reload = () => api.patches(slug).then(setFiles, (e: Error) => toast(e.message, 'err'));
  useEffect(() => void reload(), [slug]);

  const open = async (n: string) => {
    try {
      const data = await api.patch(slug, n);
      const problems = checkShape(data);
      setName(n);
      setShape(problems);
      setPatch(problems.length ? null : (data as ScenarioPatch));
      setSelected(new Set(problems.length ? [] : (data as ScenarioPatch).ops.map((_, i) => i)));
    } catch (e) {
      toast((e as Error).message, 'err');
    }
  };

  const importFile = async (f: File | undefined) => {
    if (!f) return;
    try {
      const data = JSON.parse(await f.text()) as unknown;
      const { name: saved } = await api.savePatch(slug, f.name, data);
      await reload();
      await open(saved);
      toast('提案ファイルを読み込みました');
    } catch (e) {
      toast(`読み込めません：${(e as Error).message}`, 'err');
    }
    if (input.current) input.current.value = '';
  };

  const run = useMemo(() => (patch ? dryRun(project, patch, selected) : null), [project, patch, selected]);
  const otherErrors = run ? run.issues.filter((i) => i.level === 'error' && i.code !== 'prov.unapproved') : [];
  const counts = countProv(project);
  const pending = listProv(project, 'inferred');

  const toggle = (i: number) =>
    setSelected((s) => {
      const n = new Set(s);
      if (n.has(i)) n.delete(i);
      else n.add(i);
      return n;
    });

  const apply = () => {
    if (!run || !patch) return;
    if (run.failed.length) return toast('当たらない提案があるため取り込めません。提案を作り直してください', 'err');
    if (otherErrors.length && !window.confirm(`選んだ提案を取り込むと、重大エラーが ${otherErrors.length} 件出ます（あとで直せます）。取り込みますか？`)) return;
    update((d) => replaceDraft(d, run.project));
    toast(`${selected.size}件の提案を下書きに取り込みました。承認待ちを確かめてから保存してください`);
  };

  return (
    <div className="detail" style={{ width: '100%' }}>
      <h2>AI提案（スクリーンショットから）</h2>
      <p className="muted small" style={{ margin: '4px 0 0' }}>
        Claude Code が既存アカウントのスクショから作ったシナリオ案を、確かめて取り込みます。手順は <span className="mono">docs/AI_WORKFLOW.md</span>。
        <b>AI推定は、承認するまで営業用HTMLに書き出せません。</b>
      </p>
      <div className="row" style={{ marginTop: 10 }}>
        <span className="prov prov-observed">確認済み {counts.observed}</span>
        <span className="prov prov-inferred">AI推定（承認待ち） {counts.inferred}</span>
        <span className="prov prov-approved">承認済み {counts.approved}</span>
      </div>

      <div className="section">
        <h3>
          提案ファイル
          <span className="spacer" />
          <button className="btn small" onClick={() => input.current?.click()}>
            ファイルから読み込む
          </button>
          <input ref={input} type="file" accept=".json,application/json" hidden onChange={(e) => void importFile(e.target.files?.[0])} />
        </h3>
        {files === null ? (
          <div className="muted small">読み込み中…</div>
        ) : files.length === 0 ? (
          <div className="muted small">
            まだありません。Claude Code に「スクショからシナリオ案を作って（{slug}）」と頼むと、<span className="mono">projects/{slug}/patches/</span> に置かれます。
          </div>
        ) : (
          <div className="row">
            {files.map((f) => (
              <button key={f.name} className={`btn small${f.name === name ? ' primary' : ''}`} onClick={() => void open(f.name)} title={fmtDateTime(f.savedAt)}>
                {f.name}
              </button>
            ))}
          </div>
        )}
        {shape.length > 0 && (
          <div className="inline-issues" style={{ marginTop: 8 }}>
            {shape.map((s, i) => (
              <div key={i} className="error">
                ✖ {s.index === null ? '全体' : `提案${s.index + 1}`}：{s.message}
              </div>
            ))}
          </div>
        )}
      </div>

      {patch && run && (
        <div className="section">
          <h3>
            {name}
            <span className="spacer" />
            <button className="btn small" onClick={() => setSelected(new Set(patch.ops.map((_, i) => i)))}>
              全部選ぶ
            </button>
            <button className="btn small" onClick={() => setSelected(new Set(patch.ops.flatMap((o, i) => (o.prov.state === 'observed' ? [i] : []))))}>
              確認済みだけ
            </button>
            <button className="btn small" onClick={() => setSelected(new Set())}>
              外す
            </button>
          </h3>
          <div className="small" style={{ marginBottom: 8 }}>
            {patch.summary}
          </div>
          {run.stale && <div className="prov-note prov-note-inferred small">この提案を作った後に、案件が変わっています。中身を確かめてから取り込んでください。</div>}
          <table className="ai-ops">
            <thead>
              <tr>
                <th>取り込む</th>
                <th>出所</th>
                <th>内容</th>
                <th>根拠・確認元</th>
              </tr>
            </thead>
            <tbody>
              {patch.ops.map((op, i) => {
                const failed = run.failed.find((f) => f.index === i);
                return (
                  <tr key={i} className={selected.has(i) ? '' : 'off'}>
                    <td>
                      <input type="checkbox" checked={selected.has(i)} onChange={() => toggle(i)} aria-label={`提案${i + 1}を取り込む`} />
                    </td>
                    <td>
                      <ProvBadge prov={op.prov} />
                    </td>
                    <td>
                      <b>{opLabel(op.op)}</b>　{where(project, op.path.replace(/\/-$/, ''))}
                      <div className="muted small">{summarizeValue(op.value)}</div>
                      {failed && <div className="bad small">当たりません：{failed.message}</div>}
                    </td>
                    <td className="small">{op.prov.state === 'observed' ? op.prov.source : op.prov.rationale}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          <div className="row" style={{ marginTop: 10 }}>
            <span className="small">
              選んだ {selected.size}件を取り込むと：重大エラー <b>{otherErrors.length}</b>件（承認待ちを除く）・AI推定（承認待ち） <b>{countProv(run.project).inferred}</b>件
            </span>
            <span className="spacer" />
            <button className="btn primary" disabled={selected.size === 0} onClick={apply}>
              選んだ提案を下書きに取り込む
            </button>
          </div>
          {otherErrors.length > 0 && <InlineIssues issues={otherErrors.slice(0, 8)} />}
        </div>
      )}

      <div className="section">
        <h3>
          承認待ちのAI推定（{pending.length}件）
          <span className="spacer" />
          {pending.length > 0 && (
            <button
              className="btn small"
              onClick={() => {
                if (window.confirm(`AI推定 ${pending.length}件を、すべて承認しますか？（中身を確かめてから承認してください）`)) update((d) => replaceDraft(d, approve(d, '')));
              }}
            >
              すべて承認
            </button>
          )}
        </h3>
        {pending.length === 0 ? (
          <div className="muted small">ありません。営業用HTMLを書き出せる状態です（ほかの重大エラーが無ければ）。</div>
        ) : (
          <table className="ai-ops">
            <tbody>
              {pending.map((p) => (
                <tr key={p.path}>
                  <td style={{ width: 70 }}>
                    <span className="prov prov-inferred">AI推定</span>
                  </td>
                  <td>
                    <b>{where(project, p.path)}</b>
                    <div className="muted small">{p.summary}</div>
                    <div className="small">根拠：{p.rationale ?? '（記載なし）'}</div>
                  </td>
                  <td style={{ width: 230, textAlign: 'right' }}>
                    <button className="btn small primary" onClick={() => update((d) => replaceDraft(d, approve(d, p.path)))}>
                      承認
                    </button>{' '}
                    <button className="btn small" onClick={() => go(locate(project, slug, p.path))}>
                      開いて直す
                    </button>{' '}
                    <button
                      className="btn small ghost danger"
                      onClick={() => {
                        if (window.confirm(`「${where(project, p.path)}」を却下して取り除きますか？`)) update((d) => replaceDraft(d, removeAt(d, p.path)));
                      }}
                    >
                      却下
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        {pending.length > 0 && <div className="muted small" style={{ marginTop: 6 }}>{describeIssue({ level: 'error', code: 'prov.unapproved', path: '/' })}</div>}
      </div>
    </div>
  );
}
