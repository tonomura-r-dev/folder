// 案件の編集画面の骨組み：上の帯（保存・履歴・検証・書き出し）／左の目次／中身／検証一覧。

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { Project } from '../../engine/src/index.ts';
import { personaCoverage, validate } from '../../engine/src/index.ts';
import { describeIssue } from '../../player/src/issues.ts';
import { api, ApiError, downloadJson, fmtDateTime, type ExportResponse, type HistoryEntry } from './api.ts';
import { Ctx, type EditorCtx } from './ctx.ts';
import { FlowEditor } from './flow/FlowEditor.tsx';
import { DemoView } from './DemoView.tsx';
import { AutoReplies } from './parts/AutoReplies.tsx';
import { LabelsEditor } from './parts/Labels.tsx';
import { Messages } from './parts/Messages.tsx';
import { Narration } from './parts/Narration.tsx';
import { Personas } from './parts/Personas.tsx';
import { RichMenus } from './parts/RichMenus.tsx';
import { Settings } from './parts/Settings.tsx';
import { Surveys } from './parts/Surveys.tsx';
import { WebPages } from './parts/WebPages.tsx';
import { go, href, locate, where, type Route, type Section } from './route.ts';
import { InlineIssues, Modal } from './ui/form.tsx';

type Doc = { project: Project; base: string; rev: number; savedRev: number };
type ProjectRoute = Extract<Route, { page: 'project' }>;

const NAV: { section: Section; label: string; count?: (p: Project) => number }[] = [
  { section: 'flows', label: 'シナリオ（フロー）', count: (p) => p.flows.length },
  { section: 'messages', label: '文面', count: (p) => p.messages.length },
  { section: 'labels', label: 'タグ・CV・経路', count: (p) => p.tags.length },
  { section: 'surveys', label: 'アンケート', count: (p) => p.surveys.length },
  { section: 'richMenus', label: 'リッチメニュー', count: (p) => p.richMenus.length },
  { section: 'autoReplies', label: '自動応答', count: (p) => p.autoReplies.length },
  { section: 'webPages', label: 'Webページ', count: (p) => p.webPages.length },
  { section: 'personas', label: 'ペルソナ', count: (p) => p.personas.length },
  { section: 'narration', label: '説明文', count: (p) => p.narration.length },
  { section: 'settings', label: '案件設定' },
];

const NAV_DEMO: { section: Section; label: string } = { section: 'demo', label: '▶ デモ（動作確認）' };

export function ProjectEditor({ route, toast }: { route: ProjectRoute; toast: (m: string, k?: 'ok' | 'err') => void }) {
  const { slug } = route;
  const [doc, setDoc] = useState<Doc | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [conflict, setConflict] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [issuesOpen, setIssuesOpen] = useState(true);
  const docRef = useRef<Doc | null>(null);
  docRef.current = doc;

  const load = useCallback(async () => {
    try {
      const p = await api.read(slug);
      setDoc({ project: p, base: p.meta.updatedAt, rev: 0, savedRev: 0 });
      setLoadError(null);
    } catch (e) {
      setLoadError((e as Error).message);
    }
  }, [slug]);
  useEffect(() => void load(), [load]);

  const dirty = !!doc && doc.rev !== doc.savedRev;

  const update = useCallback((fn: (draft: Project) => void) => {
    setDoc((d) => {
      if (!d) return d;
      const next = structuredClone(d.project);
      fn(next);
      return { ...d, project: next, rev: d.rev + 1 };
    });
  }, []);

  const save = useCallback(
    async (force = false) => {
      const d = docRef.current;
      if (!d || saving) return;
      setSaving(true);
      try {
        const r = await api.save(slug, d.project, force ? undefined : d.base);
        // 保存中に編集が進んでいても消さない（保存した版までを「保存済み」にする）
        setDoc((cur) =>
          cur && { ...cur, base: r.updatedAt, savedRev: d.rev, project: { ...cur.project, meta: { ...cur.project.meta, updatedAt: r.updatedAt } } },
        );
        setConflict(false);
        toast('保存しました');
      } catch (e) {
        if (e instanceof ApiError && e.status === 409) setConflict(true);
        else toast(`保存できませんでした：${(e as Error).message}`, 'err');
      } finally {
        setSaving(false);
      }
    },
    [slug, saving, toast],
  );

  // Ctrl+S で保存・未保存のまま閉じようとしたら止める
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        void save();
      }
    };
    const onLeave = (e: BeforeUnloadEvent) => {
      if (docRef.current && docRef.current.rev !== docRef.current.savedRev) e.preventDefault();
    };
    window.addEventListener('keydown', onKey);
    window.addEventListener('beforeunload', onLeave);
    return () => {
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('beforeunload', onLeave);
    };
  }, [save]);

  // 重大エラーが無いときだけ、ペルソナを台本どおり流して「誰も入らないシナリオ」も確かめる
  const issues = useMemo(() => {
    if (!doc) return [];
    const v = validate(doc.project);
    return v.some((i) => i.level === 'error') ? v : [...v, ...personaCoverage(doc.project)];
  }, [doc?.project]);
  const errors = issues.filter((i) => i.level === 'error').length;
  const warnings = issues.length - errors;

  const ctx = useMemo<EditorCtx | null>(
    () => (doc ? { slug, project: doc.project, update, issues, route, toast } : null),
    [doc?.project, slug, update, issues, route, toast],
  );

  if (loadError)
    return (
      <div className="page">
        <div className="empty">
          {loadError}
          <div style={{ marginTop: 12 }}>
            <a href="#/">一覧へ戻る</a>
          </div>
        </div>
      </div>
    );
  if (!doc || !ctx) return <div className="page muted">読み込み中…</div>;

  const p = doc.project;
  const leave = (e: React.MouseEvent) => {
    if (dirty && !window.confirm('保存していない変更があります。破棄して一覧へ戻りますか？')) e.preventDefault();
  };

  return (
    <Ctx.Provider value={ctx}>
      <div className="shell">
        <div className="topbar">
          <a className="btn ghost" href="#/" onClick={leave}>
            ← 一覧
          </a>
          <h1>{p.meta.name}</h1>
          {dirty ? <span className="dirty">● 未保存の変更あり</span> : <span className="saved">保存済み {fmtDateTime(p.meta.updatedAt)}</span>}
          <span className="spacer" />
          <button className="btn" onClick={() => setIssuesOpen((v) => !v)} title="検証結果の一覧を開く／閉じる">
            {errors > 0 ? <span className="badge err">重大 {errors}</span> : <span className="badge ok">重大 0</span>}
            {warnings > 0 && <span className="badge warn">警告 {warnings}</span>}
          </button>
          <button className="btn" onClick={() => setHistoryOpen(true)}>
            履歴
          </button>
          <button className="btn" onClick={() => downloadJson(`${slug}.json`, p)} title="いま画面にある内容（未保存の変更も含む）をJSONで保存">
            JSONで書き出し
          </button>
          <button className="btn" onClick={() => setExportOpen(true)} title="社外にも渡せる、ネット接続なしで動く単一のHTMLを Downloads に書き出す">
            営業用HTMLを書き出す
          </button>
          <button className="btn primary" disabled={!dirty || saving} onClick={() => void save()} title="Ctrl+S でも保存できます">
            {saving ? '保存中…' : '保存'}
          </button>
        </div>

        <div className="body">
          <nav className="side" aria-label="編集する部品">
            <h3>確かめる</h3>
            <a href={href({ page: 'project', slug, section: NAV_DEMO.section })} className={route.section === NAV_DEMO.section ? 'on' : ''}>
              {NAV_DEMO.label}
            </a>
            <h3>編集</h3>
            {NAV.map((n) => (
              <a key={n.section} href={href({ page: 'project', slug, section: n.section })} className={route.section === n.section ? 'on' : ''}>
                {n.label}
                {n.count && <span className="count">{n.count(p)}</span>}
              </a>
            ))}
          </nav>
          <main className="main">
            <SectionView section={route.section} />
          </main>
        </div>

        {issuesOpen && issues.length > 0 && (
          <div className="issues" aria-label="検証結果">
            <div className="issues-head">
              <h3>検証結果</h3>
              <span className="muted small">重大エラーがあると営業用HTMLは書き出せません。行をクリックすると該当箇所へ移動します。</span>
              <span className="spacer" />
              <button className="btn small ghost" onClick={() => setIssuesOpen(false)}>
                閉じる
              </button>
            </div>
            {issues.map((i, k) => (
              <div key={k} className="issue" onClick={() => go(locate(p, slug, i.path))}>
                <span className={`lv ${i.level}`}>{i.level === 'error' ? '重大' : '警告'}</span>
                <span>{describeIssue(i)}</span>
                <span className="where">{where(p, i.path)}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      {exportOpen && <ExportDialog slug={slug} dirty={dirty} errors={errors} onClose={() => setExportOpen(false)} onSave={() => void save()} />}

      {historyOpen && (
        <HistoryDialog
          slug={slug}
          dirty={dirty}
          onClose={() => setHistoryOpen(false)}
          onRestored={async () => {
            setHistoryOpen(false);
            await load();
            toast('過去の版に戻しました（戻す前の版も履歴に残っています）');
          }}
          toast={toast}
        />
      )}

      {conflict && (
        <Modal title="別の場所で更新されています" onClose={() => setConflict(false)}>
          <p>この案件は、画面を開いた後に別のタブや別の作業で保存されています。このまま保存すると、その変更を上書きします。</p>
          <p className="muted small">上書きしても、上書きされる側の版は履歴に残ります。</p>
          <div className="actions">
            <button
              className="btn"
              onClick={() => {
                if (window.confirm('この画面の未保存の変更を捨てて、保存されている最新の内容を読み込みます。よろしいですか？')) {
                  setConflict(false);
                  void load();
                }
              }}
            >
              最新を読み込み直す
            </button>
            <button className="btn primary" onClick={() => void save(true)}>
              この画面の内容で上書き
            </button>
          </div>
        </Modal>
      )}
    </Ctx.Provider>
  );
}

function SectionView({ section }: { section: Section }) {
  switch (section) {
    case 'demo':
      return <DemoView />;
    case 'flows':
      return <FlowEditor />;
    case 'messages':
      return <Messages />;
    case 'labels':
      return <LabelsEditor />;
    case 'surveys':
      return <Surveys />;
    case 'richMenus':
      return <RichMenus />;
    case 'autoReplies':
      return <AutoReplies />;
    case 'webPages':
      return <WebPages />;
    case 'personas':
      return <Personas />;
    case 'narration':
      return <Narration />;
    case 'settings':
      return <Settings />;
  }
}

function HistoryDialog({
  slug,
  dirty,
  onClose,
  onRestored,
  toast,
}: {
  slug: string;
  dirty: boolean;
  onClose: () => void;
  onRestored: () => void;
  toast: (m: string, k?: 'ok' | 'err') => void;
}) {
  const [list, setList] = useState<HistoryEntry[] | null>(null);
  const reload = () => api.history(slug).then(setList, (e: Error) => toast(e.message, 'err'));
  useEffect(() => void reload(), []);
  return (
    <Modal title="保存の履歴" onClose={onClose}>
      <p className="muted small">
        保存するたびに、それまでの版がここに残ります。日時はその版を保存した時刻です（名前の無い版は新しい順に30件まで。名前を付けた版は消えません）。
      </p>
      {list === null ? (
        <div className="muted">読み込み中…</div>
      ) : list.length === 0 ? (
        <div className="muted">まだ履歴はありません。保存すると増えます。</div>
      ) : (
        list.map((h) => (
          <div className="history-row" key={h.id}>
            <div>
              <div>
                <b>{fmtDateTime(h.savedAt)}</b> {h.label && <span className="badge ok">{h.label}</span>}
              </div>
              <div className="muted small mono">
                {h.id}　{Math.round(h.bytes / 1024)}KB
              </div>
            </div>
            <div className="row">
              <button
                className="btn small"
                onClick={async () => {
                  const label = window.prompt('この版の名前（例：提案前・先方確認版）。空にすると名前を外します。', h.label ?? '');
                  if (label === null) return;
                  await api.label(slug, h.id, label);
                  void reload();
                }}
              >
                名前を付ける
              </button>
              <button
                className="btn small primary"
                onClick={async () => {
                  const msg = dirty
                    ? '保存していない変更は消えます。この版に戻しますか？（いま保存されている版は履歴に残ります）'
                    : 'この版に戻しますか？（いま保存されている版は履歴に残ります）';
                  if (!window.confirm(msg)) return;
                  try {
                    await api.restore(slug, h.id);
                    onRestored();
                  } catch (e) {
                    toast((e as Error).message, 'err');
                  }
                }}
              >
                この版に戻す
              </button>
            </div>
          </div>
        ))
      )}
      <div className="actions">
        <button className="btn" onClick={onClose}>
          閉じる
        </button>
      </div>
    </Modal>
  );
}

/** 営業用HTMLの書き出し。保存済みの内容を書き出すので、未保存の変更があれば先に保存してもらう */
function ExportDialog({ slug, dirty, errors, onClose, onSave }: { slug: string; dirty: boolean; errors: number; onClose: () => void; onSave: () => void }) {
  const [confirmed, setConfirmed] = useState(false);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<ExportResponse | null>(null);
  const [failure, setFailure] = useState<string | null>(null);
  const run = async () => {
    setBusy(true);
    setFailure(null);
    try {
      setResult(await api.exportHtml(slug, confirmed));
    } catch (e) {
      setFailure((e as Error).message);
    } finally {
      setBusy(false);
    }
  };
  const shown = result?.issues ?? [];
  return (
    <Modal title="営業用HTMLを書き出す" onClose={onClose}>
      <p className="small">
        ネット接続なしで動く単一のHTMLを、<b>Downloads</b> フォルダに書き出します。編集機能・社内メモは入りません。
        重大エラーがあると書き出せません（いま {errors} 件）。
      </p>
      {dirty && (
        <div className="row" style={{ background: 'var(--warn-soft)', padding: 8, borderRadius: 8 }}>
          <span className="small">保存していない変更があります。書き出すのは保存済みの内容です。</span>
          <button className="btn small" onClick={onSave}>
            いま保存する
          </button>
        </div>
      )}
      <label className="row small" style={{ marginTop: 10 }}>
        <input type="checkbox" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} />
        会社名・人物・求人などは、架空または匿名化したものだけです（実在の顧客情報は入っていません）
      </label>
      {result && (
        <div className="section" style={{ marginTop: 12 }}>
          {result.ok ? (
            <>
              <b style={{ color: 'var(--accent-ink)' }}>書き出しました（{Math.round(result.bytes / 1024)}KB）</b>
              <div className="readonly" style={{ marginTop: 6, wordBreak: 'break-all' }}>
                {result.path}
              </div>
            </>
          ) : (
            <b style={{ color: 'var(--err)' }}>書き出しませんでした</b>
          )}
          {shown.length > 0 && <InlineIssues issues={shown} />}
        </div>
      )}
      {failure && <div className="bad small">書き出せませんでした：{failure}</div>}
      <div className="actions">
        <button className="btn" onClick={onClose}>
          閉じる
        </button>
        <button className="btn primary" disabled={!confirmed || busy || dirty} onClick={() => void run()}>
          {busy ? '書き出し中…' : '書き出す'}
        </button>
      </div>
    </Modal>
  );
}
