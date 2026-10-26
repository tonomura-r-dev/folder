// 案件の中身を選ぶ部品（文面・タグ・ページ…）と、条件式・動作・時間の入力。

import { useState } from 'react';
import type { Action, Expr, Id, InputEventType, Minutes, Project, Trigger } from '../../../engine/src/index.ts';
import { describeExpr, Labels } from '../../../player/src/describe.ts';
import { issuesAt, useEditor } from '../ctx.ts';
import { Checks, Field, InlineIssues, JsonField, NumberInput, Select, Text, type Option } from './form.tsx';

export const options = {
  messages: (p: Project, onlyButtons = false): Option[] => {
    const L = new Labels(p);
    return p.messages
      .filter((m) => !onlyButtons || m.bubbles.some((b) => b.kind === 'buttons' || b.kind === 'card'))
      .map((m) => ({ value: m.id, label: `${L.message(m.id)}（${m.id}）` }));
  },
  tags: (p: Project): Option[] => p.tags.map((t) => ({ value: t.id, label: t.label })),
  cvs: (p: Project): Option[] => p.cvs.map((c) => ({ value: c.id, label: c.label })),
  routes: (p: Project): Option[] => p.routes.map((r) => ({ value: r.id, label: r.label })),
  pages: (p: Project): Option[] => p.webPages.map((w) => ({ value: w.id, label: w.title })),
  flows: (p: Project): Option[] => p.flows.map((f) => ({ value: f.id, label: f.label })),
  surveys: (p: Project): Option[] => p.surveys.map((s) => ({ value: s.id, label: s.label })),
  richMenus: (p: Project): Option[] => p.richMenus.map((r) => ({ value: r.id, label: r.label })),
  assets: (p: Project): Option[] => p.assets.map((a) => ({ value: a.id, label: a.alt || a.path })),
  autoReplies: (p: Project): Option[] => p.autoReplies.map((a) => ({ value: a.id, label: a.match.keywords.join('・') || a.id })),
};

export const EVENT_OPTIONS: Option[] = [
  { value: 'tapButton', label: 'ボタンを押す' },
  { value: 'tapMenu', label: 'メニューを押す' },
  { value: 'sendText', label: 'メッセージを送る' },
  { value: 'answer', label: 'アンケートに答える' },
  { value: 'visitPage', label: 'ページを見る' },
  { value: 'submitForm', label: 'フォームを送信する' },
  { value: 'cv', label: 'CVする' },
];

// ---------- 時間 ----------

const UNITS: { value: string; label: string; minutes: number }[] = [
  { value: 'm', label: '分', minutes: 1 },
  { value: 'h', label: '時間', minutes: 60 },
  { value: 'd', label: '日', minutes: 1440 },
];

/** 分で保存し、画面では分・時間・日で入れる */
export function DurationInput({ value, onChange }: { value: Minutes; onChange: (v: Minutes) => void }) {
  const guess = value % 1440 === 0 && value > 0 ? 'd' : value % 60 === 0 && value > 0 ? 'h' : 'm';
  const [unit, setUnit] = useState(guess);
  const per = UNITS.find((u) => u.value === unit)!.minutes;
  return (
    <div className="row" style={{ flexWrap: 'nowrap' }}>
      <NumberInput value={value / per} min={0} onChange={(v) => onChange(Math.round((v ?? 0) * per))} />
      <select value={unit} onChange={(e) => setUnit(e.target.value)} style={{ width: 90 }}>
        {UNITS.map((u) => (
          <option key={u.value} value={u.value}>
            {u.label}
          </option>
        ))}
      </select>
    </div>
  );
}

// ---------- 条件式 ----------

/** 条件式。Phase 2b は JSON で書き、読めたら日本語で確認できる（ビルダー型の入力は Phase 4） */
export function ExprField({ label, value, onChange, path, hint }: { label: string; value: Expr; onChange: (e: Expr) => void; path: string; hint?: string }) {
  const { project, issues } = useEditor();
  const L = new Labels(project);
  return (
    <Field label={label} wide issues={issuesAt(issues, path)} hint={hint ?? <ExprHelp />}>
      <JsonField<Expr>
        value={value}
        onChange={onChange}
        describe={(e) => {
          try {
            return `→ ${describeExpr(e, L)}`;
          } catch {
            return '→ （読み取れない条件です）';
          }
        }}
      />
    </Field>
  );
}

function ExprHelp() {
  const [open, setOpen] = useState(false);
  return (
    <span>
      <button type="button" className="btn small ghost" onClick={() => setOpen((v) => !v)}>
        {open ? '書き方を閉じる' : '書き方の例'}
      </button>
      {open && (
        <pre className="mono" style={{ whiteSpace: 'pre-wrap', margin: '4px 0 0' }}>
          {`タグが付いている　　{"op":"hasTag","tagId":"area_tokyo"}
応募していない　　　{"op":"cv","cvId":"cv_apply","done":false}
回答が「東京」　　　{"op":"answer","surveyId":"sv_hope","questionId":"q_area","match":"any","choiceIds":["c_tokyo"]}
AかつB　　　　　　 {"op":"and","children":[ A , B ]}
AまたはB　　　　　 {"op":"or","children":[ A , B ]}
Aではない　　　　　{"op":"not","child": A }
常に　　　　　　　 {"op":"always"}`}
        </pre>
      )}
    </span>
  );
}

// ---------- 動作（ボタン・メニュー・ポップアップ） ----------

const ACTION_KINDS: Option[] = [
  { value: 'sendText', label: 'トークに文字を送る' },
  { value: 'openPage', label: 'Webページを開く' },
  { value: 'startFlow', label: 'シナリオを始める' },
  { value: 'answer', label: 'アンケートに答える' },
  { value: 'addFriend', label: '友だち追加する' },
  { value: 'switchMenu', label: 'メニューを切り替える' },
  { value: 'none', label: '何もしない' },
];

function emptyAction(kind: string): Action {
  switch (kind) {
    case 'sendText':
      return { kind, text: '' };
    case 'openPage':
      return { kind, pageId: '' };
    case 'startFlow':
      return { kind, flowId: '' };
    case 'answer':
      return { kind, surveyId: '', questionId: '', choiceId: '' };
    case 'addFriend':
      return { kind, routeId: '' };
    case 'switchMenu':
      return { kind, richMenuId: '' };
    default:
      return { kind: 'none' };
  }
}

export function ActionEditor({ value, onChange, path }: { value: Action; onChange: (a: Action) => void; path: string }) {
  const { project, issues } = useEditor();
  const at = (k: string) => issuesAt(issues, `${path}/${k}`);
  const a = value;
  return (
    <div className="fields">
      <Field label="押したときの動作">
        <Select value={a.kind} options={ACTION_KINDS} onChange={(k) => onChange(emptyAction(k))} />
      </Field>
      {a.kind === 'sendText' && (
        <Field label="送る文字" issues={at('text')} hint="自動応答のキーワードと合わせると、押したときに返信が返ります">
          <Text value={a.text} onChange={(text) => onChange({ ...a, text })} />
        </Field>
      )}
      {a.kind === 'openPage' && (
        <Field label="開くページ" issues={at('pageId')}>
          <Select value={a.pageId} empty="選んでください" options={options.pages(project)} onChange={(pageId) => onChange({ ...a, pageId })} />
        </Field>
      )}
      {a.kind === 'startFlow' && (
        <Field label="始めるシナリオ" issues={at('flowId')}>
          <Select value={a.flowId} empty="選んでください" options={options.flows(project)} onChange={(flowId) => onChange({ ...a, flowId })} />
        </Field>
      )}
      {a.kind === 'addFriend' && (
        <Field label="友だち追加の経路" issues={at('routeId')}>
          <Select value={a.routeId} empty="選んでください" options={options.routes(project)} onChange={(routeId) => onChange({ ...a, routeId })} />
        </Field>
      )}
      {a.kind === 'switchMenu' && (
        <Field label="切り替え先のメニュー" issues={at('richMenuId')}>
          <Select value={a.richMenuId} empty="選んでください" options={options.richMenus(project)} onChange={(richMenuId) => onChange({ ...a, richMenuId })} />
        </Field>
      )}
      {a.kind === 'answer' && <AnswerPicker value={a} onChange={onChange} path={path} />}
    </div>
  );
}

function AnswerPicker({ value, onChange, path }: { value: Extract<Action, { kind: 'answer' }>; onChange: (a: Action) => void; path: string }) {
  const { project, issues } = useEditor();
  const survey = project.surveys.find((s) => s.id === value.surveyId);
  const q = survey?.questions.find((x) => x.id === value.questionId);
  return (
    <>
      <Field label="アンケート">
        <Select value={value.surveyId} empty="選んでください" options={options.surveys(project)} onChange={(surveyId) => onChange({ ...value, surveyId, questionId: '', choiceId: '' })} />
      </Field>
      <Field label="質問" issues={issuesAt(issues, `${path}/questionId`)}>
        <Select
          value={value.questionId}
          empty="選んでください"
          options={(survey?.questions ?? []).map((x) => ({ value: x.id, label: x.text }))}
          onChange={(questionId) => onChange({ ...value, questionId, choiceId: '' })}
        />
      </Field>
      <Field label="選択肢" issues={issuesAt(issues, `${path}/choiceId`)}>
        <Select value={value.choiceId} empty="選んでください" options={(q?.choices ?? []).map((c) => ({ value: c.id, label: c.label }))} onChange={(choiceId) => onChange({ ...value, choiceId })} />
      </Field>
    </>
  );
}

// ---------- 入口（トリガー） ----------

const TRIGGER_KINDS: Option[] = [
  { value: 'friendAdd', label: '友だち追加' },
  { value: 'manual', label: '別のシナリオから呼ばれる' },
  { value: 'cv', label: 'CVした' },
  { value: 'keyword', label: '自動応答に反応' },
  { value: 'menuTap', label: 'メニューを押した' },
  { value: 'pageExit', label: 'ページから離脱しようとした' },
  { value: 'formSubmit', label: 'フォームを送信した' },
  { value: 'tagAdded', label: 'タグが付いた' },
];

function emptyTrigger(kind: string): Trigger {
  switch (kind) {
    case 'friendAdd':
      return { kind };
    case 'cv':
      return { kind, cvId: '' };
    case 'keyword':
      return { kind, autoReplyId: '' };
    case 'menuTap':
      return { kind, richMenuId: '', cellId: '' };
    case 'pageExit':
    case 'formSubmit':
      return { kind, pageId: '' };
    case 'tagAdded':
      return { kind, tagId: '' };
    default:
      return { kind: 'manual' };
  }
}

export function TriggerEditor({ value, onChange, path }: { value: Trigger; onChange: (t: Trigger) => void; path: string }) {
  const { project, issues } = useEditor();
  const at = (k: string) => issuesAt(issues, `${path}/${k}`);
  const t = value;
  return (
    <div className="fields">
      <Field label="いつ始まるか">
        <Select value={t.kind} options={TRIGGER_KINDS} onChange={(k) => onChange(emptyTrigger(k))} />
      </Field>
      {t.kind === 'friendAdd' && (
        <Field label="経路（どれも選ばなければ、すべての経路）" issues={at('routeIds')}>
          <Checks
            value={t.routeIds ?? []}
            options={options.routes(project)}
            onChange={(ids) => onChange(ids.length ? { kind: 'friendAdd', routeIds: ids } : { kind: 'friendAdd' })}
          />
        </Field>
      )}
      {t.kind === 'cv' && (
        <Field label="CV" issues={at('cvId')}>
          <Select value={t.cvId} empty="選んでください" options={options.cvs(project)} onChange={(cvId) => onChange({ ...t, cvId })} />
        </Field>
      )}
      {t.kind === 'keyword' && (
        <Field label="自動応答" issues={at('autoReplyId')}>
          <Select value={t.autoReplyId} empty="選んでください" options={options.autoReplies(project)} onChange={(autoReplyId) => onChange({ ...t, autoReplyId })} />
        </Field>
      )}
      {t.kind === 'menuTap' && (
        <>
          <Field label="メニュー" issues={at('richMenuId')}>
            <Select value={t.richMenuId} empty="選んでください" options={options.richMenus(project)} onChange={(richMenuId) => onChange({ ...t, richMenuId, cellId: '' })} />
          </Field>
          <Field label="区画" issues={at('cellId')}>
            <Select
              value={t.cellId}
              empty="選んでください"
              options={(project.richMenus.find((r) => r.id === t.richMenuId)?.cells ?? []).map((c) => ({ value: c.id, label: c.label }))}
              onChange={(cellId) => onChange({ ...t, cellId })}
            />
          </Field>
        </>
      )}
      {(t.kind === 'pageExit' || t.kind === 'formSubmit') && (
        <Field label="ページ" issues={at('pageId')}>
          <Select value={t.pageId} empty="選んでください" options={options.pages(project)} onChange={(pageId) => onChange({ ...t, pageId })} />
        </Field>
      )}
      {t.kind === 'tagAdded' && (
        <Field label="タグ" issues={at('tagId')}>
          <Select value={t.tagId} empty="選んでください" options={options.tags(project)} onChange={(tagId) => onChange({ ...t, tagId })} />
        </Field>
      )}
    </div>
  );
}

export function IssuesFor({ path }: { path: string }) {
  const { issues } = useEditor();
  const mine = issuesAt(issues, path);
  return mine.length ? <InlineIssues issues={mine} /> : null;
}

export type { InputEventType, Id };
