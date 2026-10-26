// 案件設定：会社名・アカウント名・配色・重複防止・画像・社内メモ。

import { issuesAt, useEditor } from '../ctx.ts';
import { Field, InlineIssues, NumberInput, ReadOnly, Text, TextArea } from '../ui/form.tsx';
import { Images } from './Images.tsx';

export function Settings() {
  const { project: p, update, issues } = useEditor();
  return (
    <div className="detail" style={{ width: '100%' }}>
      <h2>案件設定</h2>
      <div className="section">
        <h3>案件</h3>
        <div className="fields">
          <Field label="案件名（一覧に出る名前）">
            <Text value={p.meta.name} onChange={(v) => update((d) => (d.meta.name = v))} />
          </Field>
          <Field label="業種（メモ）">
            <Text value={p.meta.industry ?? ''} onChange={(v) => update((d) => (v ? (d.meta.industry = v) : delete d.meta.industry))} />
          </Field>
          <Field label="フォルダ名">
            <ReadOnly>{p.meta.slug}</ReadOnly>
          </Field>
        </div>
      </div>

      <div className="section">
        <h3>LINEアカウント</h3>
        <div className="fields">
          <Field label="会社名" hint="架空または匿名化した名前にしてください（営業用HTMLに出ます）">
            <Text value={p.brand.companyName} onChange={(v) => update((d) => (d.brand.companyName = v))} />
          </Field>
          <Field label="アカウント名（トーク画面の上に出る名前）">
            <Text value={p.brand.accountName} onChange={(v) => update((d) => (d.brand.accountName = v))} />
          </Field>
          <ColorField label="メインの色" value={p.brand.colors.primary} onChange={(v) => update((d) => (d.brand.colors.primary = v))} />
          <ColorField label="強調の色" value={p.brand.colors.accent} onChange={(v) => update((d) => (d.brand.colors.accent = v))} />
          <ColorField label="トーク画面の背景色" value={p.brand.colors.chatBg ?? '#7C9BD0'} onChange={(v) => update((d) => (d.brand.colors.chatBg = v))} />
        </div>
      </div>

      <div className="section">
        <h3>配信</h3>
        <div className="fields">
          <Field label="重複配信を防ぐ期間（時間）" hint="この時間内に同じ文面を同じ人へ2回送りません。0 で防止しない">
            <NumberInput value={p.settings.dedupeWindowMinutes / 60} min={0} onChange={(v) => update((d) => (d.settings.dedupeWindowMinutes = Math.round((v ?? 0) * 60)))} />
          </Field>
          <Field label="無料メッセージの通数（任意・目安）" hint="デモの配信履歴に「無料○通のうち△通」と出します。契約プランの数を入れてください。空なら出しません">
            <NumberInput
              value={p.settings.messageCounting?.freeQuota}
              min={0}
              onChange={(v) => update((d) => (d.settings.messageCounting = v === undefined ? { enabled: true } : { enabled: true, freeQuota: v }))}
            />
          </Field>
        </div>
      </div>

      <Images />

      <div className="section">
        <h3>社内メモ</h3>
        <Field label="社内メモ（営業用HTMLには入りません）" wide>
          <TextArea
            value={typeof p._internal?.memo === 'string' ? p._internal.memo : ''}
            onChange={(v) => update((d) => (d._internal = { ...d._internal, memo: v }))}
          />
        </Field>
      </div>
      <InlineIssues issues={issues.filter((i) => i.path.startsWith('/brand') || i.path.startsWith('/settings'))} />
    </div>
  );
}

function ColorField({ label, value, onChange }: { label: string; value: string; onChange: (v: string) => void }) {
  return (
    <Field label={label}>
      <div className="row" style={{ flexWrap: 'nowrap' }}>
        <input type="color" value={/^#[0-9a-f]{6}$/i.test(value) ? value : '#000000'} onChange={(e) => onChange(e.target.value)} style={{ width: 44, height: 32, padding: 0, border: 0 }} />
        <Text value={value} onChange={onChange} />
      </div>
    </Field>
  );
}
