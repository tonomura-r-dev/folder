// 文面の差し込み（{{user.name}} 等）。未定義のキーはそのまま残す。
// {{キー|既定の文字}} と書くと、値が無いとき（未回答など）は既定の文字になる。

import type { Bubble, Project, UserState } from './schema.ts';

export const TEXT_VAR = /\{\{\s*([\w.-]+)\s*(?:\|([^{}]*?))?\s*\}\}/g;

/** 差し込みとして解決できるキーか（検証で使う。値の有無は問わない） */
export function isKnownVar(key: string, project: Project): boolean {
  const parts = key.split('.');
  if (key === 'user.name' || key === 'brand.accountName' || key === 'brand.companyName') return true;
  if (parts[0] === 'user' && parts[1] === 'attr' && parts.length === 3) return true;
  if (parts[0] === 'answer' && parts.length === 3) {
    const survey = project.surveys.find((s) => s.id === parts[1]);
    return !!survey && survey.questions.some((q) => q.id === parts[2]);
  }
  // form.<項目ID>：いちばん最近に送信したフォームの、その項目の値（予約した日時など）
  if (parts[0] === 'form' && parts.length === 2) return project.webPages.some((p) => p.form?.fields.some((f) => f.id === parts[1]));
  return false;
}

function resolve(key: string, project: Project, user: UserState): string | undefined {
  const parts = key.split('.');
  if (key === 'user.name') return user.personName ?? user.name;
  if (key === 'brand.accountName') return project.brand.accountName;
  if (key === 'brand.companyName') return project.brand.companyName;
  if (parts[0] === 'user' && parts[1] === 'attr' && parts[2] !== undefined) {
    const v = user.attrs[parts[2]];
    return v === undefined || v === null ? undefined : String(v);
  }
  if (parts[0] === 'answer' && parts[1] !== undefined && parts[2] !== undefined) {
    const got = user.answers[parts[1]]?.[parts[2]];
    if (!got) return undefined;
    const q = project.surveys.find((s) => s.id === parts[1])?.questions.find((x) => x.id === parts[2]);
    return got.map((id) => q?.choices.find((c) => c.id === id)?.label ?? id).join(' / ');
  }
  if (parts[0] === 'form' && parts[1] !== undefined) {
    for (let i = user.behaviors.length - 1; i >= 0; i--) {
      const b = user.behaviors[i]!;
      const v = b.type === 'submitForm' ? b.detail[`values.${parts[1]}`] : undefined;
      if (v !== undefined) return v;
    }
    return undefined;
  }
  return undefined;
}

export function renderText(text: string, project: Project, user: UserState): string {
  return text.replace(TEXT_VAR, (whole, key: string, fallback?: string) => resolve(key, project, user) ?? fallback ?? whole);
}

export function renderBubble(bubble: Bubble, project: Project, user: UserState): Bubble {
  const r = (t: string) => renderText(t, project, user);
  switch (bubble.kind) {
    case 'text':
      return { kind: 'text', text: r(bubble.text) };
    case 'image':
      return bubble;
    case 'buttons':
      return { ...bubble, text: r(bubble.text) };
    case 'card':
      return bubble.body === undefined
        ? { ...bubble, title: r(bubble.title) }
        : { ...bubble, title: r(bubble.title), body: r(bubble.body) };
    case 'carousel':
      return {
        kind: 'carousel',
        cards: bubble.cards.map((c) => ({
          ...c,
          title: r(c.title),
          ...(c.body === undefined ? {} : { body: r(c.body) }),
          ...(c.facts ? { facts: c.facts.map((f) => ({ label: f.label, value: r(f.value) })) } : {}),
        })),
      };
  }
}

/** 吹き出しに含まれる文をすべて取り出す（検証用） */
export function bubbleTexts(bubble: Bubble): string[] {
  switch (bubble.kind) {
    case 'text':
      return [bubble.text];
    case 'image':
      return [];
    case 'buttons':
      return [bubble.text];
    case 'card':
      return bubble.body === undefined ? [bubble.title] : [bubble.title, bubble.body];
    case 'carousel':
      return bubble.cards.flatMap((c) => [c.title, ...(c.body === undefined ? [] : [c.body]), ...(c.facts ?? []).map((f) => f.value)]);
  }
}
