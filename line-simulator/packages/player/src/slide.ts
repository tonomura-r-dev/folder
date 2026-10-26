// 資料用 16:9 表示。1280×720 の固定キャンバスに、スマホ・説明・いまの状態・なぜ を大きめの文字で並べる。
// 画面の幅に合わせて縮小するので、そのまま切り取って16:9のスライドに貼れる（スクショ機能は作らない）。

import type { UserState } from '../../engine/src/index.ts';
import { narrationCard, recentExplanations } from './backstage.ts';
import { fmtTime } from './describe.ts';
import { phone } from './phone.ts';
import { esc, type View } from './view.ts';

export const SLIDE_W = 1280;
export const SLIDE_H = 720;

export function slide(v: View, user: UserState): string {
  const users = v.demo.session.users
    .map((u) => v.world.users[u.id])
    .filter((u): u is UserState => !!u)
    .map((u) => `<span class="lsp-chip${u.id === user.id ? ' on' : ''}">${esc(u.name)}</span>`)
    .join('');
  const tags = Object.keys(user.tags).map((t) => `<span class="lsp-tag">${esc(v.L.tag(t))}</span>`).join('');
  const whys = recentExplanations(v, user, 3)
    .map((x) => `<li class="${x.tone}"><span class="mark">${x.tone === 'ok' ? '✓' : x.tone === 'ng' ? '✗' : x.tone === 'stop' ? '■' : '→'}</span><span>${esc(x.title)}${x.reason ? `<small>${esc(x.reason)}</small>` : ''}</span></li>`)
    .join('');
  return `<div class="lsp-slide-wrap"><div class="lsp-slide" style="width:${SLIDE_W}px;height:${SLIDE_H}px">
    <div class="lsp-slide-left">${phone(v, user)}</div>
    <div class="lsp-slide-right">
      <div class="lsp-slide-head"><span class="acc">${esc(v.project.brand.accountName)}</span><span class="clock">${esc(fmtTime(v.world.now))}</span></div>
      ${users ? `<div class="lsp-slide-users">${users}</div>` : ''}
      ${narrationCard(v, user, true)}
      <div class="lsp-slide-box"><h4>${esc(user.name)}のいまの状態</h4>${tags ? `<div class="lsp-tags">${tags}</div>` : '<div class="lsp-none">タグはまだ付いていません</div>'}</div>
      <div class="lsp-slide-box"><h4>なぜ、こうなったか</h4>${whys ? `<ul class="lsp-slide-why">${whys}</ul>` : '<div class="lsp-none">まだ判定はありません</div>'}</div>
    </div>
    <div class="lsp-slide-foot">架空のデータによるデモです</div>
  </div></div>`;
}
