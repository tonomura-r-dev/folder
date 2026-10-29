// カルーセル（横に並ぶカード）：検証・ボタン・見たカードの記録・画面の表示。

import { describe, expect, it } from 'vitest';
import type { MessageDef, Project } from '../packages/engine/src/index.ts';
import { validate } from '../packages/engine/src/index.ts';
import { dispatch, newDemo, worldOf } from '../packages/player/src/controller.ts';
import { Labels } from '../packages/player/src/describe.ts';
import { phone } from '../packages/player/src/phone.ts';
import { buildTimeline } from '../packages/player/src/timeline.ts';
import { baseProject, chain, codes, node } from './helpers.ts';

const carouselMsg: MessageDef = {
  id: 'm_car',
  bubbles: [
    {
      kind: 'carousel',
      cards: [
        { id: 'c1', title: 'カードA', facts: [{ label: '場所', value: '北' }], buttons: [{ id: 'b1', label: '詳しく', action: { kind: 'openPage', pageId: 'pg' } }] },
        { id: 'c2', title: 'カードB', buttons: [{ id: 'b2', label: '詳しく', action: { kind: 'openPage', pageId: 'pg' } }] },
        { id: 'c3', title: 'カードC', buttons: [{ id: 'b3', label: '詳しく', action: { kind: 'openPage', pageId: 'pg' } }] },
      ],
    },
  ],
};

function project(): Project {
  return baseProject({
    messages: [structuredClone(carouselMsg)],
    webPages: [{ id: 'pg', kind: 'lp', title: '詳細', blocks: [{ kind: 'heading', text: '詳細' }] }],
    flows: [chain('f1', [node({ id: 'n2', type: 'message', messageId: 'm_car' }), node({ id: 'n3', type: 'end' })])],
  });
}

const view = (p: Project, d: ReturnType<typeof newDemo>) => {
  const world = worldOf(p, d);
  return { project: p, world, demo: d, ui: { menuOpen: true, picked: [], mode: 'three', tab: 'state', open: new Set<number>(), onlyNg: false, playing: false, speed: 1, status: '', fcZoom: 1, fcOptions: true, fcOpen: null, help: false }, L: new Labels(p), assetUrl: () => '' } as never;
};

describe('カルーセル', () => {
  it('検証：カードのボタンIDの重複・空のカルーセル・見たカードの台本の参照先を確かめる', () => {
    const p = project();
    expect(codes(validate(p), 'error')).toEqual([]);
    const bad = project();
    const b = bad.messages[0]!.bubbles[0]!;
    if (b.kind !== 'carousel') throw new Error('carousel');
    b.cards[1]!.buttons[0]!.id = 'b1';
    expect(codes(validate(bad), 'error')).toContain('id.duplicate');
    b.cards.length = 0;
    expect(codes(validate(bad), 'error')).toContain('node.requiredMissing');
    const script = project();
    script.personas[0]!.script = [{ atMinutes: 600, event: { type: 'viewCard', messageId: 'm_car', cardId: 'nope' } }];
    expect(codes(validate(script), 'error')).toContain('ref.missing');
  });

  it('カードのボタンを押すと動作し、見たカードは行動履歴と③の時系列に残る', () => {
    const p = project();
    let d = dispatch(newDemo(p, 'u1'), { type: 'addFriend', routeId: 'r1' });
    d = dispatch(d, { type: 'viewCard', messageId: 'm_car', cardId: 'c2' });
    d = dispatch(d, { type: 'tapButton', buttonId: 'b2' });
    const w = worldOf(p, d);
    const u = w.users[d.activeUserId]!;
    expect(u.surface).toEqual({ kind: 'page', pageId: 'pg' });
    expect(u.behaviors.find((b) => b.type === 'viewCard')!.detail).toEqual({ messageId: 'm_car', cardId: 'c2' });
    const fronts = buildTimeline(p, new Labels(p), w, [u.id]).flatMap((e) => e.front);
    expect(fronts).toContain('カルーセルのカード「カードB」を見る');
  });

  it('①では、最後に見たカードが表示され、左右のボタンでその隣のカードへ送れる', () => {
    const p = project();
    let d = dispatch(newDemo(p, 'u1'), { type: 'addFriend', routeId: 'r1' });
    let html = phone(view(p, d), worldOf(p, d).users[d.activeUserId]!);
    expect(html).toContain('1 / 3');
    expect(html).toContain('data-act="carGo" data-msg="m_car" data-card="c2" aria-label="次のカード"');
    d = dispatch(d, { type: 'viewCard', messageId: 'm_car', cardId: 'c3' });
    html = phone(view(p, d), worldOf(p, d).users[d.activeUserId]!);
    expect(html).toContain('3 / 3');
    expect(html).toContain('--i:2');
    expect(html).toMatch(/class="lsp-car-next" disabled/);
  });

  it('カードの文面にも差し込み（{{user.name}}）が効く', () => {
    const p = project();
    const b = p.messages[0]!.bubbles[0]!;
    if (b.kind === 'carousel') b.cards[0]!.title = '{{user.name}}さんへ';
    const d =dispatch(newDemo(p, 'u1'), { type: 'addFriend', routeId: 'r1' });
    const talk = worldOf(p, d).users[d.activeUserId]!.talk.find((t) => t.side === 'bot');
    expect(talk?.side === 'bot' && talk.bubble.kind === 'carousel' && talk.bubble.cards[0]!.title).toBe('u1さんへ');
  });
});
