// 案件の検証。重大エラー（error）が1つでもあれば営業用HTMLは書き出さない。
// path は JSON Pointer。画面側はこれで該当箇所へ飛び、code から日本語の説明を作る。

import type { Action, Expr, Flow, FlowNode, Id, InputEventBody, Issue, Project, Wait, WebBlock } from './schema.ts';
import { choiceButtons, portsOf } from './ports.ts';
import { bubbleTexts, isKnownVar, TEXT_VAR } from './text.ts';

export type ValidateOptions = {
  /** 画像ファイルの存在確認。エンジンはファイルを読まないので外から渡す */
  assetExists?: (path: string) => boolean;
};

/** LINE側の上限（警告に使う）。変わったらここだけ直す */
export const LINE_LIMITS = { bubblesPerMessage: 5, buttonsPerBubble: 4, textLength: 5000, stepDays: 365, menuCells: 20, carouselCards: 10 };

const EXPR_OPS = new Set(['and', 'or', 'not', 'always', 'hasTag', 'answer', 'answered', 'attr', 'route', 'friend', 'cv', 'did', 'elapsed', 'delivered']);
const CMPS = new Set(['eq', 'neq', 'in', 'gte', 'lte']);
const FEASIBILITY = new Set(['lstep', 'conditional', 'external', 'unknown']);
const EVENT_TYPES = new Set([
  'visitPage', 'exitIntent', 'dismissPopup', 'submitForm', 'addFriend', 'block', 'unblock', 'tapMenu', 'tapButton', 'sendText', 'answer', 'cv',
  'closePage', 'tapPopupCta', 'tapPageButton', 'tapLineCta', 'viewCard', 'filterList',
]);
const TIME = /^([01]?\d|2[0-3]):[0-5]\d$/;

export function validate(project: Project, options: ValidateOptions = {}): Issue[] {
  const issues: Issue[] = [];
  const err = (code: string, path: string, params?: Record<string, unknown>) =>
    issues.push(params ? { level: 'error', code, path, params } : { level: 'error', code, path });
  const warn = (code: string, path: string, params?: Record<string, unknown>) =>
    issues.push(params ? { level: 'warning', code, path, params } : { level: 'warning', code, path });

  // ---------- 索引 ----------
  const ids = <T extends { id: Id }>(list: T[]) => new Set(list.map((x) => x.id));
  const tags = ids(project.tags);
  const cvs = ids(project.cvs);
  const routes = ids(project.routes);
  const assets = ids(project.assets);
  const messages = new Map(project.messages.map((m) => [m.id, m]));
  const richMenus = new Map(project.richMenus.map((r) => [r.id, r]));
  const surveys = new Map(project.surveys.map((s) => [s.id, s]));
  const pages = ids(project.webPages);
  const flows = new Map(project.flows.map((f) => [f.id, f]));
  const autoReplies = ids(project.autoReplies);
  const allNodes = new Map<Id, { flow: Flow; node: FlowNode; path: string }>();
  const buttons = new Set<Id>();

  const tagsAdded = new Set<Id>();
  const tagsRead = new Set<Id>();
  const startedFlows = new Set<Id>();

  // ---------- ID の重複 ----------
  const unique = (list: { id: Id }[], path: string) => {
    const seen = new Set<Id>();
    list.forEach((x, i) => {
      if (seen.has(x.id)) err('id.duplicate', `${path}/${i}`, { id: x.id });
      seen.add(x.id);
    });
  };
  for (const key of ['assets', 'tags', 'cvs', 'routes', 'messages', 'richMenus', 'autoReplies', 'surveys', 'webPages', 'flows', 'personas', 'narration'] as const)
    unique(project[key], `/${key}`);
  unique(project.journey ?? [], '/journey');
  project.flows.forEach((flow, fi) =>
    flow.nodes.forEach((node, ni) => {
      const path = `/flows/${fi}/nodes/${ni}`;
      if (allNodes.has(node.id)) err('id.duplicate', path, { id: node.id });
      else allNodes.set(node.id, { flow, node, path });
    }),
  );
  const seeButton = (id: Id, path: string) => {
    if (buttons.has(id)) err('id.duplicate', path, { id });
    buttons.add(id);
  };
  project.messages.forEach((m, mi) =>
    m.bubbles.forEach((b, bi) => {
      if (b.kind === 'buttons' || b.kind === 'card') b.buttons.forEach((btn, i) => seeButton(btn.id, `/messages/${mi}/bubbles/${bi}/buttons/${i}`));
      if (b.kind === 'carousel') {
        unique(b.cards, `/messages/${mi}/bubbles/${bi}/cards`);
        b.cards.forEach((c, ci) => c.buttons.forEach((btn, i) => seeButton(btn.id, `/messages/${mi}/bubbles/${bi}/cards/${ci}/buttons/${i}`)));
      }
    }),
  );

  // ---------- 参照 ----------
  const need = (ok: boolean, path: string, kind: string, id: unknown) => {
    if (!ok) err('ref.missing', path, { kind, id });
  };
  const questionOf = (surveyId: Id, questionId: Id) => surveys.get(surveyId)?.questions.find((q) => q.id === questionId);

  const checkAction = (a: Action, path: string) => {
    switch (a.kind) {
      case 'openPage':
        return need(pages.has(a.pageId), `${path}/pageId`, 'page', a.pageId);
      case 'startFlow':
        startedFlows.add(a.flowId);
        return need(flows.has(a.flowId), `${path}/flowId`, 'flow', a.flowId);
      case 'answer': {
        const q = questionOf(a.surveyId, a.questionId);
        need(!!q, `${path}/questionId`, 'question', `${a.surveyId}/${a.questionId}`);
        if (q) need(q.choices.some((c) => c.id === a.choiceId), `${path}/choiceId`, 'choice', a.choiceId);
        return;
      }
      case 'addFriend':
        return need(routes.has(a.routeId), `${path}/routeId`, 'route', a.routeId);
      case 'switchMenu':
        return need(richMenus.has(a.richMenuId), `${path}/richMenuId`, 'richMenu', a.richMenuId);
      case 'sendText':
        if (!a.text.trim()) err('node.requiredMissing', `${path}/text`);
        return;
      case 'backToLine':
      case 'none':
        return;
      default:
        err('node.requiredMissing', path, { field: 'action' });
    }
  };

  const checkExpr = (e: Expr, path: string, depth = 0): void => {
    if (!e || typeof e !== 'object' || !EXPR_OPS.has((e as { op: string }).op) || depth > 32)
      return void err('expr.invalid', path, { reason: 'op' });
    switch (e.op) {
      case 'and':
      case 'or':
        if (!Array.isArray(e.children) || e.children.length === 0) return void err('expr.invalid', path, { reason: 'empty' });
        return e.children.forEach((c, i) => checkExpr(c, `${path}/children/${i}`, depth + 1));
      case 'not':
        return checkExpr(e.child, `${path}/child`, depth + 1);
      case 'always':
        return;
      case 'hasTag':
        tagsRead.add(e.tagId);
        return need(tags.has(e.tagId), `${path}/tagId`, 'tag', e.tagId);
      case 'answer': {
        const q = questionOf(e.surveyId, e.questionId);
        need(!!q, `${path}/questionId`, 'question', `${e.surveyId}/${e.questionId}`);
        if (!Array.isArray(e.choiceIds) || e.choiceIds.length === 0) err('expr.invalid', `${path}/choiceIds`, { reason: 'empty' });
        else if (q) e.choiceIds.forEach((c, i) => need(q.choices.some((x) => x.id === c), `${path}/choiceIds/${i}`, 'choice', c));
        if (!['any', 'all', 'none'].includes(e.match)) err('expr.invalid', `${path}/match`, { reason: 'match' });
        return;
      }
      case 'answered':
        need(surveys.has(e.surveyId), `${path}/surveyId`, 'survey', e.surveyId);
        if (e.questionId !== undefined) need(!!questionOf(e.surveyId, e.questionId), `${path}/questionId`, 'question', e.questionId);
        return;
      case 'attr':
        if (!e.key) err('expr.invalid', `${path}/key`, { reason: 'key' });
        if (!CMPS.has(e.cmp)) err('expr.invalid', `${path}/cmp`, { reason: 'cmp' });
        if ((e.cmp === 'gte' || e.cmp === 'lte') && typeof e.value !== 'number') err('expr.invalid', `${path}/value`, { reason: 'type' });
        if (e.cmp === 'in' && !Array.isArray(e.value)) err('expr.invalid', `${path}/value`, { reason: 'type' });
        return;
      case 'route':
        return e.routeIds.forEach((r, i) => need(routes.has(r), `${path}/routeIds/${i}`, 'route', r));
      case 'friend':
        if (!['not_friend', 'friend', 'blocked'].includes(e.state)) err('expr.invalid', `${path}/state`, { reason: 'state' });
        return;
      case 'cv':
        return need(cvs.has(e.cvId), `${path}/cvId`, 'cv', e.cvId);
      case 'did':
        if (!EVENT_TYPES.has(e.event)) err('expr.invalid', `${path}/event`, { reason: 'event' });
        if (e.cmp === 'countGte' && typeof e.value !== 'number') err('expr.invalid', `${path}/value`, { reason: 'type' });
        return;
      case 'elapsed':
        if (typeof e.minutes !== 'number' || e.minutes < 0) err('expr.invalid', `${path}/minutes`, { reason: 'type' });
        return;
      case 'delivered': {
        const t = e.target;
        if (!t || (t.messageId === undefined && t.flowId === undefined && t.nodeId === undefined))
          return void err('expr.invalid', `${path}/target`, { reason: 'empty' });
        if (t.messageId !== undefined) need(messages.has(t.messageId), `${path}/target/messageId`, 'message', t.messageId);
        if (t.flowId !== undefined) need(flows.has(t.flowId), `${path}/target/flowId`, 'flow', t.flowId);
        if (t.nodeId !== undefined) need(allNodes.has(t.nodeId), `${path}/target/nodeId`, 'node', t.nodeId);
        return;
      }
    }
  };

  const checkText = (text: string | undefined, path: string) => {
    if (text === undefined) return;
    if (text.length > LINE_LIMITS.textLength) warn('line.limit', path, { limit: 'textLength', value: text.length });
    for (const m of text.matchAll(TEXT_VAR)) if (!isKnownVar(m[1]!, project)) warn('text.unknownVar', path, { key: m[1] });
  };

  // 案件全体
  if (project.brand.iconAssetId) need(assets.has(project.brand.iconAssetId), '/brand/iconAssetId', 'asset', project.brand.iconAssetId);
  if (project.settings.defaultRichMenuId)
    need(richMenus.has(project.settings.defaultRichMenuId), '/settings/defaultRichMenuId', 'richMenu', project.settings.defaultRichMenuId);
  const endDay = project.settings.demoEndDay;
  if (endDay !== undefined && (!Number.isInteger(endDay) || endDay < 1 || endDay > 365)) err('node.requiredMissing', '/settings/demoEndDay');
  (project.settings.demoEntries ?? []).forEach((e, i) => {
    need(project.webPages.some((w) => w.id === e.pageId), `/settings/demoEntries/${i}/pageId`, 'page', e.pageId);
    if (!e.label.trim()) err('node.requiredMissing', `/settings/demoEntries/${i}/label`);
  });
  const start = project.settings.demoStart;
  if (start) {
    if (start.mode !== 'single' && start.mode !== 'scripted') err('node.requiredMissing', '/settings/demoStart/mode');
    if (start.personaId) need(project.personas.some((p) => p.id === start.personaId), '/settings/demoStart/personaId', 'persona', start.personaId);
    if (start.mode === 'scripted' && !project.personas.some((p) => (p.script ?? []).length > 0))
      warn('optional.empty', '/settings/demoStart', { reason: 'noScript' });
  }

  project.assets.forEach((a, i) => {
    if (!a.alt.trim()) warn('optional.empty', `/assets/${i}/alt`);
    if (options.assetExists && !options.assetExists(a.path)) err('asset.fileMissing', `/assets/${i}/path`, { path: a.path });
  });

  project.messages.forEach((m, mi) => {
    const path = `/messages/${mi}`;
    if (m.bubbles.length === 0) err('node.requiredMissing', `${path}/bubbles`);
    if (m.bubbles.length > LINE_LIMITS.bubblesPerMessage) warn('line.limit', `${path}/bubbles`, { limit: 'bubblesPerMessage', value: m.bubbles.length });
    m.bubbles.forEach((b, bi) => {
      const bp = `${path}/bubbles/${bi}`;
      bubbleTexts(b).forEach((t) => checkText(t, bp));
      if (b.kind === 'text' && !b.text.trim()) err('node.requiredMissing', `${bp}/text`);
      if (b.kind === 'image') need(assets.has(b.assetId), `${bp}/assetId`, 'asset', b.assetId);
      if (b.kind === 'card' && b.assetId) need(assets.has(b.assetId), `${bp}/assetId`, 'asset', b.assetId);
      if (b.kind === 'buttons' || b.kind === 'card') {
        if (b.buttons.length === 0) err('node.requiredMissing', `${bp}/buttons`);
        if (b.buttons.length > LINE_LIMITS.buttonsPerBubble) warn('line.limit', `${bp}/buttons`, { limit: 'buttonsPerBubble', value: b.buttons.length });
        b.buttons.forEach((btn, i) => checkAction(btn.action, `${bp}/buttons/${i}/action`));
      }
      if (b.kind === 'carousel') {
        if (b.cards.length === 0) err('node.requiredMissing', `${bp}/cards`);
        if (b.cards.length > LINE_LIMITS.carouselCards) warn('line.limit', `${bp}/cards`, { limit: 'carouselCards', value: b.cards.length });
        b.cards.forEach((c, ci) => {
          const cp = `${bp}/cards/${ci}`;
          if (!c.title.trim()) err('node.requiredMissing', `${cp}/title`);
          if (c.assetId) need(assets.has(c.assetId), `${cp}/assetId`, 'asset', c.assetId);
          if (c.buttons.length === 0) err('node.requiredMissing', `${cp}/buttons`);
          if (c.buttons.length > LINE_LIMITS.buttonsPerBubble) warn('line.limit', `${cp}/buttons`, { limit: 'buttonsPerBubble', value: c.buttons.length });
          c.buttons.forEach((btn, i) => checkAction(btn.action, `${cp}/buttons/${i}/action`));
        });
      }
    });
  });

  project.richMenus.forEach((r, ri) => {
    const path = `/richMenus/${ri}`;
    if (r.imageAssetId) need(assets.has(r.imageAssetId), `${path}/imageAssetId`, 'asset', r.imageAssetId);
    if (r.cells.length === 0) err('node.requiredMissing', `${path}/cells`);
    if (r.cells.length > LINE_LIMITS.menuCells) warn('line.limit', `${path}/cells`, { limit: 'menuCells', value: r.cells.length });
    unique(r.cells, `${path}/cells`);
    r.cells.forEach((c, ci) => {
      const { col, row, colSpan = 1, rowSpan = 1 } = c.area;
      if (col < 0 || row < 0 || col + colSpan > r.layout.cols || row + rowSpan > r.layout.rows)
        err('node.requiredMissing', `${path}/cells/${ci}/area`, { reason: 'outOfBounds' });
      checkAction(c.action, `${path}/cells/${ci}/action`);
    });
  });

  project.autoReplies.forEach((a, ai) => {
    const path = `/autoReplies/${ai}`;
    if (a.match.keywords.filter((k) => k.trim()).length === 0) err('node.requiredMissing', `${path}/match/keywords`);
    if (a.cases.length === 0 && !a.fallbackMessageId) err('node.requiredMissing', `${path}/cases`);
    a.cases.forEach((c, ci) => {
      need(messages.has(c.messageId), `${path}/cases/${ci}/messageId`, 'message', c.messageId);
      checkExpr(c.when, `${path}/cases/${ci}/when`);
    });
    if (a.fallbackMessageId) need(messages.has(a.fallbackMessageId), `${path}/fallbackMessageId`, 'message', a.fallbackMessageId);
  });

  project.surveys.forEach((s, si) => {
    const path = `/surveys/${si}`;
    if (s.questions.length === 0) err('node.requiredMissing', `${path}/questions`);
    unique(s.questions, `${path}/questions`);
    s.questions.forEach((q, qi) => {
      checkText(q.text, `${path}/questions/${qi}/text`);
      if (q.choices.length === 0) err('node.requiredMissing', `${path}/questions/${qi}/choices`);
      unique(q.choices, `${path}/questions/${qi}/choices`);
      q.choices.forEach((c, ci) => {
        const cp = `${path}/questions/${qi}/choices/${ci}`;
        (c.addTags ?? []).forEach((t, i) => (tagsAdded.add(t), need(tags.has(t), `${cp}/addTags/${i}`, 'tag', t)));
        (c.removeTags ?? []).forEach((t, i) => (tagsRead.add(t), need(tags.has(t), `${cp}/removeTags/${i}`, 'tag', t)));
      });
    });
  });

  project.webPages.forEach((p, pi) => {
    const path = `/webPages/${pi}`;
    p.blocks.forEach((b, bi) => {
      const bp = `${path}/blocks/${bi}`;
      if (b.kind === 'image') need(assets.has(b.assetId), `${bp}/assetId`, 'asset', b.assetId);
      else if (b.kind === 'button') checkAction(b.action, `${bp}/action`);
      else if (b.kind === 'list') checkList(b, bp);
      else checkText(b.text, `${bp}/text`);
    });
    if (p.form) {
      if (p.form.fields.length === 0) err('node.requiredMissing', `${path}/form/fields`);
      p.form.fields.forEach((f, fi) => {
        if (f.closesAt && f.closesAt.length !== (f.options ?? []).length) err('node.requiredMissing', `${path}/form/fields/${fi}/closesAt`, { reason: 'closesAtLength' });
      });
      if (p.form.confirm) checkText(p.form.confirm.note, `${path}/form/confirm/note`);
      const s = p.form.onSubmit;
      if (s.cvId) need(cvs.has(s.cvId), `${path}/form/onSubmit/cvId`, 'cv', s.cvId);
      if (s.nextPageId) need(pages.has(s.nextPageId), `${path}/form/onSubmit/nextPageId`, 'page', s.nextPageId);
    }
    if (p.exitPopup) {
      checkText(p.exitPopup.title, `${path}/exitPopup/title`);
      checkText(p.exitPopup.body, `${path}/exitPopup/body`);
      if (p.exitPopup.assetId) need(assets.has(p.exitPopup.assetId), `${path}/exitPopup/assetId`, 'asset', p.exitPopup.assetId);
      checkAction(p.exitPopup.cta.action, `${path}/exitPopup/cta/action`);
    }
    if (p.lineCta) need(routes.has(p.lineCta.routeId), `${path}/lineCta/routeId`, 'route', p.lineCta.routeId);
    checkText(p.lineCta?.prefill, `${path}/lineCta/prefill`);
  });

  /** 絞り込みつきの一覧：軸・選択肢・項目のID、項目が指す軸と選択肢、押したときの動作 */
  function checkList(b: Extract<WebBlock, { kind: 'list' }>, path: string) {
    unique(b.facets, `${path}/facets`);
    unique(b.items, `${path}/items`);
    if (b.items.length === 0) err('node.requiredMissing', `${path}/items`);
    checkText(b.emptyText, `${path}/emptyText`);
    b.facets.forEach((f, fi) => {
      if (f.options.length === 0) err('node.requiredMissing', `${path}/facets/${fi}/options`);
      unique(f.options, `${path}/facets/${fi}/options`);
      f.options.forEach((o, oi) => {
        if (o.tagId === undefined) return;
        tagsRead.add(o.tagId);
        need(tags.has(o.tagId), `${path}/facets/${fi}/options/${oi}/tagId`, 'tag', o.tagId);
      });
    });
    b.items.forEach((it, ii) => {
      const ip = `${path}/items/${ii}`;
      if (!it.title.trim()) err('node.requiredMissing', `${ip}/title`);
      checkText(it.title, `${ip}/title`);
      checkText(it.body, `${ip}/body`);
      if (it.assetId) need(assets.has(it.assetId), `${ip}/assetId`, 'asset', it.assetId);
      if (it.closesAt !== undefined && !(Number.isInteger(it.closesAt) && it.closesAt >= 0)) err('node.requiredMissing', `${ip}/closesAt`);
      for (const [facetId, optionIds] of Object.entries(it.values)) {
        const f = b.facets.find((x) => x.id === facetId);
        need(!!f, `${ip}/values/${facetId}`, 'facet', facetId);
        if (f) optionIds.forEach((o, oi) => need(f.options.some((x) => x.id === o), `${ip}/values/${facetId}/${oi}`, 'option', o));
      }
      checkAction(it.action, `${ip}/action`);
    });
  }

  // ---------- CV ----------
  project.cvs.forEach((c, i) => {
    if (c.prompt) checkExpr(c.prompt.when, `/cvs/${i}/prompt/when`);
  });

  // ---------- フロー ----------
  project.flows.forEach((flow, fi) => {
    const path = `/flows/${fi}`;
    if (!flow.label.trim()) warn('optional.empty', `${path}/label`);
    checkTrigger(flow, path);
    if (flow.stopWhen) checkExpr(flow.stopWhen, `${path}/stopWhen`);

    const entries = flow.nodes.filter((n) => n.type === 'entry');
    if (entries.length !== 1) err('flow.noEntry', `${path}/nodes`, { count: entries.length });

    const local = new Map(flow.nodes.map((n, ni) => [n.id, { node: n, path: `${path}/nodes/${ni}` }]));
    flow.nodes.forEach((node, ni) => checkNode(node, `${path}/nodes/${ni}`));

    // エッジ
    const used = new Set<string>();
    unique(flow.edges, `${path}/edges`);
    flow.edges.forEach((e, ei) => {
      const ep = `${path}/edges/${ei}`;
      const from = local.get(e.from.nodeId);
      if (!from) return void err('edge.danglingSource', ep, { nodeId: e.from.nodeId });
      if (!local.has(e.to)) err('edge.danglingTarget', ep, { nodeId: e.to });
      else if (local.get(e.to)!.node.type === 'entry') err('edge.danglingTarget', ep, { nodeId: e.to, reason: 'entry' });
      if (!portsOf(from.node, project).includes(e.from.port)) err('edge.invalidPort', ep, { nodeId: e.from.nodeId, port: e.from.port });
      const key = `${e.from.nodeId}\u0000${e.from.port}`;
      if (used.has(key)) err('edge.invalidPort', ep, { nodeId: e.from.nodeId, port: e.from.port, reason: 'duplicate' });
      used.add(key);
    });

    // 分岐の行き場・行き止まり
    for (const [id, { node, path: np }] of local) {
      const ports = portsOf(node, project);
      if (node.type === 'branch' && !used.has(`${id}\u0000else`)) err('branch.noExit', np);
      else if (ports.length > 0 && !ports.some((p) => used.has(`${id}\u0000${p}`))) warn('node.deadEnd', np);
    }

    // 到達できないノード
    const entry = entries[0];
    if (entry) {
      const seen = new Set<Id>([entry.id]);
      const stack = [entry.id];
      while (stack.length) {
        const id = stack.pop()!;
        for (const e of flow.edges) if (e.from.nodeId === id && !seen.has(e.to)) (seen.add(e.to), stack.push(e.to));
      }
      for (const [id, { path: np }] of local) if (!seen.has(id)) warn('node.unreachable', np);
    }

    // 停止条件のないステップ配信
    if (!flow.stopWhen && flow.nodes.some((n) => n.type === 'wait' && n.wait.kind !== 'untilEvent')) warn('flow.noStop', path);

    // 友だちになる前に始まるフローが、待機を挟まずに配信しようとしている
    if (entry && (flow.trigger.kind === 'pageExit' || flow.trigger.kind === 'formSubmit')) {
      const seen = new Set<Id>();
      const stack = [entry.id];
      while (stack.length) {
        const id = stack.pop()!;
        if (seen.has(id)) continue;
        seen.add(id);
        const n = local.get(id)?.node;
        if (!n || n.type === 'wait') continue;
        if (n.type === 'message' || n.type === 'choice' || n.type === 'image' || n.type === 'survey') {
          warn('flow.friendRequired', local.get(id)!.path);
          continue;
        }
        for (const e of flow.edges) if (e.from.nodeId === id) stack.push(e.to);
      }
    }
  });

  function checkTrigger(flow: Flow, path: string) {
    const t = flow.trigger;
    const tp = `${path}/trigger`;
    switch (t.kind) {
      case 'friendAdd':
        return (t.routeIds ?? []).forEach((r, i) => need(routes.has(r), `${tp}/routeIds/${i}`, 'route', r));
      case 'keyword':
        return need(autoReplies.has(t.autoReplyId), `${tp}/autoReplyId`, 'autoReply', t.autoReplyId);
      case 'menuTap': {
        const menu = richMenus.get(t.richMenuId);
        need(!!menu, `${tp}/richMenuId`, 'richMenu', t.richMenuId);
        if (menu) need(menu.cells.some((c) => c.id === t.cellId), `${tp}/cellId`, 'cell', t.cellId);
        return;
      }
      case 'cv':
        return need(cvs.has(t.cvId), `${tp}/cvId`, 'cv', t.cvId);
      case 'pageExit':
      case 'formSubmit':
        return need(pages.has(t.pageId), `${tp}/pageId`, 'page', t.pageId);
      case 'tagAdded':
        tagsRead.add(t.tagId);
        return need(tags.has(t.tagId), `${tp}/tagId`, 'tag', t.tagId);
      case 'manual':
        return;
    }
  }

  function checkWait(w: Wait, path: string) {
    if (w.kind === 'duration' && (!Number.isInteger(w.minutes) || w.minutes < 0)) err('node.requiredMissing', `${path}/minutes`);
    if (w.kind === 'at') {
      if (!Number.isInteger(w.day) || w.day < 0) err('node.requiredMissing', `${path}/day`);
      if (!TIME.test(w.time)) err('node.requiredMissing', `${path}/time`);
    }
    if (w.kind === 'dayAt') {
      if (!Number.isInteger(w.day) || w.day < 0) err('node.requiredMissing', `${path}/day`);
      if (!TIME.test(w.time)) err('node.requiredMissing', `${path}/time`);
      if (w.day > LINE_LIMITS.stepDays) warn('line.limit', `${path}/day`, { limit: 'stepDays', value: w.day });
    }
    if (w.kind === 'untilEvent') {
      if (!EVENT_TYPES.has(w.event)) err('node.requiredMissing', `${path}/event`);
      if (!(w.timeoutMinutes > 0)) err('node.requiredMissing', `${path}/timeoutMinutes`);
    }
  }

  function checkNode(node: FlowNode, path: string) {
    switch (node.type) {
      case 'message':
        return need(messages.has(node.messageId), `${path}/messageId`, 'message', node.messageId);
      case 'image':
        return need(assets.has(node.assetId), `${path}/assetId`, 'asset', node.assetId);
      case 'richMenu':
        return need(richMenus.has(node.richMenuId), `${path}/richMenuId`, 'richMenu', node.richMenuId);
      case 'choice': {
        const m = messages.get(node.messageId);
        need(!!m, `${path}/messageId`, 'message', node.messageId);
        if (m && choiceButtons(node, project).length === 0) err('node.requiredMissing', `${path}/messageId`, { reason: 'noButtons' });
        return;
      }
      case 'survey':
        need(surveys.has(node.surveyId), `${path}/surveyId`, 'survey', node.surveyId);
        if (node.introMessageId) need(messages.has(node.introMessageId), `${path}/introMessageId`, 'message', node.introMessageId);
        return;
      case 'tag':
        if ((node.add ?? []).length + (node.remove ?? []).length === 0) err('node.requiredMissing', path);
        (node.add ?? []).forEach((t, i) => (tagsAdded.add(t), need(tags.has(t), `${path}/add/${i}`, 'tag', t)));
        (node.remove ?? []).forEach((t, i) => (tagsRead.add(t), need(tags.has(t), `${path}/remove/${i}`, 'tag', t)));
        return;
      case 'branch':
        if (node.cases.length === 0) err('node.requiredMissing', `${path}/cases`);
        unique(node.cases, `${path}/cases`);
        return node.cases.forEach((c, i) => checkExpr(c.when, `${path}/cases/${i}/when`));
      case 'wait':
        return checkWait(node.wait, `${path}/wait`);
      case 'page':
        return need(pages.has(node.pageId), `${path}/pageId`, 'page', node.pageId);
      case 'jump':
        startedFlows.add(node.flowId);
        return need(flows.has(node.flowId), `${path}/flowId`, 'flow', node.flowId);
      case 'entry':
      case 'end':
      case 'stopAll':
        return;
    }
  }


  // ---------- 時間が進まないループ ----------
  zeroTimeCycles(project, allNodes, err);

  // ---------- ペルソナ・説明 ----------
  project.personas.forEach((p, pi) => {
    const path = `/personas/${pi}`;
    (p.initial?.tags ?? []).forEach((t, i) => (tagsAdded.add(t), need(tags.has(t), `${path}/initial/tags/${i}`, 'tag', t)));
    (p.initial?.cvs ?? []).forEach((c, i) => need(cvs.has(c), `${path}/initial/cvs/${i}`, 'cv', c));
    (p.script ?? []).forEach((s, i) => checkEvent(s.event, `${path}/script/${i}/event`));
  });
  project.narration.forEach((n, i) => {
    if (n.on.kind === 'node') need(allNodes.has(n.on.nodeId), `/narration/${i}/on/nodeId`, 'node', n.on.nodeId);
    checkText(n.title, `/narration/${i}/title`);
    checkText(n.body, `/narration/${i}/body`);
    if (!n.body.trim()) warn('optional.empty', `/narration/${i}/body`);
  });

  (project.journey ?? []).forEach((s, i) => {
    const path = `/journey/${i}`;
    if (!s.label?.trim()) err('node.requiredMissing', `${path}/label`);
    (s.nodeIds ?? []).forEach((n, j) => need(allNodes.has(n), `${path}/nodeIds/${j}`, 'node', n));
    if (s.reachedWhen) checkExpr(s.reachedWhen, `${path}/reachedWhen`);
    if (s.cvId !== undefined) need(cvs.has(s.cvId), `${path}/cvId`, 'cv', s.cvId);
    if (s.loopTo !== undefined) need((project.journey ?? []).some((x) => x.id === s.loopTo), `${path}/loopTo`, 'journey', s.loopTo);
    if (!s.nodeIds?.length && !s.reachedWhen && s.cvId === undefined && !(s.variants ?? []).some((x) => x.nodeIds?.length)) warn('optional.empty', path, { reason: 'journeyReach' });
    const refs = (p: string, nodeIds?: Id[], messageIds?: Id[], feasibility?: string, options?: { id: Id; label: string; feasibility?: string }[]) => {
      (nodeIds ?? []).forEach((n, j) => need(allNodes.has(n), `${p}/nodeIds/${j}`, 'node', n));
      (messageIds ?? []).forEach((m, j) => need(messages.has(m), `${p}/messageIds/${j}`, 'message', m));
      if (feasibility !== undefined && !FEASIBILITY.has(feasibility)) err('journey.feasibility', `${p}/feasibility`);
      unique(options ?? [], `${p}/options`);
      (options ?? []).forEach((o, j) => {
        if (!o.label?.trim()) err('node.requiredMissing', `${p}/options/${j}/label`);
        if (o.feasibility !== undefined && !FEASIBILITY.has(o.feasibility)) err('journey.feasibility', `${p}/options/${j}/feasibility`);
      });
    };
    refs(path, undefined, s.messageIds, s.feasibility, s.options);
    unique(s.variants ?? [], `${path}/variants`);
    (s.variants ?? []).forEach((x, j) => {
      if (!x.label?.trim()) err('node.requiredMissing', `${path}/variants/${j}/label`);
      checkExpr(x.when, `${path}/variants/${j}/when`);
      refs(`${path}/variants/${j}`, x.nodeIds, x.messageIds, x.feasibility, x.options);
    });
  });

  function checkEvent(ev: InputEventBody, path: string) {
    switch (ev.type) {
      case 'visitPage':
      case 'exitIntent':
      case 'dismissPopup':
      case 'submitForm':
      case 'closePage':
      case 'tapPopupCta':
      case 'tapLineCta':
        return need(pages.has(ev.pageId), `${path}/pageId`, 'page', ev.pageId);
      case 'tapPageButton': {
        const page = project.webPages.find((w) => w.id === ev.pageId);
        need(!!page, `${path}/pageId`, 'page', ev.pageId);
        const block = page?.blocks[ev.blockIndex];
        if (page && block?.kind === 'list') need(block.items.some((x) => x.id === ev.itemId), `${path}/itemId`, 'item', ev.itemId);
        else if (page) need(block?.kind === 'button', `${path}/blockIndex`, 'button', ev.blockIndex);
        return;
      }
      case 'filterList': {
        const block = project.webPages.find((w) => w.id === ev.pageId)?.blocks[ev.blockIndex];
        need(block?.kind === 'list', `${path}/blockIndex`, 'list', ev.blockIndex);
        if (block?.kind === 'list') {
          const f = block.facets.find((x) => x.id === ev.facetId);
          need(!!f, `${path}/facetId`, 'facet', ev.facetId);
          if (f && ev.optionId !== '') need(f.options.some((x) => x.id === ev.optionId), `${path}/optionId`, 'option', ev.optionId);
        }
        return;
      }
      case 'addFriend':
        return need(routes.has(ev.routeId), `${path}/routeId`, 'route', ev.routeId);
      case 'tapMenu': {
        const menu = richMenus.get(ev.richMenuId);
        need(!!menu, `${path}/richMenuId`, 'richMenu', ev.richMenuId);
        if (menu) need(menu.cells.some((c) => c.id === ev.cellId), `${path}/cellId`, 'cell', ev.cellId);
        return;
      }
      case 'tapButton':
        need(buttons.has(ev.buttonId), `${path}/buttonId`, 'button', ev.buttonId);
        if (ev.nodeId !== undefined) need(allNodes.has(ev.nodeId), `${path}/nodeId`, 'node', ev.nodeId);
        return;
      case 'answer': {
        const q = questionOf(ev.surveyId, ev.questionId);
        need(!!q, `${path}/questionId`, 'question', `${ev.surveyId}/${ev.questionId}`);
        if (q) ev.choiceIds.forEach((c, i) => need(q.choices.some((x) => x.id === c), `${path}/choiceIds/${i}`, 'choice', c));
        return;
      }
      case 'cv':
        return need(cvs.has(ev.cvId), `${path}/cvId`, 'cv', ev.cvId);
      case 'viewCard': {
        const m = messages.get(ev.messageId);
        need(!!m, `${path}/messageId`, 'message', ev.messageId);
        if (m) need(m.bubbles.some((b) => b.kind === 'carousel' && b.cards.some((c) => c.id === ev.cardId)), `${path}/cardId`, 'card', ev.cardId);
        return;
      }
      default:
        return;
    }
  }

  // ---------- 使われていないもの ----------
  project.tags.forEach((t, i) => {
    if (!tagsAdded.has(t.id) && !tagsRead.has(t.id)) warn('tag.unused', `/tags/${i}`, { id: t.id });
  });
  project.flows.forEach((f, i) => {
    if (f.trigger.kind === 'manual' && !startedFlows.has(f.id)) warn('flow.neverStarted', `/flows/${i}`, { id: f.id });
  });

  // 同じメッセージに、別々のフローから届く経路がある
  const senders = new Map<Id, Set<Id>>();
  for (const { flow, node } of allNodes.values()) {
    const ids = node.type === 'message' || node.type === 'choice' ? [node.messageId] : node.type === 'survey' && node.introMessageId ? [node.introMessageId] : [];
    for (const id of ids) (senders.get(id) ?? senders.set(id, new Set()).get(id)!).add(flow.id);
  }
  for (const [messageId, flowIds] of senders) {
    if (flowIds.size < 2) continue;
    const anyUnstopped = [...flowIds].some((id) => !flows.get(id)?.stopWhen);
    if (project.settings.dedupeWindowMinutes === 0 || anyUnstopped) {
      const mi = project.messages.findIndex((m) => m.id === messageId);
      warn('delivery.possibleDuplicate', `/messages/${mi}`, { id: messageId, flows: [...flowIds].sort() });
    }
  }

  // ---------- AIの推定が残っていないか ----------
  walkProv(project, '', (path) => err('prov.unapproved', path));

  return issues;
}

/** 時間が進まず入力も待たないループ（＝無限ループ）を探す。jump は飛び先フローの入口へ繋ぐ */
function zeroTimeCycles(
  project: Project,
  allNodes: Map<Id, { flow: Flow; node: FlowNode; path: string }>,
  err: (code: string, path: string, params?: Record<string, unknown>) => void,
) {
  const breaks = (n: FlowNode) =>
    n.type === 'choice' ||
    n.type === 'survey' ||
    (n.type === 'wait' && (n.wait.kind !== 'duration' || n.wait.minutes >= 1));
  const next = new Map<Id, Id[]>();
  for (const flow of project.flows) {
    for (const e of flow.edges) (next.get(e.from.nodeId) ?? next.set(e.from.nodeId, []).get(e.from.nodeId)!).push(e.to);
    for (const n of flow.nodes) {
      if (n.type !== 'jump') continue;
      const entry = project.flows.find((f) => f.id === n.flowId)?.nodes.find((x) => x.type === 'entry');
      if (entry) (next.get(n.id) ?? next.set(n.id, []).get(n.id)!).push(entry.id);
    }
  }
  // Tarjan の強連結成分（時間が進むノードは取り除いたグラフで探す）
  let counter = 0;
  const index = new Map<Id, number>();
  const low = new Map<Id, number>();
  const onStack = new Set<Id>();
  const stack: Id[] = [];
  const reported = new Set<Id>();
  const visit = (v: Id) => {
    index.set(v, counter);
    low.set(v, counter++);
    stack.push(v);
    onStack.add(v);
    for (const w of next.get(v) ?? []) {
      const wn = allNodes.get(w)?.node;
      if (!wn || breaks(wn)) continue;
      if (!index.has(w)) {
        visit(w);
        low.set(v, Math.min(low.get(v)!, low.get(w)!));
      } else if (onStack.has(w)) low.set(v, Math.min(low.get(v)!, index.get(w)!));
    }
    if (low.get(v) === index.get(v)) {
      const comp: Id[] = [];
      let w: Id;
      do {
        w = stack.pop()!;
        onStack.delete(w);
        comp.push(w);
      } while (w !== v);
      const selfLoop = comp.length === 1 && (next.get(v) ?? []).includes(v);
      if ((comp.length > 1 || selfLoop) && !comp.some((id) => reported.has(id))) {
        comp.forEach((id) => reported.add(id));
        const first = [...comp].sort()[0]!;
        err('flow.zeroTimeCycle', allNodes.get(first)!.path, { nodes: [...comp].sort() });
      }
    }
  };
  for (const [id, { node }] of allNodes) if (!breaks(node) && !index.has(id)) visit(id);
}

function walkProv(value: unknown, path: string, onInferred: (path: string) => void) {
  if (Array.isArray(value)) return value.forEach((v, i) => walkProv(v, `${path}/${i}`, onInferred));
  if (!value || typeof value !== 'object') return;
  const obj = value as Record<string, unknown>;
  const prov = obj['prov'] as { state?: string } | undefined;
  if (prov && prov.state === 'inferred') onInferred(path === '' ? '/' : path);
  for (const [k, v] of Object.entries(obj)) if (k !== 'prov') walkProv(v, `${path}/${k}`, onInferred);
}

export const hasErrors = (issues: Issue[]) => issues.some((i) => i.level === 'error');
