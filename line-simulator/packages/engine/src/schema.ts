// シナリオエンジンの型。docs/DATA_MODEL.md と1対1で対応させる。
// ここに業種固有の語・画面に出す文字列を置かない（tools/check-engine.ts が検査する）。

export type Id = string;
/** 仮想時間（分）。シミュレーション開始 = 0。Day = floor(分 / 1440) */
export type Minutes = number;
/** 画面に出す文。{{user.name}} などの差し込み可 */
export type Text = string;

export const MINUTES_PER_DAY = 1440;

// ---------- 出所・社内情報 ----------

export type ProvenanceState = 'observed' | 'inferred' | 'approved' | 'manual';
export type Provenance = {
  state: ProvenanceState;
  source?: string;
  rationale?: string;
  at?: string;
};
/** 社内専用。営業用HTMLでは必ず取り除く */
export type Internal = { memo?: string; [k: string]: unknown };
type Tracked = { prov?: Provenance; _internal?: Internal };

// ---------- 案件 ----------

export type Project = {
  schemaVersion: 1;
  meta: { id: Id; name: string; slug: string; industry?: string; createdAt: string; updatedAt: string };
  brand: {
    companyName: string;
    accountName: string;
    iconAssetId?: Id;
    colors: { primary: string; accent: string; chatBg?: string };
  };
  settings: {
    /** 同じメッセージを同じ人に再送しない期間。0 で無効 */
    dedupeWindowMinutes: Minutes;
    defaultRichMenuId?: Id;
    messageCounting?: { enabled: boolean; freeQuota?: number };
    /**
     * デモ（営業用HTML・編集ツールのデモ画面）を開いたときの始まり方。既定は single。
     * single = 1人を手で操作（personaId のペルソナ。無ければ先頭）／scripted = 台本のあるペルソナ全員を台本どおりに並べる
     */
    demoStart?: { mode: 'single' | 'scripted'; personaId?: Id };
    labels?: Record<string, string>;
  };
  assets: Asset[];
  tags: TagDef[];
  cvs: CvDef[];
  routes: RouteDef[];
  messages: MessageDef[];
  richMenus: RichMenu[];
  autoReplies: AutoReply[];
  surveys: Survey[];
  webPages: WebPage[];
  flows: Flow[];
  personas: Persona[];
  narration: NarrationDef[];
  /** CVまでの道筋（デモ画面の③）。表示だけに使い、エンジンの動きには関係しない。無い案件はシナリオ一覧から簡易表示を作る */
  journey?: JourneyStep[];
  _internal?: Internal;
};

export type Asset = Tracked & { id: Id; kind: 'image'; path: string; width: number; height: number; alt: string };
export type TagDef = Tracked & { id: Id; label: string; group?: string; color?: string; description?: string };
export type CvDef = Tracked & { id: Id; label: string };
export type RouteDef = Tracked & { id: Id; label: string };

export type MessageDef = Tracked & {
  id: Id;
  bubbles: Bubble[];
  /** 通数の換算。既定 1（吹き出しが複数でも1配信=1通） */
  countAs?: number;
};
export type Bubble =
  | { kind: 'text'; text: Text }
  | { kind: 'image'; assetId: Id }
  | { kind: 'buttons'; text: Text; buttons: Button[] }
  | { kind: 'card'; title: Text; body?: Text; assetId?: Id; buttons: Button[] };
export type Button = { id: Id; label: string; action: Action };

export type Action =
  | { kind: 'sendText'; text: string }
  | { kind: 'openPage'; pageId: Id }
  | { kind: 'startFlow'; flowId: Id }
  | { kind: 'answer'; surveyId: Id; questionId: Id; choiceId: Id }
  | { kind: 'addFriend'; routeId: Id }
  | { kind: 'switchMenu'; richMenuId: Id }
  | { kind: 'none' };

export type RichMenuCell = {
  id: Id;
  label: string;
  area: { col: number; row: number; colSpan?: number; rowSpan?: number };
  action: Action;
};
export type RichMenu = Tracked & {
  id: Id;
  label: string;
  layout: { cols: number; rows: number };
  imageAssetId?: Id;
  cells: RichMenuCell[];
};

export type AutoReply = Tracked & {
  id: Id;
  priority: number;
  match: { mode: 'exact' | 'contains'; keywords: string[] };
  cases: { when: Expr; messageId: Id }[];
  fallbackMessageId?: Id;
};

export type SurveyChoice = { id: Id; label: string; addTags?: Id[]; removeTags?: Id[] };
export type SurveyQuestion = { id: Id; text: Text; mode: 'single' | 'multiple'; choices: SurveyChoice[] };
export type Survey = Tracked & { id: Id; label: string; questions: SurveyQuestion[] };

export type WebBlock =
  | { kind: 'heading' | 'text'; text: Text }
  | { kind: 'image'; assetId: Id }
  | { kind: 'button'; label: string; action: Action };
export type WebPage = Tracked & {
  id: Id;
  kind: 'lp' | 'form' | 'thanks';
  title: string;
  url?: string;
  blocks: WebBlock[];
  form?: {
    fields: { id: Id; label: string; type: 'text' | 'select' | 'tel' | 'email'; options?: string[] }[];
    onSubmit: { cvId?: Id; nextPageId?: Id };
  };
  exitPopup?: {
    title: Text;
    body: Text;
    assetId?: Id;
    cta: { label: string; action: Action };
    dismissLabel: string;
  };
  lineCta?: { label: string; routeId: Id };
};

export type ScriptStep = { atMinutes: Minutes; event: InputEventBody };
export type Persona = Tracked & {
  id: Id;
  name: string;
  attrs: Record<string, string | number | boolean>;
  initial?: { friendState?: FriendState; tags?: Id[]; cvs?: Id[] };
  autoAnswers?: { surveyId: Id; questionId: Id; choiceIds: Id[] }[];
  script?: ScriptStep[];
};

export type NarrationDef = Tracked & {
  id: Id;
  on: { kind: 'node'; nodeId: Id } | { kind: 'event'; type: InputEventType };
  eyebrow?: string;
  title: Text;
  body: Text;
};

/**
 * Lステップでの実現性（営業説明用の目安）。
 * lstep = Lステップ単体で可／conditional = プラン・設定に条件あり／external = 外部ツール・Webサイト側の実装が必要／unknown = 未確認
 */
export type Feasibility = 'lstep' | 'conditional' | 'external' | 'unknown';

/** 追加提案の候補。表示するだけで、シナリオには入らず配信もされない */
export type JourneyOption = Tracked & { id: Id; label: string; summary?: Text; sample?: Text; feasibility?: Feasibility; note?: Text };

/** 分岐の行き先（例：東京 → 東京向け求人配信） */
export type JourneyVariant = Tracked & {
  id: Id;
  /** 矢印に付ける分岐の名前 */
  label: string;
  when: Expr;
  /** 行き先の施策名（無ければ label） */
  title?: string;
  purpose?: Text;
  /** この行き先の施策＝シナリオのノード（通ったら「通った経路」。文面はここから参照する） */
  nodeIds?: Id[];
  /** シナリオ外の文面（自動応答など）を参照するとき */
  messageIds?: Id[];
  feasibility?: Feasibility;
  feasibilityNote?: Text;
  options?: JourneyOption[];
};

/**
 * 道筋の1段（例：友だち追加 → アンケート → 出し分け → 後押し → CV）。最初の段が START、最後のCV地点が GOAL。
 * 到達：nodeIds のどれかを通った／reachedWhen が真／cvId をした／variants の nodeIds を通った のどれか。
 * variants があると分岐（ひし形）になり、上から順に最初に when が真になったものがその人のルート
 */
export type JourneyStep = Tracked & {
  id: Id;
  label: string;
  purpose?: Text;
  nodeIds?: Id[];
  reachedWhen?: Expr;
  cvId?: Id;
  /** 分岐のひし形に書く問い（例：地域選択済み？）。無ければ label */
  decision?: string;
  variants?: JourneyVariant[];
  messageIds?: Id[];
  feasibility?: Feasibility;
  feasibilityNote?: Text;
  options?: JourneyOption[];
};

// ---------- フロー ----------

export type Trigger =
  | { kind: 'friendAdd'; routeIds?: Id[] }
  | { kind: 'keyword'; autoReplyId: Id }
  | { kind: 'menuTap'; richMenuId: Id; cellId: Id }
  | { kind: 'cv'; cvId: Id }
  | { kind: 'pageExit'; pageId: Id }
  | { kind: 'formSubmit'; pageId: Id }
  | { kind: 'tagAdded'; tagId: Id }
  | { kind: 'manual' };

export type Flow = Tracked & {
  id: Id;
  label: string;
  trigger: Trigger;
  /** 小さいほど優先 */
  priority: number;
  concurrency: 'exclusive' | 'parallel';
  stopWhen?: Expr;
  reentry: 'never' | 'afterEnd' | 'always';
  nodes: FlowNode[];
  edges: Edge[];
};

export type Wait =
  | { kind: 'duration'; minutes: Minutes }
  /** 起点の日から day 日後の time（HH:MM）。起点の既定は friendAdd（無ければ flowStart） */
  | { kind: 'dayAt'; day: number; time: string; from?: 'friendAdd' | 'flowStart' }
  | { kind: 'untilEvent'; event: InputEventType; timeoutMinutes: Minutes };

type NodeBase = Tracked & { id: Id; label?: string; position: { x: number; y: number } };
export type FlowNode = NodeBase &
  (
    | { type: 'entry' }
    | { type: 'message'; messageId: Id }
    | { type: 'image'; assetId: Id }
    | { type: 'richMenu'; richMenuId: Id }
    | { type: 'choice'; messageId: Id; timeoutMinutes?: Minutes }
    | { type: 'survey'; surveyId: Id; introMessageId?: Id; timeoutMinutes?: Minutes }
    | { type: 'tag'; add?: Id[]; remove?: Id[] }
    | { type: 'branch'; cases: { id: Id; label: string; when: Expr }[] }
    | { type: 'wait'; wait: Wait }
    | { type: 'page'; pageId: Id }
    | { type: 'jump'; flowId: Id }
    | { type: 'end'; reason?: string }
    | { type: 'stopAll'; reason?: string }
  );
export type NodeType = FlowNode['type'];

export type Edge = Tracked & { id: Id; from: { nodeId: Id; port: string }; to: Id; label?: string };

// ---------- 条件式 ----------

export type Cmp = 'eq' | 'neq' | 'in' | 'gte' | 'lte';
export type Expr =
  | { op: 'and' | 'or'; children: Expr[] }
  | { op: 'not'; child: Expr }
  | { op: 'always' }
  | { op: 'hasTag'; tagId: Id }
  | { op: 'answer'; surveyId: Id; questionId: Id; match: 'any' | 'all' | 'none'; choiceIds: Id[] }
  | { op: 'answered'; surveyId: Id; questionId?: Id }
  | { op: 'attr'; key: string; cmp: Cmp; value: unknown }
  | { op: 'route'; routeIds: Id[] }
  | { op: 'friend'; state: FriendState }
  | { op: 'cv'; cvId: Id; done: boolean; withinMinutes?: Minutes }
  | {
      op: 'did';
      event: InputEventType;
      where?: Record<string, string>;
      cmp: 'exists' | 'countGte';
      value?: number;
      withinMinutes?: Minutes;
    }
  | { op: 'elapsed'; since: 'friendAdd' | 'flowStart' | 'lastDelivery'; cmp: 'gte' | 'lt'; minutes: Minutes }
  | { op: 'delivered'; target: { messageId?: Id; flowId?: Id; nodeId?: Id }; done: boolean; withinMinutes?: Minutes };
export type ExprOp = Expr['op'];

export type Trace = {
  op: ExprOp;
  value: boolean;
  expr: Expr;
  actual?: unknown;
  children?: Trace[];
};

// ---------- 入力イベント・セッション ----------

export type FriendState = 'not_friend' | 'friend' | 'blocked';

export type InputEventBody =
  | { type: 'visitPage'; pageId: Id }
  | { type: 'exitIntent'; pageId: Id }
  | { type: 'dismissPopup'; pageId: Id }
  | { type: 'submitForm'; pageId: Id; values: Record<string, string> }
  | { type: 'addFriend'; routeId: Id }
  | { type: 'block' }
  | { type: 'unblock' }
  | { type: 'tapMenu'; richMenuId: Id; cellId: Id }
  | { type: 'tapButton'; buttonId: Id; nodeId?: Id }
  | { type: 'sendText'; text: string }
  | { type: 'answer'; surveyId: Id; questionId: Id; choiceIds: Id[] }
  | { type: 'cv'; cvId: Id }
  // Webページ上の操作。押したときの動作は案件JSONにあるので、エンジンが解釈する
  | { type: 'closePage'; pageId: Id }
  | { type: 'tapPopupCta'; pageId: Id }
  | { type: 'tapPageButton'; pageId: Id; blockIndex: number }
  | { type: 'tapLineCta'; pageId: Id };
export type InputEventType = InputEventBody['type'];
export type InputEvent = { at: Minutes; userId: Id } & InputEventBody;
/** 時間だけを進める（全員共通） */
export type ClockEvent = { at: Minutes; userId: '*'; type: 'advanceTo' };
export type SessionEvent = InputEvent | ClockEvent;

export type Session = {
  schemaVersion: 1;
  projectId: Id;
  projectHash?: string;
  users: { id: Id; personaId: Id; displayName?: string }[];
  events: SessionEvent[];
  cursor?: Minutes;
};

// ---------- 状態（エンジンが導出する） ----------

export type FlowCursor = {
  status: 'running' | 'waiting' | 'ended' | 'stopped';
  nodeId: Id | null;
  startedAt: Minutes;
  endedAt?: Minutes;
  /** untilEvent で待っているイベント */
  awaitEvent?: InputEventType;
  runs: number;
};
export type Timer = { userId: Id; flowId: Id; nodeId: Id; fireAt: Minutes; kind: 'wait' | 'timeout' };
export type Delivery = { at: Minutes; flowId: Id | null; nodeId: Id | null; messageId: Id; count: number };
export type Surface =
  | { kind: 'talk' }
  | { kind: 'page'; pageId: Id }
  | { kind: 'popup'; pageId: Id }
  | { kind: 'survey'; surveyId: Id; questionId: Id };

export type SystemCode = 'friendAdded' | 'blocked' | 'unblocked' | 'surveyShown' | 'pageOpened';
export type TalkItem =
  | { at: Minutes; side: 'bot'; bubble: Bubble; messageId: Id; fromNodeId?: Id }
  | { at: Minutes; side: 'user'; bubble: Bubble }
  | { at: Minutes; side: 'system'; code: SystemCode; params?: Record<string, string> };

export type Behavior = { at: Minutes; type: InputEventType; detail: Record<string, string> };

export type UserState = {
  id: Id;
  personaId: Id;
  name: string;
  attrs: Record<string, unknown>;
  friendState: FriendState;
  friendAddedAt: Minutes | null;
  route: Id | null;
  tags: Record<Id, { at: Minutes; byNodeId?: Id }>;
  answers: Record<Id, Record<Id, Id[]>>;
  cvs: Record<Id, { at: Minutes; count: number }>;
  behaviors: Behavior[];
  deliveries: Delivery[];
  flows: Record<Id, FlowCursor>;
  richMenuId: Id | null;
  surface: Surface;
  talk: TalkItem[];
  messageCount: number;
  /** 通ったノードの記録（説明カード・「いま動いているシナリオ」の表示に使う） */
  trail: { at: Minutes; flowId: Id; nodeId: Id }[];
};

export type DecisionKind = 'branch' | 'stopWhen' | 'autoReply' | 'trigger' | 'deliver';
export type DecisionOutcome = 'taken' | 'skipped' | 'excluded' | 'deferred' | 'deduped' | 'stopped';
export type DecisionReason =
  | 'blocked'
  | 'not_friend'
  | 'lower_priority'
  | 'duplicate'
  | 'stop_condition'
  | 'condition_false'
  | 'reentry'
  | 'stop_all';
export type Decision = {
  at: Minutes;
  userId: Id;
  flowId: Id | null;
  nodeId: Id | null;
  kind: DecisionKind;
  outcome: DecisionOutcome;
  /** branch の case id / autoReply の case 番号 / 配信の messageId など */
  ref?: string;
  trace?: Trace;
  reason?: DecisionReason;
};

export type RuntimeWarning = { at: Minutes; code: 'runtime.stepBudget' | 'runtime.unknownRef'; userId: Id; ref: string };

export type WorldState = {
  now: Minutes;
  users: Record<Id, UserState>;
  timers: Timer[];
  decisions: Decision[];
  warnings: RuntimeWarning[];
};

// ---------- 検証・差分 ----------

export type IssueLevel = 'error' | 'warning';
export type Issue = { level: IssueLevel; code: string; path: string; params?: Record<string, unknown> };

export type PatchOp = { op: 'add' | 'replace' | 'remove'; path: string; value?: unknown; prov: Provenance };
export type ScenarioPatch = { schemaVersion: 1; basedOn?: string; summary: string; ops: PatchOp[] };
