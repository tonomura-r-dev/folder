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
    /** 自動再生で最後まで流す日数（既定 30）。30日より先の配信（休眠掘り起こし等）を見せたい案件で延ばす */
    demoEndDay?: number;
    /** デモの「はじめかた」に大きく出す入口（Webページ）。無ければ全ページを並べる。ほかのページ・経路は折りたたみに残す */
    demoEntries?: { pageId: Id; label: string }[];
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
export type CvDef = Tracked & {
  id: Id;
  label: string;
  /** デモ画面②の「その段階へ進める」ボタンの名前（例：来場予約後へ）。無ければ「<label>へ」 */
  demoAction?: string;
  /**
   * cv = ユーザーの成果（既定）／stage = 運営側が登録する段階（例：勤務確定）。
   * stage はCVの数に入れず、デモ画面では「運営側の登録」として別に出す
   */
  kind?: 'cv' | 'stage';
  /** 何度でも記録できる（例：再応募・イベントごとの勤務確定）。デモ画面②のボタンが済みにならない */
  repeatable?: boolean;
  /** false＝デモ画面②に手で記録するボタンを出さない（予約の日時など、①の画面で選ぶ内容があるCV） */
  demoButton?: boolean;
};
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
  | { kind: 'card'; title: Text; body?: Text; assetId?: Id; buttons: Button[] }
  /** 横に並ぶカード（LINEのカルーセル）。左右に送って見る。見たカードは viewCard として記録する */
  | { kind: 'carousel'; cards: CarouselCard[] };
export type Button = { id: Id; label: string; action: Action };
/** カルーセルの1枚。facts はカードに並べる項目（例：勤務地・開始時期） */
export type CarouselCard = {
  id: Id;
  title: Text;
  body?: Text;
  assetId?: Id;
  /** 画像の上に重ねる小さな札（例：デモ用サンプル） */
  badge?: string;
  facts?: { label: string; value: Text }[];
  buttons: Button[];
};

export type Action =
  | { kind: 'sendText'; text: string }
  | { kind: 'openPage'; pageId: Id }
  | { kind: 'startFlow'; flowId: Id }
  | { kind: 'answer'; surveyId: Id; questionId: Id; choiceId: Id }
  | { kind: 'addFriend'; routeId: Id }
  | { kind: 'switchMenu'; richMenuId: Id }
  /** Webページを閉じて、LINEのトークに戻る（予約完了画面の「LINEに戻る」など） */
  | { kind: 'backToLine' }
  | { kind: 'none' };

export type RichMenuCell = {
  id: Id;
  label: string;
  area: { col: number; row: number; colSpan?: number; rowSpan?: number };
  action: Action;
  /** いちばん押してほしい項目として目立たせる */
  emphasis?: boolean;
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
  | { kind: 'button'; label: string; action: Action }
  /** 絞り込みつきの一覧（求人一覧など）。絞り込みは filterList、項目を押すと tapPageButton（itemId つき）として記録する */
  | { kind: 'list'; facets: ListFacet[]; items: ListItem[]; emptyText?: Text };
/** 絞り込みの軸（例：エリア）。tagId のある選択肢は、そのタグを持つ人には最初から選ばれた状態で開く */
export type ListFacet = { id: Id; label: string; options: { id: Id; label: string; tagId?: Id }[] };
/** 一覧の1件。values は軸ごとに当てはまる選択肢（例：{ area: ['yokohama'] }） */
export type ListItem = {
  id: Id;
  title: Text;
  body?: Text;
  assetId?: Id;
  badge?: string;
  facts?: { label: string; value: Text }[];
  values: Record<Id, Id[]>;
  action: Action;
  /** 募集の締切（デモの時刻・分）。過ぎたら一覧に出さない */
  closesAt?: Minutes;
};
export type WebPage = Tracked & {
  id: Id;
  kind: 'lp' | 'form' | 'thanks';
  title: string;
  url?: string;
  blocks: WebBlock[];
  form?: {
    /** closesAt は選択肢ごとの締切（options と同じ順・null は締切なし）。過ぎた選択肢は出さない（予約の空き枠など） */
    fields: { id: Id; label: string; type: 'text' | 'select' | 'tel' | 'email'; options?: string[]; closesAt?: (Minutes | null)[] }[];
    onSubmit: { cvId?: Id; nextPageId?: Id };
    /** 送信ボタンの文字（既定：送信する） */
    submitLabel?: string;
    /** 送信の前に、入力内容の確認画面を出す（予約内容の確認など）。画面だけの動きで、送信するまでイベントにはならない */
    confirm?: { title?: string; note?: Text; submitLabel?: string };
  };
  exitPopup?: {
    title: Text;
    body: Text;
    assetId?: Id;
    cta: { label: string; action: Action };
    dismissLabel: string;
  };
  /** prefill はLINEのトーク入力欄に入れておく文（サンクスLINE誘導でフォームの内容を引き継ぐ等）。送るかはユーザーが決める */
  lineCta?: { label: string; routeId: Id; prefill?: string };
};

export type ScriptStep = { atMinutes: Minutes; event: InputEventBody };
export type Persona = Tracked & {
  id: Id;
  /** シナリオ上の役割（例：検討中さん）。画面でユーザーを見分けるのに使う */
  name: string;
  /** 架空の人物名（例：山田 一郎）。あれば文面の {{user.name}} はこちらになり、画面は「人物名＠役割」で出す */
  personName?: string;
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
  /** この段のあと、くり返しで戻る段（例：再応募 → 勤務確定）。③に「↺ 戻る」として出す */
  loopTo?: Id;
  loopLabel?: string;
  /** 主要な工程。資料用16:9の③でも「⋮ ○段」にまとめず、いつも出す */
  key?: boolean;
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
  /** 全員共通の日時（シミュレーションの Day day の time）。企画配信・イベント前日のリマインド等。過ぎていたら passed へ（無ければ終了） */
  | { kind: 'at'; day: number; time: string }
  | { kind: 'untilEvent'; event: InputEventType; timeoutMinutes: Minutes };

type NodeBase = Tracked & { id: Id; label?: string; position: { x: number; y: number } };
export type FlowNode = NodeBase &
  (
    | { type: 'entry' }
    | { type: 'message'; messageId: Id }
    | { type: 'image'; assetId: Id }
    | { type: 'richMenu'; richMenuId: Id }
    | { type: 'choice'; messageId: Id; timeoutMinutes?: Minutes }
    /** reask = 前の回答を消して、はじめから聞き直す（希望条件の変更など） */
    | { type: 'survey'; surveyId: Id; introMessageId?: Id; timeoutMinutes?: Minutes; reask?: boolean }
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
  /** カルーセルのカードを見た（左右に送って表示した） */
  | { type: 'viewCard'; messageId: Id; cardId: Id }
  | { type: 'sendText'; text: string }
  | { type: 'answer'; surveyId: Id; questionId: Id; choiceIds: Id[] }
  | { type: 'cv'; cvId: Id }
  // Webページ上の操作。押したときの動作は案件JSONにあるので、エンジンが解釈する
  | { type: 'closePage'; pageId: Id }
  | { type: 'tapPopupCta'; pageId: Id }
  /** itemId は一覧（list）の項目を押したとき */
  | { type: 'tapPageButton'; pageId: Id; blockIndex: number; itemId?: Id }
  | { type: 'tapLineCta'; pageId: Id }
  /** 一覧の絞り込みを変えた（optionId が空なら「すべて」）。行動履歴に残るだけ */
  | { type: 'filterList'; pageId: Id; blockIndex: number; facetId: Id; optionId: Id };
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
  /** 架空の人物名（ペルソナの personName）。文面の {{user.name}} に使う */
  personName?: string;
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
