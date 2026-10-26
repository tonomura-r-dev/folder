# DATA_MODEL — JSON構造・ユーザー状態・分岐ロジック

最終更新：2026-10-26（Phase 1）

型は TypeScript 風に書く。実装では `packages/engine/src/schema.ts` にこのまま置く想定。
ここに書く値の例（「東京」「応募」など）は**検証案件の案件JSONに入るもの**であり、エンジンには入らない。

---

## 1. 全体像

データは3つに分ける。混ぜない。

| 区分 | 何か | ファイル | 誰が書くか |
|---|---|---|---|
| **案件（Project）** | シナリオ設定。フロー・文面・メニュー・アンケート・タグ定義・ペルソナ | `projects/<案件>/project.json` | 編集ツール／AI差分 |
| **セッション（Session）** | デモの実行記録。入力イベント列だけ | `projects/<案件>/sessions/<名前>.json` | デモ操作 |
| **状態（State）** | ユーザーごとのタグ・配信履歴・進行状況 | 保存しない（毎回計算） | エンジン |

```
State = reduce(Project, Session.events)
```

案件を直したらセッションをそのまま再生し直せる。「Day 7 に戻る」は「イベント列をDay 7の手前で切って再計算」。

---

## 2. 共通の型

```ts
type Id = string;            // 案件内で一意。英数字と - _ のみ（例: "msg_welcome"）
type Minutes = number;       // 仮想時間。友だち追加などの基準からの「分」。整数
type Text = string;          // 画面に出す文。{{...}} の差し込み可（§9）

/** 出所（AI連携用）。すべての主要オブジェクトが任意で持つ */
type Provenance = {
  state: 'observed' | 'inferred' | 'approved' | 'manual';
  // observed = スクショで確認できた / inferred = AIが推定・提案した
  // approved = inferred を人が承認した / manual = 人が手で作った（既定）
  source?: string;           // 例: "screenshot:03_richmenu.png"
  rationale?: string;        // AIが推定した根拠
  at?: string;               // ISO日時（書き込み時刻。エンジンの判定には使わない）
};

/** 社内専用の情報。営業用HTMLでは必ず取り除かれる */
type Internal = { memo?: string; [k: string]: unknown };
```

すべての主要オブジェクトは `prov?: Provenance` と `_internal?: Internal` を持てる。

---

## 3. 案件（Project）

```ts
type Project = {
  schemaVersion: 1;
  meta: {
    id: Id; name: string; slug: string;          // slug = フォルダ名
    industry?: string;                           // 自由記述（エンジンは見ない）
    createdAt: string; updatedAt: string;
  };
  brand: {
    companyName: string;                         // "サンプルスタッフ株式会社（仮）"
    accountName: string;                         // LINE上の表示名
    iconAssetId?: Id;
    colors: { primary: string; accent: string; chatBg?: string };
  };
  settings: {
    dedupeWindowMinutes: Minutes;                // 重複配信の防止期間（FR-16）。既定 1440
    defaultDeliveryTime?: string;                // "09:00"
    messageCounting?: { enabled: boolean; freeQuota?: number }; // 通数（追加提案G-6・採否待ち）
    labels?: Record<string, string>;             // 画面用の言い換え（例: "cv" → "応募"）
  };

  assets: Asset[];
  tags: TagDef[];
  cvs: CvDef[];                                  // 応募・予約など、CVの種類
  routes: RouteDef[];                            // 友だち追加の経路
  messages: MessageDef[];
  richMenus: RichMenu[];
  autoReplies: AutoReply[];
  surveys: Survey[];
  webPages: WebPage[];
  flows: Flow[];
  personas: Persona[];
  narration: NarrationDef[];                     // 説明文（FR-54）

  _internal?: Internal;                          // 案件全体の社内メモ
};
```

### 3.1 部品

```ts
type Asset = { id: Id; kind: 'image'; path: string; width: number; height: number; alt: string };

type TagDef   = { id: Id; label: string; group?: string; color?: string; description?: string };
type CvDef    = { id: Id; label: string };                      // { id:"cv_apply", label:"応募" }
type RouteDef = { id: Id; label: string };                      // { id:"route_exit_popup", label:"離脱ポップアップ" }

type MessageDef = {
  id: Id;
  bubbles: Bubble[];                    // 1配信 = 吹き出し1〜5個（LINEの上限に合わせて検証）
  countAs?: number;                     // 通数換算（既定 1）
};
type Bubble =
  | { kind: 'text'; text: Text }
  | { kind: 'image'; assetId: Id }
  | { kind: 'buttons'; text: Text; buttons: Button[] }   // ボタン付きメッセージ
  | { kind: 'card'; title: Text; body?: Text; assetId?: Id; buttons: Button[] };

type Button = { id: Id; label: string; action: Action };

/** ボタン・メニューセル・ポップアップ等が押されたときの動作 */
type Action =
  | { kind: 'sendText'; text: string }          // ユーザーの発言としてトークに送る → 自動応答の判定へ
  | { kind: 'openPage'; pageId: Id }             // Webページへ
  | { kind: 'startFlow'; flowId: Id }            // フローを起動
  | { kind: 'answer'; surveyId: Id; questionId: Id; choiceId: Id }
  | { kind: 'addFriend'; routeId: Id }           // 友だち追加（ポップアップ・サンクス）
  | { kind: 'switchMenu'; richMenuId: Id }
  | { kind: 'none' };

type RichMenu = {
  id: Id; label: string;
  layout: { cols: number; rows: number };       // 2x3 等
  imageAssetId?: Id;                             // 無ければセルのラベルで描画
  cells: { id: Id; label: string; area: { col: number; row: number; colSpan?: number; rowSpan?: number }; action: Action }[];
};

type AutoReply = {
  id: Id; priority: number;                      // 小さいほど先に判定
  match: { mode: 'exact' | 'contains'; keywords: string[] };
  cases: { when: Expr; messageId: Id }[];        // 条件別の応答（FR-21）。上から順に最初の真
  fallbackMessageId?: Id;
};

type Survey = {
  id: Id; label: string;
  questions: {
    id: Id; text: Text; mode: 'single' | 'multiple';
    choices: { id: Id; label: string; addTags?: Id[]; removeTags?: Id[] }[];
  }[];
};

type WebPage = {
  id: Id; kind: 'lp' | 'form' | 'thanks';
  title: string; url?: string;                   // url は表示用の飾り。開かない
  blocks: ({ kind: 'heading' | 'text'; text: Text } | { kind: 'image'; assetId: Id }
           | { kind: 'button'; label: string; action: Action })[];
  form?: { fields: { id: Id; label: string; type: 'text' | 'select' | 'tel' | 'email'; options?: string[] }[];
           onSubmit: { cvId?: Id; nextPageId?: Id } };
  exitPopup?: { title: Text; body: Text; assetId?: Id; cta: { label: string; action: Action }; dismissLabel: string };
  lineCta?: { label: string; routeId: Id };      // サンクスページの「LINEで受け取る」等
};

type Persona = {
  id: Id; name: string;
  attrs: Record<string, string | number | boolean>;   // { area:"東京", age:28 }
  initial?: { friendState?: FriendState; tags?: Id[]; cvs?: Id[] };
  autoAnswers?: { surveyId: Id; questionId: Id; choiceIds: Id[] }[];  // 自動再生で台本どおり答える（FR-64）
  script?: ScriptStep[];                                               // 自動再生でこの人が取る行動
};
type ScriptStep = { atMinutes: Minutes; event: InputEvent };

type NarrationDef = {
  id: Id;
  on: { kind: 'node'; nodeId: Id } | { kind: 'event'; type: InputEvent['type'] };
  eyebrow?: string; title: Text; body: Text;     // 八十吉デモの説明カード
};
```

---

## 4. フロー（ノード・エッジ）

### 4.1 フロー

```ts
type Flow = {
  id: Id; label: string;
  trigger: Trigger;                              // 起点（FR-13）
  priority: number;                              // 小さいほど優先（FR-15）
  concurrency: 'exclusive' | 'parallel';         // 既定 exclusive
  stopWhen?: Expr;                               // 停止条件（FR-17）。毎ステップ前に評価
  reentry: 'never' | 'afterEnd' | 'always';      // 同じ人が再度入れるか。既定 never
  nodes: Node[];
  edges: Edge[];
};

type Trigger =
  | { kind: 'friendAdd'; routeIds?: Id[] }       // 指定がなければ全経路
  | { kind: 'keyword'; autoReplyId: Id }
  | { kind: 'menuTap'; richMenuId: Id; cellId: Id }
  | { kind: 'cv'; cvId: Id }
  | { kind: 'pageExit'; pageId: Id }
  | { kind: 'formSubmit'; pageId: Id }
  | { kind: 'tagAdded'; tagId: Id }
  | { kind: 'manual' };                          // 他フローの jump からのみ
```

**exclusive の意味**：同じユーザーで exclusive なフローが2本同時に「配信しようとした」とき、優先度が高い方だけが配信し、低い方は**その配信を見送る**（タイマーは残す）。見送った理由は判定記録に残す。`parallel` のフローはこの制御の対象外。

### 4.2 ノード

```ts
type NodeBase = { id: Id; label?: string; position: { x: number; y: number };
                  prov?: Provenance; _internal?: Internal };

type Node = NodeBase & (
  | { type: 'entry' }                                            // フローに1つだけ。trigger の到達点
  | { type: 'message'; messageId: Id }                           // ポート: out
  | { type: 'image'; assetId: Id }                               // ポート: out
  | { type: 'richMenu'; richMenuId: Id }                         // ポート: out
  | { type: 'choice'; messageId: Id; timeoutMinutes?: Minutes }  // ボタン選択。ポート: btn:<buttonId> / timeout
  | { type: 'survey'; surveyId: Id; timeoutMinutes?: Minutes }   // ポート: done / timeout
  | { type: 'tag'; add?: Id[]; remove?: Id[] }                   // ポート: out
  | { type: 'branch'; cases: { id: Id; label: string; when: Expr }[] } // ポート: case:<id> / else
  | { type: 'wait'; wait: Wait }                                 // ポート: out
  | { type: 'page'; pageId: Id }                                 // 外部ページ遷移。ポート: out
  | { type: 'jump'; flowId: Id }                                 // 別フローへ。ポート: なし
  | { type: 'end'; reason?: string }                             // このフローを終える
  | { type: 'stopAll'; reason?: string }                         // 全フローを止める（CV時など）
);

type Wait =
  | { kind: 'duration'; minutes: Minutes }                       // N分後
  | { kind: 'dayAt'; day: number; time: string }                 // 起点から N日目の HH:MM（G-2）
  | { kind: 'untilEvent'; event: InputEvent['type']; timeoutMinutes: Minutes }; // ポート: out / timeout
```

元の指示の10種との対応：メッセージ＝`message`／画像＝`image`／リッチメニュー表示＝`richMenu`／ボタン選択＝`choice`／アンケート＝`survey`／タグ付与＝`tag`／条件分岐＝`branch`／ステップ配信・待機＝`wait`／外部ページ遷移＝`page`／終了＝`end`。
追加は `entry`・`jump`・`stopAll` の3つ（REQUIREMENTS FR-13・FR-17）。

### 4.3 エッジ

```ts
type Edge = { id: Id; from: { nodeId: Id; port: string }; to: Id; label?: string };
```

ポートは名前付きの文字列。ノード種別ごとに許されるポートが決まっていて、検証で照合する（§8）。

---

## 5. 条件式（Expr）と判定の木（Trace）

### 5.1 条件式

```ts
type Expr =
  | { op: 'and' | 'or'; children: Expr[] }
  | { op: 'not'; child: Expr }
  | { op: 'always' }
  | { op: 'hasTag'; tagId: Id }
  | { op: 'answer'; surveyId: Id; questionId: Id; match: 'any' | 'all' | 'none'; choiceIds: Id[] }
  | { op: 'answered'; surveyId: Id; questionId?: Id }
  | { op: 'attr'; key: string; cmp: 'eq' | 'neq' | 'in' | 'gte' | 'lte'; value: unknown }
  | { op: 'route'; routeIds: Id[] }
  | { op: 'friend'; state: FriendState }
  | { op: 'cv'; cvId: Id; done: boolean; withinMinutes?: Minutes }
  | { op: 'did'; event: InputEvent['type']; where?: Record<string, string>;
      cmp: 'exists' | 'countGte'; value?: number; withinMinutes?: Minutes }        // 行動履歴
  | { op: 'elapsed'; since: 'friendAdd' | 'flowStart' | 'lastDelivery';
      cmp: 'gte' | 'lt'; minutes: Minutes }                                        // 経過時間
  | { op: 'delivered'; target: { messageId?: Id; flowId?: Id; nodeId?: Id };
      done: boolean; withinMinutes?: Minutes };                                    // 配信済み判定
```

保存例（「東京希望 かつ 応募していない かつ ブロックしていない」）：

```json
{ "op": "and", "children": [
  { "op": "answer", "surveyId": "sv_hope", "questionId": "q_area", "match": "any", "choiceIds": ["c_tokyo"] },
  { "op": "cv", "cvId": "cv_apply", "done": false },
  { "op": "not", "child": { "op": "friend", "state": "blocked" } }
]}
```

### 5.2 判定の木（除外理由を出すための仕組み）

```ts
type Trace = {
  op: Expr['op'];
  value: boolean;
  expr: Expr;              // どの式か（画面側がこれを日本語にする）
  actual?: unknown;        // 実際の値（例: 付いているタグ一覧、経過分）
  children?: Trace[];
  shortCircuit?: boolean;  // and/or で評価を打ち切ったか（打ち切らずに全部評価する設定も可）
};

function evaluate(expr: Expr, ctx: EvalContext): Trace;   // 戻り値の .value が結果
```

- **判定は必ず Trace を返す。** 真偽だけ返す関数は作らない（REQUIREMENTS C-3）
- 既定では and/or を**最後まで評価**する（「何が足りなかったか」を全部見せるため）
- 画面は Trace を木で表示し、`value:false` の葉を赤で示す →「大阪希望なので、東京向けの配信から外れました」
- エンジンは日本語を作らない。`expr` と `actual` を渡し、player 側が案件JSONのラベル（タグ名・選択肢名）を引いて文章にする

---

## 6. 入力イベントとセッション

### 6.1 入力イベント（外から入るものだけ）

```ts
type FriendState = 'not_friend' | 'friend' | 'blocked';

type InputEvent = { at: Minutes; userId: Id } & (
  | { type: 'visitPage'; pageId: Id }
  | { type: 'exitIntent'; pageId: Id }                       // 離脱操作
  | { type: 'dismissPopup'; pageId: Id }
  | { type: 'submitForm'; pageId: Id; values: Record<string, string> }
  | { type: 'addFriend'; routeId: Id }
  | { type: 'block' } | { type: 'unblock' }
  | { type: 'tapMenu'; richMenuId: Id; cellId: Id }
  | { type: 'tapButton'; nodeId: Id; buttonId: Id }
  | { type: 'sendText'; text: string }
  | { type: 'answer'; surveyId: Id; questionId: Id; choiceIds: Id[] }
  | { type: 'cv'; cvId: Id }                                 // 手動でCVを起こす（応募など）
);

type ClockEvent = { at: Minutes; userId: '*'; type: 'advanceTo' };  // 時間を進める（全員共通）
```

- `at` は**全ユーザー共通の仮想時刻**（シミュレーション開始からの分）。Day = `floor(at / 1440)`
- ユーザーごとの「友だち追加からの経過」は、`at − そのユーザーの友だち追加時刻` で計算する
- タグ付与・配信・タイマーは入力イベントではない（エンジンが導出する）

### 6.2 セッション

```ts
type Session = {
  schemaVersion: 1;
  projectId: Id;
  projectHash: string;          // 作ったときの案件の版。違えば「案件が変わっています」と出して再計算
  users: { id: Id; personaId: Id; displayName?: string }[];
  events: (InputEvent | ClockEvent)[];   // 時刻順。これがセッションの唯一の中身
  cursor?: Minutes;             // いま表示している時刻
};
```

---

## 7. 状態（エンジンが導出する）

### 7.1 ユーザー状態

```ts
type UserState = {
  id: Id; personaId: Id; name: string;
  attrs: Record<string, unknown>;
  friendState: FriendState;
  friendAddedAt: Minutes | null;
  route: Id | null;                                     // 友だち追加経路
  tags: Record<Id, { at: Minutes; byNodeId?: Id }>;
  answers: Record<Id, Record<Id, Id[]>>;                // surveyId → questionId → choiceIds
  cvs: Record<Id, { at: Minutes }>;
  behaviors: { at: Minutes; type: string; detail?: Record<string, string> }[];  // 行動履歴
  deliveries: Delivery[];                                // 配信履歴
  flows: Record<Id, FlowCursor>;                         // シナリオ進行状況
  timers: Timer[];                                       // 予定されている配信・待機
  richMenuId: Id | null;
  surface: Surface;                                      // いま画面に出ているもの
  talk: TalkItem[];                                      // トーク画面の吹き出し列
  messageCount: number;                                  // 通数（G-6 採用時）
};

type FlowCursor = { status: 'running' | 'waiting' | 'ended' | 'stopped';
                    nodeId: Id | null; startedAt: Minutes; endedAt?: Minutes; stopReason?: string };
type Timer = { flowId: Id; nodeId: Id; fireAt: Minutes; kind: 'wait' | 'timeout' };
type Delivery = { at: Minutes; flowId: Id; nodeId: Id; messageId: Id; count: number };
type Surface =
  | { kind: 'talk' } | { kind: 'page'; pageId: Id } | { kind: 'popup'; pageId: Id }
  | { kind: 'survey'; surveyId: Id; questionId: Id };
type TalkItem = { at: Minutes; side: 'bot' | 'user' | 'system'; bubble: Bubble; fromNodeId?: Id };
```

### 7.2 判定記録（なぜ届いた／なぜ届かなかった）

```ts
type Decision = {
  at: Minutes; userId: Id; flowId: Id; nodeId: Id;
  kind: 'branch' | 'stopWhen' | 'autoReply' | 'trigger' | 'deliver';
  outcome: 'taken' | 'skipped' | 'excluded' | 'deferred' | 'deduped' | 'stopped';
  trace?: Trace;
  reason?: 'blocked' | 'not_friend' | 'lower_priority' | 'duplicate' | 'stop_condition' | 'condition_false';
};

type WorldState = { now: Minutes; users: Record<Id, UserState>; decisions: Decision[]; warnings: RuntimeWarning[] };
```

「配信対象から外れた理由」は `outcome: 'excluded' | 'deferred' | 'deduped' | 'stopped'` の Decision として残る（FR-33）。
「このメッセージは誰に届く？」の横断一覧（FR-34）は、全ユーザーの deliveries と decisions を messageId で集計して作る。

### 7.3 時間進行のアルゴリズム

```
reduce(project, events):
  world = 初期化(personas)
  for ev in events（時刻順）:
      runTimersUntil(world, ev.at)       ← その時刻までに来るタイマーを先に全部処理
      apply(world, ev)                   ← 入力イベントを処理（トリガー判定・ノード実行）
  return world

runTimersUntil(world, t):
  loop:
    due = fireAt <= t のタイマー
    なければ world.now = t; 終了
    順番 = (fireAt, userId, flow.priority, nodeId) で最小の1つ   ← 決定論（C-7）
    world.now = その fireAt
    そのノードから実行を再開（step）

step(ユーザー, フロー, ノード):
  stopWhen を評価 → 真なら停止して Decision を残す
  ノードを実行（配信・タグ・分岐）
  次のノードへ。wait／入力待ち／end に着くまで繰り返す
  1回の進行で実行したノード数が上限（10,000）を超えたら停止し、無限ループとして報告
```

配信時の判定順（どれかで止まれば配信しない）：

```
1. ブロック中 / 友だちでない            → excluded（blocked / not_friend）
2. 停止条件を満たした                   → stopped
3. 重複防止期間内に同じ messageId を配信済み → deduped
4. exclusive で、より優先度の高いフローが同時刻に配信した → deferred（lower_priority）
5. 配信する → Delivery と TalkItem を追加、通数を加算
```

---

## 8. 検証（validate）

```ts
type Issue = { level: 'error' | 'warning'; code: string; path: string; params?: Record<string, unknown> };
function validate(project: Project): Issue[];
```

`path` は JSON Pointer（例 `/flows/2/nodes/5/messageId`）。編集画面はこれで該当ノードに飛ぶ。メッセージ文はエンジンでは作らず、`code` から player／editor 側が日本語にする。

### 8.1 重大エラー（書き出し禁止）

| code | 中身 |
|---|---|
| `edge.danglingTarget` | エッジの行き先ノードが存在しない |
| `edge.invalidPort` | そのノード種別に無いポートから出ている（例: `message` から `case:x`） |
| `ref.missing` | 参照先が無い（messageId・surveyId・pageId・tagId・assetId・flowId・routeId・cvId） |
| `expr.invalid` | 条件式が壊れている（未知の op・子が空の and/or・型の不一致） |
| `flow.noEntry` | フローに entry が無い／2つ以上ある |
| `flow.zeroTimeCycle` | 待機（1分以上）も入力待ちも含まないループがある（C-1） |
| `branch.noExit` | 分岐ノードに else が無く、どの case にも当てはまらない場合の行き場が無い |
| `node.requiredMissing` | 必須設定が空（メッセージの吹き出しが0個、アンケートの選択肢が0個、メニューのセルに動作が無い、wait の時間が負 など） |
| `flow.friendRequired` | 友だちでないユーザーが通る経路（ポップアップ・サンクス）から、友だち追加を経ずにLINE配信ノードへ行く |
| `prov.unapproved` | 未承認の AI 推定（`prov.state = 'inferred'`）が残っている（FR-82・追加提案） |
| `asset.fileMissing` | 画像ファイルが無い |

### 8.2 警告（書き出し可）

| code | 中身 |
|---|---|
| `tag.unused` | 定義したが、どこからも付与も参照もされないタグ |
| `node.unreachable` | entry から辿り着けないノード |
| `optional.empty` | 任意項目が未設定（説明文が無い・画像の alt が無い 等） |
| `delivery.possibleDuplicate` | 同じ messageId に届く経路が2本以上あり、重複防止期間が0か、停止条件が無い |
| `flow.noStop` | 停止条件の無いステップ配信（CV済みにも送り続ける可能性） |
| `persona.noneMatch` | どのペルソナも入口に入らないフロー（デモで見せられない） |
| `line.limit` | LINE側の上限超え（吹き出し6個以上・ステップ365日超・ボタン数超過 など） |
| `export.large` | 画像の合計が大きく、書き出しが5MBを超えそう |

---

## 9. 文面の差し込み

```
{{user.name}}              ユーザーの名前
{{user.attr.area}}         属性
{{answer.sv_hope.q_area}}  アンケート回答のラベル
{{brand.accountName}}      アカウント名
```

未定義のキーはそのまま残し、警告 `text.unknownVar` を出す。

---

## 10. シナリオ差分（AI連携の境界）

AI（初期版は Claude Code、将来はツール内AI）が出すのはこの形だけ。案件JSONを丸ごと書き換えさせない。

```ts
type ScenarioPatch = {
  schemaVersion: 1;
  basedOn?: string;                    // 対象の projectHash
  summary: string;                     // 何をしたかの一文
  ops: {
    op: 'add' | 'replace' | 'remove';
    path: string;                      // JSON Pointer（"/flows/0/nodes/-" など）
    value?: unknown;
    prov: Provenance;                  // observed か inferred。approved はAIが付けてはいけない
  }[];
};

function applyPatch(project: Project, patch: ScenarioPatch): { project: Project; issues: Issue[] };
```

流れ：

```
差分を読み込む → 画面で差分をプレビュー（追加は緑・置換は黄・削除は赤、出所バッジつき）
  → 項目ごとに 承認／却下 → 承認分だけ本体に反映し、inferred を approved に書き換える
  → 却下分は捨てる（履歴には残す）
```

---

## 11. 検証案件（サンプルスタッフ）を当てはめた例

```
タグ      : area_tokyo, area_osaka, job_office, job_light, applied
CV        : cv_apply（応募）
経路      : route_exit_popup（離脱ポップアップ）, route_thanks（応募サンクス）
アンケート : sv_hope = q_area（東京／大阪）, q_job（事務／軽作業）
           選択肢ごとに addTags: c_tokyo→area_tokyo …

flow_welcome   trigger: friendAdd（全経路） priority:10 exclusive
  entry → message(挨拶) → survey(sv_hope) → branch
        ├ case:tokyo  when hasTag(area_tokyo) → jump(flow_step_tokyo)
        ├ case:osaka  when hasTag(area_osaka) → jump(flow_step_osaka)
        └ else                               → jump(flow_step_common)

flow_step_tokyo  trigger: manual  priority:20  stopWhen: cv(cv_apply, done:true)
  entry → wait(dayAt 1 09:00) → message(東京の求人) → wait(dayAt 3 09:00)
        → branch [case: answer(q_job=事務)] → message(事務求人) / else → message(軽作業求人)
        → wait(dayAt 7 09:00) → message(再応募のお誘い) → end

flow_after_apply trigger: cv(cv_apply)  priority:5
  entry → tag(add applied) → message(応募ありがとう) → wait(dayAt 14 09:00)
        → branch [case: did(tapButton, 再応募)] … → end
```

ペルソナ：

| ID | 名前 | 入口 | 回答 | 期待される違い |
|---|---|---|---|---|
| p_tokyo | 東京希望さん | 経路A（離脱ポップアップ） | 東京・事務 | 東京のステップが届く |
| p_osaka | 大阪希望さん | 経路A | 大阪・軽作業 | 大阪のステップが届き、東京分は「条件外」で除外理由が出る |
| p_applied | 応募済みさん | 経路B（応募サンクス） | 東京・事務 | 応募済みなので東京ステップは「停止条件」で止まる |
| p_notyet | 未応募さん | 経路B 途中離脱 → 経路A | 回答しない | アンケート未回答で else（共通）に流れる |

この表の「期待される違い」が Phase 2 の検収条件になる（ROADMAP 参照）。
