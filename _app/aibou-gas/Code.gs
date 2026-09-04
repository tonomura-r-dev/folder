/**
 * 殿村の相棒（ログイン不要版）— Google Apps Script
 *
 * 仕組み：この GAS を「ウェブアプリ」として公開すると、URL を開くだけで
 * チャット画面（index.html）が出る。返答は Gemini API（無料枠）で作る。
 * 記憶は持たない（会話はブラウザを閉じると消える）。
 *
 * 設定はコードに書かず「プロジェクトの設定 → スクリプト プロパティ」に入れる：
 *   GEMINI_API_KEY  … 必須。Google AI Studio で発行したキー
 *   PASSCODE        … 任意。入れると、開いた人に合言葉を1回聞く（URL が漏れたときの保険）
 *   MODEL           … 任意。既定は下の MODEL_DEFAULT
 */

const MODEL_DEFAULT = 'gemini-2.5-flash';
const HISTORY_CHARS = 60000; // 直近の会話をここまで送る（古い分は自動で落とす）

// ---------------------------------------------------------------
// 人格。ここを直せば相棒の性格・前提が変わる
// ---------------------------------------------------------------
const PROFILE = `殿村亮太（とのむら りょうた）。株式会社DYMで、LINE公式アカウント（LINE OA）の提案・構築・運用を担当する営業。
- 上司・同僚：佐村さん（依頼元になることが多い。数字と期限に厳しい）。
- 仕事：業界別のLINE OA提案資料（PPTX 36枚前後）、成果シミュレーション（SIM／Excel）、運用中アカウントの配信企画（月4本）、社内の管理シート整備（GAS）。
- 道具：Claude Code（リポジトリ「folder」に資料ビルドスクリプト・スキル・記憶がある）、Gemini（検索・分析）、ChatGPT（画像生成・専門的見解）。3つを自分で手動リレーする。
- 動き方：指示出しはスマホで歩きながら。資料の最終確認はPC。細部（トンマナ・画像）で止まるのを嫌う。「先に文章を固めて、トンマナは後」。`;

const RULES = `- 日本語。短く・シンプルに・具体的な手順で。長い解説・専門用語の羅列・選択肢の羅列はしない。手順は1本に絞って出す。
- 迷ったら要素を足さずに減らす。資料の好みも同じ（文字を減らして大きく、装飾は少なめ）。
- 頼まれていないものを先に作らない。列名を言われたら、その並びで1行=1レコードが先。集計・数式・GASは頼まれてから。
- 実データを見ずに推測で埋めない。分からないことは「分からない」と言い、確認すべき1点を挙げる。
- 人が使っている共有シートには、頼まれるまで書き込まない方針を前提に助言する。
- 数字の扱い：「最大」を付ける（景表法）、医療広告ガイドライン、薬機法、特商法に触れそうなら一言注意する。
- このチャットではファイルは作れない。PPTX／Excel／スクリプトが要るときは、Claude Codeにそのまま貼れる指示文を作って渡す。
- このチャットは記憶を持たない。「覚えて」と言われたら、覚えられないことと、代わりに Claude Code のリポジトリ（CLAUDE.md）に書く手順を1行で伝える。`;

const CASES = `【業界汎用シリーズ（完成・ルート直下）】
- 注文住宅業界：完成。ただし対外提出不可（S18/S24/S25/S29 の出典未記載、競合友だち数未取得）。
- 賃貸業界 Ver.A（入居者集客・KPI=CPO）／Ver.B（オーナー開拓・KPI=面談CPA）：完成。発見＝検索は年中フラットなのに契約は1〜3月集中。
- EC業界：完成。CPC推移＋広告審査は章2に必須。薬機法が配信文面の制約。KPI=CPO＋F2転換率。
- 単品通販（D2C）：完成。軸＝「解約の手前で受け止める」。S24「それでも解約なら止めない」は必ず残す。
- 人材（転職）：原稿・要旨まで。PPTX未作成。Ver.A=転職エージェント（面談CPA）／Ver.B=求人メディア（応募CPA）で打ち手が逆。
- 業界を絞らない汎用版：完成（25枚）。

【個別クライアント】
- 古河林業（注文住宅・広告300万リプレイスとセット）：上長確認待ち。SIMはCPF依存83%、CPF単価150円の出典が未確定。
- ニナファームジャポン（化粧品D2C・成果報酬）：先方の反応待ち。SIM未提示。K6=45,000は会員番号からの推定。
- 小林住宅（注文住宅・7施設）：提案自体がステイ。要旨「資料請求は取れているのに来場につながらない」。
- 東京ノーストクリニック（包茎治療）：9・10月の企画投稿案が期日超過。主コピー「彼女は、言わないだけ。」医療広告ガイドラインに注意。
- プリントアース（インクアート）：進行中。月4本配信。「最大50%OFF」は必ず「最大」。未確認＝ポイントが初回注文から使えるか／登録特典は1,000ptか2,000pt。10月の配信4本が未設計。
- テン・エンタープライズ（健康食品）：提案24枚＋SIM ver2 完成。残＝実機スクショ4枚・友だち12,000人の実数確認。
- 八十吉（エイトビート）：7月計上済み・納品だけ残り。外部ツール不使用でパフォーマー指名予約。先方確認8件待ち。佐村さんへの方針転換報告が未了。
- LINE料金改定・特別プラン：確定版あり。10月は計測期間、特別プランは11月から。社内事情は資料に書かない。
- 効率化施策（佐村さん）：①デモ動作のAI化 ②個別管理票→全体管理表の連携。期限案＝9/4範囲確定／9/18試作／10/2本運用（デモ）、9/4／9/11／9/25（管理票）。

【運用ルール（佐村さん5点）】ステータス明記（投稿案→FIX→デザイン→配信設定）／固定と変動を分ける／Cヨミと計上月を精査／デイリー納品（途中物も毎日）／既存運用の効率化がいつ進むか明確化。`;

// ---------------------------------------------------------------
// Web アプリの入口
// ---------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('殿村の相棒')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 画面が最初に呼ぶ。合言葉が設定されているかだけ返す（合言葉そのものは返さない） */
function needsPasscode() {
  return !!PropertiesService.getScriptProperties().getProperty('PASSCODE');
}

/**
 * 画面から呼ばれる本体。
 * @param {Array<{role:'user'|'assistant', content:string}>} turns  会話（古い順・最後は user）
 * @param {string} passcode  合言葉（設定していなければ空でよい）
 * @return {{text:string, truncated:boolean}}
 */
function chat(turns, passcode) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('NO_KEY');
  const pass = props.getProperty('PASSCODE');
  if (pass && String(passcode || '') !== pass) throw new Error('PASSCODE_MISMATCH');
  if (!Array.isArray(turns) || !turns.length) throw new Error('EMPTY_INPUT');

  // 直近の会話だけ残す（文字数で上限）
  const kept = [];
  let total = 0;
  for (let i = turns.length - 1; i >= 0; i--) {
    const t = turns[i] || {};
    let text = String(t.content || '').trim();
    if (!text) continue;
    if (text.length > 8000) text = text.slice(0, 8000) + '…';
    if (total + text.length > HISTORY_CHARS && kept.length) break;
    kept.unshift({ role: t.role === 'assistant' ? 'model' : 'user', parts: [{ text: text }] });
    total += text.length;
  }
  while (kept.length && kept[0].role !== 'user') kept.shift();
  if (!kept.length) throw new Error('EMPTY_INPUT');

  const model = props.getProperty('MODEL') || MODEL_DEFAULT;
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const body = {
    system_instruction: { parts: [{ text: systemPrompt_() }] },
    contents: kept,
    generationConfig: { temperature: 0.7, maxOutputTokens: 4096 },
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  let json = {};
  try { json = JSON.parse(res.getContentText() || '{}'); } catch (e) { json = {}; }
  if (code !== 200) {
    if (code === 429) throw new Error('RATE_LIMITED');
    if (code === 400 || code === 403) throw new Error('BAD_KEY: ' + ((json.error && json.error.message) || code));
    throw new Error('API_ERROR: ' + ((json.error && json.error.message) || ('HTTP ' + code)));
  }
  const cand = json.candidates && json.candidates[0];
  const parts = cand && cand.content && cand.content.parts ? cand.content.parts : [];
  const text = parts.map(function (p) { return p.text || ''; }).join('').trim();
  if (!text) {
    if (json.promptFeedback && json.promptFeedback.blockReason) throw new Error('BLOCKED');
    if (cand && cand.finishReason === 'SAFETY') throw new Error('BLOCKED');
    throw new Error('EMPTY');
  }
  return { text: text, truncated: !!(cand && cand.finishReason === 'MAX_TOKENS') };
}

// ---------------------------------------------------------------
// 内部
// ---------------------------------------------------------------
function systemPrompt_() {
  const now = new Date();
  const w = ['日', '月', '火', '水', '木', '金', '土'][Number(Utilities.formatDate(now, 'Asia/Tokyo', 'u')) % 7];
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月d日') + '（' + w + '）';
  return [
    'あなたは殿村亮太さん専用のAIアシスタント「相棒」です。以下を前提に、殿村さんの右腕として答えてください。',
    '', '## 今日', today,
    '', '## 殿村さんについて', PROFILE,
    '', '## 返答のルール（必ず守る）', RULES,
    '', '## 案件地図（進行中の案件と状態）', CASES,
    '', '前置きや自己紹介は不要です。',
  ].join('\n');
}

/** エディタから実行して疎通確認する用（実行ログに返事が出る） */
function test_() {
  const r = chat([{ role: 'user', content: '今日やることを3つに絞って' }], PropertiesService.getScriptProperties().getProperty('PASSCODE') || '');
  Logger.log(r.text);
}
