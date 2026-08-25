/**
 * タスク管理シート セットアップスクリプト (Google Apps Script)
 *
 * ■ 使い方
 * 1. Google スプレッドシートを新規作成(または既存のものを開く)
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. このコードを丸ごと貼り付けて保存
 * 4. 上部の関数選択で「setupTaskSheet」を選んで「実行」
 *    (初回は権限の承認ダイアログが出るので許可してください)
 * 5. 毎朝のカレンダー通知を使う場合は、続けて「setupNotificationTrigger」を実行
 *    (スプレッドシートを開き直すと「タスク管理」メニューからも実行できます)
 *
 * ■ できあがるもの
 * - 「タスク管理シート」: メインのタスク一覧
 *   列: タスクの種類 / タスク名 / 発生日 / 完了予定日 / ガチの期日 /
 *        重要度 / 予定工数 / 実際の工数 / 納品物・所感 / そこからの学び / ステータス
 * - 「工数集計」: 種類別・月別の予定工数/実際工数の集計(自動計算)
 *
 * ■ 色分けルール(条件付き書式なので自動で変わります)
 * 優先度の高い順に:
 * 1. ガチの期日を超過した未完了タスク → タスク名が濃い赤+白文字
 * 2. 完了予定日まで 当日・超過→赤 / 3日以内→オレンジ / 7日以内→黄 (タスク名セル)
 * 3. ステータス「完了」の行 → 行全体がグレー+取り消し線
 * 4. タスクの種類ごとに行全体を色分け
 *
 * ■ 自動機能
 * - タスク名(B列)を入力すると、発生日(C列)が空なら今日の日付が自動で入る
 * - setupNotificationTrigger 実行後は、毎朝8時台に期日をチェックして、
 *   「完了予定日が3日以内 or 超過」「ガチの期日超過」の未完了タスクがあれば
 *   Googleカレンダーに通知付きの予定を自動作成(スマホのカレンダー通知が届く)
 *
 * ■ カスタマイズ
 * - タスクの種類や色 → TASK_TYPES を編集して setupTaskSheet を再実行
 * - 色が変わる日数の区切り → DEADLINE_RULES を編集して再実行
 * - 通知する残り日数 → NOTIFY_DAYS_AHEAD を編集
 * - setupTaskSheet は何度実行しても入力済みのデータは消えません
 *   (書式・プルダウン・色分けルール・集計シートだけ設定し直します)
 */

// ===== 設定(ここを編集すれば自由にカスタマイズできます) =====

const SHEET_NAME = 'タスク管理シート';
const SUMMARY_SHEET_NAME = '工数集計';

// タスクの種類と行の背景色(薄めの色を推奨)
const TASK_TYPES = [
  { name: '提案資料作成',     color: '#cfe2f3' }, // 青
  { name: '分析・リサーチ',   color: '#d9ead3' }, // 緑
  { name: '企画・ライティング', color: '#d9d2e9' }, // 紫
  { name: '商談・MTG',        color: '#fce5cd' }, // オレンジ
  { name: '事務作業',         color: '#f3f3f3' }, // グレー
  { name: '学習・研修',       color: '#fff2cc' }, // 黄
  { name: 'その他',           color: '#ead1dc' }, // ピンク
];

// 完了予定日までの残り日数と「タスク名」セルの色(上にあるルールほど優先)
const DEADLINE_RULES = [
  { days: 0, color: '#e06666' }, // 当日・超過 → 赤
  { days: 3, color: '#f6b26b' }, // 残り3日以内 → オレンジ
  { days: 7, color: '#ffd966' }, // 残り7日以内 → 黄
];

// ガチの期日を超過した未完了タスクの「タスク名」の見た目(最優先)
const HARD_DEADLINE_STYLE = { background: '#cc0000', fontColor: '#ffffff' };

// 完了した行の見た目
const DONE_STYLE = { background: '#efefef', fontColor: '#999999' };

const IMPORTANCE_OPTIONS = ['高', '中', '低'];
const STATUS_OPTIONS = ['未着手', '進行中', '完了'];
const DONE_STATUS = '完了';

// カレンダー通知: 完了予定日が今日から何日以内なら通知対象にするか
const NOTIFY_DAYS_AHEAD = 3;
// 通知用に作る予定のタイトル先頭(同じ日の古い通知予定はこれを目印に消して作り直します)
const NOTIFY_EVENT_PREFIX = '⚠タスク期日チェック';

const HEADERS = [
  'タスクの種類',   // A
  'タスク名',       // B
  '発生日',         // C
  '完了予定日',     // D
  'ガチの期日',     // E
  '重要度',         // F
  '予定工数',       // G
  '実際の工数',     // H
  '納品物・所感',   // I
  'そこからの学び', // J
  'ステータス',     // K
];

const DATA_ROWS = 1000; // 書式・ルールを適用する行数

// ===== 日常(私生活)verの設定 =====

const LIFE_SHEET_NAME = '日常タスク管理シート';

// 日常タスクの種類と行の背景色
const LIFE_TASK_TYPES = [
  { name: '家事',         color: '#d9ead3' }, // 緑
  { name: '買い物',       color: '#cfe2f3' }, // 青
  { name: '手続き・お金', color: '#fce5cd' }, // オレンジ
  { name: '健康・運動',   color: '#d0e0e3' }, // 水色
  { name: '趣味・遊び',   color: '#d9d2e9' }, // 紫
  { name: '人付き合い',   color: '#ead1dc' }, // ピンク
  { name: 'その他',       color: '#f3f3f3' }, // グレー
];

const LIFE_HEADERS = [
  'タスクの種類', // A
  'タスク名',     // B
  '考案日',       // C
  '需要度',       // D
];

const DEMAND_OPTIONS = ['高', '中', '低'];

// ===== 既存提案管理シートの設定 =====
// 「今すでに運用している顧客」に対する追加提案(アップセル)を管理するシート

const DEAL_SHEET_NAME = '既存提案管理シート';

// 提案フェーズと行の背景色(上ほど初期、下ほど成約に近い)
const DEAL_PHASES = [
  { name: '①ネタ出し',   color: '#f3f3f3' }, // グレー
  { name: '②社内確認',   color: '#fff2cc' }, // 黄
  { name: '③提案準備',   color: '#fce5cd' }, // オレンジ
  { name: '④提案済み',   color: '#cfe2f3' }, // 青
  { name: '⑤検討中',     color: '#d0e0e3' }, // 水色
  { name: '⑥受注',       color: '#d9ead3' }, // 緑
  { name: '⑦見送り',     color: '#efefef' }, // 薄グレー(取り消し線)
  { name: '⑧保留',       color: '#ead1dc' }, // ピンク
];

const DEAL_LOST_PHASE = '⑦見送り';
const DEAL_WON_PHASE = '⑥受注';

// ヨミ(受注確度)
const DEAL_PROBABILITY_OPTIONS = ['A(ほぼ確実)', 'B(有力)', 'C(可能性あり)', 'D(薄い)'];

const DEAL_HEADERS = [
  '顧客名',             // A
  '現契約・運用中サービス', // B
  '現在の月額',         // C
  '契約開始日',         // D
  '提案テーマ(追加提案)', // E
  '提案のきっかけ・根拠', // F
  'フェーズ',           // G
  '提案予定日',         // H
  '次アクション',       // I
  '次アクション期限',   // J
  'ヨミ',               // K
  '想定初期費用',       // L
  '想定月額(増分)',     // M
  '想定年間増収',       // N
  'メモ',               // O
  '最終更新日',         // P
];

// 次アクション期限までの残り日数と「次アクション」セルの色
const DEAL_ACTION_RULES = [
  { days: 0, color: '#e06666' }, // 当日・超過 → 赤
  { days: 3, color: '#f6b26b' }, // 残り3日以内 → オレンジ
  { days: 7, color: '#ffd966' }, // 残り7日以内 → 黄
];

// ===== ここから下は基本的に編集不要 =====

// ---- メニュー ----

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('タスク管理')
    .addItem('初期セットアップ(再実行OK)', 'setupTaskSheet')
    .addItem('毎朝のカレンダー通知を有効化', 'setupNotificationTrigger')
    .addItem('今すぐ期日チェック', 'notifyDeadlinesManual')
    .addToUi();
}

// ---- セットアップ本体 ----

function setupTaskSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  const isNewSheet = !sheet;
  if (isNewSheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  setupHeader_(sheet);
  setupColumnFormats_(sheet);
  setupValidations_(sheet);
  setupConditionalFormats_(sheet);
  setupSummarySheet_(ss);

  if (isNewSheet) {
    insertSampleRows_(sheet);
  }

  setupLifeSheet_(ss);
  setupDealSheet_(ss);

  ss.setActiveSheet(sheet);
  alert_('「' + SHEET_NAME + '」「' + LIFE_SHEET_NAME + '」「' + DEAL_SHEET_NAME + '」「' + SUMMARY_SHEET_NAME +
    '」のセットアップが完了しました。\n' +
    '毎朝のカレンダー通知を使う場合は「タスク管理」メニューから有効化してください。');
}

// ---- 既存提案管理シートのセットアップ ----
// すでに運用している顧客に対する「追加提案(アップセル)」を管理するシート。
// 1顧客に複数の提案テーマがある場合は、顧客名を繰り返して複数行で持つ。
// フェーズごとの行色分けと、次アクション期限が近づくとそのセルが色づく
function setupDealSheet_(ss) {
  let sheet = ss.getSheetByName(DEAL_SHEET_NAME);
  const isNewSheet = !sheet;
  if (isNewSheet) {
    sheet = ss.insertSheet(DEAL_SHEET_NAME);
  }
  const rows = DATA_ROWS;
  const numCols = DEAL_HEADERS.length;
  const lastRow = rows + 1;

  // ヘッダー
  sheet.getRange(1, 1, 1, numCols)
    .setValues([DEAL_HEADERS])
    .setBackground('#434343')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1); // 顧客名は常に見えるように

  // 列幅・表示形式
  const widths = [140, 200, 95, 95, 240, 260, 105, 95, 240, 105, 105, 105, 110, 110, 240, 95];
  widths.forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
  sheet.getRange(2, 3, rows, 1).setNumberFormat('¥#,##0');      // 現在の月額
  sheet.getRange(2, 4, rows, 1).setNumberFormat('yyyy/mm/dd');  // 契約開始日
  sheet.getRange(2, 8, rows, 1).setNumberFormat('yyyy/mm/dd');  // 提案予定日
  sheet.getRange(2, 10, rows, 1).setNumberFormat('yyyy/mm/dd'); // 次アクション期限
  sheet.getRange(2, 12, rows, 3).setNumberFormat('¥#,##0');     // 想定初期費用・想定月額・想定年間増収
  sheet.getRange(2, 16, rows, 1).setNumberFormat('yyyy/mm/dd'); // 最終更新日
  [5, 6, 9, 15].forEach(function (col) {
    sheet.getRange(2, col, rows, 1).setWrap(true); // 提案テーマ・きっかけ・次アクション・メモ
  });
  [3, 4, 7, 8, 10, 11, 16].forEach(function (col) {
    sheet.getRange(2, col, rows, 1).setHorizontalAlignment('center');
  });

  // 想定年間増収 = 想定初期費用 + 想定月額×12 を自動計算
  const revenueFormulas = [];
  for (let r = 2; r <= lastRow; r++) {
    revenueFormulas.push(['=IF(AND($L' + r + '="",$M' + r + '=""),"",N($L' + r + ')+N($M' + r + ')*12)']);
  }
  sheet.getRange(2, 14, rows, 1).setFormulas(revenueFormulas);

  // プルダウン
  const phaseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEAL_PHASES.map(function (p) { return p.name; }), true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 7, rows, 1).setDataValidation(phaseRule);

  const probRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEAL_PROBABILITY_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 11, rows, 1).setDataValidation(probRule);

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('日付を入力してください (例: 2026/08/28)')
    .build();
  sheet.getRange(2, 4, rows, 1).setDataValidation(dateRule);
  sheet.getRange(2, 8, rows, 1).setDataValidation(dateRule);
  sheet.getRange(2, 10, rows, 1).setDataValidation(dateRule);

  // 条件付き書式(先に登録したルールほど優先)
  const rules = [];
  const actionRange = sheet.getRange('I2:I' + lastRow);
  const rowRange = sheet.getRange('A2:P' + lastRow);

  // 1. 次アクション期限が近い/超過 → 次アクションのセルを色づけ(受注・見送り済みは対象外)
  DEAL_ACTION_RULES.forEach(function (r) {
    const formula = '=AND($J2<>"", $G2<>"' + DEAL_WON_PHASE + '", $G2<>"' + DEAL_LOST_PHASE + '", $J2<=TODAY()+' + r.days + ')';
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(r.color)
        .setBold(true)
        .setRanges([actionRange])
        .build()
    );
  });

  // 2. 見送りになった行 → グレー+取り消し線
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2="' + DEAL_LOST_PHASE + '"')
      .setBackground('#efefef')
      .setFontColor('#999999')
      .setStrikethrough(true)
      .setRanges([rowRange])
      .build()
  );

  // 3. フェーズごとの行色分け
  DEAL_PHASES.forEach(function (p) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$G2="' + p.name + '"')
        .setBackground(p.color)
        .setRanges([rowRange])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);

  // 運用中の顧客を初期投入(新規作成時のみ。不要なら行ごと削除してOK)
  // ※提案テーマ以降は空欄。追加提案のネタが決まったら埋めていく想定
  if (isNewSheet) {
    const seed = [
      ['ソウガク', 'LINE公式アカウント運用(配信企画・投稿制作)'],
      ['インクアート', 'LINE公式アカウント運用'],
      ['ノーストクリニック', 'LINE公式アカウント運用'],
    ];
    seed.forEach(function (row, i) {
      const r = 2 + i;
      sheet.getRange(r, 1, 1, 2).setValues([row]);
      sheet.getRange(r, 7).setValue(DEAL_PHASES[0].name); // ①ネタ出し
      sheet.getRange(r, 16).setValue(new Date());
    });
  }
}

// ---- 日常(私生活)verのセットアップ ----
// 列は「タスクの種類 / タスク名 / 考案日 / 需要度」だけのシンプル構成。
// 種類ごとの行色分け・プルダウン・考案日の自動入力が効く(期日色分け・通知・工数集計は対象外)
function setupLifeSheet_(ss) {
  let sheet = ss.getSheetByName(LIFE_SHEET_NAME);
  const isNewSheet = !sheet;
  if (isNewSheet) {
    sheet = ss.insertSheet(LIFE_SHEET_NAME);
  }
  const rows = DATA_ROWS;
  const numCols = LIFE_HEADERS.length;

  // ヘッダー
  sheet.getRange(1, 1, 1, numCols)
    .setValues([LIFE_HEADERS])
    .setBackground('#434343')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);

  // 列幅・表示形式
  const widths = [110, 320, 90, 60];
  widths.forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
  sheet.getRange(2, 3, rows, 1).setNumberFormat('yyyy/mm/dd'); // 考案日
  [1, 3, 4].forEach(function (col) {
    sheet.getRange(2, col, rows, 1).setHorizontalAlignment('center');
  });

  // プルダウン
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LIFE_TASK_TYPES.map(function (t) { return t.name; }), true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 1, rows, 1).setDataValidation(typeRule);

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('日付を入力してください (例: 2026/08/22)')
    .build();
  sheet.getRange(2, 3, rows, 1).setDataValidation(dateRule);

  const demandRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEMAND_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, rows, 1).setDataValidation(demandRule);

  // 種類ごとの行色分け
  const rowRange = sheet.getRange('A2:D' + (rows + 1));
  const rules = LIFE_TASK_TYPES.map(function (t) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="' + t.name + '"')
      .setBackground(t.color)
      .setRanges([rowRange])
      .build();
  });
  sheet.setConditionalFormatRules(rules);

  // 動作確認用サンプル(新規作成時のみ。不要なら行ごと削除でOK)
  if (isNewSheet) {
    sheet.getRange(2, 1, 2, numCols).setValues([
      ['家事', '(サンプル) 部屋の大掃除', new Date(), '中'],
      ['趣味・遊び', '(サンプル) 行きたいカフェをリストアップ', new Date(), '低'],
    ]);
  }
}

// ヘッダー行の作成と固定
function setupHeader_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setValues([HEADERS])
    .setBackground('#434343')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);
}

// 列幅・表示形式
function setupColumnFormats_(sheet) {
  const widths = [110, 240, 90, 90, 90, 60, 70, 70, 280, 280, 80];
  widths.forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  const rows = DATA_ROWS;
  // 日付列 (C:発生日, D:完了予定日, E:ガチの期日)
  sheet.getRange(2, 3, rows, 3).setNumberFormat('yyyy/mm/dd');
  // 工数列 (G, H) : 時間(h)
  sheet.getRange(2, 7, rows, 2).setNumberFormat('0.0"h"');
  // 長文列 (I, J) は折り返し表示
  sheet.getRange(2, 9, rows, 2).setWrap(true);
  // 中央寄せにする列 (A, C, D, E, F, K)
  [1, 3, 4, 5, 6, 11].forEach(function (col) {
    sheet.getRange(2, col, rows, 1).setHorizontalAlignment('center');
  });
}

// プルダウン(入力規則)
function setupValidations_(sheet) {
  const rows = DATA_ROWS;

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TASK_TYPES.map(function (t) { return t.name; }), true)
    .setAllowInvalid(true) // リスト外の値も一応許可(後から種類を増やしても困らないように)
    .build();
  sheet.getRange(2, 1, rows, 1).setDataValidation(typeRule);

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('日付を入力してください (例: 2026/08/22)')
    .build();
  sheet.getRange(2, 3, rows, 3).setDataValidation(dateRule);

  const importanceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(IMPORTANCE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, rows, 1).setDataValidation(importanceRule);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 11, rows, 1).setDataValidation(statusRule);
}

// 条件付き書式(色分けルール)。先に登録したルールほど優先される
function setupConditionalFormats_(sheet) {
  const rules = [];
  const lastRow = DATA_ROWS + 1;
  const taskNameRange = sheet.getRange('B2:B' + lastRow);
  const rowRange = sheet.getRange('A2:K' + lastRow);

  // --- 1. ガチの期日(E)を超過した未完了タスク → タスク名を濃い赤+白文字 ---
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($E2<>"", $K2<>"' + DONE_STATUS + '", $E2<TODAY())')
      .setBackground(HARD_DEADLINE_STYLE.background)
      .setFontColor(HARD_DEADLINE_STYLE.fontColor)
      .setBold(true)
      .setRanges([taskNameRange])
      .build()
  );

  // --- 2. 完了予定日(D)の残り日数によるタスク名の色 ---
  DEADLINE_RULES.forEach(function (r) {
    const formula =
      '=AND($D2<>"", $K2<>"' + DONE_STATUS + '", $D2<=TODAY()+' + r.days + ')';
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(r.color)
        .setBold(true)
        .setRanges([taskNameRange])
        .build()
    );
  });

  // --- 3. 完了した行 → 行全体グレー+取り消し線(種類の色より優先) ---
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$K2="' + DONE_STATUS + '"')
      .setBackground(DONE_STYLE.background)
      .setFontColor(DONE_STYLE.fontColor)
      .setStrikethrough(true)
      .setRanges([rowRange])
      .build()
  );

  // --- 4. タスクの種類による行全体の色分け ---
  TASK_TYPES.forEach(function (t) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A2="' + t.name + '"')
        .setBackground(t.color)
        .setRanges([rowRange])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules); // 既存ルールを置き換え
}

// 工数集計シート(種類別・月別)。数式で組むので値の変更に自動追従します
function setupSummarySheet_(ss) {
  let sheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_SHEET_NAME);
  }
  sheet.clear();

  const src = "'" + SHEET_NAME + "'";
  const last = DATA_ROWS + 1;
  const typeCol = src + '!$A$2:$A$' + last;
  const dueCol = src + '!$D$2:$D$' + last;
  const planCol = src + '!$G$2:$G$' + last;
  const actualCol = src + '!$H$2:$H$' + last;

  // --- 種類別集計 ---
  // ※セル1個ずつ書き込むと極端に遅くなるため、すべて配列で一括書き込みする
  sheet.getRange('A1').setValue('種類別集計').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A2:D2')
    .setValues([['タスクの種類', '予定工数', '実際の工数', '差分(実際-予定)']])
    .setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  const n = TASK_TYPES.length;
  sheet.getRange(3, 1, n, 1)
    .setValues(TASK_TYPES.map(function (t) { return [t.name]; }))
    .setBackgrounds(TASK_TYPES.map(function (t) { return [t.color]; }));

  const typeFormulas = TASK_TYPES.map(function (t, i) {
    const row = 3 + i;
    return [
      '=SUMIF(' + typeCol + ',$A' + row + ',' + planCol + ')',
      '=SUMIF(' + typeCol + ',$A' + row + ',' + actualCol + ')',
      '=C' + row + '-B' + row,
    ];
  });
  const totalRow = 3 + n;
  typeFormulas.push([
    '=SUM(B3:B' + (totalRow - 1) + ')',
    '=SUM(C3:C' + (totalRow - 1) + ')',
    '=C' + totalRow + '-B' + totalRow,
  ]);
  sheet.getRange(3, 2, n + 1, 3).setFormulas(typeFormulas).setNumberFormat('0.0"h"');
  sheet.getRange(totalRow, 1).setValue('合計');
  sheet.getRange(totalRow, 1, 1, 4).setFontWeight('bold');

  // --- 月別集計(完了予定日ベース・今年の1〜12月) ---
  sheet.getRange('F1').setValue('月別集計(完了予定日ベース・今年)').setFontWeight('bold').setFontSize(12);
  sheet.getRange('F2:H2')
    .setValues([['月', '予定工数', '実際の工数']])
    .setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  const monthFormulas = [];
  for (let m = 1; m <= 12; m++) {
    const row = 2 + m;
    monthFormulas.push([
      '=DATE(YEAR(TODAY()),' + m + ',1)',
      '=SUMIFS(' + planCol + ',' + dueCol + ',">="&$F' + row + ',' + dueCol + ',"<"&EDATE($F' + row + ',1))',
      '=SUMIFS(' + actualCol + ',' + dueCol + ',">="&$F' + row + ',' + dueCol + ',"<"&EDATE($F' + row + ',1))',
    ]);
  }
  sheet.getRange(3, 6, 12, 3).setFormulas(monthFormulas);
  sheet.getRange(3, 6, 12, 1).setNumberFormat('yyyy/mm').setHorizontalAlignment('center');
  sheet.getRange(3, 7, 12, 2).setNumberFormat('0.0"h"');

  const widths = [130, 100, 100, 120, 30, 90, 100, 100];
  widths.forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
}

// ---- 発生日/考案日の自動入力 ----
// タスク名(B列)を入力したとき、C列(仕事: 発生日 / 日常: 考案日)が空なら今日の日付を入れる
function onEdit(e) {
  if (!e || !e.range) return; // エディタから手動実行された場合は何もしない(編集時に自動で動く関数です)
  const range = e.range;
  const sheet = range.getSheet();
  const name_ = sheet.getName();

  // 既存提案管理シートは、行を編集したら「最終更新日」(P列)を今日にする
  if (name_ === DEAL_SHEET_NAME) {
    if (range.getColumn() === 16) return; // 最終更新日自体の編集では動かさない(無限ループ防止)
    const first = Math.max(range.getRow(), 2);
    const last = range.getRow() + range.getNumRows() - 1;
    for (let row = first; row <= last; row++) {
      if (sheet.getRange(row, 1).getValue() !== '') { // 顧客名が入っている行だけ
        sheet.getRange(row, 16).setValue(new Date()).setNumberFormat('yyyy/mm/dd');
      }
    }
    return;
  }

  if (name_ !== SHEET_NAME && name_ !== LIFE_SHEET_NAME) return;

  const startRow = Math.max(range.getRow(), 2);
  const endRow = range.getRow() + range.getNumRows() - 1;
  const startCol = range.getColumn();
  const endCol = startCol + range.getNumColumns() - 1;
  if (endRow < 2 || startCol > 2 || endCol < 2) return; // B列に触れていなければ何もしない

  for (let row = startRow; row <= endRow; row++) {
    const name = sheet.getRange(row, 2).getValue();
    const dateCell = sheet.getRange(row, 3);
    if (name !== '' && dateCell.getValue() === '') {
      dateCell.setValue(new Date()).setNumberFormat('yyyy/mm/dd');
    }
  }
}

// ---- Googleカレンダー通知 ----

// 毎朝8時台に notifyDeadlines を実行するトリガーを設定(再実行すると張り直し)
function setupNotificationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'notifyDeadlines') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('notifyDeadlines').timeBased().everyDays(1).atHour(8).create();
  alert_('毎朝8時台に期日をチェックして、対象タスクがあればGoogleカレンダーに通知予定を作るようにしました。');
}

// メニューの「今すぐ期日チェック」用
function notifyDeadlinesManual() {
  const count = notifyDeadlines();
  alert_(count === 0
    ? '期日が近い・超過している未完了タスクはありませんでした。'
    : count + '件の要対応タスクをGoogleカレンダーの予定にまとめました(約15分後に通知されます)。');
}

/**
 * 期日チェック本体。
 * 「ガチの期日を超過」「完了予定日が超過 or NOTIFY_DAYS_AHEAD日以内」の未完了タスクを集め、
 * デフォルトのGoogleカレンダーに通知(ポップアップリマインダー)付きの予定を1件作る。
 * 対象0件なら何もしない。戻り値は対象タスク数。
 */
function notifyDeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayMs = 24 * 60 * 60 * 1000;

  const hardOverdue = [];
  const dueOverdue = [];
  const dueSoon = [];

  values.forEach(function (row) {
    const name = row[1];
    const due = row[3];   // 完了予定日
    const hard = row[4];  // ガチの期日
    const status = row[10];
    if (name === '' || status === DONE_STATUS) return;

    if (hard instanceof Date) {
      const h = new Date(hard); h.setHours(0, 0, 0, 0);
      if (h < today) {
        hardOverdue.push('・' + name + ' (ガチの期日 ' + formatDate_(h) + ' を超過!)');
        return;
      }
    }
    if (due instanceof Date) {
      const d = new Date(due); d.setHours(0, 0, 0, 0);
      const daysLeft = Math.round((d - today) / dayMs);
      if (daysLeft < 0) {
        dueOverdue.push('・' + name + ' (完了予定日を' + (-daysLeft) + '日超過)');
      } else if (daysLeft <= NOTIFY_DAYS_AHEAD) {
        dueSoon.push('・' + name + ' (完了予定日まであと' + daysLeft + '日)');
      }
    }
  });

  const total = hardOverdue.length + dueOverdue.length + dueSoon.length;
  if (total === 0) return 0;

  const lines = [];
  if (hardOverdue.length) lines.push('【ガチの期日 超過】', hardOverdue.join('\n'), '');
  if (dueOverdue.length) lines.push('【完了予定日 超過】', dueOverdue.join('\n'), '');
  if (dueSoon.length) lines.push('【完了予定日まで' + NOTIFY_DAYS_AHEAD + '日以内】', dueSoon.join('\n'), '');
  lines.push('シート: ' + ss.getUrl());

  const cal = CalendarApp.getDefaultCalendar();
  // 同じ日に作った古いチェック予定は消して作り直す(重複防止)
  cal.getEventsForDay(new Date()).forEach(function (ev) {
    if (ev.getTitle().indexOf(NOTIFY_EVENT_PREFIX) === 0) {
      ev.deleteEvent();
    }
  });

  const start = new Date(Date.now() + 15 * 60 * 1000); // 15分後に通知が届くように
  const end = new Date(start.getTime() + 15 * 60 * 1000);
  const event = cal.createEvent(
    NOTIFY_EVENT_PREFIX + ': 要対応 ' + total + '件',
    start,
    end,
    { description: lines.join('\n') }
  );
  event.removeAllReminders();
  event.addPopupReminder(1);

  return total;
}

// ---- 補助 ----

function formatDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

// 完了メッセージの表示。
// getUi().alert() はスプレッドシート側のタブにダイアログを出し、閉じられるまで
// スクリプトが待ち続けて6分でタイムアウトするため、待ちが発生しないトースト通知を使う
function alert_(message) {
  Logger.log(message);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'タスク管理', 10);
  } catch (err) {
    // トリガー実行時など表示できない場面では何もしない(ログには残る)
  }
}

// 動作確認用のサンプル行(新規作成時のみ投入。不要なら行ごと削除してOK)
function insertSampleRows_(sheet) {
  const today = new Date();
  const addDays = function (n) {
    const d = new Date(today);
    d.setDate(d.getDate() + n);
    return d;
  };
  const samples = [
    ['提案資料作成', '(サンプル) LINE公式アカウント提案資料の作成', addDays(-3), addDays(2), addDays(5), '高', 8, '', '', '', '進行中'],
    ['分析・リサーチ', '(サンプル) LP分析レポートまとめ', addDays(-1), addDays(6), addDays(10), '中', 4, '', '', '', '未着手'],
    ['企画・ライティング', '(サンプル) 記事企画案の修正 ※ガチの期日超過の例', addDays(-10), addDays(-5), addDays(-2), '高', 3, '', '', '', '進行中'],
    ['学習・研修', '(サンプル) AIバナー生成ワークの復習', addDays(-7), addDays(-1), addDays(14), '低', 2, 2, '(サンプル) 復習メモを作成', '(サンプル) プロンプトの型が重要', '完了'],
  ];
  sheet.getRange(2, 1, samples.length, HEADERS.length).setValues(samples);
}
