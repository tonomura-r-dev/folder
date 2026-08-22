/**
 * タスク管理シート セットアップスクリプト (Google Apps Script)
 *
 * ■ 使い方
 * 1. Google スプレッドシートを新規作成(または既存のものを開く)
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. このコードを丸ごと貼り付けて保存
 * 4. 上部の関数選択で「setupTaskSheet」を選んで「実行」
 *    (初回は権限の承認ダイアログが出るので許可してください)
 *
 * ■ できあがるもの
 * - 「タスク管理シート」という名前のシート
 * - 列: タスクの種類 / タスク名 / 発生日 / 完了予定日 / ガチの期日 /
 *        重要度 / 予定工数 / 実際の工数 / 納品物・所感 / そこからの学び / ステータス
 * - タスクの種類ごとに行全体を色分け(条件付き書式なので自動で色が付く)
 * - 完了予定日が近づくと「タスク名」のセルだけ 黄 → オレンジ → 赤 に変化
 *   (ステータスが「完了」になると色は消える)
 * - タスクの種類 / 重要度 / ステータス はプルダウン選択
 *
 * ■ カスタマイズ
 * - タスクの種類や色を変えたい → 下の TASK_TYPES を編集して setupTaskSheet を再実行
 * - 色が変わる日数の区切りを変えたい → DEADLINE_RULES を編集して再実行
 * - setupTaskSheet は何度実行しても入力済みのデータは消えません
 *   (書式・プルダウン・色分けルールだけ設定し直します)
 */

// ===== 設定(ここを編集すれば自由にカスタマイズできます) =====

const SHEET_NAME = 'タスク管理シート';

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

const IMPORTANCE_OPTIONS = ['高', '中', '低'];
const STATUS_OPTIONS = ['未着手', '進行中', '完了'];
const DONE_STATUS = '完了'; // この状態になるとタスク名の期日色が消える

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

// ===== ここから下は基本的に編集不要 =====

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

  if (isNewSheet) {
    insertSampleRows_(sheet);
  }

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('「' + SHEET_NAME + '」のセットアップが完了しました。');
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
  // 工数列 (G, H) : 数値(時間)想定
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

// 条件付き書式(色分けルール)
function setupConditionalFormats_(sheet) {
  const rules = [];
  const lastRow = DATA_ROWS + 1;

  // --- 1. タスク名セルの期日色 (先に登録したルールが優先される) ---
  // 完了予定日(D)が入っていて、ステータス(K)が「完了」でない行が対象
  const taskNameRange = sheet.getRange('B2:B' + lastRow);
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

  // --- 2. タスクの種類による行全体の色分け ---
  const rowRange = sheet.getRange('A2:K' + lastRow);
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
    ['学習・研修', '(サンプル) AIバナー生成ワークの復習', addDays(-7), addDays(-1), addDays(14), '低', 2, 2, '(サンプル) 復習メモを作成', '(サンプル) プロンプトの型が重要', '完了'],
  ];
  sheet.getRange(2, 1, samples.length, HEADERS.length).setValues(samples);
}
