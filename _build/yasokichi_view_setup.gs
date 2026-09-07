/**
 * 八十吉「空き状況」ビュー ── ステップ1：スプレッドシートを作る
 *
 * 使い方
 *  1. 新しいGoogleスプレッドシートを作る（名前は「八十吉_予約状況」など）
 *  2. 拡張機能 → Apps Script → このファイルの中身を全部貼る → 保存
 *  3. 関数の選択で setupYasokichiSheets を選んで ▶ 実行 → 権限を許可
 *  4. スプレッドシートに戻って再読み込み → 4つのシートができていればOK
 *
 * できるシート
 *  ・予約   … 1行＝1予約。列は「予約者／日にち／時間／コース／席／指名する人」＋GAS用4列。
 *              あとでレストランボードの通知メールからGASが自動で足す。手入力もできる
 *  ・日別   … 日付ごとの予約人数と残席（数式）。「空き状況」画面はここを読む
 *  ・設定   … 総席数・営業時間など。数式とGASが参照する
 *  ・ログ   … GASの動作記録
 *
 * 何度実行しても、入っているデータは消えない（見出しと書式だけ整える）。
 */

var SHEETS = {
  RES: '予約',
  DAILY: '日別',
  CONF: '設定',
  LOG: 'ログ',
};

// 殿村さん指定の6列を先頭に。右の4列はGAS用（残席の計算と、メールの二重取り込み防止）
var RES_HEADERS = [
  '予約者',        // A
  '日にち',        // B yyyy/mm/dd
  '時間',          // C hh:mm（来店時刻）
  'コース',        // D
  '席',            // E 席数（＝人数）。残席はこの合計で出す
  '指名する人',    // F
  'ステータス',    // G 予約／来店／キャンセル（GAS用）
  '媒体',          // H ホットペッパー／食べログ／電話／その他（GAS用）
  '取込日時',      // I GASが書く。手入力なら空
  'メールID',      // J 二重取り込み防止。手入力なら空
];

var CONF_ROWS = [
  ['総席数', 30, 'お店の席数。残席＝総席数−その日の予約人数'],
  ['営業開始', '17:00', ''],
  ['営業終了', '23:00', ''],
  ['定休日', '木,日', 'カンマ区切り'],
  ['表示日数', 14, '「空き状況」画面に出す日数（今日から）'],
  ['通知メールの検索条件', 'from:(restaurant-board) newer_than:1d', 'GASがGmailを探す条件。あとで実物に合わせて直す'],
  ['最終取込', '', 'GASが自動で書く。触らない'],
];

/** メニュー */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('八十吉ビュー')
    .addItem('シートを初期化（データは消さない）', 'setupYasokichiSheets')
    .addItem('サンプル予約を入れる', 'insertSampleReservations')
    .addItem('サンプル予約を消す', 'removeSampleReservations')
    .addToUi();
}

/** 本体：4シートを作る／整える */
function setupYasokichiSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupResSheet_(ss);
  setupConfSheet_(ss);
  setupDailySheet_(ss);
  setupLogSheet_(ss);
  // 余った「シート1」は消す（空のときだけ）
  var s1 = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  log_('setup', '4シートを整えた');
  try { SpreadsheetApp.getUi().alert('できました。「予約」「日別」「設定」「ログ」の4シートを確認してください。'); } catch (e) {}
}

// ---------- 予約 ----------
function setupResSheet_(ss) {
  var sh = getOrCreate_(ss, SHEETS.RES);
  writeHeader_(sh, RES_HEADERS);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  [120, 100, 70, 160, 60, 110, 90, 110, 140, 160].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  var n = Math.max(sh.getMaxRows() - 1, 500);
  // 書式
  sh.getRange(2, 2, n, 1).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2, 3, n, 1).setNumberFormat('hh:mm');
  sh.getRange(2, 5, n, 1).setNumberFormat('0');
  sh.getRange(2, 9, n, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  // GAS用の4列は見出しをグレーにして区別
  sh.getRange(1, 7, 1, 4).setBackground('#6B7280');
  // プルダウン
  setList_(sh.getRange(2, 7, n, 1), ['予約', '来店', 'キャンセル']);
  setList_(sh.getRange(2, 8, n, 1), ['ホットペッパー', '食べログ', '電話', 'その他']);
  // キャンセル行は薄く
  var rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="キャンセル"')
    .setFontColor('#9AA0B3')
    .setRanges([sh.getRange(2, 1, n, RES_HEADERS.length)])
    .build());
  sh.setConditionalFormatRules(rules);
}

// ---------- 設定 ----------
function setupConfSheet_(ss) {
  var sh = getOrCreate_(ss, SHEETS.CONF);
  writeHeader_(sh, ['項目', '値', 'メモ']);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 360);
  // 既にある項目は上書きしない
  var existing = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (r[0]) existing[r[0]] = i + 2; });
  }
  CONF_ROWS.forEach(function (row) {
    if (existing[row[0]]) return;
    sh.appendRow(row);
  });
  sh.getRange('B2').setNumberFormat('0');
}

// ---------- 日別（数式） ----------
function setupDailySheet_(ss) {
  var sh = getOrCreate_(ss, SHEETS.DAILY);
  writeHeader_(sh, ['日にち', '曜日', '予約件数', '予約席数', '残席', '内訳（時間 席 予約者 コース 指名）']);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 50);
  sh.setColumnWidth(3, 80);
  sh.setColumnWidth(4, 80);
  sh.setColumnWidth(5, 70);
  sh.setColumnWidth(6, 420);
  // 今日から「表示日数」ぶんの日付を並べる。数式は行ごとに置く
  var days = Number(confGet_(ss, '表示日数')) || 14;
  var rows = [];
  for (var i = 0; i < days; i++) {
    var r = i + 2;
    rows.push([
      '=TODAY()+' + i,
      '=TEXT(A' + r + ',"ddd")',
      '=COUNTIFS(予約!$B:$B,A' + r + ',予約!$G:$G,"<>キャンセル",予約!$A:$A,"<>")',
      '=SUMIFS(予約!$E:$E,予約!$B:$B,A' + r + ',予約!$G:$G,"<>キャンセル")',
      '=設定!$B$2-D' + r,
      '=IFERROR(TEXTJOIN(" ／ ",TRUE,ARRAYFORMULA(IF((予約!$B$2:$B=A' + r + ')*(予約!$G$2:$G<>"キャンセル")*(予約!$A$2:$A<>""),TEXT(予約!$C$2:$C,"hh:mm")&" "&予約!$E$2:$E&"席 "&予約!$A$2:$A&IF(予約!$D$2:$D<>""," "&予約!$D$2:$D,"")&IF(予約!$F$2:$F<>"","（指名:"&予約!$F$2:$F&"）",""),""))),"")',
    ]);
  }
  sh.getRange(2, 1, days, 6).setFormulas(rows);
  sh.getRange(2, 1, days, 1).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2, 3, days, 3).setNumberFormat('0');
  // 残席が少ないと色
  var rng = sh.getRange(2, 5, days, 1);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(0).setBackground('#F4C7C3').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1, 4).setBackground('#FCE8B2').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(4).setBackground('#D9EAD3').setRanges([rng]).build(),
  ];
  sh.setConditionalFormatRules(rules);
  // 余計な行は消す
  if (sh.getMaxRows() > days + 1) sh.deleteRows(days + 2, sh.getMaxRows() - days - 1);
}

// ---------- ログ ----------
function setupLogSheet_(ss) {
  var sh = getOrCreate_(ss, SHEETS.LOG);
  writeHeader_(sh, ['日時', '種別', '内容']);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 500);
}

// ---------- サンプル ----------
var SAMPLE_TAG = 'SAMPLE';
function insertSampleReservations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.RES);
  if (!sh) { setupYasokichiSheets(); sh = ss.getSheetByName(SHEETS.RES); }
  var today = new Date(); today.setHours(0, 0, 0, 0);
  function d(n) { var x = new Date(today); x.setDate(x.getDate() + n); return x; }
  function t(h, m) { var x = new Date(1899, 11, 30, h, m); return x; }
  // 予約者, 日にち, 時間, コース, 席, 指名する人, ステータス, 媒体, 取込日時, メールID
  var rows = [
    ['高橋 様', d(1), t(17, 30), '', 2, '', '予約', '電話', '', SAMPLE_TAG],
    ['田中 様', d(1), t(18, 0), '飲み放題付き宴会コース', 4, '', '予約', 'ホットペッパー', '', SAMPLE_TAG],
    ['佐藤 様', d(1), t(18, 30), 'おまかせコース', 6, '', '予約', '食べログ', '', SAMPLE_TAG],
    ['鈴木 様', d(1), t(19, 0), '', 3, '', '予約', '電話', '', SAMPLE_TAG],
    ['伊藤 様', d(1), t(20, 0), '飲み放題付き宴会コース', 5, '', '予約', 'ホットペッパー', '', SAMPLE_TAG],
    ['渡辺 様', d(1), t(20, 30), '', 2, '', '予約', '食べログ', '', SAMPLE_TAG],
    ['山田 様', d(1), t(19, 0), 'おまかせコース', 4, 'ユウスケ', '予約', '食べログ', '', SAMPLE_TAG],
    ['中村 様', d(2), t(19, 0), '飲み放題付き宴会コース', 8, '', '予約', 'ホットペッパー', '', SAMPLE_TAG],
    ['小林 様', d(2), t(20, 0), '', 4, '', 'キャンセル', '電話', '', SAMPLE_TAG],
    ['加藤 様', d(3), t(18, 0), 'おまかせコース', 10, 'マサキ', '予約', '食べログ', '', SAMPLE_TAG],
  ];
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  log_('sample', rows.length + '件のサンプルを入れた');
}
function removeSampleReservations() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RES);
  if (!sh) return;
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 10, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === SAMPLE_TAG) sh.deleteRow(i + 2);
  }
  log_('sample', 'サンプルを消した');
}

// ---------- 共通 ----------
function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function writeHeader_(sh, headers) {
  var r = sh.getRange(1, 1, 1, headers.length);
  r.setValues([headers]);
  r.setFontWeight('bold').setBackground('#1F285A').setFontColor('#FFFFFF');
}
function setList_(range, values) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build();
  range.setDataValidation(rule);
}
function confGet_(ss, key) {
  var sh = ss.getSheetByName(SHEETS.CONF);
  if (!sh) return '';
  var last = sh.getLastRow();
  if (last < 2) return '';
  var v = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === key) return v[i][1];
  return '';
}
function log_(kind, text) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (sh) sh.appendRow([new Date(), kind, text]);
}
