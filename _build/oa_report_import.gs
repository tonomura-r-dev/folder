/**
 * LINE公式アカウント（OAM）のCSV → 月次レポートシートの OA 3タブに貼り込む。
 * 2026-10-29 殿村さん「OAのDriveフォルダ取り込みも作って」。
 *
 * 使い方：
 *  1. Googleドライブに「OAレポ取込」という名前のフォルダを作る（場所はどこでもよい。名前だけ合わせる）
 *  2. 月初に、管理画面（分析）から先月分のCSVを落としてフォルダに入れる（ファイル名はそのままでよい。中身の見出しで判別する）
 *     友だち：友だち追加数／性別／年齢×性別／地域／流入経路
 *     リッチメニュー関連：リッチメニュー／プロフィール（概要）／友だちステータス・アクセス経路（10件以下だと出ない）
 *     配信：メッセージ通数／配信ごと
 *  3. シートの「OA」メニュー →「フォルダから取り込む」
 *     貼り終えたCSVは、同じフォルダの中の「取込済」フォルダへ移す（次の月に二重で読まないため）
 *
 * ga4_report_sync.gs と同じ Apps Script プロジェクトに貼ってよい（onOpen は1つにまとめる）。
 * 触らない所：各タブ1行目の見出し（↓友達追加 など）。2行目の列名から下を、そのCSVの中身で入れ直す。
 */

const OA_FOLDER_NAME = 'OAレポ取込';
const OA_DONE_FOLDER_NAME = '取込済';

// CSVの1行目（列名）で、どのタブのどの列に貼るかを決める
// header＝CSVの列名（この並びで始まれば一致）／sheet／col＝貼り始めの列（A=1）／row＝列名を書く行
const OA_TARGETS = [
  { key: 'friends',  header: ['date', 'contacts', 'targetReaches', 'blocks'],       sheet: 'OA_ ソース',        col: 1,  row: 2 },
  { key: 'gender',   header: ['gender', 'percentage'],                               sheet: 'OA_ ソース',        col: 6,  row: 2 },
  { key: 'age',      header: ['gender', 'age', 'percentage'],                        sheet: 'OA_ ソース',        col: 9,  row: 2 },
  { key: 'area',     header: ['area', 'percentage'],                                 sheet: 'OA_ ソース',        col: 13, row: 2 },
  { key: 'path',     header: ['path', 'originPoint', 'campaign', 'gained', 'blocked'], sheet: 'OA_ ソース',      col: 16, row: 2 },
  { key: 'richmenu', header: ['date', 'richMenuId', 'cmsUrl', 'imp'],                sheet: 'OA_リッチメニュ他', col: 1,  row: 2 },
  { key: 'profile',  header: ['date', 'pageView', 'uniqueUser'],                     sheet: 'OA_リッチメニュ他', col: 11, row: 2 },
  { key: 'status',   header: ['friendStatus', 'friendStatusPercentage'],             sheet: 'OA_リッチメニュ他', col: 15, row: 2 },
  { key: 'overview', header: ['date', 'CMS_BROADCAST', 'CMS_TARGETING'],             sheet: 'OA_投稿別',         col: 1,  row: 2 },
  { key: 'broadcast', header: ['broadcastId', 'sentDate', 'cmsUrl', 'deliveredCount'], sheet: 'OA_投稿別',       col: 14, row: 2 },
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // GA4のメニュー（ga4_report_sync.gs）と同じプロジェクトなら、ここで両方作る
  if (typeof importLastMonth === 'function') {
    ui.createMenu('GA4')
      .addItem('先月分を取り込む', 'importLastMonth')
      .addItem('今月分（1日〜昨日）を取り込む', 'importThisMonth')
      .addToUi();
  }
  ui.createMenu('OA')
    .addItem('フォルダから取り込む', 'importOaFromFolder')
    .addToUi();
}

function importOaFromFolder() {
  const ui = SpreadsheetApp.getUi();
  const folders = DriveApp.getFoldersByName(OA_FOLDER_NAME);
  if (!folders.hasNext()) { ui.alert('ドライブに「' + OA_FOLDER_NAME + '」というフォルダがありません。'); return; }
  const folder = folders.next();

  // フォルダ内のCSVを読み、見出しで振り分ける（同じ種類が複数あれば行をつなげる）
  const groups = {};    // key -> {target, header, rows, files:[]}
  const unknown = [];
  const files = folder.getFilesByType(MimeType.CSV);
  const all = [];
  while (files.hasNext()) all.push(files.next());
  // 拡張子が .csv でも種類が text/plain の場合があるので、名前でも拾う
  const byName = folder.getFiles();
  while (byName.hasNext()) {
    const f = byName.next();
    if (/\.csv$/i.test(f.getName()) && !all.some(function (g) { return g.getId() === f.getId(); })) all.push(f);
  }
  if (all.length === 0) { ui.alert('「' + OA_FOLDER_NAME + '」にCSVがありません。'); return; }

  all.forEach(function (f) {
    const text = f.getBlob().getDataAsString('UTF-8').replace(/^﻿/, '');
    const rows = Utilities.parseCsv(text).filter(function (r) { return r.join('').trim() !== ''; });
    if (rows.length === 0) return;
    const header = rows[0].map(function (h) { return String(h).trim(); });
    const target = OA_TARGETS.find(function (t) {
      return t.header.every(function (h, i) { return header[i] === h; });
    });
    if (!target) { unknown.push(f.getName()); return; }
    if (!groups[target.key]) groups[target.key] = { target: target, header: header, rows: [], files: [] };
    groups[target.key].rows = groups[target.key].rows.concat(rows.slice(1));
    groups[target.key].files.push(f);
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = [];
  Object.keys(groups).forEach(function (key) {
    const g = groups[key];
    const sheet = ss.getSheetByName(g.target.sheet);
    if (!sheet) { summary.push(key + '：タブ「' + g.target.sheet + '」が無い'); return; }
    writeBlock_(sheet, g.target.col, g.target.row, g.header, g.rows);
    summary.push(g.target.sheet + ' ' + colLetter_(g.target.col) + '列：' + g.rows.length + '行（' + g.files.map(function (f) { return f.getName(); }).join('、') + '）');
  });

  // 貼ったCSVは「取込済」へ移す
  const done = getOrCreateSub_(folder, OA_DONE_FOLDER_NAME);
  Object.keys(groups).forEach(function (key) {
    groups[key].files.forEach(function (f) { f.moveTo(done); });
  });

  let msg = 'OA 取り込み完了\n' + summary.join('\n');
  if (unknown.length) msg += '\n\n見出しが合わず飛ばしたファイル：\n' + unknown.join('\n');
  ui.alert(msg);
}

/** 列名の行から下を消して、見出し＋データを入れ直す。幅はそのCSVの列数ぶんだけ。 */
function writeBlock_(sheet, col, headerRow, header, rows) {
  const width = header.length;
  const lastRow = Math.max(sheet.getLastRow(), headerRow);
  sheet.getRange(headerRow, col, lastRow - headerRow + 1, width).clearContent();
  const values = [header].concat(rows.map(function (r) {
    const out = r.slice(0, width);
    while (out.length < width) out.push('');
    return out;
  }));
  if (sheet.getMaxRows() < headerRow + values.length) {
    sheet.insertRowsAfter(sheet.getMaxRows(), headerRow + values.length - sheet.getMaxRows());
  }
  const range = sheet.getRange(headerRow, col, values.length, width);
  range.setNumberFormat('@');       // 日付（20261001）やIDを数字に化けさせない
  range.setValues(values);
}

function getOrCreateSub_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
