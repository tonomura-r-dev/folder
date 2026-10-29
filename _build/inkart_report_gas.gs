/**
 * インクアート 月次レポートシート用 GAS（これ1ファイルだけ貼ればよい）
 *  - 「GA4」メニュー：GA4から先月分（または今月分）を GA_参照元／GA_キャンペーン／GA_流入元別 に入れ直す
 *  - 「OA」メニュー：ドライブの「OAレポ取込」フォルダにあるOAMのCSVを OA_ソース／OA_リッチメニュ他／OA_投稿別 に貼る
 *
 * 入れ方：
 *  1. シート → 拡張機能 → Apps Script → 既定の Code.gs の中身を全部消して、この全文を貼って保存
 *  2. 左の「サービス」の「＋」→「Google Analytics Data API」を追加（識別子 AnalyticsData）
 *  3. シートを開き直すと「GA4」「OA」のメニューが出る。初回は権限の許可が出る
 *  4. OAを使うときは、ドライブに「OAレポ取込」フォルダを作り、OAMのCSVを入れてから「OA」→「フォルダから取り込む」
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GA4')
    .addItem('先月分を取り込む', 'importLastMonth')
    .addItem('今月分（1日〜昨日）を取り込む', 'importThisMonth')
    .addToUi();
  ui.createMenu('OA')
    .addItem('フォルダから取り込む', 'importOaFromFolder')
    .addToUi();
}

// =====================================================================
// GA4 → GA 3タブ
// =====================================================================
/**
 * GA4 → 月次レポートシート（インクアート_〇月_OAレポ）の GA 3タブに、先月分を取り込む。
 * 2026-10-29 殿村さん「GA4からGASで取ってこれるようにしたい」。決めたこと＝期間は先月固定／A4以下を消して入れ直す／
 * CV列は O=新規会員登録_20200617以降・P=purchase・R=合計収益／流入元別のA列はセッションのメディア。
 *
 * 入れ方：
 *  1. シート → 拡張機能 → Apps Script → この中身を貼って保存
 *  2. 左の「サービス」の「＋」→「Google Analytics Data API」を追加（識別子 AnalyticsData）
 *  3. シートを開き直すと「GA4」メニューが出る → 「先月分を取り込む」
 *     初回は権限の許可が出る。動かす人のGoogleアカウントが GA4 プロパティの閲覧者以上であること。
 *
 * 触らない所：見出し（1〜3行目）・W〜AA列の関数・OAの3タブ。
 */

const GA4_PROPERTY_ID = '352910939';          // GA4プリントアース（URLの p352910939）
const FILTER_COUNTRY_JAPAN = true;            // 先月の探索と同じく 国＝Japan で絞る。要らなければ false
const FIRST_DATA_ROW = 4;                     // 見出しは3行。4行目から入れる
const LAST_COL = 20;                          // A〜T まで書く（U〜V は空、W〜AA は関数なので触らない）

// 9指標（D〜L）。G「セッションあたりの平均エンゲージメント時間」は userEngagementDuration ÷ sessions で計算
const METRICS_MAIN = [
  'sessions', 'engagedSessions', 'engagementRate', 'userEngagementDuration',
  'eventsPerSession', 'eventCount', 'keyEvents', 'sessionKeyEventRate', 'totalRevenue',
];
// CVの登録名（O・P）。1回の問い合わせは指標10個までなので2回に分けて合わせる
const METRICS_CV = ['keyEvents:新規会員登録_20200617以降', 'keyEvents:purchase'];

// 3タブの定義：A・B列に入れるディメンションと、C列は日付
const TABS = [
  { sheet: 'GA_参照元 / メディア', dimA: 'sessionSourceMedium', dimB: null },
  { sheet: 'GA_キャンペーン',      dimA: 'sessionSourceMedium', dimB: 'sessionCampaignName' },
  { sheet: 'GA_流入元別',          dimA: 'sessionMedium',       dimB: 'landingPage' },   // landingPage＝クエリ無し
];


function importLastMonth() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const end = new Date(now.getFullYear(), now.getMonth(), 0);
  importRange_(start, end);
}

function importThisMonth() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  if (end < start) { SpreadsheetApp.getUi().alert('今月はまだ昨日の分がありません。'); return; }
  importRange_(start, end);
}

function importRange_(start, end) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const label = fmt_(start) + '〜' + fmt_(end);
  const summary = [];
  TABS.forEach(function (tab) {
    const sheet = ss.getSheetByName(tab.sheet);
    if (!sheet) { summary.push(tab.sheet + '：タブが無い'); return; }
    const rows = fetchTab_(tab, start, end);
    writeRows_(sheet, rows);
    summary.push(tab.sheet + '：' + rows.length + '行');
  });
  SpreadsheetApp.getUi().alert('GA4 取り込み完了（' + label + '）\n' + summary.join('\n'));
}

/** 1タブ分の行（A〜T）を作る。 */
function fetchTab_(tab, start, end) {
  const dims = [tab.dimA].concat(tab.dimB ? [tab.dimB] : []).concat(['date']);
  const main = runReportAll_(dims, METRICS_MAIN, start, end);
  const cv = runReportAll_(dims, METRICS_CV, start, end);

  // CVの結果を「ディメンションの組」で引けるようにする
  const cvMap = {};
  cv.forEach(function (r) { cvMap[r.key] = r.metrics; });

  return main.map(function (r) {
    const m = r.metrics;                       // 数値の配列（METRICS_MAIN の順）
    const c = cvMap[r.key] || [0, 0];
    const sessions = m[0];
    const avgEngage = sessions > 0 ? m[3] / sessions : 0;
    const d = r.dims;                          // [dimA, (dimB), date]
    const a = d[0];
    const b = tab.dimB ? d[1] : '';
    const date = d[d.length - 1];              // 'YYYYMMDD'
    const dateText = date.slice(0, 4) + '/' + String(Number(date.slice(4, 6))) + '/' + String(Number(date.slice(6, 8)));
    return [
      a, b, dateText,                          // A B C
      m[0], m[1], m[2], avgEngage, m[4], m[5], m[6], m[7], m[8],   // D〜L
      '',                                      // M
      m[6],                                    // N 全てのイベント（＝キーイベント）
      c[0],                                    // O 会員登録（新規会員登録_20200617以降）
      c[1],                                    // P 購入完了（purchase）
      '',                                      // Q クーポンDL（この案件では空）
      m[8],                                    // R purchase(値)＝合計収益
      '', '',                                  // S 電話Tap／T 来店（空）
    ];
  });
}

/** runReport をページ送りで全件取る。戻り＝[{key, dims:[...], metrics:[数値...]}] */
function runReportAll_(dims, metrics, start, end) {
  const out = [];
  const limit = 100000;
  let offset = 0;
  while (true) {
    const req = {
      dateRanges: [{ startDate: fmtApi_(start), endDate: fmtApi_(end) }],
      dimensions: dims.map(function (n) { return { name: n }; }),
      metrics: metrics.map(function (n) { return { name: n }; }),
      orderBys: [{ dimension: { dimensionName: 'date' } }],
      limit: limit,
      offset: offset,
      keepEmptyRows: false,
    };
    if (FILTER_COUNTRY_JAPAN) {
      req.dimensionFilter = { filter: { fieldName: 'country', stringFilter: { matchType: 'EXACT', value: 'Japan' } } };
    }
    const res = AnalyticsData.Properties.runReport(req, 'properties/' + GA4_PROPERTY_ID);
    const rows = res.rows || [];
    rows.forEach(function (row) {
      const d = row.dimensionValues.map(function (v) { return v.value; });
      const m = row.metricValues.map(function (v) { return Number(v.value); });
      out.push({ key: d.join('\u0001'), dims: d, metrics: m });
    });
    if (rows.length < limit) break;
    offset += limit;
  }
  return out;
}

/** A4以下（A〜T）を消して入れ直す。C列は文字のまま（日付に変換させない）。 */
function writeRows_(sheet, rows) {
  const lastRow = Math.max(sheet.getLastRow(), FIRST_DATA_ROW);
  sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, LAST_COL).clearContent();
  if (rows.length === 0) return;
  if (sheet.getMaxRows() < FIRST_DATA_ROW + rows.length) {
    sheet.insertRowsAfter(sheet.getMaxRows(), FIRST_DATA_ROW + rows.length - sheet.getMaxRows());
  }
  sheet.getRange(FIRST_DATA_ROW, 3, rows.length, 1).setNumberFormat('@');   // C列＝文字
  sheet.getRange(FIRST_DATA_ROW, 1, rows.length, LAST_COL).setValues(rows);
}

function fmt_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd'); }
function fmtApi_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }

// =====================================================================
// OAMのCSV → OA 3タブ
// =====================================================================
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
    const text = f.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, '');
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
