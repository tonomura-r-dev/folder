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

// メニュー。oa_report_import.gs と同じプロジェクトに入れた場合は、どちらの onOpen が動いても両方のメニューが出る
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GA4')
    .addItem('先月分を取り込む', 'importLastMonth')
    .addItem('今月分（1日〜昨日）を取り込む', 'importThisMonth')
    .addToUi();
  if (typeof importOaFromFolder === 'function') {
    ui.createMenu('OA').addItem('フォルダから取り込む', 'importOaFromFolder').addToUi();
  }
}

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
