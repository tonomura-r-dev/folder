/**
 * LINEOA提案管理シートに、月次管理ブックの既存顧客を「行として」追加するスクリプト。
 *
 * ■ 何をするか
 *   ・月次管理ブックの全媒体からエンドクライアントを集計
 *   ・提案管理シートに既にある企業は除外（重複を作らない）
 *   ・残りを「↑ここより上に行追加してください」の直前に挿入
 *   ・右端に 契約形態／実績売上／実績粗利／取込元 の4列を足す
 *
 * ■ 既存の行・列は1セルも変更しない
 *   行の挿入だけを行い、既存セルへの書き込みはしない。
 *   DYM売上・LINE売上は「提案時の見込み」の列なので、実績を入れず空欄のままにする。
 *   提案商材・業界は実績データに無いため空欄。あとから手で埋める前提。
 *
 * ■ 入れ方
 *   提案管理シートを開く → 拡張機能 → Apps Script → 全部消してこれを貼る
 *   → 保存 → シートを再読み込み → メニュー「既存顧客の取込」
 *   まず「取り込む件数を見る」で確認してから「行を追加する」を実行すること。
 */

// ========================= 設定 =========================

/** 月次管理ブックのID */
const SRC_ID = '11j9QaZ74iOp1ABNV95wq-PIDK81sWsXXxA0JYy0KCWw';

/** 書き込む先のシート（URLの gid= の数字）。0 なら開いているシート */
const TARGET_GID = 281377521;

/** 企業名の列の見出し */
const KEY_HEADER = '企業名';

/** この文字を含む行の直前に挿入する */
const TAIL_MARK = 'ここより上に行追加';

/** 右端に足す列 */
const ADD_HEADERS = ['契約形態', '実績売上', '実績粗利', '取込元'];

/** 取込元の印。あとで消したくなったらこの値で絞り込める */
const SOURCE_TAG = '月次管理ブック';

const MEDIA = [
  'AD', 'アフィ', 'CS', 'MEO', '制作', 'PR', '風評', 'タレントシェア',
  'LINE公式アカウント', 'メディア', 'ベトナム', 'ASP', 'マス広告(MA)・その他',
];
const KEIS = ['ストック', 'ショット'];

// ========================= メニュー =========================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('既存顧客の取込')
    .addItem('取り込む件数を見る', 'previewImport')
    .addItem('行を追加する', 'runImport')
    .addToUi();
}

function previewImport() { importRows_(true); }
function runImport() { importRows_(false); }

// ========================= メイン =========================

function importRows_(dryRun) {
  const ui = SpreadsheetApp.getUi();
  const sh = targetSheet_();
  const head = findHeader_(sh);
  if (!head.cols[KEY_HEADER]) {
    ui.alert('「' + KEY_HEADER + '」列が見つかりませんでした。');
    return;
  }

  const keyCol = head.cols[KEY_HEADER];
  const lastRow = sh.getLastRow();

  // すでにシートにある企業名（正規化して比較）
  const existing = {};
  sh.getRange(head.row + 1, keyCol, lastRow - head.row, 1).getValues()
    .forEach(function (r) {
      const k = normalizeName_(r[0]);
      if (k) existing[k] = true;
    });

  const recs = buildRecords_();
  const fresh = recs.filter(function (r) { return !existing[normalizeName_(r.name)]; });

  const msg = '月次管理ブックの既存顧客 ' + recs.length + '社のうち、\n' +
              'このシートにまだ無いのは ' + fresh.length + '社です。\n' +
              '（' + (recs.length - fresh.length) + '社はすでにあるので追加しません）';

  if (dryRun) {
    ui.alert('取り込む件数', msg + '\n\n※このメニューでは追加しません。', ui.ButtonSet.OK);
    return;
  }
  if (!fresh.length) {
    ui.alert('追加する行はありませんでした。', ui.ButtonSet.OK);
    return;
  }

  const ok = ui.alert('行を追加します',
    msg + '\n\n既存の行は1セルも変更しません。行の挿入だけです。\n' +
    'DYM売上・LINE売上・提案商材・業界は空欄のままにします。\n\n実行しますか？',
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  // 右端の追加列（無ければ作る）
  const addCol = {};
  let next = sh.getLastColumn() + 1;
  ADD_HEADERS.forEach(function (h) {
    if (head.cols[h]) { addCol[h] = head.cols[h]; return; }
    sh.getRange(head.row, next).setValue(h)
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864');
    addCol[h] = next;
    next++;
  });

  const at = findTailRow_(sh, keyCol, head.row);
  sh.insertRowsBefore(at, fresh.length);

  const width = sh.getLastColumn();
  const rows = fresh.map(function (r) {
    const row = new Array(width).fill('');
    row[keyCol - 1] = r.name;
    if (head.cols['既存']) row[head.cols['既存'] - 1] = '●';
    if (head.cols['商流']) row[head.cols['商流'] - 1] = r.shoryu;
    if (head.cols['外注先']) row[head.cols['外注先'] - 1] = r.agency;
    if (head.cols['担当①']) row[head.cols['担当①'] - 1] = r.tanto1;
    if (head.cols['担当②']) row[head.cols['担当②'] - 1] = r.tanto2;
    row[addCol['契約形態'] - 1] = r.keiyaku;
    row[addCol['実績売上'] - 1] = r.uri;
    row[addCol['実績粗利'] - 1] = r.rieki;
    row[addCol['取込元'] - 1] = SOURCE_TAG;
    return row;
  });
  sh.getRange(at, 1, rows.length, width).setValues(rows);
  sh.getRange(at, addCol['実績売上'], rows.length, 2).setNumberFormat('#,##0');

  ui.alert('完了',
    fresh.length + '行を追加しました。\n' +
    '追加した行は「取込元」列が「' + SOURCE_TAG + '」になっています。\n' +
    '消したくなったらそれで絞り込んでください。', ui.ButtonSet.OK);
}

// ========================= シートの読み取り =========================

function targetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (TARGET_GID) {
    const f = ss.getSheets().filter(function (s) { return s.getSheetId() === TARGET_GID; });
    if (f.length) return f[0];
  }
  return ss.getActiveSheet();
}

function findHeader_(sh) {
  const scan = sh.getRange(1, 1, Math.min(5, sh.getLastRow()), sh.getLastColumn()).getValues();
  for (let i = 0; i < scan.length; i++) {
    const cols = {};
    scan[i].forEach(function (v, j) {
      const n = String(v == null ? '' : v).replace(/\n/g, '').trim();
      if (n && !(n in cols)) cols[n] = j + 1;
    });
    if (cols[KEY_HEADER]) return { row: i + 1, cols: cols };
  }
  return { row: 1, cols: {} };
}

/** 「ここより上に行追加」の行を探す。無ければ最終行の次 */
function findTailRow_(sh, keyCol, headRow) {
  const last = sh.getLastRow();
  const vals = sh.getRange(headRow + 1, keyCol, last - headRow, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] == null ? '' : vals[i][0]);
    if (v.indexOf(TAIL_MARK) >= 0 || v === '---') return headRow + 1 + i;
  }
  return last + 1;
}

// ========================= 月次管理ブックの集計 =========================

function normalizeName_(s) {
  return String(s == null ? '' : s)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[（(\[【].*?[)）\]】]/g, '')
    .replace(/<[^>]*>/g, '')
    .replace(/(株式会社|有限会社|合同会社|一般社団法人|医療法人社団|医療法人|学校法人|公益財団法人|一般財団法人|特定非営利活動法人|弁護士法人|司法書士法人)/g, '')
    .replace(/[\s　_･・\-ー―]/g, '')
    .toLowerCase()
    .trim();
}

function toNum_(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }
function toStr_(v) { return String(v == null ? '' : v).trim(); }

/** エンドクライアント単位で1件にまとめる。売上の大きい順 */
function buildRecords_() {
  const src = SpreadsheetApp.openById(SRC_ID);
  const g = {};

  MEDIA.forEach(function (name) {
    const sh = src.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const values = sh.getRange(1, 1, sh.getLastRow(),
                               Math.min(sh.getLastColumn(), 80)).getValues();
    const h = {};
    values[0].forEach(function (v, i) {
      const k = toStr_(v).replace(/\n/g, '');
      if (k && !(k in h)) h[k] = i;
    });
    const need = ['エンドクライアント名', '社名', '計上種別', '請求額（税抜）', '利益'];
    if (need.some(function (k) { return !(k in h); })) return;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const end = toStr_(row[h['エンドクライアント名']]) || toStr_(row[h['社名']]);
      if (!end) continue;
      const kei = toStr_(row[h['計上種別']]);
      const uri = toNum_(row[h['請求額（税抜）']]);
      if (KEIS.indexOf(kei) < 0 && uri === 0) continue;

      if (!g[end]) {
        g[end] = { name: end, uri: 0, rieki: 0, kei: {}, agency: [],
                   tanto: [], shoryu: '' };
      }
      const d = g[end];
      d.uri += uri;
      d.rieki += toNum_(row[h['利益']]);
      if (KEIS.indexOf(kei) >= 0) d.kei[kei] = true;

      const shubetsu = ('種別' in h) ? toStr_(row[h['種別']]) : '';
      if (shubetsu && !d.shoryu) d.shoryu = shubetsu === '直案件' ? '直' : shubetsu;
      // 代理店経由なら請求先を外注先として残す（頭越しを防ぐため）
      const company = toStr_(row[h['社名']]);
      if (shubetsu === '代理店' && company && company !== end &&
          d.agency.indexOf(company) < 0) d.agency.push(company);

      ['アカウント', 'コンサル', '運用①', '運用②', '運用③',
       '広告主担当①', '広告主担当②', '媒体担当①', '媒体担当②'].forEach(function (c) {
        if (!(c in h)) return;
        const t = toStr_(row[h[c]]);
        if (t && d.tanto.indexOf(t) < 0) d.tanto.push(t);
      });
    }
  });

  return Object.keys(g)
    .map(function (k) {
      const d = g[k];
      const s = d.kei['ストック'], o = d.kei['ショット'];
      d.keiyaku = (s && o) ? 'ストック＋ショット' : s ? 'ストック' : o ? 'ショット' : '';
      d.agency = d.agency.join(' / ');
      d.tanto1 = d.tanto[0] || '';
      d.tanto2 = d.tanto[1] || '';
      return d;
    })
    .sort(function (a, b) { return b.uri - a.uri; });
}
