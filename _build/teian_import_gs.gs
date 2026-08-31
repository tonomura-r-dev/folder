/**
 * LINEOA提案管理シートに、月次管理ブックの既存顧客を「行として」足すスクリプト。
 *
 * ■ 入れ方
 *   提案管理シート（まずはコピーで試す）を開く
 *   → 拡張機能 → Apps Script → 中身を全部消してこれを貼る → 保存
 *   → シートを再読み込み → 上のメニューに「既存顧客の取込」が出る
 *   → まず「取り込む件数を見る」で確認 → よければ「行を追加する」
 *
 * ■ 既存の行は1セルも変更しない
 *   「↑ここより上に行追加してください」の直前に行を挿入するだけ。
 *   既にあるセルへの書き込みは一切しない。
 *
 * ■ このシートを読んで埋めるもの（エクセル版では埋められなかった部分）
 *   業界     … 同じ会社が既にシートにあれば、その業界をそのまま使う
 *   ☑        … 既にある会社なら「今の件数+1」、新規なら 1
 *   担当①②  … シートの担当欄に実際にある名字表記を集めて、それに合わせる
 *
 * ■ 空欄のままにするもの
 *   結論 / 追客 / Ac名 / 外注先 / DYM売上 / LINE売上
 *   まだ提案していないので、見込み額も結論も存在しない。提案してから手で入れる。
 */

// ========================= 設定 =========================

/** 月次管理ブックのID（読むだけ。一切変更しない） */
const SRC_ID = '11j9QaZ74iOp1ABNV95wq-PIDK81sWsXXxA0JYy0KCWw';

/** 書き込む先のシート（URLの gid= の数字）。0 にすると開いているシートを使う */
const TARGET_GID = 0;

/** 提案列に入れる年月。'' なら実行した月から自動で作る（26/8 の形） */
const TEIAN_MONTH = '';

/** 提案商材に入れる値。シートで使われている表記に合わせること */
const SHOZAI = '商流変更';

/** この文字を含む行の直前に挿入する */
const TAIL_MARK = 'ここより上に行追加';

/** 月次管理ブックの媒体シート */
const MEDIA = [
  'AD', 'アフィ', 'CS', 'MEO', '制作', 'PR', '風評', 'タレントシェア',
  'LINE公式アカウント', 'メディア', 'ベトナム', 'ASP', 'マス広告(MA)・その他',
];

/** 担当の列名。媒体によって呼び方が違う（ASPだけ広告主担当／媒体担当） */
const ACCT_COLS = ['アカウント', '広告主担当①'];
const CONS_COLS = ['コンサル', '媒体担当①'];
const ALL_TANTO = ['アカウント', 'コンサル', '運用①', '運用②', '運用③',
                   '広告主担当①', '広告主担当②', '媒体担当①', '媒体担当②'];

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

  const need = ['提案', '☑', '企業名', '既存', '商流', '提案商材', '業界', '担当①', '担当②'];
  const missing = need.filter(function (k) { return !head.cols[k]; });
  if (missing.length) {
    ui.alert('見出しが見つかりません',
      '見つからない列：' + missing.join(' / ') + '\n\n' +
      'このシートで見つかった見出し：\n' + Object.keys(head.cols).join(' / '),
      ui.ButtonSet.OK);
    return;
  }

  const known = readSheet_(sh, head);      // 業界・カウント・名字の辞書
  const recs = buildRecords_(known.surnames);
  const teian = TEIAN_MONTH || monthLabel_();

  let hitGyokai = 0;
  const data = recs.map(function (r) {
    const key = normalizeName_(r.name);
    const gyokai = known.gyokai[key] || '';
    if (gyokai) hitGyokai++;
    return {
      rec: r,
      gyokai: gyokai,
      count: (known.count[key] || 0) + 1,
    };
  });

  const dup = data.filter(function (d) { return d.count > 1; }).length;
  const msg = '月次管理ブックの既存顧客：' + data.length + '社\n' +
              '　うち ' + dup + '社は、このシートに既に行があります\n' +
              '　　（消さずに追加して、☑ を ' + '今の件数+1' + ' にします）\n' +
              '　業界が引き当てられたのは ' + hitGyokai + '社\n' +
              '　担当の名字は、このシートにある ' + known.surnames.length + '種類に合わせます';

  if (dryRun) {
    ui.alert('取り込む件数', msg + '\n\n※このメニューでは追加しません。', ui.ButtonSet.OK);
    return;
  }

  const ok = ui.alert('行を追加します',
    msg + '\n\n既存の行は1セルも変更しません。行の挿入だけです。\n' +
    '結論・追客・Ac名・外注先・DYM売上・LINE売上は空欄のままにします。\n\n実行しますか？',
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  const at = findTailRow_(sh, head.cols['企業名'], head.row);
  sh.insertRowsBefore(at, data.length);

  const width = sh.getLastColumn();
  const c = head.cols;
  const rows = data.map(function (d) {
    const row = new Array(width).fill('');
    row[c['提案'] - 1] = teian;
    row[c['☑'] - 1] = d.count;
    row[c['企業名'] - 1] = d.rec.name;
    row[c['既存'] - 1] = '●';
    row[c['商流'] - 1] = d.rec.shoryu;
    row[c['提案商材'] - 1] = SHOZAI;
    row[c['業界'] - 1] = d.gyokai;
    row[c['担当①'] - 1] = d.rec.tanto1;
    row[c['担当②'] - 1] = d.rec.tanto2;
    return row;
  });
  sh.getRange(at, 1, rows.length, width).setValues(rows);

  ui.alert('完了',
    data.length + '行を追加しました。\n' +
    '提案列が「' + teian + '」の行が今回ぶんです。\n' +
    '消したくなったら、提案列で絞り込んでください。', ui.ButtonSet.OK);
}

// ========================= 提案管理シートを読む =========================

function targetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (TARGET_GID) {
    const f = ss.getSheets().filter(function (s) { return s.getSheetId() === TARGET_GID; });
    if (f.length) return f[0];
  }
  return ss.getActiveSheet();
}

/** 見出し行を上から5行以内で探す。「企業名」がある行を見出しとみなす */
function findHeader_(sh) {
  const rows = Math.min(5, sh.getLastRow());
  const scan = sh.getRange(1, 1, rows, sh.getLastColumn()).getValues();
  for (let i = 0; i < scan.length; i++) {
    const cols = {};
    scan[i].forEach(function (v, j) {
      const n = String(v == null ? '' : v).replace(/\n/g, '').trim();
      if (n && !(n in cols)) cols[n] = j + 1;
    });
    if (cols['企業名']) return { row: i + 1, cols: cols };
  }
  return { row: 1, cols: {} };
}

/**
 * シートから3つ作る。
 *   gyokai   : 正規化した企業名 → 業界
 *   count    : 正規化した企業名 → 今ある行数（☑の付け直しに使う）
 *   surnames : 担当①②に実際に入っている表記の一覧（長い順）
 * 名字を推測せず、このシートで使われている表記そのものに合わせるのが狙い。
 */
function readSheet_(sh, head) {
  const last = sh.getLastRow();
  const gyokai = {}, count = {}, sur = {};
  if (last <= head.row) return { gyokai: gyokai, count: count, surnames: [] };

  const vals = sh.getRange(head.row + 1, 1, last - head.row, sh.getLastColumn()).getValues();
  const cName = head.cols['企業名'] - 1;
  const cGyo = head.cols['業界'] - 1;
  const cT1 = head.cols['担当①'] - 1;
  const cT2 = head.cols['担当②'] - 1;

  vals.forEach(function (r) {
    const raw = String(r[cName] == null ? '' : r[cName]).trim();
    if (raw && raw.indexOf(TAIL_MARK) < 0 && raw !== '---') {
      const key = normalizeName_(raw);
      if (key) {
        count[key] = (count[key] || 0) + 1;
        const g = String(r[cGyo] == null ? '' : r[cGyo]).trim();
        // 同じ会社で業界が入っている行があれば、それを採用する
        if (g && g !== '--' && !gyokai[key]) gyokai[key] = g;
      }
    }
    [r[cT1], r[cT2]].forEach(function (v) {
      const t = String(v == null ? '' : v).trim();
      // 「-」「--」「その他（AD）」のような人名でない値は名字辞書に入れない
      if (t && t.length <= 5 && !/^[-ー－]+$/.test(t) && t.indexOf('その他') < 0) {
        sur[t] = true;
      }
    });
  });

  // 長いものから当てる。「長谷川」を「長谷」より先に見るため
  const surnames = Object.keys(sur).sort(function (a, b) { return b.length - a.length; });
  return { gyokai: gyokai, count: count, surnames: surnames };
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

// ========================= 月次管理ブックを集計する =========================

function monthLabel_() {
  const d = new Date();
  return (d.getFullYear() % 100) + '/' + (d.getMonth() + 1);
}

/**
 * 会社名の表記ゆれを吸収する。
 * 株式会社などの法人格、空白、括弧書き、全角半角の違いを落として比べる。
 */
function normalizeName_(s) {
  return String(s == null ? '' : s)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[（(\[【].*?[)）\]】]/g, '')
    .replace(/<[^>]*>/g, '')                       // <既存> を落とす
    .replace(/^【リ】/, '')                        // 【リ】を落とす
    .replace(/_(ショット|ショットマイナス|代理店|CPF|既存マイナス).*$/, '')
    .replace(/(株式会社|有限会社|合同会社|一般社団法人|医療法人社団|医療法人|学校法人|公益財団法人|一般財団法人|特定非営利活動法人|弁護士法人|司法書士法人)/g, '')
    .replace(/[\s　_･・\-ー―]/g, '')
    .toLowerCase()
    .trim();
}

/** フルネームから、このシートで使われている名字表記を取り出す */
function toSurname_(full, surnames) {
  const s = String(full == null ? '' : full).trim();
  if (!s) return '';
  if (s.indexOf('その他') >= 0) return s;   // 人名でない値はそのまま
  for (let i = 0; i < surnames.length; i++) {
    if (s.indexOf(surnames[i]) === 0) return surnames[i];
  }
  return s;   // 見つからなければフルネームのまま（あとで手で直せる）
}

function toNum_(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }
function toStr_(v) { return String(v == null ? '' : v).trim(); }

function push_(lst, v) { if (v && lst.indexOf(v) < 0) lst.push(v); }

/** エンドクライアント単位で1件にまとめる。売上の大きい順＝アタックの優先順 */
function buildRecords_(surnames) {
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
      if (k && !(k in h)) h[k] = i;      // 同名列は左（通常ぶん）を優先
    });
    if (!('計上種別' in h) || !('請求額（税抜）' in h)) return;

    const acctCol = ACCT_COLS.filter(function (c) { return c in h; })[0];
    const consCol = CONS_COLS.filter(function (c) { return c in h; })[0];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const end = toStr_(row[h['エンドクライアント名']]) || toStr_(row[h['社名']]);
      if (!end) continue;
      const kei = toStr_(row[h['計上種別']]);
      const uri = toNum_(row[h['請求額（税抜）']]);
      if (KEIS.indexOf(kei) < 0 && uri === 0) continue;   // 空行

      if (!g[end]) {
        g[end] = { name: end, uri: 0, choku: false, agency: [],
                   acct: '', cons: '' };
      }
      const d = g[end];
      d.uri += uri;

      // 商流はシートでは「経由している代理店の会社名そのもの」。直なら「直」
      const shubetsu = ('種別' in h) ? toStr_(row[h['種別']]) : '';
      const company = toStr_(row[h['社名']]);
      if (shubetsu.indexOf('直') >= 0) {
        d.choku = true;
      } else if (shubetsu.indexOf('代理店') >= 0 && company && company !== end) {
        push_(d.agency, company);
      }

      if (!d.acct && acctCol) d.acct = toStr_(row[h[acctCol]]);
      if (!d.cons && consCol) d.cons = toStr_(row[h[consCol]]);
      // 担当が片方も取れていなければ、他の役割から埋める
      if (!d.acct || !d.cons) {
        ALL_TANTO.forEach(function (c) {
          if (!(c in h)) return;
          const t = toStr_(row[h[c]]);
          if (!t) return;
          if (!d.acct) d.acct = t;
          else if (!d.cons && t !== d.acct) d.cons = t;
        });
      }
    }
  });

  return Object.keys(g)
    .map(function (k) {
      const d = g[k];
      d.shoryu = d.choku ? '直' : (d.agency.join(' / ') || '不明');
      d.tanto1 = toSurname_(d.acct, surnames);
      d.tanto2 = toSurname_(d.cons, surnames);
      return d;
    })
    .sort(function (a, b) { return b.uri - a.uri; });
}
