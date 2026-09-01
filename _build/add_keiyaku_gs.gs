/**
 * LINEOA提案管理シートに「契約形態」列を入れるスクリプト（Google Apps Script）
 *
 * 月次管理ブックから ストック／ショット を読み、企業名で引き当てて
 * 提案管理シートに1列だけ書き込む。
 *
 * ■ 入れ方
 *   提案管理シートを開く → 拡張機能 → Apps Script → 中身を全部消してこれを貼る
 *   → 保存 → シートを再読み込み → メニュー「契約形態」→「契約形態を入れる」
 *
 * ■ 安全のために
 *   ・書き込むのは「契約形態」列の1列だけ。他の列には一切触れない
 *   ・その列が無ければ、表の右端に新しく作る（既存の列は動かさない）
 *   ・月次管理ブックは読むだけ
 *   ・実行前に「何件ヒットしたか」を出して、確認してから書き込む
 */

// ========================= 設定 =========================

/** 月次管理ブックのID */
const SRC_ID = '11j9QaZ74iOp1ABNV95wq-PIDK81sWsXXxA0JYy0KCWw';

/** 書き込む先のシート（URLの gid= の数字）。0 にすると開いているシートを使う */
const TARGET_GID = 281377521;

/** 企業名が入っている列の見出し。表記が違うときはここを直す */
const KEY_HEADER = '企業名';

/** 書き込む列の見出し */
const OUT_HEADER = '契約形態';

/** 月次管理ブックの媒体シート */
const MEDIA = [
  'AD', 'アフィ', 'CS', 'MEO', '制作', 'PR', '風評', 'タレントシェア',
  'LINE公式アカウント', 'メディア', 'ベトナム', 'ASP', 'マス広告(MA)・その他',
];

const KEIS = ['ストック', 'ショット'];

// ========================= メニュー =========================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('契約形態')
    .addItem('契約形態を入れる', 'fillKeiyaku')
    .addItem('ヒット状況だけ見る', 'dryRun')
    .addToUi();
}

// ========================= メイン =========================

function dryRun() { run_(true); }
function fillKeiyaku() { run_(false); }

function run_(dryRunOnly) {
  const ui = SpreadsheetApp.getUi();
  const sh = targetSheet_();
  const map = buildMap_();

  const head = findHeader_(sh);
  const keyCol = head.cols[KEY_HEADER];
  if (!keyCol) {
    ui.alert('「' + KEY_HEADER + '」という見出しの列が見つかりませんでした。\n' +
             '見つかった見出し：' + Object.keys(head.cols).join(' / '));
    return;
  }

  const lastRow = sh.getLastRow();
  const names = sh.getRange(head.row + 1, keyCol, lastRow - head.row, 1).getValues();

  const out = [];
  let hit = 0, blank = 0;
  names.forEach(function (r) {
    const name = String(r[0] == null ? '' : r[0]).trim();
    if (!name || name.indexOf('ここより上に行追加') >= 0 || name === '---') {
      out.push(['']); blank++; return;
    }
    const v = map[normalizeName_(name)] || '';
    if (v) hit++;
    out.push([v]);
  });

  const target = names.length - blank;
  const msg = '対象 ' + target + '行のうち ' + hit + '行がヒットしました（' +
              (target ? Math.round(hit / target * 100) : 0) + '%）。\n' +
              'ヒットしなかった行は空欄のままになります。';

  if (dryRunOnly) {
    ui.alert('ヒット状況', msg + '\n\n※このメニューでは書き込みません。', ui.ButtonSet.OK);
    return;
  }

  const ok = ui.alert('契約形態を書き込みます',
    msg + '\n\n「' + OUT_HEADER + '」列にだけ書き込みます。\n' +
    '他の列には一切触れません。実行しますか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  let outCol = head.cols[OUT_HEADER];
  if (!outCol) {
    outCol = sh.getLastColumn() + 1;  // 表の右端に新しく作る
    sh.getRange(head.row, outCol).setValue(OUT_HEADER)
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864');
  }
  sh.getRange(head.row + 1, outCol, out.length, 1).setValues(out);

  ui.alert('完了', hit + '行に契約形態を入れました。', ui.ButtonSet.OK);
}

// ========================= 対象シート =========================

function targetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (TARGET_GID) {
    const found = ss.getSheets().filter(function (s) {
      return s.getSheetId() === TARGET_GID;
    });
    if (found.length) return found[0];
  }
  return ss.getActiveSheet();
}

/** 見出し行を上から5行以内で探す。KEY_HEADER がある行を見出しとみなす */
function findHeader_(sh) {
  const scan = sh.getRange(1, 1, Math.min(5, sh.getLastRow()),
                           sh.getLastColumn()).getValues();
  for (let i = 0; i < scan.length; i++) {
    const cols = {};
    scan[i].forEach(function (v, j) {
      const name = String(v == null ? '' : v).replace(/\n/g, '').trim();
      if (name && !(name in cols)) cols[name] = j + 1;
    });
    if (cols[KEY_HEADER]) return { row: i + 1, cols: cols };
  }
  return { row: 1, cols: {} };
}

// ========================= 引き当て表を作る =========================

/**
 * 会社名の表記ゆれを吸収する。
 * 株式会社などの法人格、空白、括弧書き、全角半角の違いを落として比べる。
 */
function normalizeName_(s) {
  return String(s)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[（(\[【].*?[)）\]】]/g, '')          // 括弧書きを落とす
    .replace(/<[^>]*>/g, '')                        // <既存> などを落とす
    .replace(/(株式会社|有限会社|合同会社|一般社団法人|医療法人社団|医療法人|学校法人|公益財団法人|一般財団法人|特定非営利活動法人|弁護士法人|司法書士法人)/g, '')
    .replace(/[\s　_･・\-ー―]/g, '')
    .toLowerCase()
    .trim();
}

/** 月次管理ブックから 会社名 → 契約形態 の対応表を作る */
function buildMap_() {
  const src = SpreadsheetApp.openById(SRC_ID);
  const acc = {};  // 正規化名 → {ストック:true, ショット:true}

  MEDIA.forEach(function (name) {
    const sh = src.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;

    const values = sh.getRange(1, 1, sh.getLastRow(),
                               Math.min(sh.getLastColumn(), 80)).getValues();
    const head = {};
    values[0].forEach(function (v, i) {
      const k = String(v == null ? '' : v).replace(/\n/g, '').trim();
      if (k && !(k in head)) head[k] = i;
    });
    if (!('計上種別' in head)) return;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const kei = String(row[head['計上種別']] || '').trim();
      if (KEIS.indexOf(kei) < 0) continue;

      ['エンドクライアント名', '社名'].forEach(function (col) {
        if (!(col in head)) return;
        const raw = String(row[head[col]] || '').trim();
        if (!raw) return;
        const key = normalizeName_(raw);
        if (!key) return;
        if (!acc[key]) acc[key] = {};
        acc[key][kei] = true;
      });
    }
  });

  const map = {};
  Object.keys(acc).forEach(function (k) {
    const d = acc[k];
    map[k] = (d['ストック'] && d['ショット']) ? 'ストック＋ショット'
           : d['ストック'] ? 'ストック'
           : d['ショット'] ? 'ショット' : '';
  });
  return map;
}
