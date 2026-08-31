/**
 * 提案管理シートの「右端に5列だけ足す」スクリプト。
 *
 * 佐村さんの最初の依頼（商材カテゴリ×契約形態×売上×粗利×担当）を、
 * 月次管理ブックから企業名で引き当てて、今あるシートの右に並べる。
 *
 * ■ 行は1行も足さない・動かさない
 *   ・行の挿入／削除／並べ替えを一切しない
 *   ・今ある列（A〜O）にも書き込まない
 *   ・書くのは右端に新しく作る5列だけ
 *
 * ■ 入れ方
 *   提案管理シート（まずはコピーで試す）を開く
 *   → 拡張機能 → Apps Script → 中身を全部消してこれを貼る → 保存
 *   → シートを再読み込み → メニュー「実績5列」が出る
 *   → ①ヒット状況を見る（書き込まない） → ②5列を入れる → （戻すなら）③消す
 */

// ========================= 設定 =========================

/** 月次管理ブックのID（読むだけ。一切変更しない） */
const SRC_ID = '11j9QaZ74iOp1ABNV95wq-PIDK81sWsXXxA0JYy0KCWw';

/** 書き込む先のシート（URLの gid= の数字）。0 にすると開いているシートを使う */
const TARGET_GID = 0;

/** 企業名が入っている列の見出し */
const KEY_HEADER = '企業名';

/**
 * 足す5列。順番どおりに右へ並ぶ。
 * 「実績」と付けているのは、シートのDYM売上（提案額）と混ざらないようにするため。
 */
const OUT_COLS = ['商材カテゴリ', '契約形態', '実績売上', '実績粗利', '実績担当'];

/** この文字を含む行は飛ばす（シート最下部の目印の行） */
const TAIL_MARK = 'ここより上に行追加';

/** 月次管理ブックの媒体シート */
const MEDIA = [
  'AD', 'アフィ', 'CS', 'MEO', '制作', 'PR', '風評', 'タレントシェア',
  'LINE公式アカウント', 'メディア', 'ベトナム', 'ASP', 'マス広告(MA)・その他',
];

/** 担当の列名。媒体によって呼び方が違う（ASPだけ広告主担当／媒体担当） */
const TANTO_COLS = ['アカウント', 'コンサル', '運用①', '運用②', '運用③',
                    '広告主担当①', '広告主担当②', '媒体担当①', '媒体担当②'];

const KEIS = ['ストック', 'ショット'];

/** 消すときのために、入れた列の位置を控えておくキー */
const UNDO_KEY = 'last5cols';

// ========================= メニュー =========================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('実績5列')
    .addItem('① ヒット状況を見る（書き込まない）', 'preview5cols')
    .addItem('② 5列を入れる', 'add5cols')
    .addSeparator()
    .addItem('③ 入れた5列を消す', 'remove5cols')
    .addToUi();
}

function preview5cols() { run_(true); }
function add5cols() { run_(false); }

// ========================= メイン =========================

function run_(dryRun) {
  const ui = SpreadsheetApp.getUi();
  const sh = targetSheet_();
  const head = findHeader_(sh);
  if (!head.cols[KEY_HEADER]) {
    ui.alert('「' + KEY_HEADER + '」列が見つかりませんでした。\n\n' +
             '見つかった見出し：\n' + Object.keys(head.cols).join(' / '),
             ui.ButtonSet.OK);
    return;
  }

  const last = sh.getLastRow();
  const names = sh.getRange(head.row + 1, head.cols[KEY_HEADER],
                            last - head.row, 1).getValues();
  const map = buildMap_();

  // 1行ずつ、企業名で引き当てる
  const out = [];
  let target = 0, hit = 0;
  const misses = [];
  names.forEach(function (r) {
    const name = String(r[0] == null ? '' : r[0]).trim();
    if (!name || name.indexOf(TAIL_MARK) >= 0 || name === '---') {
      out.push(['', '', '', '', '']);       // 目印の行と空行は触らない
      return;
    }
    target++;
    const d = map[normalizeName_(name)];
    if (!d) {
      out.push(['', '', '', '', '']);
      if (misses.length < 15) misses.push(name);
      return;
    }
    hit++;
    out.push([d.shozai.join(' / '), keiyakuLabel_(d.kei),
              d.uri, d.rieki, d.tanto.join(' / ')]);
  });

  const pct = target ? Math.round(hit / target * 100) : 0;
  const msg = 'このシートの ' + target + '行のうち、\n' +
              '月次管理ブックに同じ会社があったのは ' + hit + '行です（' + pct + '%）。\n' +
              '見つからなかった行は5列とも空欄になります。\n\n' +
              '引き当てられなかった例：\n　' +
              (misses.length ? misses.slice(0, 8).join('\n　') : 'なし');

  if (dryRun) {
    ui.alert('ヒット状況', msg + '\n\n※このメニューでは書き込みません。', ui.ButtonSet.OK);
    return;
  }

  const at = sh.getLastColumn() + 1;   // 右端の次の列から
  const colName = colLetter_(at) + '〜' + colLetter_(at + OUT_COLS.length - 1);

  const ok = ui.alert('5列を入れます',
    msg + '\n\n【今あるデータに何が起きるか】\n' +
    '　・行は1行も足さない／消さない／動かさない\n' +
    '　・今ある列（A〜' + colLetter_(at - 1) + '）にも書き込まない\n' +
    '　・' + colName + '列を新しく作って、そこにだけ書きます\n' +
    '　・メニュー「③ 入れた5列を消す」で戻せます\n\n' +
    '実行しますか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  // 見出し
  sh.getRange(head.row, at, 1, OUT_COLS.length).setValues([OUT_COLS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#808080')
    .setHorizontalAlignment('center');
  // 中身
  sh.getRange(head.row + 1, at, out.length, OUT_COLS.length).setValues(out);
  // 金額の書式
  sh.getRange(head.row + 1, at + 2, out.length, 2).setNumberFormat('¥#,##0');

  PropertiesService.getDocumentProperties().setProperty(UNDO_KEY, JSON.stringify({
    gid: sh.getSheetId(), at: at, count: OUT_COLS.length, headRow: head.row,
  }));

  ui.alert('完了',
    colName + '列に ' + hit + '行ぶん入れました。\n' +
    '行は1行も動いていません。\n\n' +
    '戻すときはメニュー「③ 入れた5列を消す」。', ui.ButtonSet.OK);
}

// ========================= 消す =========================

/**
 * 入れた5列だけを消す。
 * 見出しが OUT_COLS のままかを先に確かめ、違えば何もしない
 * （列を動かした後などに、関係ない列を巻き込まないため）。
 */
function remove5cols() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(UNDO_KEY);
  if (!raw) {
    ui.alert('消せる列がありません', 'このスクリプトで入れた記録が無い状態です。',
             ui.ButtonSet.OK);
    return;
  }
  const rec = JSON.parse(raw);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(
    function (s) { return s.getSheetId() === rec.gid; })[0];
  if (!sh) {
    ui.alert('消せません', '入れたシートが見つかりませんでした。', ui.ButtonSet.OK);
    return;
  }

  const got = sh.getRange(rec.headRow, rec.at, 1, rec.count).getValues()[0]
    .map(function (v) { return String(v == null ? '' : v).trim(); });
  const same = got.every(function (v, i) { return v === OUT_COLS[i]; });
  if (!same) {
    ui.alert('中止しました',
      colLetter_(rec.at) + '列からの見出しが、入れたときと違っています。\n' +
      '　今：' + got.join(' / ') + '\n' +
      '　入れたとき：' + OUT_COLS.join(' / ') + '\n\n' +
      '関係ない列を消さないよう、何もしませんでした。手で消してください。',
      ui.ButtonSet.OK);
    return;
  }

  const ok = ui.alert('入れた5列を消します',
    colLetter_(rec.at) + '〜' + colLetter_(rec.at + rec.count - 1) + '列を削除します。\n' +
    '（' + OUT_COLS.join(' / ') + '）\n\n実行しますか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  sh.deleteColumns(rec.at, rec.count);
  props.deleteProperty(UNDO_KEY);
  ui.alert('消しました', rec.count + '列を削除しました。', ui.ButtonSet.OK);
}

// ========================= シートを読む =========================

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
    if (cols[KEY_HEADER]) return { row: i + 1, cols: cols };
  }
  return { row: 1, cols: {} };
}

/** 1始まりの列番号を A / AB のような文字にする */
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - 1 - m) / 26;
  }
  return s;
}

// ========================= 月次管理ブックを集計 =========================

/**
 * 会社名の表記ゆれを吸収する。
 * 法人格・空白・括弧書き・全角半角の違いと、シート側の接頭辞を落として比べる。
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

function toNum_(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }
function toStr_(v) { return String(v == null ? '' : v).trim(); }
function push_(lst, v) { if (v && lst.indexOf(v) < 0) lst.push(v); }

function keiyakuLabel_(kei) {
  const s = kei['ストック'], o = kei['ショット'];
  return (s && o) ? 'ストック＋ショット' : s ? 'ストック' : o ? 'ショット' : '';
}

/** 月次管理ブックを 正規化した会社名 → 5列ぶんの中身 の表にする */
function buildMap_() {
  const src = SpreadsheetApp.openById(SRC_ID);
  const g = {};

  MEDIA.forEach(function (sheetName) {
    const sh = src.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const values = sh.getRange(1, 1, sh.getLastRow(),
                               Math.min(sh.getLastColumn(), 80)).getValues();
    const h = {};
    values[0].forEach(function (v, i) {
      const k = toStr_(v).replace(/\n/g, '');
      if (k && !(k in h)) h[k] = i;      // 同名列は左（通常ぶん）を優先
    });
    if (!('計上種別' in h) || !('請求額（税抜）' in h)) return;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const end = toStr_(row[h['エンドクライアント名']]) || toStr_(row[h['社名']]);
      if (!end) continue;
      const kei = toStr_(row[h['計上種別']]);
      const uri = toNum_(row[h['請求額（税抜）']]);
      if (KEIS.indexOf(kei) < 0 && uri === 0) continue;   // 空行

      // 会社名は表記ゆれがあるので、正規化したものをキーにしてまとめる
      const key = normalizeName_(end);
      if (!key) continue;
      if (!g[key]) g[key] = { uri: 0, rieki: 0, kei: {}, shozai: [], tanto: [] };
      const d = g[key];
      d.uri += uri;
      d.rieki += toNum_(row[h['利益']]);
      if (KEIS.indexOf(kei) >= 0) d.kei[kei] = true;
      if ('商材' in h) push_(d.shozai, toStr_(row[h['商材']]));
      TANTO_COLS.forEach(function (c) {
        if (c in h) push_(d.tanto, toStr_(row[h[c]]));
      });
    }
  });
  return g;
}
