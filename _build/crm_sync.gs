/**
 * CRMスプシ組み込み用スクリプト（Google Apps Script）
 *
 * 月次管理ブックの全媒体シートを読んで、CRMスプシに
 * 「ストック」「ショット」の2タブを自動生成する。
 *
 * ■ 入れ方
 *   CRMスプシを開く → 拡張機能 → Apps Script → 中身を全部消してこれを貼る
 *   → 保存 → CRMスプシを再読み込み → メニューに「案件データ」が出る
 *
 * ■ 使い方
 *   メニュー「案件データ」→「最新に更新」を押すだけ。
 *   初回だけ Google の承認画面が出る（自分のアカウントで実行する許可）。
 *
 * ■ 安全のために
 *   ・書き込むのは OUT_PREFIX で始まる専用タブだけ。CRMの既存タブには触らない。
 *   ・専用タブのA1に目印を書く。目印が無いタブには上書きせず中止する。
 *   ・月次管理ブックは読むだけ。一切変更しない。
 */

// ========================= 設定 =========================

/** 月次管理ブックのID（URLの /d/ と /edit の間） */
const SRC_ID = '11j9QaZ74iOp1ABNV95wq-PIDK81sWsXXxA0JYy0KCWw';

/** 自動生成タブの名前の頭につける文字。これ以外のタブには絶対に書き込まない */
const OUT_PREFIX = '【自動】';

/** 上書きしてよいタブかを見分けるための目印。A1に入る */
const MARKER = '※このタブはスクリプトが自動更新します。手入力しないでください※';

/**
 * 取り込む媒体シート。
 * LINE公式アカウントだけにしたいときは、この配列を
 *   ['LINE公式アカウント']
 * に書き換えるだけでよい。
 */
const MEDIA = [
  'AD', 'アフィ', 'CS', 'MEO', '制作', 'PR', '風評', 'タレントシェア',
  'LINE公式アカウント', 'メディア', 'ベトナム', 'ASP', 'マス広告(MA)・その他',
];

/** 担当の列名。媒体によって呼び方が違う（ASPだけ広告主担当／媒体担当） */
const ROLE_NAMES = [
  'アカウント', 'コンサル', '運用①', '運用②', '運用③',
  '広告主担当①', '広告主担当②', '媒体担当①', '媒体担当②',
];

const KEIS = ['ストック', 'ショット'];

const HEADERS = [
  'エンドクライアント名', '社名（請求先）', '売上', '粗利', '粗利率',
  '案件数', '媒体', '商材カテゴリ', '担当', '最新対象月', 'もう一方の契約形態',
];

// ========================= メニュー =========================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('案件データ')
    .addItem('最新に更新', 'updateAll')
    .addToUi();
}

// ========================= メイン =========================

function updateAll() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert(
    '案件データの更新',
    OUT_PREFIX + 'ストック と ' + OUT_PREFIX + 'ショット の2タブを作り直します。\n' +
    'それ以外のタブには一切書き込みません。\n\n実行しますか？',
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  const recs = readSource_();
  if (!recs.length) {
    ui.alert('データが1件も取れませんでした。月次管理ブックのIDとシート名を確認してください。');
    return;
  }

  KEIS.forEach(function (kei) {
    writeSheet_(kei, aggregate_(recs, kei));
  });

  ui.alert('更新しました',
    '案件数 ' + recs.length + '件を取り込みました。\n' +
    '更新日時：' + Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm'),
    ui.ButtonSet.OK);
}

// ========================= 読み取り =========================

/** 見出しの表記ゆれ（改行・前後の空白）を吸収する */
function normHeader_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').trim();
}

function toNum_(v) {
  return typeof v === 'number' && isFinite(v) ? v : 0;
}

function toStr_(v) {
  return String(v == null ? '' : v).trim();
}

/**
 * 月次管理ブックの全媒体シートから 1行=1案件 のレコードを取り出す。
 * 列の並びは媒体ごとに違うので、位置ではなく見出し名で解決する。
 */
function readSource_() {
  const src = SpreadsheetApp.openById(SRC_ID);
  const out = [];

  MEDIA.forEach(function (name) {
    const sh = src.getSheetByName(name);
    if (!sh) return;  // シートが無ければ黙って飛ばす（媒体の増減に耐えるため）

    const lastRow = sh.getLastRow();
    const lastCol = Math.min(sh.getLastColumn(), 80);
    if (lastRow < 2) return;

    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const head = {};
    values[0].forEach(function (v, i) {
      const key = normHeader_(v);
      if (key && !(key in head)) head[key] = i;  // 同名列は左側（通常ぶん）を優先
    });

    const need = ['社名', 'エンドクライアント名', '商材', '計上種別',
                  '請求額（税抜）', '利益'];
    if (need.some(function (k) { return !(k in head); })) return;

    const roleIdx = ROLE_NAMES
      .filter(function (n) { return n in head; })
      .map(function (n) { return head[n]; });

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (row[0] === '' || row[0] == null) continue;

      const kei = toStr_(row[head['計上種別']]);
      const uri = toNum_(row[head['請求額（税抜）']]);
      if (!kei && uri === 0) continue;  // 空行

      const tanto = [];
      roleIdx.forEach(function (i) {
        const t = toStr_(row[i]);
        if (t && tanto.indexOf(t) === -1) tanto.push(t);
      });

      out.push({
        media: name,
        company: toStr_(row[head['社名']]),
        end: toStr_(row[head['エンドクライアント名']]) || toStr_(row[head['社名']]),
        cat: toStr_(row[head['商材']]),
        kei: kei,
        month: '対象月' in head ? toStr_(row[head['対象月']]) : '',
        uri: uri,
        rieki: toNum_(row[head['利益']]),
        tanto: tanto,
      });
    }
  });

  return out;
}

// ========================= 集計 =========================

/** 「8月分」などから並べ替え用の数値を作る。判別できないものは0 */
function monthKey_(s) {
  const m = String(s).match(/\d+/);
  return m ? Number(m[0]) : 0;
}

function pushUniq_(arr, v) {
  if (v && arr.indexOf(v) === -1) arr.push(v);
}

/** 契約形態 kei の案件だけを、エンドクライアント名ごとに1行にまとめる */
function aggregate_(recs, kei) {
  const other = (kei === 'ストック') ? 'ショット' : 'ストック';

  // そのエンドがどの契約形態を持っているか（片方だけ見て取りこぼさないため）
  const has = {};
  recs.forEach(function (r) {
    if (KEIS.indexOf(r.kei) === -1) return;
    if (!has[r.end]) has[r.end] = {};
    has[r.end][r.kei] = true;
  });

  const g = {};
  recs.forEach(function (r) {
    if (r.kei !== kei || !r.end) return;
    if (!g[r.end]) {
      g[r.end] = { uri: 0, rieki: 0, n: 0, company: [], media: [], cat: [],
                   tanto: [], month: [] };
    }
    const d = g[r.end];
    d.n++;
    d.uri += r.uri;
    d.rieki += r.rieki;
    pushUniq_(d.company, r.company);
    pushUniq_(d.media, r.media);
    pushUniq_(d.cat, r.cat);
    pushUniq_(d.month, r.month);
    r.tanto.forEach(function (t) { pushUniq_(d.tanto, t); });
  });

  return Object.keys(g)
    .sort(function (a, b) { return g[b].uri - g[a].uri; })
    .map(function (k) {
      const d = g[k];
      const latest = d.month.slice().sort(function (a, b) {
        return monthKey_(b) - monthKey_(a);
      })[0] || '';
      return [
        k,
        d.company.join(' / '),
        d.uri,
        d.rieki,
        d.uri ? d.rieki / d.uri : '',
        d.n,
        d.media.join(' / '),
        d.cat.join(' / '),
        d.tanto.join(' / '),
        latest,
        (has[k] && has[k][other]) ? (other + 'あり') : '',
      ];
    });
}

// ========================= 書き込み =========================

/**
 * 専用タブを作り直す。
 * 目印の無いタブ（＝人が作ったタブ）には絶対に書き込まない。
 */
function writeSheet_(kei, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = OUT_PREFIX + kei;
  let sh = ss.getSheetByName(name);

  if (sh) {
    const a1 = toStr_(sh.getRange('A1').getValue());
    if (a1 !== MARKER) {
      throw new Error(
        '「' + name + '」タブのA1に目印がありません。\n' +
        '人が作ったタブを上書きしないよう中止しました。\n' +
        'このタブを消すか、名前を変えてから再実行してください。');
    }
    sh.clear();
  } else {
    sh = ss.insertSheet(name);  // 末尾に新規追加。既存タブの位置は動かさない
  }

  const stamp = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');
  const dual = rows.filter(function (r) { return r[10]; }).length;

  sh.getRange('A1').setValue(MARKER);
  sh.getRange('A2').setValue(
    kei + '｜1行=1エンドクライアント／金額は' + kei + 'ぶんのみ・税抜／' +
    rows.length + '社（うち' + dual + '社はもう一方の契約形態も保有）／更新 ' + stamp);

  sh.getRange(3, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864')
    .setVerticalAlignment('middle').setWrap(true);

  if (rows.length) {
    sh.getRange(4, 1, rows.length, HEADERS.length).setValues(rows);
    sh.getRange(4, 3, rows.length, 2).setNumberFormat('#,##0');
    sh.getRange(4, 5, rows.length, 1).setNumberFormat('0.0%');
    sh.getRange(4, 6, rows.length, 1).setNumberFormat('0');
  }

  sh.setFrozenRows(3);
  sh.getRange(3, 1, rows.length + 1, HEADERS.length).createFilter();
  [220, 200, 100, 95, 70, 60, 130, 180, 190, 90, 120]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}
