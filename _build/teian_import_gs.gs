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
 * ■ どの列に何を入れるか
 *   提案     実行した月（26/8 の形）
 *   ☑        既にシートにある会社なら「今の件数+1」、新規なら 1
 *   企業名   エンドクライアント名
 *   既存     ●
 *   商流     いま経由している代理店の会社名。直の実績があれば「直」、無ければ「不明」
 *   外注先   ① 同じ会社が既にシートにあればその外注先 ② 無ければ月次ブックの「発注先名」
 *   提案商材 SHOZAI の値（既定は 商流変更）
 *   業界     ① 同じ会社が既にシートにあればその業界 ② 無ければ社名から判定
 *   担当①②  シートの担当欄にある名字表記に合わせる
 *   DYM売上  全媒体の実績売上（税抜）
 *   LINE売上 うち媒体「LINE公式アカウント」ぶん
 *
 * ■ 空欄のままにするもの
 *   結論 / 追客 / Ac名 … 提案してから手で入れる列なので触らない
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
const LINE_MEDIA = 'LINE公式アカウント';

/**
 * 社名から業界を当てる規則。シートに同じ会社が無かったときだけ使う。
 * 上から順に見て最初に当たったものを採用するので、並び順が結果を決める。
 * 業界名は提案管理シートで使われている語彙に合わせること。
 */
const GYOKAI_RULES = [
  ['医療', /クリニック|医院|病院|歯科|皮膚科|眼科|整形外科|美容外科|接骨院|鍼灸|医療法人|診療所|CLINIC|Clinic/],
  ['製薬', /製薬|薬品工業|薬品株式会社|薬工/],
  ['士業', /弁護士|司法書士|税理士|行政書士|社会保険労務士|法律事務所|法務事務所|綜合法律|会計事務所|探偵/],
  ['教育', /大学|学園|学院|専門学校|高等学校|高校|中学|予備校|ゼミナール|学習|塾|スクール|学校法人|教育|保育|幼稚園|アカデミー|自動車学校|カレッジ/],
  ['金融', /信用金庫|信用組合|信用保証|銀行|証券|生命保険|損害保険|保険|カード|クレジット|ファイナンス|信販|信金/],
  ['冠婚葬祭', /セレモニー|葬儀|葬祭|冠婚|ブライダル|ウェディング|メモリアル|霊園|互助会|教会/],
  ['宿泊', /ホテル|旅館|リゾート|温泉|ホテルズ|ホステル|HOTEL|Hotel/],
  ['旅行', /旅行|ツーリズム|トラベル/],
  ['観光', /観光協会|観光開発|観光事業|観光公社/],
  ['自治体', /市役所|町役場|村役場|区役所|県庁|府庁|労働局|職業安定所|ハローワーク/],
  ['団体', /農業協同組合|生活協同組合|漁業協同組合|生活協同|協同組合|生協|商工会|商工会議所|振興協会|振興会|連合会|協議会|事業団|公益財団法人|公益社団法人|一般財団法人|一般社団法人|特定非営利活動法人|NPO|組合|労働組合|福祉会/],
  ['自動車', /自動車|モータース|モビリティ|トヨタ|日産|ホンダ|スズキ|ダイハツ|カローラ|レンタリース|レンタカー|車輌|車輛|中古車/],
  ['ペット', /ペット|動物病院|動物/],
  ['不動産', /工務店|ハウス|住宅|不動産|建設|建築|ホーム|地所|リアルティ|ハウジング|住建|建工|ホームズ|レジデンシャル|都市開発|賃貸|リフォーム|設計|林業/],
  ['住宅設備', /住宅設備|建材|サッシ|給湯|カーテン|じゅうたん|インテリア|家具|板金|ガーデン|エクステリア/],
  ['インフラ', /ガス|電力|電気|エネルギー|石油|燃料|運輸|運送|物流|航空|鉄道|電鉄|交通|バス|通信|水道|高速道路|空港/],
  ['メディア', /放送|テレビ|新聞|出版|メディア|ラジオ|エフエム|ケーブル|印刷|書店|書籍|雑誌|コミュニケーションズ|広告/],
  ['アミューズ', /ゴルフ|スポーツ|フィットネス|ジム|スイミング|ボウル|競輪|競馬|競艇|BOATRACE|レジャー|遊園|パチンコ|アミューズ|エンタテインメント|エンターテインメント|劇場|シネマ|映画|ダンス|カラオケ|サウナ|スパ|温浴/],
  ['人材', /人材|スタッフ|キャリア|求人|派遣|転職|就職|リクルート|ワークス|ジョブ|ナース|エージェント/],
  ['買取', /買取|質屋|リサイクル|中古|かんてい局|オークション/],
  ['美容', /美容|エステ|サロン|ヘアー|ヘア|ネイル|化粧品|コスメ|ビューティ|BEAUTY|Beauty|脱毛/],
  ['食品', /食品|製菓|製パン|乳業|フーズ|食肉|酒造|醸造|農産|牧場|水産|畜産|菓子|パン|味噌|醤油|青果|鮮魚|ベーカリー|スイーツ|ファーム|農園|養蜂|食料/],
  ['飲食', /レストラン|居酒屋|焼肉|ラーメン|らーめん|寿司|鮨|カフェ|ダイニング|飲食|食堂|ビストロ|うどん|そば|ピザ|バーガー|喫茶/],
  ['アパレル', /アパレル|衣料|繊維|靴|シューズ|バッグ|ジュエリー|時計|眼鏡|メガネ|ランドセル|ユニフォーム|ファッション|ブティック/],
  ['生活サービス', /クリーニング|引越|清掃|警備|理容|便利屋|修理|メンテナンス/],
  ['その他店舗', /薬局|ドラッグ|スーパー|ストア|百貨店|商店|市場|ショップ|マート|センター|モール|プラザ|専門店|ホームセンター/],
  ['健康食品', /健康食品|サプリ|プロテイン|青汁/],
  ['生活用品', /生活用品|雑貨|日用品|文具|タオル|寝具|家電/],
  ['宅配', /宅配|通販|デリバリー/],
  ['BtoB', /商事|産業|工業|製作所|機械|部品|資材|問屋|システム|ソリューション|テクノロジー/],
];

function guessGyokai_(name) {
  const s = String(name == null ? '' : name);
  if (!s) return '';
  for (let i = 0; i < GYOKAI_RULES.length; i++) {
    if (GYOKAI_RULES[i][1].test(s)) return GYOKAI_RULES[i][0];
  }
  return '';
}

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

  let gFromSheet = 0, gFromName = 0, hitGaichu = 0, kinFromSheet = 0;
  const data = recs.map(function (r) {
    const key = normalizeName_(r.name);
    // 業界は ①シートに同じ会社があればその値 ②無ければ社名から判定
    let gyokai = known.gyokai[key] || '';
    if (gyokai) { gFromSheet++; }
    else { gyokai = guessGyokai_(r.name); if (gyokai) gFromName++; }
    // 外注先も ①シート ②月次ブックの発注先名 の順
    const gaichu = known.gaichu[key] || r.gaichu;
    if (gaichu) hitGaichu++;
    // 金額も ①シートの直近の提案額 ②無ければ実績。
    // シートの金額は「提案額」で、実績とは桁が違う。同じ会社の過去の提案額が
    // あるならそちらのほうが、既存行と並べたときに意味が揃う。
    const hasKin = (key in known.dym) || (key in known.line);
    if (hasKin) kinFromSheet++;
    const dym = (key in known.dym) ? known.dym[key] : r.uri;
    const line = (key in known.line) ? known.line[key] : r.lineUri;

    return { rec: r, gyokai: gyokai, gaichu: gaichu, dym: dym, line: line,
             count: (known.count[key] || 0) + 1 };
  });

  const dup = data.filter(function (d) { return d.count > 1; }).length;
  const msg = '月次管理ブックの既存顧客：' + data.length + '社\n' +
              '　うち ' + dup + '社は、このシートに既に行があります\n' +
              '　　（消さずに追加して、☑ を「今の件数+1」にします）\n' +
              '　業界　：シートから ' + gFromSheet + '社 ／ 社名から ' + gFromName + '社\n' +
              '　外注先：' + hitGaichu + '社（分からないものは空欄）\n' +
              '　金額　：シートの過去の提案額から ' + kinFromSheet + '社 ／ ' +
              '残りは実績を入れます\n' +
              '　担当の名字は、このシートにある ' + known.surnames.length + '種類に合わせます';

  if (dryRun) {
    ui.alert('取り込む件数', msg + '\n\n※このメニューでは追加しません。', ui.ButtonSet.OK);
    return;
  }

  const ok = ui.alert('行を追加します',
    msg + '\n\n既存の行は1セルも変更しません。行の挿入だけです。\n' +
    '結論・追客・Ac名は空欄のままにします（提案してから手で入れる列なので）。\n' +
    'DYM売上・LINE売上には実績を入れます。\n\n実行しますか？',
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
    row[c['外注先'] - 1] = d.gaichu;
    row[c['提案商材'] - 1] = SHOZAI;
    row[c['業界'] - 1] = d.gyokai;
    row[c['担当①'] - 1] = d.rec.tanto1;
    row[c['担当②'] - 1] = d.rec.tanto2;
    row[c['DYM売上'] - 1] = d.dym;
    row[c['LINE売上'] - 1] = d.line;
    return row;
  });
  sh.getRange(at, 1, rows.length, width).setValues(rows);
  // 金額の書式は既存行に合わせる
  sh.getRange(at, c['DYM売上'], rows.length, 1).setNumberFormat('¥#,##0');
  sh.getRange(at, c['LINE売上'], rows.length, 1).setNumberFormat('¥#,##0');

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
 * シートから引き当て表を作る。
 *   gyokai / gaichu / dym / line : 正規化した企業名 → その列の値
 *   count    : 正規化した企業名 → 今ある行数（☑の付け直しに使う）
 *   surnames : 担当①②に実際に入っている表記の一覧（長い順）
 *
 * 同じ会社が複数行あるときは**下の行（新しい提案）を優先**する。
 * 上から順に見て、値が入っていれば上書きしていくだけ。
 * 空欄・`--`・`#REF!` は「値なし」として飛ばす。
 */
function readSheet_(sh, head) {
  const last = sh.getLastRow();
  const out = { gyokai: {}, gaichu: {}, dym: {}, line: {},
                count: {}, surnames: [] };
  if (last <= head.row) return out;

  const vals = sh.getRange(head.row + 1, 1, last - head.row, sh.getLastColumn()).getValues();
  const cName = head.cols['企業名'] - 1;
  const cGyo = head.cols['業界'] - 1;
  const cGai = head.cols['外注先'] - 1;
  const cDym = head.cols['DYM売上'] - 1;
  const cLine = head.cols['LINE売上'] - 1;
  const cT1 = head.cols['担当①'] - 1;
  const cT2 = head.cols['担当②'] - 1;
  const sur = {};

  function text(v) {
    const s = String(v == null ? '' : v).trim();
    return (!s || s === '--' || s === '-' || s.indexOf('#REF') === 0) ? '' : s;
  }

  vals.forEach(function (r) {
    const raw = String(r[cName] == null ? '' : r[cName]).trim();
    if (raw && raw.indexOf(TAIL_MARK) < 0 && raw !== '---') {
      const key = normalizeName_(raw);
      if (key) {
        out.count[key] = (out.count[key] || 0) + 1;
        const g = text(r[cGyo]);
        if (g) out.gyokai[key] = g;
        const o = text(r[cGai]);
        if (o) out.gaichu[key] = o;
        // 金額は数値のときだけ拾う。#REF! や文字列は無視する
        if (typeof r[cDym] === 'number' && isFinite(r[cDym])) out.dym[key] = r[cDym];
        if (typeof r[cLine] === 'number' && isFinite(r[cLine])) out.line[key] = r[cLine];
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
  out.surnames = Object.keys(sur).sort(function (a, b) { return b.length - a.length; });
  return out;
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
        g[end] = { name: end, uri: 0, lineUri: 0, choku: false, agency: [],
                   gaichuList: [], acct: '', cons: '' };
      }
      const d = g[end];
      d.uri += uri;
      if (name === LINE_MEDIA) d.lineUri += uri;
      // 外注先。CS・制作・PR・タレントシェアにしか「発注先名」の列が無い
      if ('発注先名' in h) push_(d.gaichuList, toStr_(row[h['発注先名']]));

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
      d.gaichu = d.gaichuList.join(' / ');
      d.tanto1 = toSurname_(d.acct, surnames);
      d.tanto2 = toSurname_(d.cons, surnames);
      return d;
    })
    .sort(function (a, b) { return b.uri - a.uri; });
}
