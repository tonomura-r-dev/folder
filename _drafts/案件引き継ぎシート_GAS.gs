/**
 * 案件引き継ぎシート セットアップスクリプト
 *
 * 使い方
 * 1. 対象のGoogleスプレッドシートを開く
 * 2. 拡張機能 > Apps Script を開き、このファイルの中身を貼り付けて保存
 * 3. 上部の関数選択で setupHandoffWorkbook を選び「実行」（初回は権限承認が必要）
 * 4. 実行後、「⓪設定」「案件テンプレート」の2タブが生成される
 * 5. スプレッドシートを再読み込みすると「案件引き継ぎ」メニューが出るので、
 *    ②新規案件タブを作成 で案件ごとのタブをワンクリック複製できる
 * 6. 「⓪設定」タブの社員名簿・Lパターンなど、社内固有の値は各自入力すること
 *    （このスクリプトは仕組みだけ作る。中身の実データはスクリプトからは分からないため）
 */

const CONFIG = {
  MASTER_SHEET_NAME: '⓪設定',
  TEMPLATE_SHEET_NAME: '案件テンプレート',
};

const COLORS = {
  header共通: '#3467B2',
  headerWeb広告: '#D9661F',
  headerLINEOA: '#00897B',
  subHeader: '#8E4EC6',
  requiredLabel: '#FBE2E1',
  condRequiredLabel: '#FCE5CD',
  optionalLabel: '#EFEFEF',
  warn: '#F4CCCC',
};

const COL = { LABEL: 1, VALUE_START: 2, NOTE: 10 };
const LAST_COL = 10;

const MASTER_LISTS = [
  { name: 'LIST_直代理店', header: '直/代理店', values: ['直', '代理店'] },
  { name: 'LIST_契約形態', header: '契約形態', values: ['自動更新', '単年契約', '都度契約'] },
  { name: 'LIST_請求書タイミング', header: '請求書提出タイミング', values: ['月末締め翌月10日払い', '都度払い', 'その他'] },
  { name: 'LIST_レポートタイミング', header: 'レポート提出タイミング', values: ['週次', '月次', '随時', 'その他'] },
  { name: 'LIST_MTG周期', header: '定例MTG周期', values: ['毎週', '隔週', '月1回', '不定期'] },
  { name: 'LIST_KPI種別', header: 'KPI種別', values: ['CPA', 'CPO', 'ROAS', 'CV数', '友だち数', 'その他'] },
  { name: 'LIST_指名広告可否', header: '指名広告の可否', values: ['可', '不可', '条件付き可'] },
  { name: 'LIST_リマケ可否', header: 'リマーケティング拡張可否', values: ['可', '不可', '要相談'] },
  { name: 'LIST_広告文変更スパン', header: '広告文の変更スパン', values: ['週1回', '月1回', '随時', 'キャンペーン都度'] },
  { name: 'LIST_契約プラン', header: 'LINEOA契約プラン', values: ['コミュニケーションプラン', 'ライトプラン', 'フリープラン'] },
  { name: 'LIST_配信頻度', header: '定期配信の頻度', values: ['週1', '月2', '不定期', 'その他'] },
  { name: 'LIST_リッチメニュー頻度', header: 'リッチメニュー更新頻度', values: ['固定', '月次更新', 'キャンペーン都度'] },
  { name: 'LIST_社員', header: '営業担当・運用担当（社員名簿）', values: ['（ここに社員名を1行1名で追加してください）'] },
  { name: 'LIST_Lパターン', header: 'Lパターン（意味確認中）', values: ['（用途確認後に選択肢を追加してください）'] },
];

const COMMON_FIELDS_BASIC = [
  { label: '案件名', required: true, type: 'text' },
  { label: 'クライアント名', required: true, type: 'text' },
  { label: 'エンドクライアント名', required: false, type: 'text', note: 'クライアントが代理店/仲介の場合の実広告主' },
  { label: '直/代理店 案件', required: true, type: 'dropdown', list: 'LIST_直代理店' },
  { label: 'マージン（%）', required: true, type: 'number' },
  { label: 'Lパターン', required: true, type: 'dropdown', list: 'LIST_Lパターン', note: '意味を要確認。分かり次第⓪設定タブの選択肢を確定してください' },
  { label: 'アカウント名', required: true, type: 'text' },
  { label: '営業担当', required: true, type: 'dropdown', list: 'LIST_社員' },
];

const COMMON_FIELDS_PRODUCT = [
  { label: '商材内容', required: true, type: 'text' },
  { label: '目標（KPI）主指標／目標値', required: true, type: 'dropdown-number', list: 'LIST_KPI種別' },
];

const COMMON_FIELDS_OPS = [
  { label: 'GTM', required: true, type: 'text', note: 'コンテナIDまたはメールアドレス' },
  { label: 'GA', required: true, type: 'text', note: 'メールアドレスとプロパティ名' },
  { label: '予算（グロス/月）', required: true, type: 'number' },
  { label: '予算（ネット/月）', required: true, type: 'number' },
  { label: '請求書提出タイミング', required: true, type: 'dropdown', list: 'LIST_請求書タイミング' },
  { label: '請求書明細表URL', required: true, type: 'url' },
  { label: '都度払い用特記欄', required: 'conditional', type: 'text', note: '都度払い契約の場合のみ。入金済み予算・掲載期間目安の日付' },
  { label: 'レポート提出タイミング', required: true, type: 'dropdown', list: 'LIST_レポートタイミング' },
  { label: '特殊レポートの有無', required: true, type: 'checkbox', note: '有りの場合は右の備考欄に内容を記載' },
  { label: '定例MTG実施周期', required: true, type: 'dropdown', list: 'LIST_MTG周期' },
  { label: '備考（都度発生タスク等）', required: false, type: 'text' },
];

const AD_FIELDS_PRE = [
  { label: 'Marchant Center', required: 'conditional', type: 'text', note: 'メールアドレス＋アカウント名。EC・ショッピング広告利用時のみ' },
  { label: '指名広告の掲載可否', required: true, type: 'dropdown', list: 'LIST_指名広告可否' },
];

const AD_FIELDS_POST = [
  { label: '除外KWリストの利用有無', required: true, type: 'checkbox' },
  { label: '新規キャンペーンでの除外KW共通設定有無', required: true, type: 'checkbox' },
  { label: 'リマーケティング広告の拡張可否', required: true, type: 'dropdown', list: 'LIST_リマケ可否' },
  { label: '掲載時間', required: false, type: 'text' },
  { label: '配信地域', required: true, type: 'text', note: '指定なしの場合は「指定なし」と明記' },
  { label: '性別・年齢設定', required: true, type: 'text', note: '指定なしの場合は「指定なし」と明記' },
  { label: 'プレースメント設定', required: false, type: 'text' },
  { label: '広告文の変更スパン', required: true, type: 'dropdown', list: 'LIST_広告文変更スパン' },
  { label: '広告文の最終変更日', required: true, type: 'date' },
  { label: 'CVテスト_申込者名', required: false, type: 'text' },
  { label: 'CVテスト_商品', required: false, type: 'text' },
  { label: 'CVテスト_支払い方法', required: false, type: 'text' },
  { label: 'CVテスト_再テスト前に先方確認が必要か', required: false, type: 'checkbox' },
];

const LINEOA_BASIC_FIELDS = [
  { label: 'LINE公式アカウント名／ベーシックID', required: true, type: 'text' },
  { label: '管理画面アクセス権限保有者', required: true, type: 'text' },
  { label: '契約プラン', required: true, type: 'dropdown', list: 'LIST_契約プラン' },
];

const LINEOA_DISTRIBUTION_FIELDS = [
  { label: 'あいさつメッセージ設定の有無', required: true, type: 'checkbox' },
  { label: 'あいさつメッセージ内容／リンク', required: 'conditional', type: 'url' },
  { label: 'ステップ配信（シナリオ配信）の有無', required: true, type: 'checkbox' },
  { label: 'シナリオ設計図の保管場所', required: 'conditional', type: 'url' },
  { label: '定期配信（一斉配信）の頻度', required: true, type: 'dropdown', list: 'LIST_配信頻度' },
  { label: '配信ネタ管理表のリンク', required: false, type: 'url' },
];

const LINEOA_SEGMENT_FIELDS = [
  { label: '連携外部ツールの有無', required: true, type: 'checkbox', note: 'Lステップ 等' },
  { label: '連携ツール名・契約者', required: 'conditional', type: 'text' },
  { label: 'タグ付けルールの有無', required: true, type: 'checkbox' },
  { label: 'タグ一覧・付与ルールのリンク', required: 'conditional', type: 'url' },
  { label: 'セグメント配信の有無', required: true, type: 'checkbox' },
  { label: 'セグメント条件・配信ルール', required: 'conditional', type: 'text' },
];

const LINEOA_SUPPORT_FIELDS = [
  { label: '有人チャット対応の有無', required: true, type: 'checkbox' },
  { label: '有人対応ルール', required: 'conditional', type: 'text', note: '対応時間帯／一次対応者／エスカレーション先' },
  { label: '自動応答（キーワード応答）の有無', required: true, type: 'checkbox' },
];

const LINEOA_RICHMENU_FIELDS = [
  { label: 'リッチメニュー枚数・出し分け有無', required: true, type: 'number' },
  { label: 'リッチメニュー更新頻度', required: true, type: 'dropdown', list: 'LIST_リッチメニュー頻度' },
  { label: 'デザインデータの保管場所', required: true, type: 'url' },
  { label: 'クーポン／ショップカード利用有無', required: false, type: 'checkbox' },
  { label: 'LINE広告連携有無', required: false, type: 'checkbox', note: '「有」の場合はWeb広告ブロックの運用媒体マトリクス側をメイン記載とする' },
  { label: '特記事項', required: false, type: 'text' },
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('案件引き継ぎ')
    .addItem('①初期セットアップ（⓪設定・テンプレート作成）', 'setupHandoffWorkbook')
    .addItem('②新規案件タブを作成', 'createNewProjectTab')
    .addToUi();
}

function setupHandoffWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupMasterSheet_(ss);
  setupTemplateSheet_(ss);
  SpreadsheetApp.getUi().alert(
    'セットアップが完了しました。\n' +
      '「⓪設定」タブでプルダウンの中身（社員名簿など）を入力してから、\n' +
      '「案件引き継ぎ」メニュー→「②新規案件タブを作成」で案件タブを作ってください。'
  );
}

function createNewProjectTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const template = ss.getSheetByName(CONFIG.TEMPLATE_SHEET_NAME);
  if (!template) {
    ui.alert('先に「①初期セットアップ」を実行してください。');
    return;
  }
  const res = ui.prompt('新規案件タブを作成', '案件名を入力してください（タブ名になります）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const name = res.getResponseText().trim();
  if (!name) return;
  if (ss.getSheetByName(name)) {
    ui.alert('同名のタブが既にあります。');
    return;
  }
  const newSheet = template.copyTo(ss);
  newSheet.setName(name);
  ss.setActiveSheet(newSheet);
}

function setupMasterSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(CONFIG.MASTER_SHEET_NAME, 0);

  MASTER_LISTS.forEach((list, i) => {
    const col = i + 1;
    sheet.getRange(1, col).setValue(list.header);
    setStyle_(sheet.getRange(1, col), { bold: true, bg: '#E8E8E8' });
    sheet.getRange(2, col, list.values.length, 1).setValues(list.values.map((v) => [v]));
    sheet.setColumnWidth(col, 200);
    // 200行分の余白を確保し、あとから選択肢を追加してもプルダウンが自動追従するようにする
    ss.setNamedRange(list.name, sheet.getRange(2, col, 200, 1));
  });

  sheet.setFrozenRows(1);
}

function setupTemplateSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(CONFIG.TEMPLATE_SHEET_NAME);

  sheet.setColumnWidth(COL.LABEL, 230);
  for (let c = COL.VALUE_START; c < COL.NOTE; c++) sheet.setColumnWidth(c, 110);
  sheet.setColumnWidth(COL.NOTE, 320);

  const requiredCells = [];
  let row = 1;

  row = writeBlockHeader_(sheet, row, '■共通情報', COLORS.header共通);
  COMMON_FIELDS_BASIC.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));
  row = writeCheckboxGroup_(sheet, row, '運用媒体区分', true, ['Web広告', 'LINEOA'], 'チェックした媒体のブロックのみ入力してください');
  row = writeRepeatingTable_(sheet, row, { title: '先方担当者', required: true, columns: ['氏名', 'リテラシー(1-5)'], repeatCount: 3, note: '足りない場合は行を挿入' });
  COMMON_FIELDS_PRODUCT.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));
  row = writeRepeatingTable_(sheet, row, { title: 'ベンチマーク競合', required: false, columns: ['社名', 'URL'], repeatCount: 5 });
  COMMON_FIELDS_OPS.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));

  row++;
  row = writeBlockHeader_(sheet, row, '■Web広告（運用媒体区分でチェックした場合のみ入力）', COLORS.headerWeb広告);
  row = writeMediaStatusMatrix_(sheet, row);
  AD_FIELDS_PRE.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));
  row = writeRepeatingTable_(sheet, row, { title: 'ビッグワード停止設定', required: false, columns: ['KW', '停止理由', '再開提案可否'], repeatCount: 3, note: '停止中のものがある場合のみ' });
  AD_FIELDS_POST.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));
  row = writeCheckboxGroup_(sheet, row, 'デバイス設定', true, ['PC', 'スマホ', 'タブレット']);
  row = writeCheckboxGroup_(sheet, row, 'P-max・SNS等の最適化拡張可否', true, ['P-max', 'LINE', 'Facebook・Instagram']);

  row++;
  row = writeBlockHeader_(sheet, row, '■LINEOA（新設・運用媒体区分でチェックした場合のみ入力）', COLORS.headerLINEOA);
  row = writeSubHeader_(sheet, row, '基本情報・契約');
  LINEOA_BASIC_FIELDS.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));

  const rowLimit = row;
  row = writeField_(sheet, row, { label: '月間メッセージ通数上限', required: true, type: 'number' }, requiredCells);
  const rowUsed = row;
  row = writeField_(sheet, row, { label: '当月メッセージ消化数', required: true, type: 'number' }, requiredCells);
  const rowRate = row;
  row = writeField_(sheet, row, { label: 'メッセージ消化率', required: false, type: 'text', note: '自動計算（編集不要）' }, requiredCells);
  sheet.getRange(rowRate, COL.VALUE_START).setFormula(`=IFERROR(B${rowUsed}/B${rowLimit},"")`).setNumberFormat('0%');
  addNumberWarning_(sheet, sheet.getRange(rowRate, COL.VALUE_START), 0.8);
  row = writeField_(sheet, row, { label: '追加メッセージ購入の有無', required: true, type: 'checkbox' }, requiredCells);
  row = writeField_(sheet, row, { label: '友だち数', required: true, type: 'number', note: '取得日を右の備考欄に併記' }, requiredCells);

  row = writeSubHeader_(sheet, row, '配信・シナリオ');
  LINEOA_DISTRIBUTION_FIELDS.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));

  row = writeSubHeader_(sheet, row, '顧客管理（タグ・セグメント）');
  LINEOA_SEGMENT_FIELDS.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));

  row = writeSubHeader_(sheet, row, '接客対応');
  LINEOA_SUPPORT_FIELDS.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));

  row = writeSubHeader_(sheet, row, 'リッチメニュー・その他');
  LINEOA_RICHMENU_FIELDS.forEach((f) => (row = writeField_(sheet, row, f, requiredCells)));

  addRequiredEmptyFormatting_(sheet, requiredCells);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function writeBlockHeader_(sheet, row, title, bg) {
  const range = sheet.getRange(row, 1, 1, LAST_COL);
  range.merge().setValue(title);
  setStyle_(range, { bg, bold: true, color: '#FFFFFF', fontSize: 12 });
  sheet.setRowHeight(row, 28);
  return row + 1;
}

function writeSubHeader_(sheet, row, title) {
  const range = sheet.getRange(row, 1, 1, LAST_COL);
  range.merge().setValue('　' + title);
  setStyle_(range, { bg: COLORS.subHeader, bold: true, color: '#FFFFFF', fontSize: 10 });
  return row + 1;
}

function labelBg_(required) {
  if (required === true) return COLORS.requiredLabel;
  if (required === 'conditional') return COLORS.condRequiredLabel;
  return COLORS.optionalLabel;
}

function writeField_(sheet, row, field, requiredCellsAccumulator) {
  const labelCell = sheet.getRange(row, COL.LABEL);
  labelCell.setValue(field.label);
  setStyle_(labelCell, { bg: labelBg_(field.required), bold: true });

  const parts = field.type.split('-');
  parts.forEach((part, i) => {
    const cell = sheet.getRange(row, COL.VALUE_START + i);
    if (part === 'checkbox') cell.insertCheckboxes();
    else if (part === 'dropdown') applyDropdown_(cell, field.list);
    else if (part === 'date') cell.setNumberFormat('yyyy/mm/dd');
    else if (part === 'number') cell.setNumberFormat('#,##0');
  });

  if (field.required === true) {
    requiredCellsAccumulator.push(sheet.getRange(row, COL.VALUE_START));
  }

  if (field.note) {
    const noteCell = sheet.getRange(row, COL.NOTE);
    noteCell.setValue(field.note);
    setStyle_(noteCell, { color: '#888888', fontSize: 9, fontStyle: 'italic' });
  }

  return row + 1;
}

function writeCheckboxGroup_(sheet, row, label, required, options, note) {
  const labelCell = sheet.getRange(row, COL.LABEL);
  labelCell.setValue(label);
  setStyle_(labelCell, { bg: labelBg_(required), bold: true });

  let col = COL.VALUE_START;
  options.forEach((opt) => {
    const optLabelCell = sheet.getRange(row, col);
    optLabelCell.setValue(opt);
    setStyle_(optLabelCell, { fontSize: 9, color: '#555555' });
    sheet.getRange(row, col + 1).insertCheckboxes();
    col += 2;
  });

  if (note) {
    const noteCell = sheet.getRange(row, COL.NOTE);
    noteCell.setValue(note);
    setStyle_(noteCell, { color: '#888888', fontSize: 9, fontStyle: 'italic' });
  }

  return row + 1;
}

function writeRepeatingTable_(sheet, row, { title, required, columns, repeatCount, note }) {
  sheet.getRange(row, COL.LABEL).setValue(title);
  setStyle_(sheet.getRange(row, COL.LABEL), { bg: labelBg_(required), bold: true });
  columns.forEach((colName, i) => {
    const c = sheet.getRange(row, COL.VALUE_START + i);
    c.setValue(colName);
    setStyle_(c, { fontSize: 9, color: '#555555', bold: true });
  });
  if (note) {
    sheet.getRange(row, COL.NOTE).setValue(note);
    setStyle_(sheet.getRange(row, COL.NOTE), { color: '#888888', fontSize: 9, fontStyle: 'italic' });
  }
  row++;

  for (let i = 0; i < repeatCount; i++) {
    sheet.getRange(row, COL.VALUE_START, 1, columns.length).setBorder(true, true, true, true, false, false);
    row++;
  }
  return row;
}

function writeMediaStatusMatrix_(sheet, row) {
  const media = ['Google広告', 'YSA（Yahoo検索）', 'YDA（Yahooディスプレイ）', 'Facebook広告', 'LINE広告', 'TikTok広告', 'X（Twitter）広告'];

  sheet.getRange(row, COL.LABEL).setValue('運用媒体');
  setStyle_(sheet.getRange(row, COL.LABEL), { bg: COLORS.requiredLabel, bold: true });
  ['稼働中', '過去実施(停止中)', '補足'].forEach((h, i) => {
    const c = sheet.getRange(row, COL.VALUE_START + i);
    c.setValue(h);
    setStyle_(c, { fontSize: 9, bold: true, color: '#555555' });
  });
  row++;

  media.forEach((m) => {
    sheet.getRange(row, COL.LABEL).setValue('　' + m);
    sheet.getRange(row, COL.VALUE_START).insertCheckboxes();
    sheet.getRange(row, COL.VALUE_START + 1).insertCheckboxes();
    row++;
  });

  return row;
}

function applyDropdown_(cell, listName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listRange = ss.getRangeByName(listName);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(listRange, true).setAllowInvalid(false).build();
  cell.setDataValidation(rule);
}

function addRequiredEmptyFormatting_(sheet, ranges) {
  if (!ranges.length) return;
  const rule = SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground(COLORS.warn).setRanges(ranges).build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function addNumberWarning_(sheet, range, threshold) {
  const rule = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(threshold).setBackground(COLORS.warn).setRanges([range]).build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function setStyle_(range, { bg, bold, color, fontSize, align, fontStyle } = {}) {
  if (bg) range.setBackground(bg);
  if (bold !== undefined) range.setFontWeight(bold ? 'bold' : 'normal');
  if (color) range.setFontColor(color);
  if (fontSize) range.setFontSize(fontSize);
  if (align) range.setHorizontalAlignment(align);
  if (fontStyle) range.setFontStyle(fontStyle);
}
