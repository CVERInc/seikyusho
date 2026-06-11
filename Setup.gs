/**
 * 請求書自動化システム — 初期セットアップ
 *
 * 使い方:
 * 1. 初回デプロイ時、Apps Script エディタから `runSetup()` を実行
 * 2. データ Sheet・3 つの PDF テンプレート・設定ページ・番号管理ページが自動作成される
 * 3. 完了後、Apps Script の「プロジェクトのプロパティ」に SPREADSHEET_ID が保存される
 */

const SHEET_NAMES = {
  SUBMISSIONS: '申請データ',
  SETTINGS: '設定',
  COUNTER: '番号管理',
  TPL_OVERSEAS: 'TPL_海外',
  TPL_WITH_WHT: 'TPL_源泉あり',
  TPL_WITHOUT_WHT: 'TPL_源泉なし'
};

const SUBMISSION_COLUMNS = [
  '審査ステータス',
  '源泉適用',
  'タイムスタンプ',
  '言語',
  '居住地',
  '氏名',
  '住所',
  '電話番号',
  'メールアドレス',
  '振込銀行名',
  '支店名',
  '口座名義',
  '口座番号',
  '通貨',
  '請求明細(JSON)',
  '備考',
  '請求書番号',
  'PDF生成日',
  'PDF URL',
  'メモ'
];

/**
 * 🚀 ワンショット初期セットアップ（BYOG 導入用）。
 *    One-shot first-time setup (for BYOG self-deployment).
 *
 * 正しい順序で初回セットアップをまとめて実行する。冪等（何度実行しても安全）：
 * Runs the whole first-time setup in the right order — idempotent, safe to re-run:
 *   1. runSetup()                 — Sheet・設定・番号管理・3 PDF テンプレートを作成/修復
 *   2. installSheetMenuTrigger()  — 「請求書」カスタムメニューを有効化
 *   3. installStatusEditTrigger() — 承認/差戻/源泉適用の自動処理トリガーを配線
 *   4. populateUsageSheet()       — 「使い方」シートを生成
 *
 * 完了後 / after this: 「設定」シートに会社情報と notification_email を記入 →
 * デプロイ → 新しいデプロイ → ウェブアプリ（実行: 自分 / アクセス: 全員）→
 * /exec URL を「設定」の form_url に貼り、populateUsageSheet() を再実行。
 */
function setup() {
  const url = runSetup();              // runSetup 内で「使い方」シートも生成される
  installSheetMenuTrigger();
  installStatusEditTrigger();
  Logger.log('✅ seikyusho セットアップ完了 / setup complete.\n  ' + url +
    '\n  次へ / next: 「設定」シート記入 → Web アプリとしてデプロイ → form_url を保存。');
  return url;
}

/**
 * Sheet を開いた時に onOpen を自動実行するためのトリガー設定
 * スタンドアロンスクリプトでは simple trigger が起動しないため、installable trigger を使用する
 */
function installSheetMenuTrigger() {
  // 容器バインド（テンプレートのコピー）では simple onOpen が自動起動するため、
  // installable な onOpen トリガーは不要。入れると onOpen が二重に走りメニューが重複する。
  let active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active) {
    Logger.log('バインド済み: simple onOpen が動作するため installable onOpen トリガーはスキップします。');
    return;
  }

  const ss = getMainSpreadsheet_();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  Logger.log('メニュー用 onOpen トリガーをインストールしました。Sheet を再読み込みすると「請求書」メニューが表示されます。');
}

/**
 * セル編集時に handleStatusEdit を呼び出すトリガー。
 * 「審査ステータス」列が変更されたとき、申請者へ通知メールを送る。
 */
function installStatusEditTrigger() {
  const ss = getMainSpreadsheet_();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleStatusEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('handleStatusEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('審査ステータス編集用トリガーをインストールしました。');
}

function runSetup() {
  let ss;

  // 容器バインド（テンプレートを「コピーを作成」した副本）なら、
  // 紐づくスプレッドシート自身を主シートにし、その ID を必ず書き込む。
  // （コピー時に旧 ID が残っていても上書きするので、誤って master を指さない）
  let active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }

  if (active) {
    ss = active;
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', active.getId());
    Logger.log('バインド先のスプレッドシートを使用: ' + ss.getUrl());
  } else {
    // standalone（CLI / 手動デプロイ）: 従来どおり既存 ID を開くか新規作成
    const existingId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (existingId) {
      try {
        ss = SpreadsheetApp.openById(existingId);
        Logger.log('既存のスプレッドシートを使用: ' + ss.getUrl());
      } catch (e) {
        ss = null;
      }
    }
    if (!ss) {
      ss = SpreadsheetApp.create('請求書管理');
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
      Logger.log('新規スプレッドシート作成: ' + ss.getUrl());
    }
  }

  setupSubmissionsSheet_(ss);
  setupSettingsSheet_(ss);
  setupCounterSheet_(ss);
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_OVERSEAS, { withConsumptionTax: false, withWithholdingTax: false });
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_WITH_WHT, { withConsumptionTax: true, withWithholdingTax: true });
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_WITHOUT_WHT, { withConsumptionTax: true, withWithholdingTax: false });

  // 使い方ガイドも常に用意する（テンプレートを複製した副本にもそのまま入るように）
  populateUsageSheet();

  removeDefaultBlankSheets_(ss);

  Logger.log('セットアップ完了！\n  Spreadsheet URL: ' + ss.getUrl());
  return ss.getUrl();
}

/**
 * 既定の空シートを言語非依存で削除する。
 * 新規スプレッドシートの既定シート名は言語で変わる（日「シート1」/英「Sheet1」/
 * 繁中「工作表1」/簡中「工作表1」/韓「시트1」…）。名前で決め打ちせず、
 * 「自分たちのシート以外で空のもの」を削除することで全ロケールに対応する。
 */
function removeDefaultBlankSheets_(ss) {
  const ours = [
    SHEET_NAMES.SUBMISSIONS, SHEET_NAMES.SETTINGS, SHEET_NAMES.COUNTER,
    SHEET_NAMES.TPL_OVERSEAS, SHEET_NAMES.TPL_WITH_WHT, SHEET_NAMES.TPL_WITHOUT_WHT,
    '使い方'
  ];
  ss.getSheets().forEach(sh => {
    if (ours.indexOf(sh.getName()) >= 0) return;        // 自分たちのシートは残す
    // 空シートのみ削除（データのある利用者のシートには触れない）。最低1枚は残す。
    if (sh.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });
}

function setupSubmissionsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.SUBMISSIONS);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SUBMISSION_COLUMNS.length).setValues([SUBMISSION_COLUMNS]);
    sheet.getRange(1, 1, 1, SUBMISSION_COLUMNS.length)
      .setBackground('#1f8e3d')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(12, 350);
    sheet.setColumnWidth(13, 200);
    sheet.setColumnWidth(18, 250);
  }

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .clearDataValidations();

  const reviewCol = SUBMISSION_COLUMNS.indexOf('審査ステータス') + 1;
  const withholdingCol = SUBMISSION_COLUMNS.indexOf('源泉適用') + 1;

  const reviewRange = sheet.getRange(2, reviewCol, sheet.getMaxRows() - 1, 1);
  const reviewRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['pending', 'approved', 'rejected'], true)
    .setAllowInvalid(false)
    .build();
  reviewRange.setDataValidation(reviewRule);

  const withholdingRange = sheet.getRange(2, withholdingCol, sheet.getMaxRows() - 1, 1);
  const withholdingRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['yes', 'no', 'N/A'], true)
    .setAllowInvalid(false)
    .build();
  withholdingRange.setDataValidation(withholdingRule);

  // 電話番号・口座番号は先頭の「0」が数値化で消えないよう、列をプレーンテキスト書式にする。
  const phoneCol = SUBMISSION_COLUMNS.indexOf('電話番号') + 1;
  const accountNumberCol = SUBMISSION_COLUMNS.indexOf('口座番号') + 1;
  sheet.getRange(2, phoneCol, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  sheet.getRange(2, accountNumberCol, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
}

function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (sheet) return;

  sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
  const data = [
    ['key', 'value', '備考'],
    ['company_name_ja-JP', 'YOUR_COMPANY_NAME_JA_JP', '請求書宛先（日本語）'],
    ['company_name_en-US', 'YOUR_COMPANY_NAME_EN_US', 'Recipient name (English)'],
    ['company_name_zh-TW', 'YOUR_COMPANY_NAME_ZH_TW', '請求書宛先（台湾華語・繁体字表記）'],
    ['company_name_es-ES', 'YOUR_COMPANY_NAME_ES_ES', 'Nombre del destinatario (Español)'],
    ['company_address', 'YOUR_COMPANY_ADDRESS', '本店所在地'],
    ['corporate_number', '', '法人番号'],
    ['representative', 'YOUR_REPRESENTATIVE_NAME', '代表者'],
    ['form_url', '', '【任意】申請者に渡すフォームの公開URL（デプロイ後の /exec 形式。空欄なら自動取得するが /dev 版になる場合あり）'],
    ['qualified_invoice_number', '', '適格請求書発行事業者登録番号 (例: T1234567890123、未登録なら空欄)'],
    ['show_corporate_number', 'no', 'PDF に法人番号を表示する (yes/no、インボイス未登録時は no 推奨)'],
    ['notification_email', '', '【入力必須】請求書送付先メールアドレス（管理者用）'],
    ['email_subject_approved', '請求書 {{INVOICE_NO}} が承認されました', '【承認時】申請者宛メール件名'],
    ['email_body_approved', '{{APPLICANT_NAME}}様\n\nご請求ありがとうございました。\n\n請求が承認されましたので、お支払いまで今しばらくお待ちください。', '【承認時】申請者宛メール本文（{{APPLICANT_NAME}} {{INVOICE_NO}} 使用可）'],
    ['email_subject_rejected', '請求書 {{INVOICE_NO}} は差戻となりました', '【差戻時】申請者宛メール件名'],
    ['email_body_rejected', '{{APPLICANT_NAME}}様\n\nご請求ありがとうございました。\n\n請求内容に不備がありましたので差戻いたします。\n\n差戻事由を担当者にご確認ください。\n\nご確認いただきましたら請求の再申請をお願いいたします。', '【差戻時】申請者宛メール本文（{{APPLICANT_NAME}} {{INVOICE_NO}} 使用可）'],
    ['consumption_tax_rate', '0.10', '消費税率 (10%)'],
    ['withholding_tax_rate', '0.1021', '源泉所得税率 (10.21%)'],
    ['withholding_threshold', '1000000', '源泉所得税率の基準額 (100万円)'],
    ['withholding_tax_rate_over', '0.2042', '100万円超部分の源泉所得税率 (20.42%)'],
    ['pdf_folder_id', '', '【任意】PDF保存先 Drive フォルダ ID（空欄なら My Drive 直下）']
  ];
  sheet.getRange(1, 1, data.length, 3).setValues(data);
  sheet.getRange(1, 1, 1, 3).setBackground('#1f8e3d').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 350);
  sheet.setFrozenRows(1);
}

function setupCounterSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.COUNTER);
  if (sheet) return;

  sheet = ss.insertSheet(SHEET_NAMES.COUNTER);
  sheet.getRange(1, 1, 1, 2).setValues([['年月 (YYYYMM)', '最終番号']]);
  sheet.getRange(1, 1, 1, 2).setBackground('#1f8e3d').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * テンプレートシート作成（固定レイアウト版）
 *
 * 設計方針: 申請者情報は「固定枠＋1行縮小表示」で、内容が長くても折り返さず
 * フォントを縮めて収める（孤字・左右ズレを防止し、どの請求書も同じ版面になる）。
 * 住所だけは全幅1行を割り当て、長い日本語住所も1行で収まるようにする。
 * 明細の品名のみ、長ければ折り返して複数行（パディングは行高で確保）。
 *
 * 共有列: A/G=余白、B(ラベル/品名左)・C(値/品名右)・D(ラベル2/数量)・E(値2/単価)・F(値2/金額)。
 */
function setupTemplateSheet_(ss, sheetName, options) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);
  sheet.setHiddenGridlines(true);

  const GREEN = '#1f8e3d';
  const LIGHT = '#e8f5e9';
  const HEAD = '#d9ead3';
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;

  sheet.setColumnWidth(1, 24);   // A 左余白
  sheet.setColumnWidth(2, 100);  // B ラベル / 品名(左半)
  sheet.setColumnWidth(3, 210);  // C 値   / 品名(右半)
  sheet.setColumnWidth(4, 80);   // D ラベル2 / 数量
  sheet.setColumnWidth(5, 100);  // E 値2 / 単価
  sheet.setColumnWidth(6, 120);  // F 値2 / 金額
  sheet.setColumnWidth(7, 24);   // G 右余白

  // ── タイトル＋発行情報 ──
  sheet.getRange('B2:C2').merge();
  sheet.getRange('B2').setValue('請求書').setFontSize(26).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange('E2').setValue('作成日：').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange('F2').setValue('{{ISSUE_DATE}}').setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange('E3').setValue('NO：').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange('F3').setValue('{{INVOICE_NO}}').setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(2, 40);
  sheet.setRowHeight(3, 24);

  // ── 宛先 ──
  sheet.getRange('B5:F5').merge();
  sheet.getRange('B5').setValue('{{CLIENT_NAME}}　様')
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  // 枠線は合併範囲全体に対して設定する（アンカー単一セルに掛けると下端＝合併内部になり描画されない）
  sheet.getRange('B5:F5').setBorder(false, false, true, false, false, false, '#000000', SOLID);
  sheet.setRowHeight(5, 36);
  sheet.setRowHeight(6, 12);  // spacer

  // ── 申請者情報（住所は全幅1行、その他は整列グリッド・縦中央）──
  const label = (a1, text) => sheet.getRange(a1).setValue(text)
    .setHorizontalAlignment('right').setFontWeight('bold').setVerticalAlignment('middle');
  const value = (a1, text) => sheet.getRange(a1).setValue(text)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  label('B7', '住所：');     sheet.getRange('C7:F7').merge(); value('C7', '{{APPLICANT_ADDRESS}}');
  label('B8', '氏名：');     value('C8', '{{APPLICANT_NAME}}');
  label('D8', '電話番号：'); sheet.getRange('E8:F8').merge(); value('E8', '{{APPLICANT_PHONE}}');
  label('B9', '振込銀行：'); value('C9', '{{BANK_NAME}}');
  label('D9', '支店：');     sheet.getRange('E9:F9').merge(); value('E9', '{{BRANCH_NAME}}');
  label('B10', '口座名義：'); value('C10', '{{ACCOUNT_NAME}}');
  label('D10', '口座番号：'); sheet.getRange('E10:F10').merge(); value('E10', '{{ACCOUNT_NUMBER}}');
  for (let r = 7; r <= 10; r++) sheet.setRowHeight(r, 30);
  sheet.getRange('B10:F10').setBorder(false, false, true, false, false, false, GREEN, SOLID);
  sheet.setRowHeight(11, 14);  // spacer

  // ── 案内文 ──
  sheet.getRange('B12').setValue('下記の通りご請求申し上げます').setFontSize(11).setVerticalAlignment('middle');
  sheet.setRowHeight(12, 24);

  // ── 明細表ヘッダー ──
  const tableStartRow = 13;
  sheet.getRange('B13:C13').merge();
  sheet.getRange('B13').setValue('品名').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('D13').setValue('数量').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('E13').setValue('単価').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('F13').setValue('金額').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('B13:F13').setBackground(GREEN).setFontColor('#ffffff').setFontWeight('bold');
  sheet.setRowHeight(13, 28);

  // ── 明細行（最大 20 行ぶん用意。生成時に使用ぶんだけ表示し残りは hideRows で隠す。品名は B:C 合併で広く）──
  const totalItemRows = 20;
  for (let i = 0; i < totalItemRows; i++) {
    const row = tableStartRow + 1 + i;
    sheet.getRange(row, 2, 1, 2).merge();  // B:C 品名
    sheet.getRange(row, 2, 1, 5).setBackground(i % 2 === 0 ? LIGHT : '#ffffff');
    sheet.getRange(row, 2).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.getRange(row, 4).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(row, 5).setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.getRange(row, 6).setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 26);
  }

  // ── 合計欄（明細表の直後。ラベルは D:E 合併、値は F）──
  let r = tableStartRow + 1 + totalItemRows;
  const putTotal = (text, valueCell, emphasize) => {
    sheet.getRange(r, 4, 1, 2).merge();
    sheet.getRange(r, 4).setValue(text).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontWeight('bold').setBackground(emphasize ? GREEN : HEAD).setFontColor(emphasize ? '#ffffff' : '#000000');
    const v = sheet.getRange(r, 6).setValue(valueCell).setHorizontalAlignment('right').setVerticalAlignment('middle');
    if (emphasize) v.setFontWeight('bold').setBorder(true, true, true, true, false, false, GREEN, SOLID);
    sheet.setRowHeight(r, 26);
    r++;
  };
  putTotal('小計', '{{SUBTOTAL}}', false);
  // 「消費税 N%」ラベルは PDF 生成時に実際の税率へ自動補正される（fillTemplateValues_）。
  putTotal(options.withConsumptionTax ? '消費税 10%' : '消費税 0%',
           options.withConsumptionTax ? '{{CONSUMPTION_TAX}}' : '0', false);
  if (options.withWithholdingTax) putTotal('源泉所得税', '{{WITHHOLDING_TAX}}', false);
  putTotal('合計', '{{GRAND_TOTAL}}', true);

  // ── 備考（固定高の枠・折り返し可）──
  const noteLabelRow = r + 1;
  sheet.getRange(noteLabelRow, 2).setValue('備考：').setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(noteLabelRow, 24);
  const noteBoxRow = noteLabelRow + 1;
  const noteBox = sheet.getRange(noteBoxRow, 2, 4, 5);  // B:F × 4 行
  noteBox.merge();
  // 枠線は合併範囲全体に掛ける（アンカー単一セルだと下・右が合併内部になり描画されない）
  noteBox.setBorder(true, true, true, true, false, false, GREEN, SOLID);
  sheet.getRange(noteBoxRow, 2).setValue('{{NOTES}}')
    .setVerticalAlignment('top').setHorizontalAlignment('left').setWrap(true);
  for (let k = 0; k < 4; k++) sheet.setRowHeight(noteBoxRow + k, 22);

  // ── 発行元（B:F 全幅・小さめ）──
  let ir = noteBoxRow + 5;
  ['{{ISSUER_LINE_1}}', '{{ISSUER_LINE_2}}', '{{ISSUER_LINE_3}}'].forEach(ph => {
    sheet.getRange(ir, 2, 1, 5).merge();
    sheet.getRange(ir, 2).setValue(ph).setFontSize(9).setFontColor('#666666').setVerticalAlignment('middle');
    ir++;
  });
}

/**
 * 【メンテナンス】PDF テンプレート（TPL_海外 / TPL_源泉あり / TPL_源泉なし）の 3 枚だけを
 * 最新レイアウトで作り直す。申請データ・設定・カウンター・使い方シートには一切触れない。
 * 版面リニューアルを既存の運用コピーへ反映するため、エディタから手動実行する
 * （末尾アンダースコアなし＝実行ドロップダウンに表示）。
 */
function rebuildTemplates() {
  const ss = getMainSpreadsheet_();
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_OVERSEAS, { withConsumptionTax: false, withWithholdingTax: false });
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_WITH_WHT, { withConsumptionTax: true, withWithholdingTax: true });
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_WITHOUT_WHT, { withConsumptionTax: true, withWithholdingTax: false });
  const msg = '✅ テンプレート3種を最新レイアウトで再生成しました（申請データ・設定・カウンターは変更なし）。';
  Logger.log(msg);
  try { SpreadsheetApp.getActive().toast(msg, 'rebuildTemplates 完了', 8); } catch (e) {}
  return msg;
}

function getMainSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID が未設定です。先に runSetup() を実行してください。');
  return SpreadsheetApp.openById(id);
}

/**
 * 「使い方」シートに使用ガイドを書き込み、フォーマットを適用する
 * 一回だけ実行すれば OK。再実行すると内容が上書きされる。
 */
function populateUsageSheet() {
  const ss = getMainSpreadsheet_();
  let sheet = ss.getSheetByName('使い方');
  if (!sheet) sheet = ss.insertSheet('使い方');

  sheet.clear();
  sheet.clearFormats();

  // 設定シートから動的に値を取得
  const settings = getSettings_();
  const companyName = settings['company_name_ja-JP'] || '請求書システム';
  const adminEmail = settings.notification_email || '担当者のメールアドレス';
  const formUrl = settings.form_url || ScriptApp.getService().getUrl();
  const sheetUrl = ss.getUrl();
  const driveUrl = settings.pdf_folder_id ? ('https://drive.google.com/drive/folders/' + settings.pdf_folder_id) : '(未設定)';

  // [テキスト, タイプ, 行スパン] — スパン省略時は 1
  const rows = [
    [companyName + ' 請求書システム — 使い方', 'title', 3],
    ['', 'spacer'],
    ['申請者からの請求書申請を自動で受付・PDF生成・送信するシステム', 'body'],
    ['担当者は基本「メールを受け取って承認または差戻を選ぶ」だけで完結します', 'body'],
    ['', 'spacer'],
    ['📅 全体フロー', 'h1', 2],
    ['1. 申請者がフォーム送信', 'body'],
    ['2. スプレッドシートに自動記録(審査ステータス: pending)', 'body'],
    ['3. 請求書番号を自動採番(例:202605-005)', 'body'],
    ['4. PDF自動生成(居住地・源泉適用に応じて3パターンの様式を自動選択)', 'body'],
    ['5. 担当者にメール通知(PDF添付)', 'body'],
    ['6. 担当者がシート上で「承認」または「差戻」を選択', 'body'],
    ['7. 申請者に結果メール自動送信(担当者にもCC)', 'body'],
    ['8. PDFは Drive フォルダに自動保存', 'body'],
    ['', 'spacer'],
    ['📎 ① 請求書発行フォーム(申請者向け)', 'h1', 2],
    [formUrl, 'url'],
    ['', 'spacer'],
    ['用途:申請者が請求書の申請を行うウェブフォーム', 'body'],
    ['使う人:申請者(請求する側)', 'body'],
    ['対応言語:日本語 / English / 繁體中文 / Español(右上で切替可)', 'body'],
    ['送付方法:このURLを申請者に送る', 'body'],
    ['入力項目:居住地、氏名、住所、電話、メール、振込銀行情報、通貨(JPY/TWD/USD/EUR)、請求明細、備考', 'body'],
    ['', 'spacer'],
    ['⚠️ 担当者がこのURLを開くことは基本ありません。申請者に渡すだけ。', 'note', 2],
    ['', 'spacer'],
    ['📎 ② 請求書管理シート(このスプレッドシート)', 'h1', 2],
    [sheetUrl, 'url'],
    ['', 'spacer'],
    ['用途:全申請データの一元管理・承認/差戻の判定', 'body'],
    ['使う人:担当者', 'body'],
    ['アクセス頻度:申請があるたびに承認/差戻を行う', 'body'],
    ['', 'spacer'],
    ['【シート構成】', 'h2'],
    ['  申請データ … 全申請の記録(メイン)', 'body'],
    ['  設定 … 会社名・税率・通知メールアドレス・通知メール文面など', 'body'],
    ['  番号管理 … 月別の請求書番号採番カウンター(触らない)', 'body'],
    ['  TPL_海外 / TPL_源泉あり / TPL_源泉なし … PDF生成用テンプレート(触らない)', 'body'],
    ['', 'spacer'],
    ['【担当者が操作する場面】', 'h2'],
    ['', 'spacer'],
    ['A. 申請を承認する場合(基本操作)', 'h2'],
    ['  1. 「申請データ」シートで該当行を確認', 'body'],
    ['  2. メールに添付されたPDFの内容を確認', 'body'],
    ['  3. A列「審査ステータス」を approved に変更', 'body'],
    ['  4. 申請者に自動で承認メールが送信される(PDF添付・担当者にCC)', 'body'],
    ['', 'spacer'],
    ['B. 源泉徴収の適用を変更する場合(日本国内居住者のみ)', 'h2'],
    ['  1. B列「源泉適用」をプルダウンから yes に変更', 'body'],
    ['  2. PDFが自動的に源泉控除後の金額で再生成される(数十秒)', 'body'],
    ['  3. その後 A列を approved にすれば、最新PDFで承認メールが届く', 'body'],
    ['', 'spacer'],
    ['C. 申請内容に不備があり、差戻する場合', 'h2'],
    ['  1. A列「審査ステータス」を rejected に変更', 'body'],
    ['  2. 申請者に自動で差戻通知メールが送信される(担当者にCC・PDFなし)', 'body'],
    ['  3. 必要に応じて申請者に直接連絡し、再申請を依頼', 'body'],
    ['', 'spacer'],
    ['D. 過去の申請を検索したい場合', 'h2'],
    ['  ⌘+FかCtrl+Fで氏名・請求書番号で検索可能', 'body'],
    ['', 'spacer'],
    ['📎 ③ PDFまとめ(Google Driveフォルダ)', 'h1', 2],
    [driveUrl, 'url'],
    ['', 'spacer'],
    ['用途:全PDF請求書の保管庫', 'body'],
    ['使う人:担当者・税理士・会計', 'body'],
    ['ファイル形式:{請求書番号}_{氏名}_請求書.pdf', 'body'],
    ['  例:202605-004_山田太郎_請求書.pdf', 'body'],
    ['', 'spacer'],
    ['【このフォルダの使い方】', 'h2'],
    ['  税理士・会計事務所への月次提出時:このフォルダのリンクを共有するだけで全PDFが渡せるか、ダウンロードして別方法で送る', 'body'],
    ['  過去の請求書再発行:ファイル名で検索(請求書番号 or 氏名)', 'body'],
    ['  バックアップ不要:Google Driveに自動保存されているため安心', 'body'],
    ['', 'spacer'],
    ['⚠️ PDFの直接編集・削除は禁止(管理シートとの整合性が崩れます)。修正が必要な場合は管理シート側で再生成してください。', 'note', 2],
    ['', 'spacer'],
    ['📩 受信メールについて', 'h1', 2],
    ['担当者・申請者のメールボックスに、3種類のメールが自動で届きます。', 'body'],
    ['', 'spacer'],
    ['① 申請受付時 → 担当者へ', 'h2'],
    ['  件名:[' + companyName + '] 請求書 202605-XXX(申請者名様)', 'body'],
    ['  本文:申請内容の案内 + 承認/差戻のお願い + 請求書フォルダのリンク', 'body'],
    ['  添付:PDF請求書 1件', 'body'],
    ['', 'spacer'],
    ['② 承認時 → 申請者へ(担当者にCC)', 'h2'],
    ['  件名:請求書 202605-XXX が承認されました', 'body'],
    ['  本文:承認通知文', 'body'],
    ['  添付:PDF請求書', 'body'],
    ['', 'spacer'],
    ['③ 差戻時 → 申請者へ(担当者にCC)', 'h2'],
    ['  件名:請求書 202605-XXX は差戻となりました', 'body'],
    ['  本文:差戻通知文', 'body'],
    ['  添付:なし', 'body'],
    ['', 'spacer'],
    ['※ ②③ のメール文面は「設定」シートの email_subject_* / email_body_* で編集できます', 'body'],
    ['', 'spacer'],
    ['🆘 トラブル時の対応', 'h1', 2],
    ['❓ メールが届かない', 'h2'],
    ['  → 管理シートに該当申請があるか確認 → 担当者(技術)へ連絡', 'body'],
    ['', 'spacer'],
    ['❓ PDFの内容が間違っている', 'h2'],
    ['  → 管理シートで源泉適用などを修正 → 行を選択 → メニューから「再生成」', 'body'],
    ['', 'spacer'],
    ['❓ 申請者からフォームが送れないと連絡', 'h2'],
    ['  → URL有効期限切れの可能性 → 担当者(技術)へ連絡', 'body'],
    ['', 'spacer'],
    ['❓ その他不明点', 'h2'],
    ['  → 担当者(技術)へ連絡', 'body'],
  ];

  // 内容は B 列に配置。A 列・C 列は余白として確保（視覚的に窮屈にならないように）
  const CONTENT_COL = 2;

  const valuesToSet = [];
  rows.forEach(r => {
    const span = r[2] || 1;
    valuesToSet.push([r[0]]);
    for (let i = 1; i < span; i++) valuesToSet.push(['']);
  });

  sheet.getRange(1, CONTENT_COL, valuesToSet.length, 1).setValues(valuesToSet);

  let pos = 1;
  rows.forEach(r => {
    const text = r[0];
    const type = r[1];
    const span = r[2] || 1;
    const contentRange = sheet.getRange(pos, CONTENT_COL, span, 1);

    // title / h1 のみ背景色を A:C 全体に拡張(バナー感を出す)。note は B 列のみ。
    if (type === 'title' || type === 'h1') {
      const bg = type === 'title' ? '#1f8e3d' : '#d9ead3';
      sheet.getRange(pos, 1, span, 3).setBackground(bg);
    } else if (type === 'note') {
      contentRange.setBackground('#fff2cc');
    }

    if (span > 1) contentRange.merge();
    contentRange.setWrap(true).setVerticalAlignment('middle');

    if (type === 'title') {
      contentRange.setFontSize(22).setFontWeight('bold').setFontColor('#ffffff').setHorizontalAlignment('center');
    } else if (type === 'h1') {
      contentRange.setFontSize(15).setFontWeight('bold').setFontColor('#0b5b1f');
    } else if (type === 'h2') {
      contentRange.setFontSize(12).setFontWeight('bold').setFontColor('#1f8e3d');
    } else if (type === 'url') {
      contentRange.setWrap(false);
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(text)
        .setLinkUrl(text)
        .setTextStyle(SpreadsheetApp.newTextStyle()
          .setForegroundColor('#1155cc')
          .setUnderline(true)
          .setFontFamily('Courier New')
          .build())
        .build();
      contentRange.setRichTextValue(richText);
    } else if (type === 'note') {
      contentRange.setFontColor('#a06b00').setFontWeight('bold').setHorizontalAlignment('center');
    } else {
      contentRange.setFontSize(11);
    }

    pos += span;
  });

  sheet.setRowHeights(1, valuesToSet.length, 32);

  // 3カラム構成：A 列(余白) / B 列(内容) / C 列(余白)
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 1100);
  sheet.setColumnWidth(3, 30);

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(3);

  Logger.log('使い方シートに内容を書き込みました');
}

/**
 * 既存「設定」シートに不足している設定キーをすべて追加する。
 * (新機能追加時に既存環境を同期するための統合マイグレーション関数)
 */
function addMissingSettings() {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error('設定シートが見つかりません');

  const allDefaults = [
    ['form_url', '', '【任意】申請者に渡すフォームの公開URL（デプロイ後の /exec 形式。空欄なら自動取得するが /dev 版になる場合あり）'],
    ['company_name_es-ES', '', 'Nombre del destinatario (Español)'],
    ['qualified_invoice_number', '', '適格請求書発行事業者登録番号 (例: T1234567890123、未登録なら空欄)'],
    ['show_corporate_number', 'no', 'PDF に法人番号を表示する (yes/no、インボイス未登録時は no 推奨)'],
    ['email_subject_approved', '請求書 {{INVOICE_NO}} が承認されました', '【承認時】申請者宛メール件名'],
    ['email_body_approved', '{{APPLICANT_NAME}}様\n\nご請求ありがとうございました。\n\n請求が承認されましたので、お支払いまで今しばらくお待ちください。', '【承認時】申請者宛メール本文（{{APPLICANT_NAME}} {{INVOICE_NO}} 使用可）'],
    ['email_subject_rejected', '請求書 {{INVOICE_NO}} は差戻となりました', '【差戻時】申請者宛メール件名'],
    ['email_body_rejected', '{{APPLICANT_NAME}}様\n\nご請求ありがとうございました。\n\n請求内容に不備がありましたので差戻いたします。\n\n差戻事由を担当者にご確認ください。\n\nご確認いただきましたら請求の再申請をお願いいたします。', '【差戻時】申請者宛メール本文（{{APPLICANT_NAME}} {{INVOICE_NO}} 使用可）']
  ];

  const existing = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().map(r => r[0]);
  const additions = allDefaults.filter(d => existing.indexOf(d[0]) < 0);

  if (additions.length === 0) {
    Logger.log('追加すべき設定はありません(全て既存)');
    return;
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, 3).setValues(additions);
  Logger.log('設定を ' + additions.length + ' 件追加しました: ' + additions.map(a => a[0]).join(', '));
}

/**
 * 既存の TPL_* と「設定」シート内の旧 placeholder {{TEACHER_*}} を
 * 新 placeholder {{APPLICANT_*}} に一括置換する。
 * 一回だけ実行すれば OK。実行後は Code.gs 内の後方互換マッピングを削除しても良い。
 */
function migratePlaceholders() {
  const ss = getMainSpreadsheet_();
  // 部分一致で置換（cell 内のどこにあっても置換）
  const substringMapping = {
    '{{TEACHER_NAME}}': '{{APPLICANT_NAME}}',
    '{{TEACHER_ADDRESS}}': '{{APPLICANT_ADDRESS}}',
    '{{TEACHER_PHONE}}': '{{APPLICANT_PHONE}}',
    '先生宛': '申請者宛',
    '先生に': '申請者に',
    '先生へ': '申請者へ',
    '先生（': '申請者（',
    '先生(': '申請者('
  };
  // 完全一致のみ置換（cell の内容と完全一致した場合のみ。BCP 47 化のため再実行しても二重適用されない）
  const exactMapping = {
    'company_name_ja': 'company_name_ja-JP',
    'company_name_en': 'company_name_en-US',
    'company_name_zh': 'company_name_zh-TW'
  };

  // 対象シート: TPL_海外 / TPL_源泉あり / TPL_源泉なし / 設定
  const targets = [
    SHEET_NAMES.TPL_OVERSEAS,
    SHEET_NAMES.TPL_WITH_WHT,
    SHEET_NAMES.TPL_WITHOUT_WHT,
    SHEET_NAMES.SETTINGS
  ];

  let totalReplaced = 0;
  targets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const range = sheet.getDataRange();
    const values = range.getValues();
    let sheetReplaced = 0;
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cell = values[r][c];
        if (typeof cell !== 'string') continue;
        let replaced = cell;
        // 完全一致マッピング(cell 全体が古いキーと一致する場合のみ)
        if (exactMapping.hasOwnProperty(replaced)) {
          replaced = exactMapping[replaced];
          sheetReplaced++;
        }
        // 部分一致マッピング(cell 内のどこにあっても置換)
        Object.keys(substringMapping).forEach(oldKey => {
          if (replaced.indexOf(oldKey) >= 0) {
            replaced = replaced.split(oldKey).join(substringMapping[oldKey]);
            sheetReplaced++;
          }
        });
        if (replaced !== cell) {
          values[r][c] = replaced;
        }
      }
    }
    if (sheetReplaced > 0) {
      range.setValues(values);
      Logger.log(name + ': ' + sheetReplaced + ' 箇所を更新');
      totalReplaced += sheetReplaced;
    }
  });

  Logger.log('合計 ' + totalReplaced + ' 箇所の placeholder を更新しました');
}

/**
 * 既存テンプレート(TPL_海外・TPL_源泉あり・TPL_源泉なし)のフッター 3 行
 * (発行元・住所・法人番号)を placeholder に置き換える。
 * 「発行元」セルを基準に、その下 2 行を含めて変換する。
 */
function convertTemplatesToFooterPlaceholders() {
  const ss = getMainSpreadsheet_();
  const templateNames = [SHEET_NAMES.TPL_OVERSEAS, SHEET_NAMES.TPL_WITH_WHT, SHEET_NAMES.TPL_WITHOUT_WHT];

  templateNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let anchorRow = -1, anchorCol = -1;

    outer: for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const cell = String(data[r][c] || '');
        if (cell.indexOf('発行元') >= 0) {
          anchorRow = r;
          anchorCol = c;
          break outer;
        }
      }
    }

    if (anchorRow < 0) {
      Logger.log(name + ': 発行元 セルが見つかりません - スキップ');
      return;
    }

    sheet.getRange(anchorRow + 1, anchorCol + 1).setValue('{{ISSUER_LINE_1}}').setFontSize(9).setFontColor('#666666');
    sheet.getRange(anchorRow + 2, anchorCol + 1).setValue('{{ISSUER_LINE_2}}').setFontSize(9).setFontColor('#666666');
    sheet.getRange(anchorRow + 3, anchorCol + 1).setValue('{{ISSUER_LINE_3}}').setFontSize(9).setFontColor('#666666');
    Logger.log(name + ': フッター 3 行を placeholder に変換しました');
  });
}

/**
 * 全テストデータを削除する(申請データ・番号管理・Driveフォルダ内のPDF)
 * 実行前に確認ダイアログを表示する。元に戻せないので注意。
 */
function clearAllTestData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ 全テストデータ削除確認',
    '以下を全て削除します：\n\n' +
    '① 申請データシートの全行(ヘッダー以外)\n' +
    '② 番号管理シートのカウンター(ヘッダー以外)\n' +
    '③ 請求書PDF Driveフォルダ内の全PDF\n\n' +
    '元に戻せません。実行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }

  const ss = getMainSpreadsheet_();

  // 1. 申請データを削除
  const subSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  let clearedRows = 0;
  if (subSheet && subSheet.getLastRow() > 1) {
    clearedRows = subSheet.getLastRow() - 1;
    subSheet.getRange(2, 1, clearedRows, subSheet.getLastColumn()).clearContent();
  }

  // 2. 番号管理を削除
  const counterSheet = ss.getSheetByName(SHEET_NAMES.COUNTER);
  if (counterSheet && counterSheet.getLastRow() > 1) {
    counterSheet.getRange(2, 1, counterSheet.getLastRow() - 1, counterSheet.getLastColumn()).clearContent();
  }

  // 3. Driveフォルダ内のPDFを削除
  const settings = getSettings_();
  let deletedFiles = 0;
  if (settings.pdf_folder_id) {
    try {
      const folder = DriveApp.getFolderById(settings.pdf_folder_id);
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        file.setTrashed(true);
        deletedFiles++;
      }
    } catch (e) {
      Logger.log('フォルダアクセス失敗: ' + e.message);
    }
  }

  ui.alert(
    '✅ 削除完了',
    '申請データ：' + clearedRows + '行クリア\n' +
    '番号管理：クリア完了\n' +
    'PDFファイル：' + deletedFiles + '件をゴミ箱へ移動',
    ui.ButtonSet.OK
  );

  Logger.log('テストデータ削除完了 - 申請: ' + clearedRows + '行, PDF: ' + deletedFiles + '件');
}

/**
 * 既存の申請データシートのカラム順を SUBMISSION_COLUMNS の順に揃える
 * ヘッダー名でマッチングして既存データを保持したまま並べ替える
 */
function migrateColumnOrder() {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  if (!sheet) throw new Error('申請データシートが見つかりません');

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const oldHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const newHeaders = SUBMISSION_COLUMNS;

  const colMap = newHeaders.map(name => oldHeaders.indexOf(name));

  sheet.getRange(1, 1, sheet.getMaxRows(), Math.max(lastCol, newHeaders.length))
    .clearDataValidations();

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const newData = data.map(row => colMap.map(oldIdx => oldIdx >= 0 ? row[oldIdx] : ''));
    sheet.getRange(2, 1, newData.length, newHeaders.length).setValues(newData);
  }

  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, newHeaders.length)
    .setBackground('#1f8e3d').setFontColor('#ffffff').setFontWeight('bold');

  if (lastCol > newHeaders.length) {
    sheet.deleteColumns(newHeaders.length + 1, lastCol - newHeaders.length);
  }

  setupSubmissionsSheet_(ss);
  Logger.log('カラム順の移行が完了しました。');
}

function getSettings_() {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const settings = {};
  values.forEach(row => { if (row[0]) settings[row[0]] = row[1]; });
  return settings;
}
