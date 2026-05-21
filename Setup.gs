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
 * Sheet を開いた時に onOpen を自動実行するためのトリガー設定
 * スタンドアロンスクリプトでは simple trigger が起動しないため、installable trigger を使用する
 */
function installSheetMenuTrigger() {
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

  setupSubmissionsSheet_(ss);
  setupSettingsSheet_(ss);
  setupCounterSheet_(ss);
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_OVERSEAS, { withConsumptionTax: false, withWithholdingTax: false });
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_WITH_WHT, { withConsumptionTax: true, withWithholdingTax: true });
  setupTemplateSheet_(ss, SHEET_NAMES.TPL_WITHOUT_WHT, { withConsumptionTax: true, withWithholdingTax: false });

  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  Logger.log('セットアップ完了！\n  Spreadsheet URL: ' + ss.getUrl());
  return ss.getUrl();
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
 * テンプレートシート作成
 * 既存のPDFひな形のレイアウトをGoogle Sheetsで再現
 */
function setupTemplateSheet_(ss, sheetName, options) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  sheet.setHiddenGridlines(true);

  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 30);

  sheet.getRange('B2:E2').merge();
  sheet.getRange('B2').setValue('請求書')
    .setFontSize(28).setFontWeight('bold').setHorizontalAlignment('left');

  sheet.getRange('D3').setValue('作成日：').setHorizontalAlignment('right');
  sheet.getRange('E3').setValue('{{ISSUE_DATE}}').setHorizontalAlignment('left');
  sheet.getRange('D4').setValue('NO：').setHorizontalAlignment('right');
  sheet.getRange('E4').setValue('{{INVOICE_NO}}').setHorizontalAlignment('left');

  sheet.getRange('B6').setValue('{{CLIENT_NAME}}')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('C6').setValue('様').setFontSize(12).setHorizontalAlignment('left');

  sheet.getRange('D6').setValue('住所：').setHorizontalAlignment('right');
  sheet.getRange('E6').setValue('{{APPLICANT_ADDRESS}}').setHorizontalAlignment('left');

  sheet.getRange('B8').setValue('振込銀行名：').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('C8:C8').merge();
  sheet.getRange('B8:C8').setBorder(false, false, true, false, false, false, '#1f8e3d', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('C8').setValue('{{BANK_NAME}}');

  sheet.getRange('B9').setValue('支店名：').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('C9').setValue('{{BRANCH_NAME}}');
  sheet.getRange('B9:C9').setBorder(false, false, true, false, false, false, '#1f8e3d', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('B10').setValue('口座名義：').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('C10').setValue('{{ACCOUNT_NAME}}');
  sheet.getRange('B10:C10').setBorder(false, false, true, false, false, false, '#1f8e3d', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('B11').setValue('口座番号：').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('C11').setValue('{{ACCOUNT_NUMBER}}');
  sheet.getRange('B11:C11').setBorder(false, false, true, false, false, false, '#1f8e3d', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('D8').setValue('氏名：').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('E8').setValue('{{APPLICANT_NAME}}');

  sheet.getRange('D9').setValue('電話番号：').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('E9').setValue('{{APPLICANT_PHONE}}');

  sheet.getRange('B14').setValue('下記の通りご請求申し上げます').setFontSize(11);

  sheet.getRange('B15').setValue('当月請求額').setFontWeight('bold')
    .setBackground('#d9ead3').setHorizontalAlignment('center');
  sheet.getRange('B16').setValue('{{TOTAL_AMOUNT}}')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, false, false, '#1f8e3d', SpreadsheetApp.BorderStyle.SOLID);

  const tableStartRow = 18;
  sheet.getRange(tableStartRow, 2, 1, 4).setValues([['品名', '数量', '単価', '金額']])
    .setBackground('#1f8e3d').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  const totalItemRows = 27;
  for (let i = 0; i < totalItemRows; i++) {
    const row = tableStartRow + 1 + i;
    sheet.getRange(row, 1).setValue(i + 1).setHorizontalAlignment('center').setFontColor('#666666').setFontSize(9);
    sheet.getRange(row, 2, 1, 4).setBackground(i % 2 === 0 ? '#e8f5e9' : '#ffffff');
    sheet.getRange(row, 3).setHorizontalAlignment('right');
    sheet.getRange(row, 4).setHorizontalAlignment('right');
    sheet.getRange(row, 5).setHorizontalAlignment('right');
  }

  const summaryStartRow = tableStartRow + 1 + totalItemRows;
  sheet.getRange(summaryStartRow, 4).setValue('小計').setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center');
  sheet.getRange(summaryStartRow, 5).setValue('{{SUBTOTAL}}').setHorizontalAlignment('right');

  let nextRow = summaryStartRow + 1;
  if (options.withConsumptionTax) {
    sheet.getRange(nextRow, 4).setValue('消費税 10%').setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center');
    sheet.getRange(nextRow, 5).setValue('{{CONSUMPTION_TAX}}').setHorizontalAlignment('right');
    nextRow++;
  } else {
    sheet.getRange(nextRow, 4).setValue('消費税 0%').setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center');
    sheet.getRange(nextRow, 5).setValue('0').setHorizontalAlignment('right');
    nextRow++;
  }

  if (options.withWithholdingTax) {
    sheet.getRange(nextRow, 4).setValue('源泉所得税').setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center');
    sheet.getRange(nextRow, 5).setValue('{{WITHHOLDING_TAX}}').setHorizontalAlignment('right');
    nextRow++;
  }

  sheet.getRange(nextRow, 4).setValue('合計').setFontWeight('bold').setBackground('#1f8e3d').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.getRange(nextRow, 5).setValue('{{GRAND_TOTAL}}').setFontWeight('bold').setHorizontalAlignment('right')
    .setBorder(true, true, true, true, false, false, '#1f8e3d', SpreadsheetApp.BorderStyle.SOLID);

  const noteRow = nextRow + 3;
  sheet.getRange(noteRow, 2).setValue('備考：').setFontWeight('bold');
  sheet.getRange(noteRow + 1, 2, 5, 4).merge();
  sheet.getRange(noteRow + 1, 2).setValue('{{NOTES}}')
    .setVerticalAlignment('top').setHorizontalAlignment('left').setWrap(true)
    .setBorder(true, true, true, true, false, false, '#1f8e3d', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(noteRow + 8, 2).setValue('{{ISSUER_LINE_1}}').setFontSize(9).setFontColor('#666666');
  sheet.getRange(noteRow + 9, 2).setValue('{{ISSUER_LINE_2}}').setFontSize(9).setFontColor('#666666');
  sheet.getRange(noteRow + 10, 2).setValue('{{ISSUER_LINE_3}}').setFontSize(9).setFontColor('#666666');
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
    ['対応言語:日本語 / English / 繁體中文(右上で切替可)', 'body'],
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
