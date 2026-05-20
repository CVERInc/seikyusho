/**
 * 請求書自動化システム — メイン処理
 *
 * 機能：
 *  - Web App として申請者用フォームを提供（doGet）
 *  - フォーム送信を受け取り Sheet に保存（submitInvoiceRequest）
 *  - 管理者が承認 → PDF 生成 → メール送付（generateInvoiceForRow）
 *  - 一括処理用メニュー（onOpen）
 */

const WITHHOLDING_THRESHOLD = 1000000;

/* ──────────────────── Web App エントリーポイント ──────────────────── */

function doGet(e) {
  const requestedLang = (e && e.parameter && e.parameter.lang) || '';
  const settings = getSettings_();
  const companyNames = {
    'ja-JP': settings['company_name_ja-JP'] || '',
    'en-US': settings['company_name_en-US'] || '',
    'zh-TW': settings['company_name_zh-TW'] || ''
  };
  const supportedLangs = Object.keys(companyNames);
  // BCP 47 locale を解決する: exact match → prefix match → デフォルト ja-JP
  let lang = 'ja-JP';
  if (supportedLangs.indexOf(requestedLang) >= 0) {
    lang = requestedLang;
  } else if (requestedLang) {
    const prefix = requestedLang.split('-')[0];
    const matched = supportedLangs.find(l => l === prefix || l.indexOf(prefix + '-') === 0);
    if (matched) lang = matched;
  }
  const t = HtmlService.createTemplateFromFile('index');
  t.lang = lang;
  t.companyNames = companyNames;
  const titleName = companyNames['ja-JP'] || companyNames['en-US'] || 'Invoice';
  return t.evaluate()
    .setTitle(titleName + ' 請求書発行')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ──────────────────── フォーム送信ハンドラ ──────────────────── */

/**
 * Web App フォームから呼ばれる
 * @param {Object} payload - フォームデータ
 * @returns {Object} { ok, message }
 */
function submitInvoiceRequest(payload) {
  try {
    const validated = validatePayload_(payload);
    const ss = getMainSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);

    const defaultWithholding = validated.residence === 'overseas' ? 'N/A' : 'no';

    const rowData = {
      '審査ステータス': 'pending',
      '源泉適用': defaultWithholding,
      'タイムスタンプ': new Date(),
      '言語': validated.language,
      '居住地': validated.residence,
      '氏名': validated.name,
      '住所': validated.address,
      '電話番号': validated.phone,
      'メールアドレス': validated.email,
      '振込銀行名': validated.bankName,
      '支店名': validated.branchName,
      '口座名義': validated.accountName,
      '口座番号': validated.accountNumber,
      '通貨': validated.currency,
      '請求明細(JSON)': JSON.stringify(validated.items),
      '備考': validated.notes,
      '請求書番号': '',
      'PDF生成日': '',
      'PDF URL': '',
      'メモ': '自動生成'
    };

    const row = SUBMISSION_COLUMNS.map(col => rowData[col] !== undefined ? rowData[col] : '');

    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();

    try {
      generateInvoiceForRow(rowNumber);
    } catch (pdfErr) {
      Logger.log('Auto PDF generation failed: ' + pdfErr.stack);
      notifyAdminOfPdfFailure_(validated, rowNumber, pdfErr);
    }

    return { ok: true, message: 'submitted' };
  } catch (err) {
    Logger.log('submitInvoiceRequest error: ' + err.stack);
    return { ok: false, message: String(err.message || err) };
  }
}

function validatePayload_(p) {
  if (!p) throw new Error('payload missing');
  const required = ['name', 'address', 'phone', 'email', 'bankName', 'branchName', 'accountName', 'accountNumber', 'residence', 'currency'];
  required.forEach(k => { if (!p[k]) throw new Error('Missing field: ' + k); });

  if (['japan', 'overseas'].indexOf(p.residence) < 0) throw new Error('Invalid residence');
  if (!Array.isArray(p.items) || p.items.length === 0) throw new Error('No items');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) throw new Error('Invalid email format');

  const items = p.items
    .filter(it => it && it.name && Number(it.quantity) > 0 && Number(it.unitPrice) >= 0)
    .map(it => ({
      name: String(it.name).slice(0, 200),
      quantity: Number(it.quantity),
      unitPrice: Number(it.unitPrice)
    }));
  if (items.length === 0) throw new Error('No valid items');

  return {
    language: ['ja-JP', 'en-US', 'zh-TW'].indexOf(p.language) >= 0 ? p.language : 'ja-JP',
    residence: p.residence,
    name: String(p.name).slice(0, 100),
    address: String(p.address).slice(0, 300),
    phone: String(p.phone).slice(0, 50),
    email: String(p.email).slice(0, 200),
    bankName: String(p.bankName).slice(0, 100),
    branchName: String(p.branchName).slice(0, 100),
    accountName: String(p.accountName).slice(0, 100),
    accountNumber: String(p.accountNumber).slice(0, 50),
    currency: ['JPY', 'TWD', 'USD'].indexOf(p.currency) >= 0 ? p.currency : 'JPY',
    items: items,
    notes: String(p.notes || '').slice(0, 500)
  };
}

function notifyAdminOfPdfFailure_(payload, rowNumber, error) {
  const settings = getSettings_();
  const to = settings.notification_email;
  if (!to) return;

  const companyName = settings['company_name_ja-JP'] || 'Invoice System';
  const subject = '[' + companyName + '] ⚠️ 請求書PDF自動生成失敗：' + payload.name;
  const body = [
    'PDFの自動生成に失敗しました。手動対応をお願いします。',
    '',
    '氏名：' + payload.name,
    '居住地：' + (payload.residence === 'japan' ? '日本国内' : '日本国外'),
    '通貨：' + payload.currency,
    '',
    'エラー内容：',
    String(error.message || error),
    '',
    'スプレッドシートで該当行（行 ' + rowNumber + '）を確認し、',
    'メニュー「請求書 → ➡️ 選んだ行のPDFを生成」を手動実行してください。',
    '',
    'スプレッドシート：' + getMainSpreadsheet_().getUrl()
  ].join('\n');

  MailApp.sendEmail({ to: to, subject: subject, body: body });
}

/* ──────────────────── PDF 生成 ──────────────────── */

/**
 * メニューから選択された行（最終行）に対してPDF生成
 */
function generateInvoiceForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.SUBMISSIONS) {
    SpreadsheetApp.getUi().alert('申請データシートで実行してください。');
    return;
  }
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('データ行を選択してください。');
    return;
  }
  try {
    const result = generateInvoiceForRow(row);
    SpreadsheetApp.getUi().alert('PDF生成完了：' + result.invoiceNo + '\n' + result.pdfUrl);
  } catch (err) {
    SpreadsheetApp.getUi().alert('エラー：' + err.message);
  }
}

/**
 * すべての pending → approved 済み行を一括処理
 */
function generateAllApprovedInvoices() {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, SUBMISSION_COLUMNS.length).getValues();

  let processed = 0;
  let errors = [];
  data.forEach((row, idx) => {
    const reviewStatus = row[SUBMISSION_COLUMNS.indexOf('審査ステータス')];
    const invoiceNo = row[SUBMISSION_COLUMNS.indexOf('請求書番号')];
    if (reviewStatus === 'approved' && !invoiceNo) {
      try {
        generateInvoiceForRow(idx + 2);
        processed++;
      } catch (err) {
        errors.push('行 ' + (idx + 2) + ': ' + err.message);
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    processed + ' 件の請求書を生成しました。' + (errors.length ? '\nエラー：\n' + errors.join('\n') : '')
  );
}

/**
 * 指定行の請求書PDFを生成
 * @param {number} rowNumber - 申請データシートの行番号
 * @param {string} [overrideInvoiceNo] - 既存番号を再利用する場合に指定(再生成用)
 */
function generateInvoiceForRow(rowNumber, overrideInvoiceNo) {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  const row = sheet.getRange(rowNumber, 1, 1, SUBMISSION_COLUMNS.length).getValues()[0];

  const getCol = name => row[SUBMISSION_COLUMNS.indexOf(name)];

  const residence = getCol('居住地');
  const reviewStatus = getCol('審査ステータス');
  const existingInvoiceNo = getCol('請求書番号');
  const withholding = getCol('源泉適用');

  if (existingInvoiceNo && !overrideInvoiceNo) throw new Error('既にPDF生成済みです: ' + existingInvoiceNo);
  if (reviewStatus === 'rejected') throw new Error('差戻済みの行はPDF生成できません');
  if (residence === 'japan' && withholding !== 'yes' && withholding !== 'no') {
    throw new Error('源泉適用を yes/no で指定してください');
  }

  const items = JSON.parse(getCol('請求明細(JSON)'));
  const currency = getCol('通貨');
  const subtotal = items.reduce((s, it) => s + it.quantity * it.unitPrice, 0);

  let consumptionTax = 0;
  let withholdingTax = 0;
  let templateName = SHEET_NAMES.TPL_OVERSEAS;

  if (residence === 'japan') {
    consumptionTax = Math.floor(subtotal * 0.10);
    if (withholding === 'yes') {
      withholdingTax = calculateWithholdingTax_(subtotal);
      templateName = SHEET_NAMES.TPL_WITH_WHT;
    } else {
      templateName = SHEET_NAMES.TPL_WITHOUT_WHT;
    }
  }

  const grandTotal = subtotal + consumptionTax - withholdingTax;
  const invoiceNo = overrideInvoiceNo || generateInvoiceNumber_();
  const settings = getSettings_();

  const issuerLine1 = settings['company_name_ja-JP'] ? ('発行元：' + settings['company_name_ja-JP']) : '';
  const issuerLine2 = settings.company_address ? ('〒 ' + settings.company_address) : '';
  let issuerLine3 = '';
  if (settings.qualified_invoice_number) {
    issuerLine3 = '登録番号：' + settings.qualified_invoice_number;
  } else if (String(settings.show_corporate_number || '').toLowerCase() === 'yes' && settings.corporate_number) {
    issuerLine3 = '法人番号：' + settings.corporate_number;
  }

  const pdfBlob = renderInvoicePDF_({
    templateName: templateName,
    invoiceNo: invoiceNo,
    issueDate: formatDate_(new Date()),
    clientName: settings['company_name_ja-JP'] || 'Company',
    applicantName: getCol('氏名'),
    applicantAddress: getCol('住所'),
    applicantPhone: getCol('電話番号'),
    bankName: getCol('振込銀行名'),
    branchName: getCol('支店名'),
    accountName: getCol('口座名義'),
    accountNumber: getCol('口座番号'),
    items: items,
    currency: currency,
    subtotal: subtotal,
    consumptionTax: consumptionTax,
    withholdingTax: withholdingTax,
    grandTotal: grandTotal,
    notes: getCol('備考'),
    issuerLine1: issuerLine1,
    issuerLine2: issuerLine2,
    issuerLine3: issuerLine3
  });

  const file = saveToDrive_(pdfBlob, invoiceNo, getCol('氏名'));
  sendInvoiceEmail_(file, invoiceNo, getCol('氏名'), grandTotal, currency, settings);

  sheet.getRange(rowNumber, SUBMISSION_COLUMNS.indexOf('請求書番号') + 1).setValue(invoiceNo);
  sheet.getRange(rowNumber, SUBMISSION_COLUMNS.indexOf('PDF生成日') + 1).setValue(new Date());
  sheet.getRange(rowNumber, SUBMISSION_COLUMNS.indexOf('PDF URL') + 1).setValue(file.getUrl());

  return { invoiceNo: invoiceNo, pdfUrl: file.getUrl() };
}

/* ──────────────────── 請求書番号生成 ──────────────────── */

function generateInvoiceNumber_() {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.COUNTER);
  const now = new Date();
  const ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');

  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === ym) { rowIdx = i; break; }
  }

  let seq;
  if (rowIdx < 0) {
    sheet.appendRow([ym, 1]);
    seq = 1;
  } else {
    seq = Number(data[rowIdx][1]) + 1;
    sheet.getRange(rowIdx + 1, 2).setValue(seq);
  }

  return ym + '-' + String(seq).padStart(3, '0');
}

/* ──────────────────── 源泉所得税の計算 ──────────────────── */

function calculateWithholdingTax_(amount) {
  if (amount <= WITHHOLDING_THRESHOLD) {
    return Math.floor(amount * 0.1021);
  }
  return Math.floor(WITHHOLDING_THRESHOLD * 0.1021 + (amount - WITHHOLDING_THRESHOLD) * 0.2042);
}

/* ──────────────────── PDFレンダリング ──────────────────── */

function renderInvoicePDF_(data) {
  const ss = getMainSpreadsheet_();
  const template = ss.getSheetByName(data.templateName);
  if (!template) throw new Error('テンプレートが見つかりません: ' + data.templateName);

  const tempSheet = template.copyTo(ss);
  tempSheet.setName('_temp_' + Date.now());

  try {
    fillTemplateValues_(tempSheet, data);
    SpreadsheetApp.flush();

    const pdfBlob = exportSheetAsPDF_(ss.getId(), tempSheet.getSheetId(), data.invoiceNo);
    return pdfBlob;
  } finally {
    ss.deleteSheet(tempSheet);
  }
}

function fillTemplateValues_(sheet, data) {
  const replacements = {
    '{{ISSUE_DATE}}': data.issueDate,
    '{{INVOICE_NO}}': data.invoiceNo,
    '{{CLIENT_NAME}}': data.clientName,
    '{{APPLICANT_NAME}}': data.applicantName,
    '{{APPLICANT_ADDRESS}}': data.applicantAddress,
    '{{APPLICANT_PHONE}}': data.applicantPhone,
    '{{BANK_NAME}}': data.bankName,
    '{{BRANCH_NAME}}': data.branchName,
    '{{ACCOUNT_NAME}}': data.accountName,
    '{{ACCOUNT_NUMBER}}': data.accountNumber,
    '{{TOTAL_AMOUNT}}': formatCurrency_(data.grandTotal, data.currency),
    '{{SUBTOTAL}}': formatCurrency_(data.subtotal, data.currency),
    '{{CONSUMPTION_TAX}}': formatCurrency_(data.consumptionTax, data.currency),
    '{{WITHHOLDING_TAX}}': '-' + formatCurrency_(data.withholdingTax, data.currency),
    '{{GRAND_TOTAL}}': formatCurrency_(data.grandTotal, data.currency),
    '{{NOTES}}': data.notes || (data.residence === 'overseas' ? '海外居住者の為非課税' : ''),
    '{{ISSUER_LINE_1}}': data.issuerLine1 || '',
    '{{ISSUER_LINE_2}}': data.issuerLine2 || '',
    '{{ISSUER_LINE_3}}': data.issuerLine3 || ''
  };

  const range = sheet.getDataRange();
  const values = range.getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = values[r][c];
      if (typeof cell === 'string' && cell.indexOf('{{') >= 0) {
        let replaced = cell;
        Object.keys(replacements).forEach(key => {
          replaced = replaced.split(key).join(replacements[key]);
        });
        values[r][c] = replaced;
      }
    }
  }
  range.setValues(values);

  const itemTable = findItemTablePosition_(sheet);
  if (!itemTable) {
    throw new Error('テンプレートに「品名/数量/単価/金額」ヘッダーが見つかりません');
  }

  const usedRows = Math.min(data.items.length, itemTable.totalItemRows);

  for (let i = 0; i < usedRows; i++) {
    const it = data.items[i];
    const row = itemTable.itemStartRow + i;
    sheet.getRange(row, itemTable.nameCol).setValue(it.name);
    sheet.getRange(row, itemTable.qtyCol).setValue(it.quantity);
    sheet.getRange(row, itemTable.priceCol).setValue(formatCurrency_(it.unitPrice, data.currency));
    sheet.getRange(row, itemTable.amountCol).setValue(formatCurrency_(it.quantity * it.unitPrice, data.currency));
  }

  if (usedRows < itemTable.totalItemRows) {
    sheet.hideRows(itemTable.itemStartRow + usedRows, itemTable.totalItemRows - usedRows);
  }
}

/**
 * テンプレート内の明細表ヘッダー（品名/数量/単価/金額）の位置を自動検出する
 * 小計の位置から明細表の最大行数も決定する
 */
function findItemTablePosition_(sheet) {
  const data = sheet.getDataRange().getValues();

  let headerRow = -1;
  let nameCol = -1, qtyCol = -1, priceCol = -1, amountCol = -1;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    let foundName = -1, foundQty = -1, foundPrice = -1, foundAmount = -1;
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim();
      if (cell === '品名') foundName = c + 1;
      else if (cell === '数量') foundQty = c + 1;
      else if (cell === '単価') foundPrice = c + 1;
      else if (cell === '金額') foundAmount = c + 1;
    }
    if (foundName > 0 && foundQty > 0 && foundPrice > 0 && foundAmount > 0) {
      headerRow = r + 1;
      nameCol = foundName;
      qtyCol = foundQty;
      priceCol = foundPrice;
      amountCol = foundAmount;
      break;
    }
  }

  if (headerRow < 0) return null;

  let subtotalRow = -1;
  for (let r = headerRow; r < data.length; r++) {
    const row = data[r];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim();
      if (cell === '小計') {
        subtotalRow = r + 1;
        break;
      }
    }
    if (subtotalRow > 0) break;
  }

  const totalItemRows = subtotalRow > 0 ? (subtotalRow - headerRow - 1) : 27;

  return {
    headerRow,
    itemStartRow: headerRow + 1,
    totalItemRows,
    nameCol, qtyCol, priceCol, amountCol
  };
}

function exportSheetAsPDF_(spreadsheetId, sheetId, fileName) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' + [
    'format=pdf',
    'size=A4',
    'portrait=true',
    'fitw=true',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',
    'gid=' + sheetId,
    'top_margin=0.50',
    'bottom_margin=0.50',
    'left_margin=0.50',
    'right_margin=0.50'
  ].join('&');

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: ' + response.getContentText().slice(0, 200));
  }

  return response.getBlob().setName(fileName + '.pdf');
}

function saveToDrive_(blob, invoiceNo, applicantName) {
  const fileName = invoiceNo + '_' + applicantName + '_請求書.pdf';
  blob.setName(fileName);
  return ensurePdfFolder_().createFile(blob);
}

function ensurePdfFolder_() {
  const settings = getSettings_();
  if (settings.pdf_folder_id) {
    try {
      return DriveApp.getFolderById(settings.pdf_folder_id);
    } catch (e) {
      Logger.log('既存の pdf_folder_id が無効です。再作成します。');
    }
  }

  const folder = DriveApp.createFolder('請求書PDF');
  saveSettingValue_('pdf_folder_id', folder.getId());
  Logger.log('PDF保存フォルダを作成しました: ' + folder.getUrl());
  return folder;
}

function saveSettingValue_(key, value) {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, '']);
}

function sendInvoiceEmail_(file, invoiceNo, applicantName, grandTotal, currency, settings) {
  const to = settings.notification_email;
  if (!to) {
    Logger.log('notification_email 未設定のためメール送信をスキップ');
    return;
  }

  const companyName = settings['company_name_ja-JP'] || 'Invoice System';
  const subject = '[' + companyName + '] 請求書 ' + invoiceNo + '（' + applicantName + '様）';
  const body = [
    '申請者からの請求書申請を受け付け、PDFを自動生成しました。',
    '',
    '請求書番号：' + invoiceNo,
    '申請者：' + applicantName,
    '請求額：' + formatCurrency_(grandTotal, currency),
    '',
    '※ 日本国内居住者はデフォルトで「源泉所得税なし」として処理されています。',
    '   源泉適用を変更する場合：',
    '   1. スプレッドシートで「源泉適用」を yes に変更',
    '   2. 該当行を選択し、メニュー「請求書 → 🔄 選んだ行のPDFを再生成」を実行',
    '',
    'PDFを添付しております。ご確認のうえ、承認または差戻をお願いいたします。',
    '',
    '━━━━━━━━━━━━━━━━━━━━',
    '請求書フォルダ：' + (settings.pdf_folder_id ? 'https://drive.google.com/drive/folders/' + settings.pdf_folder_id : '(未設定)')
  ].join('\n');

  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: body,
    attachments: [file.getBlob()]
  });
}

/* ──────────────────── 審査ステータス変更時の通知 ──────────────────── */

/**
 * Installable onEdit トリガーから呼ばれる。
 * 申請データシートの A 列(審査ステータス)が approved または rejected に
 * 変更されたときに、申請者へメール通知を送る。
 */
function handleStatusEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.SUBMISSIONS) return;

  const editedCol = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 2) return;

  const statusCol = SUBMISSION_COLUMNS.indexOf('審査ステータス') + 1;
  const withholdingCol = SUBMISSION_COLUMNS.indexOf('源泉適用') + 1;

  // A 列(審査ステータス) 変更 → 申請者へ承認/差戻通知
  if (editedCol === statusCol) {
    const newValue = String(e.value || '').trim();
    if (newValue !== 'approved' && newValue !== 'rejected') return;
    try {
      sendApplicantStatusNotification_(row, newValue);
    } catch (err) {
      Logger.log('通知送信エラー(行 ' + row + '): ' + err.stack);
    }
    return;
  }

  // B 列(源泉適用) 変更 → 既に請求書番号が発番済みなら自動再生成
  // (担当者が編集権限のないPDFを操作するとPERMISSION_DENIEDになるため、
  //  オーナー権限で動くトリガー内で処理する)
  if (editedCol === withholdingCol) {
    const invoiceNoCol = SUBMISSION_COLUMNS.indexOf('請求書番号') + 1;
    const existingInvoiceNo = sheet.getRange(row, invoiceNoCol).getValue();
    if (!existingInvoiceNo) return;  // PDF 未生成 → 何もしない
    try {
      regenerateInvoice_(row);
      Logger.log('行 ' + row + ': 源泉適用変更により PDF を自動再生成しました');
    } catch (err) {
      Logger.log('自動再生成エラー(行 ' + row + '): ' + err.stack);
    }
    return;
  }
}

function sendApplicantStatusNotification_(rowNumber, status) {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  const row = sheet.getRange(rowNumber, 1, 1, SUBMISSION_COLUMNS.length).getValues()[0];
  const getCol = name => row[SUBMISSION_COLUMNS.indexOf(name)];

  const applicantEmail = getCol('メールアドレス');
  const applicantName = getCol('氏名');
  const invoiceNo = getCol('請求書番号') || '(番号未発行)';
  const pdfUrl = getCol('PDF URL');

  if (!applicantEmail) {
    Logger.log('行 ' + rowNumber + ': メールアドレス未設定のため通知をスキップ');
    return;
  }

  const settings = getSettings_();
  const subjectTpl = status === 'approved' ? settings.email_subject_approved : settings.email_subject_rejected;
  const bodyTpl = status === 'approved' ? settings.email_body_approved : settings.email_body_rejected;

  if (!subjectTpl || !bodyTpl) {
    Logger.log('行 ' + rowNumber + ': メールテンプレート未設定のためスキップ');
    return;
  }

  const fillTemplate = s => String(s || '')
    .split('{{INVOICE_NO}}').join(invoiceNo)
    .split('{{APPLICANT_NAME}}').join(applicantName);

  const subject = fillTemplate(subjectTpl);
  const body = fillTemplate(bodyTpl);

  const options = {};
  if (status === 'approved' && pdfUrl) {
    try {
      const match = String(pdfUrl).match(/[-\w]{25,}/);
      if (match) {
        options.attachments = [DriveApp.getFileById(match[0]).getBlob()];
      }
    } catch (e) {
      Logger.log('PDF添付失敗(無視して続行): ' + e.message);
    }
  }
  if (settings.notification_email) {
    options.cc = settings.notification_email;
  }

  MailApp.sendEmail(Object.assign({
    to: applicantEmail,
    subject: subject,
    body: body
  }, options));

  Logger.log('行 ' + rowNumber + ': 申請者(' + applicantEmail + ')へ ' + status + ' 通知を送信 (CC: ' + (settings.notification_email || 'なし') + ')');
}

/* ──────────────────── ユーティリティ ──────────────────── */

function formatCurrency_(amount, currency) {
  const n = Math.round(Number(amount) || 0);
  const formatted = n.toLocaleString('ja-JP');
  if (currency === 'JPY') return '¥' + formatted;
  if (currency === 'TWD') return 'NT$ ' + formatted;
  if (currency === 'USD') return '$ ' + formatted;
  return formatted;
}

function formatDate_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年M月d日');
}

/* ──────────────────── スプレッドシート起動時のメニュー ──────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('請求書')
    .addItem('➡️ 選んだ行のPDFを生成', 'generateInvoiceForActiveRow')
    .addItem('🔄 選んだ行のPDFを再生成', 'regenerateInvoiceForActiveRow')
    .addItem('✅ approved 全件PDF生成', 'generateAllApprovedInvoices')
    .addSeparator()
    .addItem('⚠️ 初期セットアップ実行', 'runSetup')
    .addItem('❌ 全テストデータ削除', 'clearAllTestData')
    .addToUi();
}

/**
 * 指定行のPDFを再生成（番号据え置き・旧PDFゴミ箱移動）。
 * メニュー操作とトリガーの両方から呼ばれる共通ロジック。
 */
function regenerateInvoice_(rowNumber) {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  const invoiceNoCol = SUBMISSION_COLUMNS.indexOf('請求書番号') + 1;
  const pdfDateCol = SUBMISSION_COLUMNS.indexOf('PDF生成日') + 1;
  const pdfUrlCol = SUBMISSION_COLUMNS.indexOf('PDF URL') + 1;

  const existingInvoiceNo = sheet.getRange(rowNumber, invoiceNoCol).getValue();
  const existingPdfUrl = sheet.getRange(rowNumber, pdfUrlCol).getValue();

  if (existingPdfUrl) {
    try {
      const match = String(existingPdfUrl).match(/[-\w]{25,}/);
      if (match) DriveApp.getFileById(match[0]).setTrashed(true);
    } catch (e) {
      Logger.log('既存PDF削除に失敗(無視して続行): ' + e.message);
    }
  }

  sheet.getRange(rowNumber, invoiceNoCol).setValue('');
  sheet.getRange(rowNumber, pdfDateCol).setValue('');
  sheet.getRange(rowNumber, pdfUrlCol).setValue('');

  return generateInvoiceForRow(rowNumber, existingInvoiceNo || undefined);
}

function regenerateInvoiceForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.SUBMISSIONS) {
    SpreadsheetApp.getUi().alert('申請データシートで実行してください。');
    return;
  }
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('データ行を選択してください。');
    return;
  }
  try {
    const result = regenerateInvoice_(row);
    SpreadsheetApp.getUi().alert('PDF再生成完了：' + result.invoiceNo + '\n' + result.pdfUrl);
  } catch (err) {
    SpreadsheetApp.getUi().alert('エラー：' + err.message);
  }
}
