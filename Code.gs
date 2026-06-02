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
    'zh-TW': settings['company_name_zh-TW'] || '',
    'es-ES': settings['company_name_es-ES'] || ''
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

    // 電話番号・口座番号は先頭の「0」が数値化で消えないよう、必ず文字列として保存し直す。
    // （appendRow はセル書式により "0987..." を数値 987... に丸めてしまうため）
    forceTextCells_(sheet, rowNumber, {
      '電話番号': validated.phone,
      '口座番号': validated.accountNumber
    });

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

/**
 * 指定行の指定カラムを「書式: プレーンテキスト(@)」で保存し直す。
 * 先頭ゼロを保持したい電話番号・口座番号などに使う（数値化で 0 が消えるのを防ぐ）。
 * @param {Sheet} sheet
 * @param {number} rowNumber
 * @param {Object} colNameToValue - { カラム名: 文字列値 }
 */
function forceTextCells_(sheet, rowNumber, colNameToValue) {
  Object.keys(colNameToValue).forEach(colName => {
    const idx = SUBMISSION_COLUMNS.indexOf(colName);
    if (idx < 0) return;
    const value = colNameToValue[colName];
    const cell = sheet.getRange(rowNumber, idx + 1);
    cell.setNumberFormat('@');
    cell.setValue(String(value == null ? '' : value));
  });
}

/**
 * 【一回限りのメンテナンス】既存データの電話番号・口座番号を修復する。
 * エディタから関数を選んで手動実行する（末尾アンダースコアなし＝実行ドロップダウンに表示）。
 *
 *  - 電話番号: 数値で保存されている = 先頭ゼロが落ちている状態。先頭に "0" を補って文字列化する。
 *    （日本・台湾の電話番号は必ず 0 始まりのため、数値化で落ちた 1 桁を安全に復元できる）
 *  - 口座番号: 数値なら文字列化のみ。落ちたゼロの桁数は特定できないため自動補完はせず、
 *    要確認リストとしてログに出す。
 *  - 両列をプレーンテキスト書式に設定し、今後の手入力・自動保存でもゼロが消えないようにする。
 *
 * 冪等: 既に文字列のセルは触らない。何度実行しても安全。変更内容は全てログに残す。
 */
function fixExistingContactData() {
  const ss = getMainSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  if (!sheet) throw new Error('申請データシートが見つかりません');

  const phoneCol = SUBMISSION_COLUMNS.indexOf('電話番号') + 1;
  const acctCol = SUBMISSION_COLUMNS.indexOf('口座番号') + 1;

  // 今後の入力でもゼロが消えないよう、列をテキスト書式にする
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, phoneCol, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    sheet.getRange(2, acctCol, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データ行がありません。列の書式のみ設定しました。');
    return 'no-data';
  }

  const phoneFixes = [];
  const acctReview = [];
  for (let row = 2; row <= lastRow; row++) {
    const pCell = sheet.getRange(row, phoneCol);
    const pVal = pCell.getValue();
    if (typeof pVal === 'number') {
      const restored = '0' + String(pVal);
      pCell.setNumberFormat('@').setValue(restored);
      phoneFixes.push('行' + row + ': ' + pVal + ' → ' + restored);
    }
    const aCell = sheet.getRange(row, acctCol);
    const aVal = aCell.getValue();
    if (typeof aVal === 'number') {
      const asText = String(aVal);
      aCell.setNumberFormat('@').setValue(asText);
      acctReview.push('行' + row + ': ' + asText + '（先頭ゼロが必要か要確認）');
    }
  }

  const report =
    '✅ fixExistingContactData 完了\n\n' +
    '電話番号 先頭ゼロ復元: ' + phoneFixes.length + ' 件\n' + (phoneFixes.join('\n') || '（なし）') +
    '\n\n口座番号 文字列化: ' + acctReview.length + ' 件\n' + (acctReview.join('\n') || '（なし）');
  Logger.log(report);
  return report;
}

/**
 * 【診断用・読み取り専用】各テンプレートの住所セル({{APPLICANT_ADDRESS}})の
 * 合併状態・寸法・フォント・行高をログ出力する。エディタから手動実行。
 * 住所が PDF で切れる原因（合併の向きなど）を特定するために使う。
 */
function inspectAddressCell() {
  const ss = getMainSpreadsheet_();
  const names = [SHEET_NAMES.TPL_OVERSEAS, SHEET_NAMES.TPL_WITH_WHT, SHEET_NAMES.TPL_WITHOUT_WHT];
  const out = [];
  names.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { out.push(name + ': シートなし'); return; }
    const vals = sh.getDataRange().getValues();
    let found = null;
    for (let r = 0; r < vals.length && !found; r++) {
      for (let c = 0; c < vals[r].length; c++) {
        if (typeof vals[r][c] === 'string' && vals[r][c].indexOf('{{APPLICANT_ADDRESS}}') >= 0) {
          found = { r: r + 1, c: c + 1 };
          break;
        }
      }
    }
    if (!found) { out.push(name + ': {{APPLICANT_ADDRESS}} 見つからず'); return; }
    const cell = sh.getRange(found.r, found.c);
    let info = name + ': 住所=' + cell.getA1Notation() + ' merged=' + cell.isPartOfMerge();
    if (cell.isPartOfMerge()) {
      const m = cell.getMergedRanges()[0];
      info += ' range=' + m.getA1Notation() + '(' + m.getNumRows() + '行x' + m.getNumColumns() + '列)';
    }
    info += ' colW=' + sh.getColumnWidth(found.c) + ' rowH=' + sh.getRowHeight(found.r) + ' font=' + cell.getFontSize();
    out.push(info);
  });
  const report = out.join('\n');
  Logger.log(report);
  return report;
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
    language: ['ja-JP', 'en-US', 'zh-TW', 'es-ES'].indexOf(p.language) >= 0 ? p.language : 'ja-JP',
    residence: p.residence,
    name: String(p.name).slice(0, 100),
    address: String(p.address).slice(0, 300),
    phone: String(p.phone).slice(0, 50),
    email: String(p.email).slice(0, 200),
    bankName: String(p.bankName).slice(0, 100),
    branchName: String(p.branchName).slice(0, 100),
    accountName: String(p.accountName).slice(0, 100),
    accountNumber: String(p.accountNumber).slice(0, 50),
    currency: ['JPY', 'TWD', 'USD', 'EUR'].indexOf(p.currency) >= 0 ? p.currency : 'JPY',
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
 * @param {{skipEmail?: boolean}} [options] - skipEmail=true で管理者通知メールを送らない(一括再生成用)
 */
function generateInvoiceForRow(rowNumber, overrideInvoiceNo, options) {
  options = options || {};
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

  const settings = getSettings_();
  const taxRates = resolveTaxRates_(settings);

  let consumptionTax = 0;
  let withholdingTax = 0;
  let templateName = SHEET_NAMES.TPL_OVERSEAS;

  if (residence === 'japan') {
    consumptionTax = Math.floor(subtotal * taxRates.consumptionRate);
    if (withholding === 'yes') {
      withholdingTax = calculateWithholdingTax_(subtotal, taxRates);
      templateName = SHEET_NAMES.TPL_WITH_WHT;
    } else {
      templateName = SHEET_NAMES.TPL_WITHOUT_WHT;
    }
  }

  const grandTotal = subtotal + consumptionTax - withholdingTax;
  const invoiceNo = overrideInvoiceNo || generateInvoiceNumber_();

  // 消費税ラベルを実際に適用した税率で表示する（国内は設定税率、海外は非課税 0%）。
  // PDF 上の「消費税 N%」表記が設定・計算と必ず一致するようにする。
  const consumptionRatePct = residence === 'japan' ? taxRates.consumptionRate * 100 : 0;
  const consumptionRateLabel = formatTaxRatePercent_(consumptionRatePct) + '%';

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
    residence: residence,
    subtotal: subtotal,
    consumptionTax: consumptionTax,
    consumptionRateLabel: consumptionRateLabel,
    withholdingTax: withholdingTax,
    grandTotal: grandTotal,
    notes: getCol('備考'),
    issuerLine1: issuerLine1,
    issuerLine2: issuerLine2,
    issuerLine3: issuerLine3
  });

  const file = saveToDrive_(pdfBlob, invoiceNo, getCol('氏名'));
  if (!options.skipEmail) {
    sendInvoiceEmail_(file, invoiceNo, getCol('氏名'), grandTotal, currency, settings);
  }

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

/**
 * 設定シートの税率を数値として解決する。
 * 値が未設定・非数値・0 以下の場合は、日本の法定デフォルトにフォールバックする
 * （セルの消し忘れ等で誤った税額を出さないための安全策）。
 * @param {Object} settings - getSettings_() の返り値
 * @returns {{consumptionRate:number, withholdingRate:number, withholdingThreshold:number, withholdingRateOver:number}}
 */
function resolveTaxRates_(settings) {
  settings = settings || {};
  const num = (v, fallback) => {
    const n = Number(v);
    return (isFinite(n) && n > 0) ? n : fallback;
  };
  return {
    consumptionRate: num(settings.consumption_tax_rate, 0.10),
    withholdingRate: num(settings.withholding_tax_rate, 0.1021),
    withholdingThreshold: num(settings.withholding_threshold, WITHHOLDING_THRESHOLD),
    withholdingRateOver: num(settings.withholding_tax_rate_over, 0.2042)
  };
}

/**
 * 源泉所得税を計算する。
 * @param {number} amount - 課税対象額（小計）
 * @param {Object} [rates] - resolveTaxRates_() の返り値。未指定なら法定デフォルト。
 */
function calculateWithholdingTax_(amount, rates) {
  rates = rates || resolveTaxRates_({});
  if (amount <= rates.withholdingThreshold) {
    return Math.floor(amount * rates.withholdingRate);
  }
  return Math.floor(
    rates.withholdingThreshold * rates.withholdingRate +
    (amount - rates.withholdingThreshold) * rates.withholdingRateOver
  );
}

/**
 * 税率(小数または百分率)を表示用のパーセント数値文字列にする。
 * 浮動小数点の誤差を丸め、末尾の余分な 0 を落とす。
 * 例: 10.000000000000002 → "10"、 8 → "8"、 10.5 → "10.5"
 * @param {number} pct - 百分率の値（例: 10 は 10%）
 */
function formatTaxRatePercent_(pct) {
  return String(Math.round((Number(pct) || 0) * 100) / 100);
}

/**
 * 税額計算の回帰テスト。Apps Script エディタから手動実行する（関数を選んで「実行」）。
 * 設定シートの税率が実際に計算へ反映されること、不正値が安全にフォールバックすることを検証する。
 * 注: 末尾アンダースコアを付けない（付けると実行ドロップダウンに表示されないため）。
 */
function test_taxCalculation() {
  const assert = (cond, msg) => { if (!cond) throw new Error('FAIL: ' + msg); };

  // 1. デフォルト税率（設定が空でも法定値）
  const def = resolveTaxRates_({});
  assert(def.consumptionRate === 0.10, 'default consumption 10%');
  assert(def.withholdingRate === 0.1021, 'default withholding 10.21%');
  assert(def.withholdingThreshold === 1000000, 'default threshold 100万');
  assert(def.withholdingRateOver === 0.2042, 'default over 20.42%');

  // 2. 消費税：設定変更が計算に反映される（これがバグの本丸）
  const r8 = resolveTaxRates_({ consumption_tax_rate: '0.08' });
  assert(Math.floor(1000000 * r8.consumptionRate) === 80000, 'consumption follows setting: 8% of 100万 = 80000');
  assert(Math.floor(1000000 * def.consumptionRate) === 100000, 'consumption default: 10% of 100万 = 100000');

  // 3. 源泉：100万以下は単一税率（500000*0.1021=51050）
  assert(calculateWithholdingTax_(500000, def) === 51050, 'withholding under threshold');

  // 4. 源泉：100万超は二段階（1,000,000*0.1021 + 200,000*0.2042 = 142940）
  assert(calculateWithholdingTax_(1200000, def) === 142940, 'withholding over threshold (two-tier)');

  // 5. 源泉率・基準額の設定変更が反映される（500000*0.05=25000、基準額200万に引上げ）
  const rw = resolveTaxRates_({ withholding_tax_rate: '0.05', withholding_threshold: '2000000' });
  assert(calculateWithholdingTax_(500000, rw) === 25000, 'withholding follows setting');

  // 6. 不正値・空欄はデフォルトへフォールバック（誤課税防止）
  const bad = resolveTaxRates_({ consumption_tax_rate: '', withholding_tax_rate: 'abc', withholding_threshold: '0' });
  assert(bad.consumptionRate === 0.10, 'empty → default');
  assert(bad.withholdingRate === 0.1021, 'non-numeric → default');
  assert(bad.withholdingThreshold === 1000000, 'zero → default');

  // 7. パーセント表示の整形（浮動小数点誤差・末尾ゼロ除去）
  assert(formatTaxRatePercent_(0.10 * 100) === '10', 'format 10% (fp-safe)');
  assert(formatTaxRatePercent_(0.08 * 100) === '8', 'format 8% (fp-safe)');
  assert(formatTaxRatePercent_(0) === '0', 'format 0%');
  assert(formatTaxRatePercent_(10.5) === '10.5', 'format 10.5%');

  // 8. 消費税ラベルが税率に追従する（PDF 表記＝設定＝計算 の一致）
  assert(formatTaxRatePercent_(def.consumptionRate * 100) + '%' === '10%', 'label default 10%');
  assert(formatTaxRatePercent_(r8.consumptionRate * 100) + '%' === '8%', 'label follows setting 8%');

  // 9. 「消費税 N%」固定ラベル検出パターン（fillTemplateValues_ と同一）
  const LABEL_RE = /^消費税\s*[0-9]+(?:\.[0-9]+)?\s*%$/;
  assert(LABEL_RE.test('消費税 10%'), 'matches 消費税 10%');
  assert(LABEL_RE.test('消費税 0%'), 'matches 消費税 0%');
  assert(LABEL_RE.test('消費税 8%'), 'matches 消費税 8%');
  assert(!LABEL_RE.test('消費税の説明は別途 10% と記載'), 'does not match free text containing 10%');
  assert(!LABEL_RE.test('源泉所得税'), 'does not match withholding label');

  Logger.log('✅ test_taxCalculation: 全アサーション通過');
  return 'OK';
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

/**
 * 折り返しテキストに必要な高さ(px)をざっくり見積もる。
 * 利用可能幅(px)とフォントサイズから 1 行に入る文字数を推定し、行数 × 行高で算出する。
 * CJK・全角はフォント幅、ASCII は約半角として概算。切れ防止のため余裕を多めに取る。
 * @param {number} availWidthPx - セル(合併時は合計)の幅
 * @param {number} fontPt - フォントサイズ(pt)
 * @param {string} text - 表示テキスト
 */
function estimateWrappedLines_(availWidthPx, fontPt, text) {
  const fontPx = (fontPt || 10) * 1.33;          // pt → px 概算
  const cjk = /[ᄀ-￿]/;                // CJK・全角はフォント幅、その他は約半角
  const usable = Math.max(availWidthPx - 8, 10);  // セル内パディングぶん控える
  // 明示的な改行(\n)で行を分け、各行ごとに自動折り返し行数を加算する。
  // （\n を無視すると複数行テキスト＝備考などが大幅に過小評価され、末尾が切れる）
  let total = 0;
  String(text).split('\n').forEach(segment => {
    let textPx = 0;
    for (let i = 0; i < segment.length; i++) {
      textPx += cjk.test(segment[i]) ? fontPx : fontPx * 0.6;
    }
    total += Math.max(1, Math.ceil(textPx / usable));
  });
  return Math.max(1, total);
}

function estimateWrappedHeight_(availWidthPx, fontPt, text) {
  const fontPx = (fontPt || 10) * 1.33;
  const lines = estimateWrappedLines_(availWidthPx, fontPt, text);
  const lineHeightPx = Math.ceil(fontPx * 1.6) + 4;
  return lines * lineHeightPx + 10;               // 上下の余白を多めに
}

/**
 * セルの合併ブロック情報を返す。合併していなければ単一セルとして扱う。
 * @returns {{topRow:number, numRows:number, leftCol:number, numCols:number, width:number}}
 */
function mergedBlock_(sheet, row, col) {
  const cell = sheet.getRange(row, col);
  if (cell.isPartOfMerge()) {
    const m = cell.getMergedRanges()[0];
    let width = 0;
    for (let i = 0; i < m.getNumColumns(); i++) width += sheet.getColumnWidth(m.getColumn() + i);
    return { topRow: m.getRow(), numRows: m.getNumRows(), leftCol: m.getColumn(), numCols: m.getNumColumns(), width: width };
  }
  return { topRow: row, numRows: 1, leftCol: col, numCols: 1, width: sheet.getColumnWidth(col) };
}

/**
 * ブロックの総高さが neededPx に満たなければ、不足ぶんを最終行に加える。
 * 縦合併・横合併・非合併のいずれでも安全。高さは縮めない（増やすだけ）。
 */
function ensureBlockHeight_(sheet, block, neededPx) {
  let current = 0;
  for (let i = 0; i < block.numRows; i++) current += sheet.getRowHeight(block.topRow + i);
  if (neededPx > current) {
    const lastRow = block.topRow + block.numRows - 1;
    sheet.setRowHeight(lastRow, sheet.getRowHeight(lastRow) + (neededPx - current));
  }
}

/**
 * テキストが 1 行で availWidthPx に収まる最大フォント(pt)を探す。
 * defaultPt から 0.5pt 刻みで minPt まで縮める。
 * @returns {{pt:number, fits:boolean}} fits=false は minPt でも収まらないことを示す
 */
function fitFontOneLine_(availWidthPx, text, defaultPt, minPt) {
  const cjk = /[ᄀ-￿]/;
  const widthAt = pt => {
    const px = pt * 1.33;
    let w = 0;
    for (let i = 0; i < text.length; i++) w += cjk.test(text[i]) ? px : px * 0.6;
    return w;
  };
  const usable = Math.max(availWidthPx - 8, 10);
  let pt = defaultPt;
  while (pt > minPt && widthAt(pt) > usable) pt -= 0.5;
  if (pt < minPt) pt = minPt;
  return { pt: pt, fits: widthAt(pt) <= usable };
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

  // 住所・氏名など、欄幅を超えると末尾が切れてしまう長文プレースホルダー。
  // 折り返しを有効化し、該当行の高さを内容に合わせて広げる（PDF切れ対策）。
  const WRAP_PLACEHOLDERS = [
    '{{APPLICANT_ADDRESS}}', '{{APPLICANT_NAME}}',
    '{{BANK_NAME}}', '{{BRANCH_NAME}}', '{{ACCOUNT_NAME}}', '{{NOTES}}',
    '{{ISSUER_LINE_1}}', '{{ISSUER_LINE_2}}', '{{ISSUER_LINE_3}}'
  ];

  // 1 行に収めたい数字列（口座番号など）。折り返すと不格好なので、
  // セル幅に収まるようフォントを縮小して 1 行表示する（shrink-to-fit）。
  const SHRINK_PLACEHOLDERS = ['{{ACCOUNT_NUMBER}}'];

  // 数字に見えるが「文字列」として表示すべきプレースホルダー。
  // setValues で General 書式のセルに "0987..." を書くと再び数値化され先頭ゼロが落ちるため、
  // 値を書き込む前にセルをプレーンテキスト書式(@)にしておく。
  const TEXT_PLACEHOLDERS = ['{{APPLICANT_PHONE}}', '{{ACCOUNT_NUMBER}}'];

  // 「消費税 10%」のような固定税率ラベルを、実際に適用した税率へ自動補正するためのパターン。
  // セル全体がこの形のときだけ置換し、備考など他テキストには触れない（誤置換防止）。
  // これにより、旧テンプレートでもマイグレーション不要で表記が設定・計算と必ず一致する。
  const CONSUMPTION_LABEL_RE = /^消費税\s*[0-9]+(?:\.[0-9]+)?\s*%$/;

  const range = sheet.getDataRange();
  const values = range.getValues();
  const wrapCells = [];    // {row, col, text} 折り返し対象セル
  const shrinkCells = [];  // {row, col, text} 1 行に縮小表示するセル
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = values[r][c];
      if (typeof cell !== 'string') continue;
      if (cell.indexOf('{{') >= 0) {
        const needsWrap = WRAP_PLACEHOLDERS.some(p => cell.indexOf(p) >= 0);
        const needsShrink = SHRINK_PLACEHOLDERS.some(p => cell.indexOf(p) >= 0);
        const needsText = TEXT_PLACEHOLDERS.some(p => cell.indexOf(p) >= 0);
        let replaced = cell;
        Object.keys(replacements).forEach(key => {
          replaced = replaced.split(key).join(replacements[key]);
        });
        values[r][c] = replaced;
        // 先頭ゼロを保持: 値を書き込む前にセルをテキスト書式にする（setValues の再数値化対策）
        if (needsText) {
          sheet.getRange(r + 1, c + 1).setNumberFormat('@');
        }
        if (needsWrap) {
          sheet.getRange(r + 1, c + 1).setWrap(true);
          wrapCells.push({ row: r + 1, col: c + 1, text: replaced });
        }
        if (needsShrink) {
          shrinkCells.push({ row: r + 1, col: c + 1, text: replaced });
        }
      } else if (data.consumptionRateLabel && CONSUMPTION_LABEL_RE.test(cell.trim())) {
        values[r][c] = '消費税 ' + data.consumptionRateLabel;
      }
    }
  }
  range.setValues(values);

  // 折り返し後の実レイアウトを確定させてから行高を調整する。
  SpreadsheetApp.flush();

  // 口座番号など 1 行で見せたい数字列:
  // まず右隣の空セルへ広げて（合併）標準フォントを保ち、それでも収まらなければ縮小、
  // 縮小の下限でも無理なら折り返しにフォールバックする。
  shrinkCells.forEach(sc => {
    // 右隣が空セルなら最大 2 列ぶんまで広げ、フォントを縮めずに済むようにする
    let cols = 1;
    while (cols < 3) {
      const nb = values[sc.row - 1] ? values[sc.row - 1][sc.col - 1 + cols] : undefined;
      if (nb === '' && !sheet.getRange(sc.row, sc.col + cols).isPartOfMerge()) cols++;
      else break;
    }
    if (cols > 1) sheet.getRange(sc.row, sc.col, 1, cols).merge();

    const anchor = sheet.getRange(sc.row, sc.col);
    const block = mergedBlock_(sheet, sc.row, sc.col);
    let fontPt = 10;
    try { fontPt = anchor.getFontSize() || 10; } catch (e) {}
    const fit = fitFontOneLine_(block.width, sc.text, fontPt, 7);
    anchor.setFontSize(fit.fits ? fit.pt : 7).setWrap(!fit.fits);
    if (!fit.fits) {
      anchor.setVerticalAlignment('top');
      ensureBlockHeight_(sheet, block, estimateWrappedHeight_(block.width, 7, sc.text));
    }
  });

  // 折り返しセル: 実際に複数行になるものだけ行高を確保し、値とラベル(左隣)を上揃えにする。
  // 単一行で収まるものは既定のまま（=ラベルと同じ下揃え）にして「値だけ浮く」のを防ぐ。
  wrapCells.forEach(wc => {
    const block = mergedBlock_(sheet, wc.row, wc.col);
    let fontPt = 10;
    try { fontPt = sheet.getRange(wc.row, wc.col).getFontSize() || 10; } catch (e) {}
    if (estimateWrappedLines_(block.width, fontPt, wc.text) >= 2) {
      sheet.getRange(wc.row, wc.col).setVerticalAlignment('top');
      if (wc.col > 1) sheet.getRange(wc.row, wc.col - 1).setVerticalAlignment('top'); // ラベルも上揃え
      ensureBlockHeight_(sheet, block, estimateWrappedHeight_(block.width, fontPt, wc.text));
    }
  });

  const itemTable = findItemTablePosition_(sheet);
  if (!itemTable) {
    throw new Error('テンプレートに「品名/数量/単価/金額」ヘッダーが見つかりません');
  }

  const usedRows = Math.min(data.items.length, itemTable.totalItemRows);
  if (data.items.length > itemTable.totalItemRows) {
    Logger.log('⚠️ 明細が ' + data.items.length + ' 件あり、テンプレート上限 ' +
      itemTable.totalItemRows + ' 行を超えています。' +
      (data.items.length - itemTable.totalItemRows) + ' 件が PDF に表示されません。');
  }

  const nameWidth = sheet.getColumnWidth(itemTable.nameCol);
  for (let i = 0; i < usedRows; i++) {
    const it = data.items[i];
    const row = itemTable.itemStartRow + i;
    // 品名は長いと欄幅で切れるため折り返し、必要なら行高を確保する
    sheet.getRange(row, itemTable.nameCol).setValue(it.name).setWrap(true).setVerticalAlignment('top');
    sheet.getRange(row, itemTable.qtyCol).setValue(it.quantity);
    sheet.getRange(row, itemTable.priceCol).setValue(formatCurrency_(it.unitPrice, data.currency));
    sheet.getRange(row, itemTable.amountCol).setValue(formatCurrency_(it.quantity * it.unitPrice, data.currency));
    if (estimateWrappedLines_(nameWidth, 10, String(it.name)) >= 2) {
      ensureBlockHeight_(
        sheet,
        { topRow: row, numRows: 1, leftCol: itemTable.nameCol, numCols: 1, width: nameWidth },
        estimateWrappedHeight_(nameWidth, 10, String(it.name))
      );
    }
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
  if (currency === 'EUR') return '€' + formatted;
  return formatted;
}

function formatDate_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年M月d日');
}

/* ──────────────────── スプレッドシート起動時のメニュー ──────────────────── */

function onOpen() {
  // メニューは最優先で構築する（ここで例外が出ると「請求書」メニュー自体が消えるため）。
  // 初回セットアップは runSetup ではなく setup() を呼ぶ（タブ生成＋ステータス編集トリガー設置まで一括）。
  SpreadsheetApp.getUi()
    .createMenu('請求書')
    .addItem('🚀 初期セットアップ（最初に1回）', 'setup')
    .addSeparator()
    .addItem('➡️ 選んだ行のPDFを生成', 'generateInvoiceForActiveRow')
    .addItem('🔄 選んだ行のPDFを再生成（複数行可）', 'regenerateSelectedInvoices')
    .addItem('✅ approved 全件PDF生成', 'generateAllApprovedInvoices')
    .addSeparator()
    .addItem('❌ 全テストデータ削除', 'clearAllTestData')
    .addToUi();

  // コピー直後など未設定の場合は、初期セットアップへ誘導するトーストを出す（best-effort）。
  try {
    if (!isConfigured_()) {
      SpreadsheetApp.getActive().toast(
        '「請求書」メニュー →「🚀 初期セットアップ」を1回実行してください。',
        'ようこそ！まずはじめに', 10);
    }
  } catch (e) { /* simple trigger の制約等で失敗しても無視 */ }
}

/**
 * 初期セットアップ済みか（SPREADSHEET_ID が設定されているか）を返す。
 * テンプレートをコピーした直後はスクリプトプロパティが空なので false になる。
 */
function isConfigured_() {
  return !!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

/**
 * 指定行のPDFを再生成（番号据え置き・旧PDFゴミ箱移動）。
 * メニュー操作とトリガーの両方から呼ばれる共通ロジック。
 * @param {number} rowNumber - 申請データシートの行番号
 * @param {{skipEmail?: boolean}} [options] - 一括再生成時はメール送信を抑止
 */
function regenerateInvoice_(rowNumber, options) {
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

  return generateInvoiceForRow(rowNumber, existingInvoiceNo || undefined, options);
}

/**
 * 選択中の行（1 行でも複数行でも可）のPDFを再生成する。
 * Cmd/Ctrl での複数選択・範囲ドラッグ・全選択にも対応。
 * 番号発行済みの行のみ再生成（番号は据え置き）。番号未発行の行はスキップする。
 * 実行前に「管理者へ通知メールを送るか」を確認ダイアログで選べる（大量送信の暴発防止）。
 */
function regenerateSelectedInvoices() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.SUBMISSIONS) {
    ui.alert('申請データシートで実行してください。');
    return;
  }

  // 選択中の全行を集約（複数レンジ・複数行選択に対応）
  const rangeList = sheet.getActiveRangeList();
  const rowSet = {};
  if (rangeList) {
    rangeList.getRanges().forEach(rng => {
      const start = rng.getRow();
      for (let i = 0; i < rng.getNumRows(); i++) {
        const row = start + i;
        if (row >= 2) rowSet[row] = true;  // ヘッダー行は除外
      }
    });
  }
  const rows = Object.keys(rowSet).map(Number).sort((a, b) => a - b);
  if (rows.length === 0) {
    ui.alert('再生成するデータ行を選択してください（複数行可）。');
    return;
  }

  // メール送信の有無を選ばせる: はい=送信する / いいえ=送信しない / キャンセル=中止
  const resp = ui.alert(
    'PDF再生成',
    rows.length + ' 行のPDFを再生成します。\n' +
    '（番号は据え置き、未発行の行はスキップ）\n\n' +
    '管理者へ最新PDFを通知メールで送信しますか？\n\n' +
    '［はい］送信する　／　［いいえ］送信しない　／　［キャンセル］中止',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (resp !== ui.Button.YES && resp !== ui.Button.NO) return;  // キャンセル・×で中止
  const skipEmail = (resp === ui.Button.NO);

  const invoiceNoCol = SUBMISSION_COLUMNS.indexOf('請求書番号') + 1;
  let done = 0, skipped = 0;
  const errors = [];
  rows.forEach(row => {
    const invoiceNo = sheet.getRange(row, invoiceNoCol).getValue();
    if (!invoiceNo) { skipped++; return; }  // 未生成の行はスキップ
    try {
      regenerateInvoice_(row, { skipEmail: skipEmail });
      done++;
    } catch (err) {
      errors.push('行 ' + row + ': ' + err.message);
    }
  });

  ui.alert(
    '再生成 ' + done + ' 行完了' + (skipEmail ? '（メール送信なし）' : '（メール送信あり）') + '。' +
    (skipped ? '\n（番号未発行のためスキップ: ' + skipped + ' 行）' : '') +
    (errors.length ? '\n\nエラー:\n' + errors.join('\n') : '')
  );
}
