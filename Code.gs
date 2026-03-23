const SOURCE_SHEET_CANDIDATES = ['Индвидуальные показатели', 'Индивидуальные показатели'];
const RECIPIENT_SHEET_NAME = 'Отправка';
const BONUS_SHEET_PREFIX = 'Премии+';
const MENU_NAME = 'Премии';
const MONTH_NAMES_RU = [
  'январь',
  'февраль',
  'март',
  'апрель',
  'май',
  'июнь',
  'июль',
  'август',
  'сентябрь',
  'октябрь',
  'ноябрь',
  'декабрь',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('Сформировать месячные премии', 'generateMonthlyBonuses')
    .addItem('Отправить премии', 'sendMonthlyBonuses')
    .addToUi();
}

function generateMonthlyBonuses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const monthCode = getMonthCodeFromActiveCell_(spreadsheet);
  const sourceSheet = getSourceSheet_(spreadsheet);
  const sourceData = sourceSheet.getDataRange().getValues();

  if (sourceData.length < 2) {
    throw new Error('На листе с индивидуальными показателями нет данных для обработки.');
  }

  const headerMap = createHeaderMap_(sourceData[0], [
    'Месяц',
    'Врач',
    'Клиника',
    'Премия ИТОГО (округл)',
  ]);

  const rows = sourceData
    .slice(1)
    .filter((row) => String(row[headerMap['Месяц']]).trim() === monthCode)
    .filter((row) => isBonusValueFilled_(row[headerMap['Премия ИТОГО (округл)']]))
    .map((row) => [
      row[headerMap['Врач']],
      row[headerMap['Клиника']],
      normalizeBonusValue_(row[headerMap['Премия ИТОГО (округл)']]),
    ]);

  if (!rows.length) {
    throw new Error(`По коду месяца ${monthCode} не найдено строк с ненулевой премией.`);
  }

  const targetSheetName = BONUS_SHEET_PREFIX + monthCode;
  const existingSheet = spreadsheet.getSheetByName(targetSheetName);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }

  const targetSheet = spreadsheet.insertSheet(targetSheetName);
  const output = [['Врач', 'Клиника', 'Премия ИТОГО (округл)']].concat(rows);
  targetSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  targetSheet.getRange(2, 3, rows.length, 1).setNumberFormat('0');
  targetSheet.autoResizeColumns(1, output[0].length);

  SpreadsheetApp.getUi().alert(`Лист ${targetSheetName} сформирован. Строк выгружено: ${rows.length}.`);
}

function sendMonthlyBonuses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const monthCode = getMonthCodeFromActiveCell_(spreadsheet);
  const bonusSheetName = BONUS_SHEET_PREFIX + monthCode;
  const bonusSheet = spreadsheet.getSheetByName(bonusSheetName);

  if (!bonusSheet) {
    throw new Error(`Сначала сформируйте лист ${bonusSheetName}.`);
  }

  const recipients = getRecipients_();
  if (!recipients.length) {
    throw new Error('На листе "Отправка" не найдены адреса в колонке "Кому".');
  }

  const monthInfo = getMonthInfo_(monthCode);
  const subject = `Премия врачей ${monthInfo.monthName} ${monthInfo.fullYear}`;
  const body = [
    'Добрый день!',
    '',
    `Во вложении файл с премиями врачей за ${monthInfo.monthName} ${monthInfo.fullYear}.`,
    '',
    'Письмо сформировано автоматически.',
  ].join('\n');

  const tempSpreadsheet = SpreadsheetApp.create(`${bonusSheetName}_temp_export`);
  let tempFile;

  try {
    const copiedSheet = bonusSheet.copyTo(tempSpreadsheet).setName(bonusSheetName);
    const defaultSheet = tempSpreadsheet.getSheets().find((sheet) => sheet.getSheetId() !== copiedSheet.getSheetId());
    if (defaultSheet) {
      tempSpreadsheet.deleteSheet(defaultSheet);
    }

    const blob = exportSpreadsheetAsXlsx_(tempSpreadsheet.getId(), bonusSheetName + '.xlsx');
    MailApp.sendEmail({
      to: recipients.join(','),
      subject,
      body,
      attachments: [blob],
    });
  } finally {
    tempFile = DriveApp.getFileById(tempSpreadsheet.getId());
    if (tempFile) {
      tempFile.setTrashed(true);
    }
  }

  SpreadsheetApp.getUi().alert(`Письмо отправлено: ${recipients.join(', ')}.`);
}

function getMonthCodeFromActiveCell_(spreadsheet) {
  const activeRange = spreadsheet.getActiveRange();
  if (!activeRange) {
    throw new Error('Выберите ячейку с кодом месяца перед запуском команды.');
  }

  const rawValue = String(activeRange.getDisplayValue()).trim();
  if (!/^\d{4}$/.test(rawValue)) {
    throw new Error('В активной ячейке должен быть 4-значный код месяца в формате ГГММ, например 2604.');
  }

  getMonthInfo_(rawValue);
  return rawValue;
}

function getMonthInfo_(monthCode) {
  const yearCode = Number(monthCode.slice(0, 2));
  const monthIndex = Number(monthCode.slice(2, 4));
  if (monthIndex < 1 || monthIndex > 12) {
    throw new Error(`Некорректный код месяца: ${monthCode}. Последние две цифры должны быть от 01 до 12.`);
  }

  return {
    monthName: MONTH_NAMES_RU[monthIndex - 1],
    shortYear: String(yearCode).padStart(2, '0'),
    fullYear: 2000 + yearCode,
  };
}

function getSourceSheet_(spreadsheet) {
  const sourceSheet = SOURCE_SHEET_CANDIDATES
    .map((name) => spreadsheet.getSheetByName(name))
    .find(Boolean);

  if (!sourceSheet) {
    throw new Error('Не найден лист "Индвидуальные показатели".');
  }

  return sourceSheet;
}

function getRecipients_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(RECIPIENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('Не найден лист "Отправка".');
  }

  const data = sheet.getDataRange().getValues();
  if (!data.length) {
    return [];
  }

  const headerMap = createHeaderMap_(data[0], ['Кому']);
  const seen = {};

  return data
    .slice(1)
    .map((row) => String(row[headerMap['Кому']]).trim())
    .filter((email) => email)
    .filter((email) => {
      const key = email.toLowerCase();
      if (seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });
}

function createHeaderMap_(headers, requiredHeaders) {
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[String(header).trim()] = index;
  });

  requiredHeaders.forEach((header) => {
    if (headerMap[header] === undefined) {
      throw new Error(`Не найдена колонка "${header}".`);
    }
  });

  return headerMap;
}

function isBonusValueFilled_(value) {
  if (value === '' || value === null || value === undefined) {
    return false;
  }

  const normalized = normalizeBonusValue_(value);
  return normalized !== '0';
}

function normalizeBonusValue_(value) {
  if (typeof value === 'number') {
    return String(Math.round(value));
  }

  const digits = String(value).replace(/\D+/g, '');
  const normalized = digits.replace(/^0+(?=\d)/, '');
  return normalized || '0';
}

function exportSpreadsheetAsXlsx_(spreadsheetId, fileName) {
  const url = `https://www.googleapis.com/drive/v3/files/${spreadsheetId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
    muteHttpExceptions: true,
  });

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error(`Не удалось экспортировать XLSX. Код ответа: ${responseCode}.`);
  }

  return response.getBlob().setName(fileName);
}
