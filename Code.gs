const SOURCE_SHEET_CANDIDATES = ['Индвидуальные показатели', 'Индивидуальные показатели'];
const RECIPIENT_SHEET_NAME = 'Отправка';
const BONUS_SHEET_PREFIX = 'Премии ';
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
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('TEMED')
    .addItem('Отобрать врачей', 'selectDoctors')
    .addItem('Сформировать сообщения ГВ (премии)', 'buildClinicBonusMessages')
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

  const headerMap = createHeaderAliasMap_(sourceData[0], {
    month: ['Месяц', 'Код месяца'],
    doctor: ['Врач', 'Фамилия', 'ФИО'],
    clinic: ['Клиника'],
    totalBonus: ['Премия ИТОГО (округл)'],
  });

  const rows = sourceData
    .slice(1)
    .filter((row) => String(row[headerMap.month]).trim() === monthCode)
    .filter((row) => isBonusValueFilled_(row[headerMap.totalBonus]))
    .map((row) => [
      row[headerMap.doctor],
      row[headerMap.clinic],
      normalizeBonusValue_(row[headerMap.totalBonus]),
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
  const bonusValues = bonusSheet.getDataRange().getDisplayValues();
  const emailContent = buildBonusEmailContent_(monthInfo, bonusValues);

  const attachment = buildSheetCsvBlob_(bonusValues, bonusSheetName + '.csv');
  MailApp.sendEmail({
    to: recipients.join(','),
    subject,
    body: emailContent.textBody,
    htmlBody: emailContent.htmlBody,
    attachments: [attachment],
  });

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

function createHeaderAliasMap_(headers, requiredAliases) {
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[String(header).trim()] = index;
  });

  const resolvedMap = {};
  Object.keys(requiredAliases).forEach((key) => {
    const match = requiredAliases[key].find((header) => headerMap[header] !== undefined);
    if (match === undefined) {
      throw new Error(`Не найдена колонка. Поддерживаемые варианты: ${requiredAliases[key].join(', ')}.`);
    }
    resolvedMap[key] = headerMap[match];
  });

  return resolvedMap;
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

function buildSheetCsvBlob_(values, fileName) {
  const csv = values
    .map((row) => row.map(escapeCsvValue_).join(';'))
    .join('\r\n');

  return Utilities.newBlob('\uFEFF' + csv, 'text/csv', fileName);
}

function buildBonusEmailContent_(monthInfo, values) {
  const instructionsText = [
    'Файл CSV из приложения можно открыть в Excel через меню Файл → Открыть или просто перетащить файл в окно программы. Если данные отображаются некорректно — укажите разделитель столбцов ; (точка с запятой) при импорте.',
    'В Google Таблицах откройте таблицы, выберите Файл → Импорт → Загрузка и загрузите CSV-файл, при необходимости также выберите разделитель ;.',
  ];
  const copyBlock = buildCopyFriendlyTableText_(values);
  const htmlTable = buildHtmlTable_(values);

  return {
    textBody: [
      'Добрый день!',
      '',
      ...instructionsText,
      '',
      `Во вложении файл с премиями врачей за ${monthInfo.monthName} ${monthInfo.fullYear}.`,
      '',
      'Таблица для копирования:',
      copyBlock,
      '',
      'Письмо сформировано автоматически.',
    ].join('\n'),
    htmlBody: [
      '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#202124;">',
      '<p>Добрый день!</p>',
      '<p>Файл CSV из приложения можно открыть в Excel через меню Файл → Открыть или просто перетащить файл в окно программы. Если данные отображаются некорректно — укажите разделитель столбцов <b>;</b> (точка с запятой) при импорте.</p>',
      '<p>В Google Таблицах откройте таблицы, выберите Файл → Импорт → Загрузка и загрузите CSV-файл, при необходимости также выберите разделитель <b>;</b>.</p>',
      `<p>Во вложении файл с премиями врачей за ${escapeHtml_(monthInfo.monthName)} ${escapeHtml_(monthInfo.fullYear)}.</p>`,
      '<p><b>Таблица для копирования:</b></p>',
      `<pre style="margin:0 0 16px;padding:12px;background:#f6f8fa;border:1px solid #d0d7de;border-radius:6px;font-family:Consolas,Monaco,monospace;font-size:13px;white-space:pre-wrap;">${escapeHtml_(copyBlock)}</pre>`,
      htmlTable,
      '<p>Письмо сформировано автоматически.</p>',
      '</div>',
    ].join(''),
  };
}

function buildCopyFriendlyTableText_(values) {
  return values
    .map((row) => row.map((value) => String(value ?? '')).join('\t'))
    .join('\n');
}

function buildHtmlTable_(values) {
  if (!values.length) {
    return '';
  }

  const headerCells = values[0]
    .map((value) => `<th style="border:1px solid #d0d7de;padding:8px 10px;background:#f6f8fa;text-align:left;">${escapeHtml_(value)}</th>`)
    .join('');
  const bodyRows = values
    .slice(1)
    .map((row) => '<tr>' + row
      .map((value) => `<td style="border:1px solid #d0d7de;padding:8px 10px;">${escapeHtml_(value)}</td>`)
      .join('') + '</tr>')
    .join('');

  return [
    '<table style="border-collapse:collapse;border-spacing:0;margin:0 0 16px;font-family:Arial,sans-serif;font-size:14px;">',
    `<thead><tr>${headerCells}</tr></thead>`,
    `<tbody>${bodyRows}</tbody>`,
    '</table>',
  ].join('');
}

function escapeCsvValue_(value) {
  const normalized = String(value ?? '').replace(/"/g, '""');
  return '"' + normalized + '"';
}

function escapeHtml_(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function selectDoctors() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sourceSheet = ss.getSheetByName('Отчет по врачу v2');
  var targetSheet = ss.getSheetByName('Индвидуальные показатели');

  if (!sourceSheet) {
    ui.alert('Лист "Отчет по врачу v2" не найден!');
    return;
  }
  if (!targetSheet) {
    ui.alert('Лист "Индвидуальные показатели" не найден!');
    return;
  }

  function normMonth(m) {
    if (m === null || m === '') return '';
    var s = m.toString().trim();
    if (/^\d+$/.test(s) && s.length < 4) s = s.padStart(4, '0');
    return s;
  }

  function normSurname(s) {
    if (s === null || s === '') return '';
    return s.toString().trim();
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function pickIndex(headerIndex, possibleNames) {
    for (var k = 0; k < possibleNames.length; k++) {
      var name = possibleNames[k];
      if (Object.prototype.hasOwnProperty.call(headerIndex, name)) {
        return headerIndex[name];
      }
    }
    return -1;
  }

  var srcLastRow = sourceSheet.getLastRow();
  if (srcLastRow < 2) {
    ui.alert('В "Отчет по врачу v2" нет данных.');
    return;
  }

  var srcValues = sourceSheet.getRange(1, 1, srcLastRow, 8).getValues();

  var stopRow = -1;
  for (var i = 0; i < srcValues.length; i++) {
    if (srcValues[i][1] && srcValues[i][1].toString().trim() === 'Итого') {
      stopRow = i;
      break;
    }
  }

  if (stopRow === -1) {
    ui.alert('Не найдено значение "Итого" для ограничения диапазона!');
    return;
  }

  var revenueByMonthDoctor = new Map();
  for (var r0 = 0; r0 < stopRow; r0++) {
    var m0 = normMonth(srcValues[r0][0]);
    var s0 = normSurname(srcValues[r0][1]);
    var h0 = srcValues[r0][7];
    if (!m0 || !s0) continue;
    var add = typeof h0 === 'number' ? h0 : 0;
    var k0 = m0 + '|' + s0;
    revenueByMonthDoctor.set(k0, (revenueByMonthDoctor.get(k0) || 0) + add);
  }

  var doctors = [];
  var seen = new Set();

  for (var r = 0; r < stopRow; r++) {
    var monthCode = normMonth(srcValues[r][0]);
    var surname = normSurname(srcValues[r][1]);
    var valueC = srcValues[r][2];

    if (!monthCode || !surname) continue;

    if (typeof valueC === 'number' && valueC > 4 && !surname.toUpperCase().startsWith('ТЕМЕД')) {
      var key = monthCode + '|' + surname;
      if (!seen.has(key)) {
        seen.add(key);
        doctors.push({ month: monthCode, surname: surname });
      }
    }
  }

  if (doctors.length === 0) {
    ui.alert('Врачи не найдены!');
    return;
  }

  var tgtLastCol = targetSheet.getLastColumn();
  if (tgtLastCol === 0) {
    ui.alert('Лист "Индвидуальные показатели" пустой (нет заголовков).');
    return;
  }

  var headers = targetSheet.getRange(1, 1, 1, tgtLastCol).getValues()[0];
  var headerIndex = {};
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h !== null && h !== '') {
      headerIndex[h.toString().trim()] = c;
    }
  }

  var colMonth = pickIndex(headerIndex, ['Код месяца', 'Месяц']);
  var colSurname = pickIndex(headerIndex, ['Фамилия', 'Врач', 'ФИО']);
  var colRole = pickIndex(headerIndex, ['Должность']);
  var colClinic = pickIndex(headerIndex, ['Клиника']);
  var colShare = pickIndex(headerIndex, ['% от общей выручки']);
  var colMinTotal = pickIndex(headerIndex, ['Минимальная общая выручка']);
  var colTotalRevenue = pickIndex(headerIndex, ['Общая выручка', 'Общая  выручка']);

  var missingCols = [];
  if (colMonth === -1) missingCols.push('Код месяца (или Месяц)');
  if (colSurname === -1) missingCols.push('Фамилия/Врач/ФИО');
  if (colRole === -1) missingCols.push('Должность');
  if (colClinic === -1) missingCols.push('Клиника');
  if (colShare === -1) missingCols.push('% от общей выручки');
  if (colMinTotal === -1) missingCols.push('Минимальная общая выручка');
  if (colTotalRevenue === -1) missingCols.push('Общая выручка');

  if (missingCols.length > 0) {
    ui.alert('Не найдены обязательные столбцы в "Индвидуальные показатели":\n- ' + missingCols.join('\n- '));
    return;
  }

  var tgtLastRow = targetSheet.getLastRow();
  var targetData = tgtLastRow >= 2
    ? targetSheet.getRange(1, 1, tgtLastRow, tgtLastCol).getValues()
    : [headers];

  var existingMonthNums = [];
  var existingMonthStrSet = new Set();

  for (var tr = 1; tr < targetData.length; tr++) {
    var m = normMonth(targetData[tr][colMonth]);
    if (m) {
      existingMonthStrSet.add(m);
      var mn = parseInt(m, 10);
      if (!isNaN(mn)) existingMonthNums.push(mn);
    }
  }
  existingMonthNums.sort(function(a, b) { return a - b; });

  function getPrevMonth(monthStr) {
    var cur = parseInt(monthStr, 10);
    if (isNaN(cur)) return '';
    var prev = null;
    for (var i2 = 0; i2 < existingMonthNums.length; i2++) {
      if (existingMonthNums[i2] < cur) {
        prev = existingMonthNums[i2];
      } else {
        break;
      }
    }
    if (prev === null) return '';
    return prev.toString().padStart(4, '0');
  }

  var prevLookup = new Map();
  for (var tr2 = 1; tr2 < targetData.length; tr2++) {
    var m2 = normMonth(targetData[tr2][colMonth]);
    var s2 = normSurname(targetData[tr2][colSurname]);
    if (!m2 || !s2) continue;

    var key2 = m2 + '|' + s2;
    if (!prevLookup.has(key2)) {
      prevLookup.set(key2, {
        role: targetData[tr2][colRole],
        clinic: targetData[tr2][colClinic],
        share: targetData[tr2][colShare],
        minTotal: targetData[tr2][colMinTotal],
      });
    }
  }

  var monthsToWriteSet = new Set();
  doctors.forEach(function(doctor) {
    monthsToWriteSet.add(doctor.month);
  });
  var monthsToWrite = Array.from(monthsToWriteSet);

  var existingMonths = [];
  monthsToWrite.forEach(function(month) {
    if (existingMonthStrSet.has(month)) existingMonths.push(month);
  });

  if (existingMonths.length > 0) {
    var resp = ui.alert(
      'Найдены данные за месяц(ы) с кодом: ' + existingMonths.join(', ') + '\n\nПерезаписать существующие данные?',
      ui.ButtonSet.YES_NO
    );
    if (resp === ui.Button.NO) return;

    var existingSet = new Set(existingMonths);
    var lastRowNow = targetSheet.getLastRow();
    if (lastRowNow >= 2) {
      var dataNow = targetSheet.getRange(1, 1, lastRowNow, tgtLastCol).getValues();
      for (var rr = dataNow.length - 1; rr >= 1; rr--) {
        var mm = normMonth(dataNow[rr][colMonth]);
        if (mm && existingSet.has(mm)) {
          targetSheet.deleteRow(rr + 1);
        }
      }
    }
  }

  var rows = [];
  var missingPrev = [];
  var missingRowIdx = [];

  for (var d = 0; d < doctors.length; d++) {
    var mCur = doctors[d].month;
    var sCur = doctors[d].surname;
    var prevMonth = getPrevMonth(mCur);
    var row = new Array(tgtLastCol).fill('');

    row[colMonth] = mCur;
    row[colSurname] = sCur;

    var revKey = mCur + '|' + sCur;
    row[colTotalRevenue] = revenueByMonthDoctor.get(revKey) || 0;

    if (prevMonth) {
      var lookup = prevLookup.get(prevMonth + '|' + sCur);
      if (lookup) {
        row[colRole] = lookup.role;
        row[colClinic] = lookup.clinic;
        row[colShare] = lookup.share;
        row[colMinTotal] = lookup.minTotal;
      } else {
        missingPrev.push({ month: mCur, surname: sCur, prevMonth: prevMonth });
        missingRowIdx.push(rows.length);
      }
    } else {
      missingPrev.push({ month: mCur, surname: sCur, prevMonth: '—' });
      missingRowIdx.push(rows.length);
    }

    rows.push(row);
  }

  var insertStartRow = Math.max(targetSheet.getLastRow() + 1, 2);
  targetSheet.getRange(insertStartRow, 1, rows.length, tgtLastCol).setValues(rows);

  if (missingRowIdx.length > 0) {
    var paleYellow = '#FFF2CC';

    missingRowIdx.forEach(function(idx) {
      targetSheet.getRange(insertStartRow + idx, colSurname + 1, 1, 1).setBackground(paleYellow);
    });

    var html = '<div style="font-family:Arial; font-size:13px;">' +
      '<b>Врачи, которых нет в предыдущем месяце</b><br><br>' +
      '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
      '<tr><th>Месяц</th><th>ФИО</th><th>Предыдущий месяц</th></tr>';

    missingPrev.forEach(function(item) {
      html += '<tr>' +
        '<td>' + escapeHtml(item.month) + '</td>' +
        '<td>' + escapeHtml(item.surname) + '</td>' +
        '<td>' + escapeHtml(item.prevMonth) + '</td>' +
        '</tr>';
    });

    html += '</table><br>' +
      '<div style="color:#666;">Подсвечена только ячейка с ФИО.</div>' +
      '</div>';

    ui.showModalDialog(
      HtmlService.createHtmlOutput(html).setWidth(650).setHeight(450),
      'Нет в предыдущем месяце'
    );
  }

  ui.alert('Отобрано врачей: ' + doctors.length);
}

function buildClinicBonusMessages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  function normMonth(m) {
    if (m === null || m === '') return '';
    var s = String(m).trim();
    if (/^\d+$/.test(s) && s.length < 4) s = s.padStart(4, '0');
    return s;
  }

  function pickIndex(headerIndex, names) {
    for (var i = 0; i < names.length; i++) {
      if (Object.prototype.hasOwnProperty.call(headerIndex, names[i])) {
        return headerIndex[names[i]];
      }
    }
    return -1;
  }

  function firstWord(s) {
    if (s === null || s === '') return '';
    var text = String(s).trim();
    if (!text) return '';
    return text.split(/\s+/)[0];
  }

  function toNumber(v) {
    if (v === null || v === '' || typeof v === 'undefined') return 0;
    if (typeof v === 'number') return v;
    var normalized = String(v).replace(/\s/g, '').replace(',', '.');
    var result = Number(normalized);
    return isNaN(result) ? 0 : result;
  }

  function fmtK(n) {
    var k = toNumber(n) / 1000;
    var formatted = k.toFixed(1);
    if (formatted.endsWith('.0')) formatted = formatted.slice(0, -2);
    return formatted + 'к';
  }

  function fmtDiffK(diff, hasPrevValue) {
    if (!hasPrevValue) return '(—)';
    var numericDiff = toNumber(diff);
    var sign = numericDiff >= 0 ? '+' : '-';
    return '(' + sign + fmtK(Math.abs(numericDiff)) + ')';
  }

  function isChiefRole(roleVal) {
    if (roleVal === null || roleVal === '') return false;
    return String(roleVal).trim().toLowerCase() === 'главный врач';
  }

  var active = ss.getActiveRange();
  if (!active) {
    ui.alert('Нет активной ячейки.');
    return;
  }

  var monthCode = normMonth(active.getValue());
  if (!monthCode || !/^\d{4}$/.test(monthCode)) {
    ui.alert('В активной ячейке должен быть код месяца из 4 цифр (например 0123 или 1234).');
    return;
  }

  var srcSheet = getSourceSheet_(ss);
  if (!srcSheet) {
    ui.alert('Лист "Индвидуальные показатели" не найден!');
    return;
  }

  var msgSheet = ss.getSheetByName('Сообщения') || ss.insertSheet('Сообщения');
  var lastRow = srcSheet.getLastRow();
  var lastCol = srcSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) {
    ui.alert('Лист "Индвидуальные показатели" пустой или без данных.');
    return;
  }

  var data = srcSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  var headerIndex = {};
  headers.forEach(function(header, idx) {
    if (header !== null && header !== '') headerIndex[String(header).trim()] = idx;
  });

  var colMonth = pickIndex(headerIndex, ['Месяц', 'Код месяца']);
  var colClinic = pickIndex(headerIndex, ['Клиника']);
  var colDoctor = pickIndex(headerIndex, ['Врач', 'Фамилия', 'ФИО']);
  var colRole = pickIndex(headerIndex, ['Должность']);
  var colPrem = pickIndex(headerIndex, ['Премия ИТОГО (округл)']);

  var missing = [];
  if (colMonth === -1) missing.push('Месяц');
  if (colClinic === -1) missing.push('Клиника');
  if (colDoctor === -1) missing.push('Врач');
  if (colRole === -1) missing.push('Должность');
  if (colPrem === -1) missing.push('Премия ИТОГО (округл)');
  if (missing.length) {
    ui.alert('На листе "Индвидуальные показатели" не найдены столбцы:\n- ' + missing.join('\n- '));
    return;
  }

  var monthNums = [];
  var monthSet = new Set();
  for (var r = 1; r < data.length; r++) {
    var monthValue = normMonth(data[r][colMonth]);
    if (monthValue && /^\d{4}$/.test(monthValue) && !monthSet.has(monthValue)) {
      monthSet.add(monthValue);
      var monthNumber = parseInt(monthValue, 10);
      if (!isNaN(monthNumber)) monthNums.push(monthNumber);
    }
  }
  monthNums.sort(function(a, b) { return a - b; });

  var curNum = parseInt(monthCode, 10);
  var prevNum = null;
  for (var j = 0; j < monthNums.length; j++) {
    if (monthNums[j] < curNum) {
      prevNum = monthNums[j];
    } else {
      break;
    }
  }
  var prevMonthCode = prevNum === null ? '' : String(prevNum).padStart(4, '0');

  var curByClinic = new Map();
  var prevByClinicDoctor = new Map();

  for (var rowIndex = 1; rowIndex < data.length; rowIndex++) {
    var currentMonth = normMonth(data[rowIndex][colMonth]);
    var clinic = data[rowIndex][colClinic];
    var doctor = data[rowIndex][colDoctor];
    var role = data[rowIndex][colRole];
    var prem = toNumber(data[rowIndex][colPrem]);

    if (!clinic || !doctor) continue;
    if (isChiefRole(role)) continue;

    var clinicName = String(clinic).trim();
    var doctorFull = String(doctor).trim();
    if (!clinicName || !doctorFull) continue;

    if (currentMonth === monthCode) {
      if (!curByClinic.has(clinicName)) curByClinic.set(clinicName, new Map());
      var docMap = curByClinic.get(clinicName);
      docMap.set(doctorFull, (docMap.get(doctorFull) || 0) + prem);
    }

    if (prevMonthCode && currentMonth === prevMonthCode) {
      var prevKey = clinicName + '|' + doctorFull;
      prevByClinicDoctor.set(prevKey, (prevByClinicDoctor.get(prevKey) || 0) + prem);
    }
  }

  msgSheet.clear();
  msgSheet.getRange(1, 1, 1, 2).setValues([['Клиника', 'Сообщение']]);

  if (curByClinic.size === 0) {
    ui.alert('По коду месяца ' + monthCode + ' (без "Главный врач") ничего не найдено.');
    return;
  }

  var out = [];
  var clinics = Array.from(curByClinic.keys());
  clinics.sort(function(a, b) { return a.localeCompare(b, 'ru'); });

  clinics.forEach(function(clinicName) {
    var lines = [];
    lines.push('Премия на согласование за месяц (код месяца): ' + monthCode);
    lines.push('Клиника: ' + clinicName);

    var docMap = curByClinic.get(clinicName);
    var doctors = Array.from(docMap.keys());
    doctors.sort(function(a, b) {
      return firstWord(a).localeCompare(firstWord(b), 'ru');
    });

    doctors.forEach(function(docFull) {
      var curSum = docMap.get(docFull) || 0;
      var surname = firstWord(docFull);
      var key = clinicName + '|' + docFull;
      var hasPrev = !!(prevMonthCode && prevByClinicDoctor.has(key));
      var prevSum = hasPrev ? prevByClinicDoctor.get(key) || 0 : 0;
      var diff = curSum - prevSum;

      lines.push(surname + ' ' + fmtK(curSum) + ' ' + fmtDiffK(diff, hasPrev));
    });

    out.push([clinicName, lines.join('\n')]);
  });

  msgSheet.getRange(2, 1, out.length, 2).setValues(out);

  ui.alert(
    'Сформировано сообщений по клиникам: ' + out.length +
    (prevMonthCode ? ' (сравнение с ' + prevMonthCode + ')' : ' (предыдущего месяца в данных нет)')
  );
}
