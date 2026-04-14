/**
 * Google Apps Script — Преміальний бекенд для анкети v4.0
 * 
 * ІНСТРУКЦІЯ:
 * ─────────────────────────────────────────────────────
 * 1. Відкрийте Google Таблицю
 * 2. Розширення → Apps Script
 * 3. Видаліть ВСЕ старе → вставте ЦЕЙ код
 * 4. Збережіть (Ctrl+S)
 * 5. Виберіть функцію «setupSheet» → натисніть ▶ Запустити
 * 6. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Скопіюйте URL → вставте в js/script.js
 * ─────────────────────────────────────────────────────
 */

// ═══════════════════════════════════════════════════
//  КОНФІГУРАЦІЯ
// ═══════════════════════════════════════════════════

/** Email для сповіщень. Порожньо = вимкнено */
const NOTIFICATION_EMAIL = '';

/** Макс. довжина поля */
const MAX_FIELD_LENGTH = 5000;

/** Маппінг полів */
const FIELD_MAP = [
  { key: 'pib',               col: 'ПІБ клієнта',                     width: 200 },
  { key: 'birthday',          col: 'День народження',                   width: 140 },
  { key: 'instagram',         col: 'Instagram',                         width: 200 },
  { key: 'sphere_experience', col: 'Сфера та досвід роботи',            width: 300 },
  { key: 'reason_purchase',   col: 'Причина придбання послуги',         width: 300 },
  { key: 'decision_point',    col: 'Що стало кінцевою точкою рішення',  width: 300 },
  { key: 'need_to_close',     col: 'Яку потребу хоче закрити',          width: 300 },
  { key: 'dream_goal',        col: 'Мрія або ціль',                     width: 300 },
  { key: 'expectations',      col: 'Очікування від співпраці',          width: 300 },
  { key: 'skills_to_learn',   col: 'Теми які хоче вивчити',             width: 300 },
  { key: 'page_difficulties', col: 'Труднощі ведення сторінки',         width: 300 },
  { key: 'past_experience',   col: 'Минулий досвід з послугами',        width: 300 },
  { key: 'additional_notes',  col: 'Додаткові зауваження',              width: 250 }
];

// Палітра кольорів
var C = {
  dark:      '#1e1e2e',
  accent:    '#4C84FF',
  accentBg:  '#e8eeff',
  white:     '#FFFFFF',
  light:     '#f7f8fc',
  gray:      '#f1f3f9',
  border:    '#e0e4ee',
  textDark:  '#1a1a2e',
  textMid:   '#555770',
  textLight: '#8b8fa8',
  green:     '#dcfce7',
  greenText: '#166534',
  pibBg:     '#dbeafe',
  pibText:   '#1e40af',
  emptyBg:   '#fef2f2',
  emptyText: '#b91c1c',
};


// ═══════════════════════════════════════════════════
//  УТИЛІТИ
// ═══════════════════════════════════════════════════

function sanitize(value) {
  if (typeof value !== 'string') return '';
  var clean = value.trim();
  // Захист від Formula injection
  if (/^[=+\-@\t\r]/.test(clean)) {
    clean = "'" + clean;
  }
  if (clean.length > MAX_FIELD_LENGTH) {
    clean = clean.substring(0, MAX_FIELD_LENGTH);
  }
  return clean;
}

function getMainSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Шукаємо лист «Відповіді» (може називатись по-різному)
  var sheet = ss.getSheetByName('Відповіді');
  if (!sheet) {
    // Спробуємо знайти перший лист
    sheet = ss.getSheets()[0];
  }
  return sheet;
}


// ═══════════════════════════════════════════════════
//  ГОЛОВНА ФУНКЦІЯ: SETUPSHEET
//  Запустіть ОДИН РАЗ перед Deploy
// ═══════════════════════════════════════════════════

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; // Беремо перший лист

  // Перейменувати лист
  sheet.setName('Відповіді');

  // Повністю очистити весь лист
  sheet.clear();
  sheet.clearFormats();
  sheet.clearNotes();

  // --- Рядок 1: НАЗВА ТАБЛИЦІ (мерж) ---
  var totalCols = FIELD_MAP.length + 2; // №, Час, + 13 полів = 15

  var titleRange = sheet.getRange(1, 1, 1, totalCols);
  titleRange.merge();
  titleRange.setValue('📋  ДІАГНОСТИЧНА АНКЕТА КЛІЄНТА  —  Відповіді');
  titleRange.setBackground(C.dark);
  titleRange.setFontColor(C.white);
  titleRange.setFontSize(14);
  titleRange.setFontWeight('bold');
  titleRange.setFontFamily('Arial');
  titleRange.setHorizontalAlignment('center');
  titleRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);

  // --- Рядок 2: ЗАГОЛОВКИ КОЛОНОК ---
  var headers = ['№', 'Дата та час'];
  for (var i = 0; i < FIELD_MAP.length; i++) {
    headers.push(FIELD_MAP[i].col);
  }

  var headerRange = sheet.getRange(2, 1, 1, totalCols);
  headerRange.setValues([headers]);
  headerRange.setBackground(C.accent);
  headerRange.setFontColor(C.white);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setFontFamily('Arial');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);
  sheet.setRowHeight(2, 42);

  // --- Ширина колонок ---
  sheet.setColumnWidth(1, 45);   // №
  sheet.setColumnWidth(2, 145);  // Дата та час
  for (var j = 0; j < FIELD_MAP.length; j++) {
    sheet.setColumnWidth(j + 3, FIELD_MAP[j].width);
  }

  // --- Закріпити рядки 1-2 ---
  sheet.setFrozenRows(2);

  // --- Захист заголовків ---
  var prot = sheet.getRange(1, 1, 2, totalCols).protect();
  prot.setDescription('Заголовки');
  prot.setWarningOnly(true);

  // --- Підготовка області даних (рядки 3-500) ---
  var dataArea = sheet.getRange(3, 1, 498, totalCols);
  dataArea.setFontFamily('Arial');
  dataArea.setFontSize(10);
  dataArea.setVerticalAlignment('top');
  dataArea.setWrap(true);

  // --- Бордери на заголовках ---
  headerRange.setBorder(true, true, true, true, true, true, C.dark, SpreadsheetApp.BorderStyle.SOLID);

  // --- Видалити зайві листи ---
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    if (sheets[s].getName() !== 'Відповіді' && sheets.length > 1) {
      try { ss.deleteSheet(sheets[s]); } catch(e) {}
    }
  }

  SpreadsheetApp.flush();
  Logger.log('✅ Таблицю налаштовано! Тепер Deploy → New deployment.');
}


// ═══════════════════════════════════════════════════
//  ОБРОБКА POST-ЗАПИТІВ
// ═══════════════════════════════════════════════════

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);

    var sheet = getMainSheet();

    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseError) {
      data = e.parameter || {};
    }

    // Валідація
    if (!data.pib || data.pib.trim() === '') {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'error', message: 'ПІБ обовʼязкове' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Часовий штамп
    var timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd.MM.yyyy  HH:mm'
    );

    // Номер заявки (рядок 1 = назва, рядок 2 = заголовки, дані з рядка 3)
    var lastRow = sheet.getLastRow();
    var rowNumber = Math.max(1, lastRow - 1); // -1 бо 2 рядки заголовків

    // Збираємо рядок
    var row = [rowNumber, timestamp];
    for (var i = 0; i < FIELD_MAP.length; i++) {
      row.push(sanitize(data[FIELD_MAP[i].key] || ''));
    }

    // Вставляємо
    sheet.appendRow(row);

    // Форматуємо рядок
    var newRowNum = sheet.getLastRow();
    formatRow(sheet, newRowNum);

    SpreadsheetApp.flush();
    lock.releaseLock();

    // Email
    if (NOTIFICATION_EMAIL) {
      sendEmail(data, timestamp, rowNumber);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', number: rowNumber })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', version: '4.0' })
  ).setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════
//  ФОРМАТУВАННЯ НОВОГО РЯДКА
// ═══════════════════════════════════════════════════

function formatRow(sheet, rowNum) {
  var totalCols = FIELD_MAP.length + 2;
  var range = sheet.getRange(rowNum, 1, 1, totalCols);

  // Зебра
  var isEven = (rowNum % 2 === 0);
  range.setBackground(isEven ? C.gray : C.white);

  // Шрифт
  range.setFontFamily('Arial');
  range.setFontSize(10);
  range.setFontColor(C.textDark);
  range.setVerticalAlignment('top');
  range.setWrap(true);

  // Мінімальна висота рядка
  sheet.setRowHeight(rowNum, 36);

  // --- Колонка № ---
  var numCell = sheet.getRange(rowNum, 1);
  numCell.setBackground(C.accentBg);
  numCell.setFontColor(C.accent);
  numCell.setFontWeight('bold');
  numCell.setHorizontalAlignment('center');
  numCell.setVerticalAlignment('middle');

  // --- Колонка Час ---
  var timeCell = sheet.getRange(rowNum, 2);
  timeCell.setBackground(C.green);
  timeCell.setFontColor(C.greenText);
  timeCell.setFontSize(9);
  timeCell.setHorizontalAlignment('center');
  timeCell.setVerticalAlignment('middle');

  // --- Колонка ПІБ (3) ---
  var pibCell = sheet.getRange(rowNum, 3);
  pibCell.setBackground(C.pibBg);
  pibCell.setFontColor(C.pibText);
  pibCell.setFontWeight('bold');
  pibCell.setFontSize(11);

  // --- Колонка Instagram (5) — зробити клікабельним ---
  var igCell = sheet.getRange(rowNum, 5);
  var igVal = igCell.getValue();
  if (igVal && String(igVal).length > 0) {
    var igStr = String(igVal).trim();
    var igUrl = igStr;
    // Нормалізуємо URL
    if (igStr.charAt(0) === '@') {
      igUrl = 'https://instagram.com/' + igStr.substring(1);
    } else if (igStr.indexOf('http') !== 0) {
      igUrl = 'https://instagram.com/' + igStr;
    }
    // Використовуємо RichTextValue замість HYPERLINK формули
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(igStr)
      .setLinkUrl(igUrl)
      .build();
    igCell.setRichTextValue(richText);
    igCell.setFontColor(C.accent);
  }

  // --- Позначити пусті обов'язкові поля ---
  // Колонки 3-14 (ПІБ до Минулий досвід) — обов'язкові
  // Колонка 15 (Додаткові) — необов'язкова
  for (var col = 3; col <= totalCols - 1; col++) {
    var cell = sheet.getRange(rowNum, col);
    var val = cell.getValue();
    if (val === '' || val === null || val === undefined) {
      cell.setBackground(C.emptyBg);
      cell.setValue('—');
      cell.setFontColor(C.emptyText);
      cell.setHorizontalAlignment('center');
      cell.setFontStyle('italic');
    }
  }

  // --- Тонкий бордер знизу ---
  range.setBorder(
    false, false, true, false, false, false,
    C.border, SpreadsheetApp.BorderStyle.SOLID
  );
}


// ═══════════════════════════════════════════════════
//  EMAIL СПОВІЩЕННЯ
// ═══════════════════════════════════════════════════

function sendEmail(data, timestamp, number) {
  try {
    var name = data.pib || 'Невідомий';
    var subject = '📋 Заявка #' + number + ': ' + name;

    var body = 'Нова заявка #' + number + ' о ' + timestamp + '\n\n';
    for (var i = 0; i < FIELD_MAP.length; i++) {
      var val = data[FIELD_MAP[i].key] || '—';
      body += FIELD_MAP[i].col + ':\n' + val + '\n\n';
    }
    body += '━━━━━━━━━━━━━━━━━━━━━━━━\n';
    body += 'Таблиця: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();

    MailApp.sendEmail({ to: NOTIFICATION_EMAIL, subject: subject, body: body });
  } catch (err) {
    Logger.log('Email error: ' + err);
  }
}
