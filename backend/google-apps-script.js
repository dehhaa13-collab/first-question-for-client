/**
 * ─────────────────────────────────────────────────────────
 * Google Apps Script v5.0 — Повністю автоматичний бекенд!
 * ─────────────────────────────────────────────────────────
 * ВАМ БІЛЬШЕ НЕ ТРЕБА НІЧОГО ЗАПУСКАТИ ВРУЧНУ!
 * 
 * 1. Вставте цей код у редактор Apps Script.
 * 2. Зробіть Deploy -> Manage deployments -> Редагувати (олівець) -> New version -> Deploy.
 * 3. Все! При наступній відправці з сайту, скрипт САМ створить
 *    всі красиві колонки, кольори та ширину у вашій таблиці.
 * ─────────────────────────────────────────────────────────
 */

var NOTIFICATION_EMAIL = '';
var MAX_FIELD_LENGTH = 5000;

var FIELD_MAP = [
  { key: 'pib',               col: 'ПІБ клієнта',                     width: 200 },
  { key: 'birthday',          col: 'День народження',                   width: 130 },
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

var C = {
  dark:      '#1e1e2e',
  accent:    '#4C84FF',
  accentBg:  '#e8eeff',
  white:     '#FFFFFF',
  light:     '#f7f8fc',
  gray:      '#fcfcfd',
  border:    '#e0e4ee',
  textDark:  '#1a1a2e',
  green:     '#ecfdf5',
  greenText: '#065f46',
  pibBg:     '#eff6ff',
  pibText:   '#1e40af',
  emptyBg:   '#fff1f2',
  emptyText: '#be123c',
};

function sanitize(value) {
  if (typeof value !== 'string') return '';
  var clean = value.trim();
  if (/^[=+\-@\t\r]/.test(clean)) clean = "'" + clean;
  if (clean.length > MAX_FIELD_LENGTH) clean = clean.substring(0, MAX_FIELD_LENGTH);
  return clean;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Знаходимо або створюємо лист "Анкети"
    var sheet = ss.getSheetByName('Анкети') || ss.insertSheet('Анкети');
    
    // АВТОМАТИЧНЕ НАЛАШТУВАННЯ ТАБЛИЦІ (якщо його ще нема)
    var isInitialized = sheet.getRange(1, 1).getValue() === '📋 ДІАГНОСТИЧНА АНКЕТА';
    if (!isInitialized) {
      setupBeautifulHeaders(sheet);
    }
    
    var data;
    try { data = JSON.parse(e.postData.contents); } 
    catch (err) { data = e.parameter || {}; }

    if (!data.pib || data.pib.trim() === '') {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'ПІБ обовʼязкове' })).setMimeType(ContentService.MimeType.JSON);
    }

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    var lastRow = Math.max(2, sheet.getLastRow()); // Заголовки займають 2 рядки
    var rowNumber = lastRow - 1; 

    // Формуємо масив даних для рядка
    var row = [rowNumber, timestamp];
    for (var i = 0; i < FIELD_MAP.length; i++) {
      row.push(sanitize(data[FIELD_MAP[i].key] || ''));
    }

    // Вставляємо новий рядок
    sheet.appendRow(row);
    var newRowNum = sheet.getLastRow();
    
    // Фарбуємо новий рядок
    formatRow(sheet, newRowNum);

    // Видаляємо дефолтний Лист1 якщо він пустий
    var sheet1 = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
    if (sheet1 && ss.getSheets().length > 1 && sheet1.getLastRow() === 0) {
      ss.deleteSheet(sheet1);
    }

    SpreadsheetApp.flush();
    lock.releaseLock();
    
    // Відправляємо імейл якщо налаштовано
    if (NOTIFICATION_EMAIL) {
      sendEmail(data, timestamp, rowNumber);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', number: rowNumber })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    if (lock) lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', version: '5.0-AUTO' })).setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────
// МАЛЮЄМО КРАСИВІ ЗАГОЛОВКИ (ТІЛЬКИ 1 РАЗ)
// ─────────────────────────────────────────────────────────
function setupBeautifulHeaders(sheet) {
  // Якщо юзер вже щось наклікав в рядок 1 (тестові дані) — зсуваємо їх вниз
  if (sheet.getLastRow() > 0) {
    sheet.insertRowsBefore(1, 2);
  }
  
  var totalCols = FIELD_MAP.length + 2;
  
  // -- РЯДОК 1: ТЕМНИЙ БАНЕР --
  var titleRange = sheet.getRange(1, 1, 1, totalCols);
  titleRange.merge();
  titleRange.setValue('📋 ДІАГНОСТИЧНА АНКЕТА');
  titleRange.setBackground(C.dark);
  titleRange.setFontColor(C.white);
  titleRange.setFontSize(14);
  titleRange.setFontWeight('bold');
  titleRange.setHorizontalAlignment('center');
  titleRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);
  
  // -- РЯДОК 2: СИНІ ЗАГОЛОВКИ КОЛОНОК --
  var headers = ['№', 'Дата та час'];
  for (var i = 0; i < FIELD_MAP.length; i++) headers.push(FIELD_MAP[i].col);
  
  var headerRange = sheet.getRange(2, 1, 1, totalCols);
  headerRange.setValues([headers]);
  headerRange.setBackground(C.accent);
  headerRange.setFontColor(C.white);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);
  sheet.setRowHeight(2, 44);
  
  // -- ШИРИНА КОЛОНОК --
  sheet.setColumnWidth(1, 45); 
  sheet.setColumnWidth(2, 120); 
  for (var j = 0; j < FIELD_MAP.length; j++) {
    sheet.setColumnWidth(j + 3, FIELD_MAP[j].width);
  }
  
  // Заморожуємо верхні 2 рядки
  sheet.setFrozenRows(2);
}

// ─────────────────────────────────────────────────────────
// ФАРБУВАННЯ РЯДКА З ДАНИМИ КЛІЄНТА
// ─────────────────────────────────────────────────────────
function formatRow(sheet, rowNum) {
  var totalCols = FIELD_MAP.length + 2;
  var range = sheet.getRange(rowNum, 1, 1, totalCols);

  var isEven = (rowNum % 2 === 0);
  range.setBackground(isEven ? C.gray : C.white);
  range.setFontFamily('Arial');
  range.setFontSize(10);
  range.setFontColor(C.textDark);
  range.setVerticalAlignment('top');
  range.setWrap(true);
  sheet.setRowHeight(rowNum, 40);

  // №
  sheet.getRange(rowNum, 1)
    .setBackground(C.accentBg)
    .setFontColor(C.accent)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Час
  sheet.getRange(rowNum, 2)
    .setBackground(C.green)
    .setFontColor(C.greenText)
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // ПІБ (Колонка 3)
  sheet.getRange(rowNum, 3)
    .setBackground(C.pibBg)
    .setFontColor(C.pibText)
    .setFontWeight('bold')
    .setFontSize(11);

  // Instagram (Колонка 5) - Клікабельне посилання
  var igCell = sheet.getRange(rowNum, 5);
  var igVal = String(igCell.getValue() || '').trim();
  if (igVal !== '') {
    var igUrl = igVal.charAt(0) === '@' ? 'https://instagram.com/' + igVal.substring(1) : (igVal.indexOf('http') !== 0 ? 'https://instagram.com/' + igVal : igVal);
    var richText = SpreadsheetApp.newRichTextValue().setText(igVal).setLinkUrl(igUrl).build();
    igCell.setRichTextValue(richText);
    igCell.setFontColor(C.accent);
  }

  // Пусті клітинки (червоний мінус)
  for (var c = 3; c <= totalCols - 1; c++) {
    var cell = sheet.getRange(rowNum, c);
    if (String(cell.getValue() || '').trim() === '') {
      cell.setBackground(C.emptyBg);
      cell.setValue('—');
      cell.setFontColor(C.emptyText);
      cell.setHorizontalAlignment('center');
      cell.setFontStyle('italic');
    }
  }

  // Бордер
  range.setBorder(false, false, true, false, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID);
}

function sendEmail(data, ts, num) {
  try {
    var body = 'Нова заявка #' + num + ' о ' + ts + '\n\n';
    for (var i = 0; i < FIELD_MAP.length; i++) body += FIELD_MAP[i].col + ':\n' + (data[FIELD_MAP[i].key] || '—') + '\n\n';
    MailApp.sendEmail({ to: NOTIFICATION_EMAIL, subject: '📋 Заявка #' + num + ': ' + (data.pib || 'Невідомий'), body: body });
  } catch (err) { }
}
