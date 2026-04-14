/**
 * Google Apps Script — Преміальний бекенд для діагностичної анкети v3.0
 * 
 * ✨ Можливості:
 * - Красива стилізована таблиця з кольоровими заголовками
 * - Автоматичне форматування кожного нового рядка (зебра)
 * - Instagram автоматично стає клікабельним посиланням
 * - Авто-ширина колонок під тип контенту
 * - Захист заголовків від випадкового редагування
 * - Підрахунок кількості заявок у комірці Dashboard
 * - Email-нотифікація (опціонально)
 * - CSV/Formula injection захист
 * 
 * ІНСТРУКЦІЯ ПО ВСТАНОВЛЕННЮ:
 * ─────────────────────────────────────────────────────
 * 1. Створіть нову Google Таблицю (або відкрийте існуючу)
 * 2. Натисніть: Розширення → Apps Script
 * 3. Видаліть весь існуючий код і вставте цей скрипт
 * 4. Збережіть (Ctrl+S)
 * 5. Оберіть функцію "setupSheet" у випадаючому списку зверху → натисніть ▶ Запустити
 *    (При першому запуску потрібно надати дозвіл доступу до таблиці)
 * 6. Натисніть: Розгортання (Deploy) → Нове розгортання (New deployment)
 * 7. Тип: Веб-застосунок (Web app)
 * 8. Виконувати від імені: Мене (Me)
 * 9. Хто має доступ: Усі (Anyone) ⚠️ Обов'язково!
 * 10. Натисніть "Розгорнути" (Deploy)
 * 11. Скопіюйте URL веб-застосунку
 * 12. Вставте цей URL у файл js/script.js → змінна WEBHOOK_URL
 * ─────────────────────────────────────────────────────
 */

// ═══════════════════════════════════════════════════
//  КОНФІГУРАЦІЯ
// ═══════════════════════════════════════════════════

/**
 * Email для нотифікацій (опціонально).
 * Залиште порожнім '', щоб вимкнути email-нотифікації.
 */
const NOTIFICATION_EMAIL = '';

/** Максимальна довжина одного поля (захист від спаму) */
const MAX_FIELD_LENGTH = 5000;

/** Назва листа з відповідями */
const SHEET_NAME = '📋 Відповіді';

/** Назва листа-дашборду */
const DASHBOARD_NAME = '📊 Dashboard';

/**
 * Маппінг полів: ключ з фронтенду → назва колонки + тип + ширина.
 */
const FIELD_MAP = [
  { key: 'pib',              header: 'ПІБ',                          icon: '👤', width: 200, type: 'short' },
  { key: 'birthday',         header: 'День народження',               icon: '🎂', width: 130, type: 'short' },
  { key: 'instagram',        header: 'Instagram',                     icon: '📱', width: 220, type: 'link'  },
  { key: 'sphere_experience', header: 'Сфера та досвід',              icon: '💼', width: 280, type: 'long'  },
  { key: 'reason_purchase',  header: 'Причина придбання',             icon: '🎯', width: 280, type: 'long'  },
  { key: 'decision_point',   header: 'Кінцева точка рішення',         icon: '⚡', width: 280, type: 'long'  },
  { key: 'need_to_close',    header: 'Яку потребу закрити',           icon: '🔑', width: 280, type: 'long'  },
  { key: 'dream_goal',       header: 'Мрія / ціль',                   icon: '🌟', width: 280, type: 'long'  },
  { key: 'expectations',     header: 'Очікування',                    icon: '🎪', width: 280, type: 'long'  },
  { key: 'skills_to_learn',  header: 'Хоче вивчити',                  icon: '📚', width: 280, type: 'long'  },
  { key: 'page_difficulties', header: 'Труднощі зі сторінкою',       icon: '🔧', width: 280, type: 'long'  },
  { key: 'past_experience',  header: 'Минулий досвід',                icon: '📝', width: 280, type: 'long'  },
  { key: 'additional_notes', header: 'Додатково',                     icon: '💬', width: 250, type: 'long'  }
];

// ═══════════════════════════════════════════════════
//  КОЛЬОРОВА ПАЛІТРА
// ═══════════════════════════════════════════════════
const COLORS = {
  headerBg:      '#1a1c2e',   // Темно-синій фон заголовків
  headerText:    '#FFFFFF',
  headerAccent:  '#4C84FF',   // Accent для номерів
  rowEven:       '#f8f9ff',   // Світлий рядок (парний)
  rowOdd:        '#FFFFFF',   // Білий рядок (непарний)
  borderColor:   '#e2e4f0',   // М'який бордер
  numberBg:      '#eef2ff',   // Фон колонки №
  numberText:    '#4C84FF',   // Колір номеру
  timestampBg:   '#f0fdf4',   // Фон часу (м'який зелений)
  timestampText: '#166534',
  sectionPib:    '#dbeafe',   // Фон ПІБ (синій)
  linkText:      '#4C84FF',   // Колір посилань
  emptyCell:     '#fef2f2',   // Фон пустої клітинки
  dashboardBg:   '#1e293b',   // Дашборд
  dashboardText: '#f8fafc',
};

// ═══════════════════════════════════════════════════
//  УТИЛІТИ
// ═══════════════════════════════════════════════════

/** Санітизація: захист від CSV/Formula injection + обрізання */
function sanitize(value) {
  if (typeof value !== 'string') return '';
  let clean = value.trim();
  if (/^[=+\-@\t\r]/.test(clean)) {
    clean = "'" + clean;
  }
  if (clean.length > MAX_FIELD_LENGTH) {
    clean = clean.substring(0, MAX_FIELD_LENGTH) + '…';
  }
  return clean;
}

/** Отримати або створити лист за назвою */
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ═══════════════════════════════════════════════════
//  ОБРОБКА ЗАПИТІВ
// ═══════════════════════════════════════════════════

function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const sheet = getOrCreateSheet(SHEET_NAME);
    
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseError) {
      data = e.parameter || {};
    }

    // Валідація
    if (!data.pib || data.pib.trim() === '') {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'error', message: 'Поле ПІБ обов\'язкове' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const timestamp = Utilities.formatDate(
      new Date(), 
      Session.getScriptTimeZone(), 
      'dd.MM.yyyy HH:mm'
    );

    // Номер заявки
    const lastRow = sheet.getLastRow();
    const rowNumber = lastRow; // Заголовок = рядок 1, тому lastRow = номер заявки

    // Формуємо рядок
    const row = [rowNumber, timestamp];
    FIELD_MAP.forEach(field => {
      row.push(sanitize(data[field.key] || ''));
    });

    sheet.appendRow(row);

    // Стилізація нового рядка
    const newRow = sheet.getLastRow();
    formatDataRow(sheet, newRow);

    SpreadsheetApp.flush();
    lock.releaseLock();

    // Оновити дашборд
    updateDashboard();

    // Email-нотифікація
    if (NOTIFICATION_EMAIL) {
      sendNotificationEmail(data, timestamp, rowNumber);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', number: rowNumber, timestamp: timestamp })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('doPost error:', error.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ 
      status: 'ok', 
      message: 'Вебхук працює! POST для відправки даних.',
      version: '3.0'
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════
//  ФОРМАТУВАННЯ РЯДКА
// ═══════════════════════════════════════════════════

function formatDataRow(sheet, rowNum) {
  const totalCols = FIELD_MAP.length + 2; // +2 для №, Час
  const range = sheet.getRange(rowNum, 1, 1, totalCols);
  
  // Зебра: парний/непарний
  const isEven = (rowNum % 2 === 0);
  const bgColor = isEven ? COLORS.rowEven : COLORS.rowOdd;
  range.setBackground(bgColor);

  // Шрифт
  range.setFontFamily('Inter, Arial, sans-serif');
  range.setFontSize(10);
  range.setVerticalAlignment('top');
  range.setWrap(true);

  // № колонка — акцентний колір
  const numCell = sheet.getRange(rowNum, 1);
  numCell.setBackground(COLORS.numberBg);
  numCell.setFontColor(COLORS.numberText);
  numCell.setFontWeight('bold');
  numCell.setHorizontalAlignment('center');
  numCell.setVerticalAlignment('middle');

  // Час — м'який зелений
  const timeCell = sheet.getRange(rowNum, 2);
  timeCell.setBackground(COLORS.timestampBg);
  timeCell.setFontColor(COLORS.timestampText);
  timeCell.setFontSize(9);
  timeCell.setHorizontalAlignment('center');
  timeCell.setVerticalAlignment('middle');

  // ПІБ — виділити синім фоном
  const pibCell = sheet.getRange(rowNum, 3);
  pibCell.setBackground(COLORS.sectionPib);
  pibCell.setFontWeight('bold');

  // Instagram — зробити клікабельним посиланням
  const igCell = sheet.getRange(rowNum, 5); // колонка 5 = Instagram
  const igValue = igCell.getValue();
  if (igValue && typeof igValue === 'string') {
    let url = igValue;
    if (igValue.startsWith('@')) {
      url = 'https://instagram.com/' + igValue.substring(1);
    } else if (!igValue.startsWith('http')) {
      url = 'https://instagram.com/' + igValue;
    }
    // Створити формулу HYPERLINK
    igCell.setFormula(`=HYPERLINK("${url}", "${igValue}")`);
    igCell.setFontColor(COLORS.linkText);
  }

  // Позначити пусті клітинки (крім "Додатково" — воно не обов'язкове)
  for (let col = 3; col <= totalCols - 1; col++) { // -1 щоб пропустити "Додатково"
    const cell = sheet.getRange(rowNum, col);
    const val = cell.getValue();
    if (val === '' || val === null) {
      cell.setBackground(COLORS.emptyCell);
      cell.setValue('—');
      cell.setFontColor('#9ca3af');
      cell.setHorizontalAlignment('center');
    }
  }

  // Бордери
  range.setBorder(
    false, false, true, false, false, false,
    COLORS.borderColor, SpreadsheetApp.BorderStyle.SOLID
  );
}

// ═══════════════════════════════════════════════════
//  DASHBOARD
// ═══════════════════════════════════════════════════

function updateDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dash = ss.getSheetByName(DASHBOARD_NAME);
    if (!dash) return; // Дашборд не створений — пропускаємо

    const mainSheet = ss.getSheetByName(SHEET_NAME);
    if (!mainSheet) return;

    const totalResponses = Math.max(0, mainSheet.getLastRow() - 1);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');

    dash.getRange('B3').setValue(totalResponses);
    dash.getRange('B4').setValue(now);
  } catch(e) {
    console.error('Dashboard update error:', e);
  }
}

// ═══════════════════════════════════════════════════
//  EMAIL
// ═══════════════════════════════════════════════════

function sendNotificationEmail(data, timestamp, number) {
  try {
    const name = data.pib || 'Невідомий';
    const instagram = data.instagram || '—';
    
    const subject = `📋 Заявка #${number}: ${name}`;
    
    let body = `Нова заявка #${number} отримана о ${timestamp}\n\n`;
    body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
    
    FIELD_MAP.forEach((field, index) => {
      const value = data[field.key] || '—';
      body += `${field.icon} ${field.header}:\n${value}\n\n`;
    });
    
    body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
    body += `Таблиця: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      body: body
    });
  } catch (error) {
    console.error('Email error:', error.toString());
  }
}

// ═══════════════════════════════════════════════════
//  ПОЧАТКОВЕ НАЛАШТУВАННЯ (запустити ОДИН РАЗ)
// ═══════════════════════════════════════════════════

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ─── 1. Лист з відповідями ──────────────────────
  const sheet = getOrCreateSheet(SHEET_NAME);
  
  // Очищаємо
  sheet.clear();
  
  // Заголовки
  const headers = ['№', '⏱ Час'];
  FIELD_MAP.forEach((field, index) => {
    headers.push(`${field.icon} ${field.header}`);
  });

  const totalCols = headers.length;

  // Рядок 1: Заголовки
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setValues([headers]);

  // Стилізація заголовків
  headerRange.setBackground(COLORS.headerBg);
  headerRange.setFontColor(COLORS.headerText);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setFontFamily('Inter, Arial, sans-serif');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);

  // Висота заголовка
  sheet.setRowHeight(1, 50);

  // Закріпити заголовок
  sheet.setFrozenRows(1);

  // Ширина колонки №
  sheet.setColumnWidth(1, 50);
  // Ширина колонки Час
  sheet.setColumnWidth(2, 140);
  
  // Ширина решти колонок з FIELD_MAP
  FIELD_MAP.forEach((field, index) => {
    sheet.setColumnWidth(index + 3, field.width);
  });

  // Бордер під заголовком
  headerRange.setBorder(
    true, true, true, true, true, true,
    COLORS.headerBg, SpreadsheetApp.BorderStyle.SOLID
  );

  // Нижній бордер заголовка — акцентний
  headerRange.setBorder(
    false, false, true, false, false, false,
    COLORS.headerAccent, SpreadsheetApp.BorderStyle.SOLID_THICK
  );

  // Захист заголовка від редагування
  const protection = sheet.getRange(1, 1, 1, totalCols).protect();
  protection.setDescription('Заголовки — не редагувати');
  protection.setWarningOnly(true);

  // Встановити wrap для всіх майбутніх даних
  sheet.getRange(2, 1, 998, totalCols).setWrap(true);
  sheet.getRange(2, 1, 998, totalCols).setVerticalAlignment('top');
  sheet.getRange(2, 1, 998, totalCols).setFontFamily('Inter, Arial, sans-serif');
  sheet.getRange(2, 1, 998, totalCols).setFontSize(10);

  // ─── 2. Dashboard ──────────────────────────────
  const dash = getOrCreateSheet(DASHBOARD_NAME);
  dash.clear();

  // Фон дашборда
  dash.getRange(1, 1, 20, 5).setBackground(COLORS.dashboardBg);

  // Заголовок
  const dashTitle = dash.getRange('A1:E1');
  dashTitle.merge();
  dashTitle.setValue('📊 DASHBOARD — Діагностична анкета');
  dashTitle.setFontSize(16);
  dashTitle.setFontWeight('bold');
  dashTitle.setFontColor(COLORS.dashboardText);
  dashTitle.setHorizontalAlignment('center');
  dashTitle.setVerticalAlignment('middle');
  dash.setRowHeight(1, 56);

  // Метрики
  const labelStyle = {
    fontSize: 11,
    fontColor: '#94a3b8',
    fontWeight: 'normal'
  };
  const valueStyle = {
    fontSize: 28,
    fontColor: '#FFFFFF',
    fontWeight: 'bold'
  };

  // "Всього заявок"
  const labelCell1 = dash.getRange('A3');
  labelCell1.setValue('📝 Всього заявок:');
  labelCell1.setFontSize(labelStyle.fontSize);
  labelCell1.setFontColor(labelStyle.fontColor);

  const valueCell1 = dash.getRange('B3');
  valueCell1.setValue(0);
  valueCell1.setFontSize(valueStyle.fontSize);
  valueCell1.setFontColor('#4ade80'); // Зелений
  valueCell1.setFontWeight(valueStyle.fontWeight);
  valueCell1.setHorizontalAlignment('left');

  // "Остання заявка"
  const labelCell2 = dash.getRange('A4');
  labelCell2.setValue('🕐 Остання заявка:');
  labelCell2.setFontSize(labelStyle.fontSize);
  labelCell2.setFontColor(labelStyle.fontColor);

  const valueCell2 = dash.getRange('B4');
  valueCell2.setValue('—');
  valueCell2.setFontSize(13);
  valueCell2.setFontColor(COLORS.dashboardText);

  // Ширина колонок дашборда
  dash.setColumnWidth(1, 200);
  dash.setColumnWidth(2, 250);

  // Роз'яснення
  const noteCell = dash.getRange('A6:E6');
  noteCell.merge();
  noteCell.setValue('⬆️ Дані оновлюються автоматично при кожній новій заявці');
  noteCell.setFontSize(10);
  noteCell.setFontColor('#64748b');
  noteCell.setFontStyle('italic');

  // ─── 3. Видалити стандартний "Аркуш 1" ─────────
  const defaultSheet = ss.getSheetByName('Аркуш1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Лист1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch(e) {
      // Якщо не вдалося видалити — не страшно
    }
  }

  // Зробити лист відповідей активним
  ss.setActiveSheet(sheet);

  SpreadsheetApp.flush();
  Logger.log('✅ Таблицю налаштовано! Тепер зробіть Deploy → New deployment.');
  Logger.log('📊 Dashboard створено на окремому листі.');
}
