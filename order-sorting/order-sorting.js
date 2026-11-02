/**
 * @file order-sorting.js
 * @description Контекстный диспетчер сортировки в Google Таблицах (Google Apps Script).
 * @version 1.2.1
 * 
 * Логика маршрутизации:
 * 1. Листы *_Вова и *_Света:
 *    - Двухуровневая сортировка: срочные (зеленые) наверх, внутри — по возрастанию объема.
 * 2. Все остальные листы (общие, сводные, архивные, например 10.26, 11.26):
 *    - Сортировка по возрастанию столбца «Дата и время» (или «Дата и время сдачи»).
 *    - Пустые строки смещаются в самый низ.
 */

/**
 * Единая точка входа. Назначается на горячую клавишу (Ctrl+Alt+Shift+1), кнопку или пункт меню.
 */
function handleSortOrAction() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName().trim();

  // Проверяем окончание имени листа на _Вова или _Света (регистронезависимо)
  if (/_Вова$/i.test(sheetName) || /_Света$/i.test(sheetName)) {
    sortOrdersByUrgencyAndVolume(sheet);
  } else {
    sortByDateTimeAscending(sheet);
  }
}

/**
 * Алиас обратной совместимости для ранее настроенных макросов.
 */
function sortOrders() {
  handleSortOrAction();
}

/**
 * Режим 1 (*_Вова, *_Света):
 * - Зеленые (срочные) наверх.
 * - Внутри каждой группы — сортировка по возрастанию объема (столбец B).
 * - Пустые строки вниз.
 */
function sortOrdersByUrgencyAndVolume(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;

  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 1, numRows, lastCol);

  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontWeights = range.getFontWeights();
  const fontColors = range.getFontColors();

  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const orderNum = String(values[i][0] || '').trim();
    const volRaw = values[i][1];

    let volume = 0;
    if (typeof volRaw === 'number') {
      volume = volRaw;
    } else if (volRaw) {
      const cleanVol = String(volRaw).replace(/\s+/g, '').replace(',', '.');
      const parsed = parseFloat(cleanVol);
      volume = isNaN(parsed) ? 0 : parsed;
    }

    rows.push({
      values: values[i],
      backgrounds: backgrounds[i],
      fontWeights: fontWeights[i],
      fontColors: fontColors[i],
      isEmpty: orderNum === '',
      isGreen: isGreenColor(backgrounds[i][0]),
      volume: volume,
      originalIndex: i
    });
  }

  rows.sort((a, b) => {
    if (a.isEmpty !== b.isEmpty) return a.isEmpty ? 1 : -1;
    if (a.isEmpty && b.isEmpty) return 0;

    if (a.isGreen !== b.isGreen) return a.isGreen ? -1 : 1;
    if (a.volume !== b.volume) return a.volume - b.volume;

    return a.originalIndex - b.originalIndex;
  });

  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.backgrounds));
  range.setFontWeights(rows.map(r => r.fontWeights));
  range.setFontColors(rows.map(r => r.fontColors));

  sheet.getParent().toast(`[${sheet.getName()}] Отсортировано по срочности и объему`, 'Готово', 2);
}

/**
 * Режим 2 (Все остальные листы: 10.26, 11.26 и др.):
 * - Автопоиск столбца «Дата и время» / «Дата и время сдачи» в строке заголовков (по умолчанию столбец C).
 * - Сортировка по возрастанию даты и времени.
 * - Пустые строки гарантированно смещаются в самый низ.
 */
function sortByDateTimeAscending(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;

  // Динамически ищем индекс столбца с датой и временем в шапке (строка 1)
  const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let dateColIndex = -1;

  for (let c = 0; c < headerValues.length; c++) {
    const header = String(headerValues[c] || '').trim().toLowerCase();
    if (header.includes('дата и время') || header.includes('сдачи') || header.includes('deadline')) {
      dateColIndex = c;
      break;
    }
  }

  // Если точного заголовка не найдено, берем стандартный столбец C (индекс 2)
  if (dateColIndex === -1) {
    dateColIndex = 2; // C
  }

  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 1, numRows, lastCol);

  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontWeights = range.getFontWeights();
  const fontColors = range.getFontColors();

  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const orderNum = String(values[i][0] || '').trim();
    const dateRaw = values[i][dateColIndex];

    let timestamp = 0;
    let hasDate = false;

    if (dateRaw instanceof Date && !isNaN(dateRaw.getTime())) {
      timestamp = dateRaw.getTime();
      hasDate = true;
    } else if (dateRaw) {
      const parsed = parseDateString(String(dateRaw).trim());
      if (parsed) {
        timestamp = parsed.getTime();
        hasDate = true;
      }
    }

    rows.push({
      values: values[i],
      backgrounds: backgrounds[i],
      fontWeights: fontWeights[i],
      fontColors: fontColors[i],
      isEmpty: orderNum === '' && !hasDate,
      hasDate: hasDate,
      timestamp: timestamp,
      originalIndex: i
    });
  }

  rows.sort((a, b) => {
    // 1. Пустые строки всегда в самый низ
    if (a.isEmpty !== b.isEmpty) return a.isEmpty ? 1 : -1;
    if (a.isEmpty && b.isEmpty) return 0;

    // 2. Строки без даты смещаются под строки с датой
    if (a.hasDate !== b.hasDate) return a.hasDate ? -1 : 1;

    // 3. Сортировка по возрастанию даты и времени
    if (a.timestamp !== b.timestamp) return a.timestamp - b.timestamp;

    // 4. Сохранение относительного порядка при равных датах
    return a.originalIndex - b.originalIndex;
  });

  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.backgrounds));
  range.setFontWeights(rows.map(r => r.fontWeights));
  range.setFontColors(rows.map(r => r.fontColors));

  sheet.getParent().toast(`[${sheet.getName()}] Отсортировано по дате и времени`, 'Готово', 2);
}

/**
 * Вспомогательный парсер строковых дат вида "ДД.ММ.ГГГГ ЧЧ:ММ" или "ДД.ММ.ГГГГ"
 */
function parseDateString(str) {
  const match = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (match) {
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1;
    const year = parseInt(match[3], 10);
    const hours = match[4] ? parseInt(match[4], 10) : 0;
    const minutes = match[5] ? parseInt(match[5], 10) : 0;
    return new Date(year, month, day, hours, minutes);
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Проверка на зеленый цвет ячейки (Excel #92D050 или палитра Google).
 */
function isGreenColor(hex) {
  if (!hex || hex === '#ffffff' || hex === '#000000' || hex === 'none') return false;
  hex = hex.replace('#', '').trim().toLowerCase();
  if (hex.length === 3) {
    hex = hex.split('').map(c => c + c).join('');
  }
  if (hex.length !== 6) return false;

  if (hex === '92d050') return true;

  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);

  return g > 110 && g > r * 1.08 && g > b * 1.08;
}

/**
 * Добавление пункта меню при открытии таблицы.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Автоматизация')
    .addItem('Сортировать (авто-режим листа)', 'handleSortOrAction')
    .addToUi();
}
