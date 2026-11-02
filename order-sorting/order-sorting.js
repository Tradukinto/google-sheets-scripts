/**
 * @file order-sorting.js
 * @description Контекстный диспетчер и сортировка заказов в Google Таблицах (Google Apps Script).
 * @version 1.1.0
 * 
 * Логика маршрутизации:
 * - Если активный лист заканчивается на "_Вова" или "_Света" -> выполняется двухуровневая сортировка заказов.
 * - Если активен любой другой лист -> управление передается в runOtherSheetAction().
 */

/**
 * Единая точка входа. Назначается на горячую клавишу (Ctrl+Alt+Shift+1), кнопку или пункт меню.
 */
function handleSortOrAction() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName().trim();

  // Проверяем, заканчивается ли имя листа на _Вова или _Света (регистронезависимо)
  if (/_Вова$/i.test(sheetName) || /_Света$/i.test(sheetName)) {
    sortOrders(sheet);
  } else {
    runOtherSheetAction(sheet);
  }
}

/**
 * Двухуровневая сортировка заказов для листов *_Вова и *_Света:
 * 1. Зеленые (срочные) заказы поднимаются наверх.
 * 2. Внутри групп — сортировка по возрастанию объема (столбец B).
 * 3. Пустые строки автоматически смещаются вниз.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Активный лист.
 */
function sortOrders(sheet) {
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

  // Двухуровневая стабильная сортировка
  rows.sort((a, b) => {
    if (a.isEmpty !== b.isEmpty) return a.isEmpty ? 1 : -1;
    if (a.isEmpty && b.isEmpty) return 0;

    if (a.isGreen !== b.isGreen) return a.isGreen ? -1 : 1;
    if (a.volume !== b.volume) return a.volume - b.volume;

    return a.originalIndex - b.originalIndex;
  });

  // Пакетная запись обратно
  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.backgrounds));
  range.setFontWeights(rows.map(r => r.fontWeights));
  range.setFontColors(rows.map(r => r.fontColors));

  sheet.getParent().toast(`Отсортировано заказов: ${rows.filter(r => !r.isEmpty).length}`, sheet.getName(), 2);
}

/**
 * Функция для обработки всех остальных листов.
 * Сюда вы вставите или подключите код второго скрипта.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Активный лист.
 */
function runOtherSheetAction(sheet) {
  // TODO: Поместите сюда логику для других листов
  sheet.getParent().toast(`Действие для обычного листа "${sheet.getName()}"`, 'Информация', 3);
}

/**
 * Проверка на зеленый цвет (Excel #92D050 или палитра Google).
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
    .addItem('Запустить действие листа', 'handleSortOrAction')
    .addToUi();
}
