/**
 * @file order-sorting.js
 * @description Двухуровневая сортировка заказов в Google Таблицах (Google Apps Script).
 * @version 1.0.0
 * 
 * Логика работы:
 * 1. Зеленые (срочные/приоритетные) заказы поднимаются наверх.
 * 2. Внутри каждой группы (зеленые и обычные) заказы сортируются по возрастанию объема (столбец B).
 * 3. Пустые строки автоматически смещаются в самый низ таблицы.
 * 4. Полностью сохраняются все выпадающие списки (Data Validation), начертание шрифта и цвета ячеек.
 */

/**
 * Главная функция сортировки. Назначается на кнопку или пункт меню.
 */
function sortOrders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Если данных нет (только шапка или пустой лист) — выходим
  if (lastRow <= 1) return;

  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 1, numRows, lastCol);

  // Считываем значения, цвета заливки, начертание шрифта и цвет текста
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontWeights = range.getFontWeights();
  const fontColors = range.getFontColors();

  // Формируем список строк с метаданными для устойчивой сортировки
  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const orderNum = String(values[i][0] || '').trim();
    const volRaw = values[i][1];

    // Корректный парсинг объема (число или текстовые форматы вида "1 200", "5,5")
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
    // 1. Пустые строки всегда смещаются в конец
    if (a.isEmpty !== b.isEmpty) return a.isEmpty ? 1 : -1;
    if (a.isEmpty && b.isEmpty) return 0;

    // 2. Уровень 1: Зеленые (срочные) строки наверх
    if (a.isGreen !== b.isGreen) return a.isGreen ? -1 : 1;

    // 3. Уровень 2: По возрастанию объема
    if (a.volume !== b.volume) return a.volume - b.volume;

    // 4. Сохранение исходного порядка строк при равных параметрах
    return a.originalIndex - b.originalIndex;
  });

  // Запись отсортированных данных обратно в таблицу одним пакетом
  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.backgrounds));
  range.setFontWeights(rows.map(r => r.fontWeights));
  range.setFontColors(rows.map(r => r.fontColors));
}

/**
 * Проверка, является ли цвет ячейки зеленым.
 * Распознает стандартный цвет Excel (#92D050) и любые оттенки зеленого из палитры Google Sheets.
 * 
 * @param {string} hex - Шестнадцатеричный код цвета ячейки.
 * @returns {boolean} True, если цвет зеленый.
 */
function isGreenColor(hex) {
  if (!hex || hex === '#ffffff' || hex === '#000000' || hex === 'none') return false;
  hex = hex.replace('#', '').trim().toLowerCase();
  if (hex.length === 3) {
    hex = hex.split('').map(c => c + c).join('');
  }
  if (hex.length !== 6) return false;

  // Прямое совпадение со стандартным цветом срочных заказов из Excel
  if (hex === '92d050') return true;

  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);

  // Оттенки зеленого: зеленый канал выражен и доминирует над красным и синим
  return g > 110 && g > r * 1.08 && g > b * 1.08;
}

/**
 * Автоматическое добавление пункта в верхнее меню Google Таблицы при открытии книги.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Автоматизация')
    .addItem('Сортировать заказы', 'sortOrders')
    .addToUi();
}
