/**
 * @file auto-buffer-rows.js
 * @description Автоматическое поддержание буфера пустых строк (по умолчанию 10) в Google Таблицах.
 * @version 1.0.0
 * 
 * Логика работы:
 * - Отслеживает редактирование через простой триггер onEdit(e).
 * - Находит последнюю заполненную строку заказа (по колонке A).
 * - Находит строку начала итогового подвала (по тексту "Исходная выручка").
 * - Если пустых строк между заказами и подвалом становится меньше 10,
 *   автоматически вставляет недостающие строки прямо перед подвалом,
 *   копируя форматы, формулы и правила валидации из строки выше.
 */

function onEdit(e) {
  const sheet = e ? e.range.getSheet() : SpreadsheetApp.getActiveSheet();

  const BUFFER_ROWS = 10; // Минимальный резерв пустых строк до конца таблицы
  const START_ROW = 2;    // Первая строка с данными (после шапки)
  const ORDER_COL = 1;    // Колонка A («Номер заказа»)

  const maxRows = sheet.getMaxRows();
  const colAValues = sheet.getRange(1, ORDER_COL, maxRows, 1).getValues();

  // 1. Поиск строки начала подвала ("Исходная выручка")
  let footerRow = 0;
  const searchValues = sheet.getRange(1, 1, maxRows, sheet.getLastColumn()).getValues();
  for (let r = START_ROW; r <= maxRows; r++) {
    const rowText = searchValues[r - 1].join(" ");
    if (rowText.indexOf("Исходная выручка") !== -1) {
      footerRow = r;
      break;
    }
  }

  // Если подвал не найден текстом, считаем концом таблицы границу листа
  if (footerRow === 0) {
    footerRow = maxRows + 1;
  }

  // 2. Поиск последней строки с заполненным заказом (выше подвала)
  let lastOrderRow = START_ROW - 1;
  for (let r = footerRow - 1; r >= START_ROW; r--) {
    const val = colAValues[r - 1][0];
    if (val !== "" && val !== null && val !== undefined) {
      lastOrderRow = r;
      break;
    }
  }

  // 3. Расчет текущего резерва свободных строк
  const currentEmptyRows = (footerRow - 1) - lastOrderRow;

  // 4. Добавление строк при нехватке
  if (currentEmptyRows < BUFFER_ROWS) {
    const rowsToAdd = BUFFER_ROWS - currentEmptyRows;

    // Вставляем строки перед блоком итогов (сдвигая подвал вниз)
    sheet.insertRowsBefore(footerRow, rowsToAdd);

    // Копируем формулы, выпадающие списки и стили из шаблона выше
    const templateRow = footerRow - 1;
    const sourceRange = sheet.getRange(templateRow, 1, 1, sheet.getLastColumn());
    const targetRange = sheet.getRange(footerRow, 1, rowsToAdd, sheet.getLastColumn());

    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    // Очищаем колонку заказа в новых строках, чтобы они считались пустыми
    sheet.getRange(footerRow, ORDER_COL, rowsToAdd, 1).clearContent();
  }
}
