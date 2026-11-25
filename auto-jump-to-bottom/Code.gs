function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const targetSheets = sheets.slice(-3);
  const currentActive = ss.getActiveSheet();

  for (const sheet of targetSheets) {
    const values = sheet.getRange("A:A").getValues();
    let lastRow = 0;
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][0] !== "") {
        lastRow = i + 1;
        break;
      }
    }
    sheet.getRange(lastRow + 1, 1).activate();
  }

  if (targetSheets.some(s => s.getSheetId() === currentActive.getSheetId())) {
    currentActive.activate();
  }
}
