function generarQuery() {
  const sheetName = "KOMMO";
  const targetSheetName = "QUERY";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName)
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  // Formula to generate
  const formula = `=QUERY(${sheetName}!A1:${columnToLetter(lastColumn)}${lastRow}, "SELECT * WHERE ${columnToLetter(lastColumn)} is not null",1)`;

  // Write formula to target sheet
  targetSheet.getRange("A1").setValue(formula);
}