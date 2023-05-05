function agregarColumna(columnaTelefono) {
  var hojaPrincipal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KOMMO") // cambia "Hoja Principal" por el nombre de tu hoja principal
  var ultimaColumna = hojaPrincipal.getLastColumn();
  var nuevaColumna = ultimaColumna + 1;

  hojaPrincipal.getRange(1, nuevaColumna).setValue("Merge # GUIA"); // cambia "Nueva Columna" por el nombre que desees para la nueva columna
  columnaTelefono = columnToLetter(columnaTelefono)
  var ultimaFila = hojaPrincipal.getLastRow();
  var rangoFormula = hojaPrincipal.getRange(2, nuevaColumna, ultimaFila - 1, 1);
  var formula = `=VLOOKUP(${columnaTelefono}2 ,EFFI!E:F,1,FALSE)`; // cambia "Effi" por el nombre de tu hoja de búsqueda y "E:F" por las columnas que deseas buscar

  rangoFormula.setFormula(formula);
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  console.log(letter, 'letter', column)

  return letter;
}

