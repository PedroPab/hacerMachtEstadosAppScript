function encontrarColumnaTelefono() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KOMMO")
  var filaEncabezados = 1;
  var rangoEncabezados = hoja.getRange(filaEncabezados, 1, 1, hoja.getLastColumn());
  var buscadorTexto = rangoEncabezados.createTextFinder("Teléfono celular (contacto)");
  var columnaTelefono = buscadorTexto.findNext().getColumn();
  Logger.log("La columna del teléfono es: " + columnaTelefono);
  return columnaTelefono
}


function parsearTelefono() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KOMMO")
  var range = sheet.getDataRange();
  var values = range.getValues();

  // Encuentra el índice de la columna "AN"
  var telefonoColumnIndex = encontrarColumnaTelefono()

  // Añade una nueva columna después de la columna "AN"
  var nuevaColumnaIndex = telefonoColumnIndex + 1
  sheet.insertColumnAfter(telefonoColumnIndex);
  sheet.getRange(1, nuevaColumnaIndex).setValue(`Telefono Parseado ${nuevaColumnaIndex}`);

  // Itera por cada celda de la columna "AN"
  for (var row = 1; row < values.length; row++) {
    var telefono = values[row][telefonoColumnIndex - 1];
    // Verifica si el valor de la celda tiene el formato '+57...' o '+...'
    //quitamos el '+
    if (typeof (telefono) == 'string'){
      const hoal = telefono.split('+57')
      telefono = hoal[1]
    }
    //si trae el 57 se lo   quitamos
    sheet.getRange(row + 1, nuevaColumnaIndex).setValue(telefono);
  }
  console.log('escribimo en la culumna')

  return nuevaColumnaIndex
}

