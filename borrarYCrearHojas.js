function borrarHojaDeCalculo(hojaName) {
  var nombreHoja = hojaName; // Cambia esto por el nombre de la hoja que quieres eliminar
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(hoja);
}

function crearNuevaHojaDeCalculo(hojaName) {
  var nombreHoja = hojaName;
  SpreadsheetApp.getActiveSpreadsheet().insertSheet(nombreHoja);
}

function borrarYCrear(hojaName) {
  borrarHojaDeCalculo(hojaName)
  crearNuevaHojaDeCalculo(hojaName)
}


function borrarHojas() {
  borrarYCrear('KOMMO')
  borrarYCrear('QUERY')
  borrarYCrear('EFFI')
}