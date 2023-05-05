function actualizarColumnaG() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KOMMO")
  var ultimaFila = hoja.getLastRow();
  var columnaError = hoja.getRange("BF1:BG" + ultimaFila).getValues();
  var columnaG = hoja.getRange("G1:G" + ultimaFila);
  var valoresColumnaG = columnaG.getValues();

  for (var i = 0; i < valoresColumnaG.length; i++) {
    if (columnaError[i][0] == "#N/A") {
    } else {
      valoresColumnaG[i][0] = "NOVEDAD";

    }
  }

  columnaG.setValues(valoresColumnaG);
}


function obtenerLetraUltimaColumna() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KOMMO")
  var ultimaColumna = hoja.getLastColumn();
  console.log('ultimaColumna', ultimaColumna)
  Logger.log(columnToLetter(ultimaColumna)); // muestra la letra de la última columna en el registro

}

function actualizarEstatusLead() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KOMMO")
  var ultimaFila = hoja.getLastRow();
  var ultimaColumna = hoja.getLastColumn();
  var culumnaLetra = columnToLetter(ultimaColumna)
  var range = hoja.getRange(`A1:${culumnaLetra}${ultimaFila}`);
  var data = range.getValues();

  // Obtener índice de la columna Merge # GUIA
  var mergeColumnIndex = allarColumna(data, "Merge # GUIA");

  // Obtener índice de la columna Estatus del lead
  var estatusColumnIndex = allarColumna(data, "Estatus del lead");

  // Verificar si se encontró la columna Merge # GUIA
  if (mergeColumnIndex === -1) {
    Logger.log("No se encontró la columna Merge # GUIA");
    return;
  }

  // Verificar si se encontró la columna Estatus del lead
  if (estatusColumnIndex === -1) {
    Logger.log("No se encontró la columna Estatus del lead");
    return;
  }

  // Recorrer filas y actualizar Estatus del lead si Merge # GUIA tiene un valor
  for (var i = 1; i < data.length; i++) {
    var mergeValue = data[i][mergeColumnIndex];
    var estatusValue = data[i][estatusColumnIndex];

    if (mergeValue !== "" && estatusValue !== "NOVEDAD") {
      hoja.getRange(i + 1, estatusColumnIndex + 1).setValue("NOVEDAD");
    }
  }
}
function allarColumna(data, nombreColumna) {
  return data[0].indexOf(nombreColumna);
}

