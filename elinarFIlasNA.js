function eliminarFilasConError() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KOMMO")
  // Cambia "Nombre de la hoja" por el nombre de la hoja que deseas modificar.
  var ultimaFila = hoja.getLastRow();
  var valoresBG = hoja.getRange("BG2:BG" + ultimaFila).getValues();
  var filasAEliminar = [];

  for (var i = 0; i < valoresBG.length; i++) {
    if (valoresBG[i][0] == "#N/A") {
      filasAEliminar.push(i + 2);
      console.log('elimini')
    }
  }

  if (filasAEliminar.length > 0) {
    var intervalos = agruparIntervalos(filasAEliminar);

    for (var i = intervalos.length - 1; i >= 0; i--) {
      var filaInicial = intervalos[i][0];
      var cantidadFilas = intervalos[i][1];
      hoja.deleteRows(filaInicial, cantidadFilas);
    }
  }
}

function agruparIntervalos(filas) {
  var intervalos = [];

  for (var i = 0; i < filas.length; i++) {
    var filaActual = filas[i];
    var cantidadFilas = 1;

    while (filas[i + cantidadFilas] == filaActual + cantidadFilas) {
      cantidadFilas++;
    }

    intervalos.push([filaActual, cantidadFilas]);
    i += cantidadFilas - 1;
  }

  return intervalos;
}
