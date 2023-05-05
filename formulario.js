function onFormSubmit1(e) {
  crearMerger()
}


function onFormSubmit(e) {
  const { idKommo, idEffi } = obtenerUltimaFilaFormRespuestas()

  const ssKommoId = converirAHojaDeCalculo(idKommo)
  const ssKommo = SpreadsheetApp.openById(ssKommoId)

  const ssEffiId = converirAHojaDeCalculo(idEffi)
  const ssEffi = SpreadsheetApp.openById(ssEffiId)

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var nameHojaOrigenKommo = 'KOMMO'
  var nameHojaOrigenEffi = 'EFFI'

  const traladarInfoKommo = copiarInfoDeHoja(ssKommo, ssKommo.getSheetName(), ss, nameHojaOrigenKommo)
  const traladarInfoEffi = copiarInfoDeHoja(ssEffi, ssEffi.getSheetName(), ss, nameHojaOrigenEffi)

  //borammos los archivos para que no se nos llene en el drive
  const archivoKomoBorrados = borraArchivo(idKommo)
  const archivoEffiBorrados = borraArchivo(idEffi)

  crearMerger()
}

function borraArchivo(id) {
  var archivo = DriveApp.getFileById(id)
  archivo.setTrashed(true);
}

function copiarInfoDeHoja(ssOrigen, nameHojaOrigen, ssDestino, nameHojaDestino) {
  //borrmos la hoja de en la que vamos ha hacer la copia
  borrarYCrear(nameHojaDestino)

  var hojaOrigen = ssOrigen.getSheetByName(nameHojaOrigen);
  var hojaDestino = ssDestino.getSheetByName(nameHojaDestino);

  var datosHojaOrigen = hojaOrigen.getDataRange().getValues();

  var filaInicial = 1; // Cambia esta línea para establecer la fila inicial donde se pegarán los datos.
  var columnaInicial = 1; // Cambia esta línea para establecer la columna inicial donde se pegarán los datos.
  hojaDestino.getRange(filaInicial, columnaInicial, datosHojaOrigen.length, datosHojaOrigen[0].length).setValues(datosHojaOrigen);
}

function converirAHojaDeCalculo(idXlsx) {
  var archivoXlsx = DriveApp.getFileById(idXlsx)
  var archivoGoogleSheets = Drive.Files.copy({}, archivoXlsx.getId(), { convert: true });
  //borramos el documeno que no queremos
  archivoXlsx.setTrashed(true);

  return archivoGoogleSheets.id
}


function obtenerUltimaFilaFormRespuestas() {
  // Obtener la hoja de respuestas del formulario
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1")
  var range = sheet.getDataRange();
  var values = range.getValues();

  var hojaKommo = values[values.length - 1][1]
  var hojaEffi = values[values.length - 1][2]

  let idKommo = hojaKommo.split('=')[1]
  let idEffi = hojaEffi.split('=')[1]

  console.log(idKommo, 'idKommo', idEffi, 'idEffi')

  return { idKommo, idEffi };
}



function copiarHojaDeCalculo(idOrigen, nombreHoja, idDestino) {
  // Obtener la hoja de cálculo de origen y destino
  var libroOrigen = SpreadsheetApp.openById(idOrigen);
  var libroDestino = SpreadsheetApp.openById(idDestino);
  var hojaOrigen = libroOrigen.getSheetByName(nombreHoja);

  // Obtener la última columna de la hoja de origen
  var ultimaColumna = hojaOrigen.getLastColumn();

  // Copiar la hoja de origen a la hoja de destino
  var hojaDestino = hojaOrigen.copyTo(libroDestino);

  // Cambiar el nombre de la hoja de destino
  var nombreDestino = hojaDestino.getName();
  hojaDestino.setName(nombreDestino + '_copia');

  // Insertar una nueva columna en la hoja de destino
  hojaDestino.insertColumnAfter(ultimaColumna);

  // Escribir en la nueva columna de la hoja de destino
  var rangoDestino = hojaDestino.getRange('I1:I');
  rangoDestino.setValue('Novedad');
}
