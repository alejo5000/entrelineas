// Contador de reacciones de Entre Líneas — Google Apps Script
// 1) Creá una Google Sheet con una hoja llamada "votos" y encabezados: fecha | nota | reaccion | delta
// 2) Extensiones > Apps Script > pegá este código.
// 3) Implementar > Nueva implementación > Aplicación web > Ejecutar como: yo / Acceso: cualquier persona.
// 4) Copiá la URL resultante en la constante ENDPOINT del HTML de la revista.

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var validas = ['orgullo', 'sirvio', 'mas'];
    if (/^n\d+$/.test(d.nota) && validas.indexOf(d.reaccion) > -1 && (d.delta === 1 || d.delta === -1)) {
      SpreadsheetApp.getActive().getSheetByName('votos')
        .appendRow([new Date(), d.nota, d.reaccion, d.delta]);
    }
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput('ok');
}

function doGet() {
  var filas = SpreadsheetApp.getActive().getSheetByName('votos').getDataRange().getValues();
  var res = {};
  for (var i = 1; i < filas.length; i++) {
    var nota = filas[i][1], r = filas[i][2], delta = Number(filas[i][3]);
    res[nota] = res[nota] || {};
    res[nota][r] = Math.max(0, (res[nota][r] || 0) + delta);
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}
