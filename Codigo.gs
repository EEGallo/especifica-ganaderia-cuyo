// Cuestionario Ganadero – CCIA SR (9-10-2026)
// Pegar este código en el proyecto de Apps Script e implementarlo como aplicación web.

// Hoja de cálculo donde se guardan las respuestas (ID tomado de su URL)
var ID_PLANILLA = '1exwaCW3Y6Md3QBOY234wKyTezC0npmkkuG0E0Iz0UOQ';
var NOMBRE_HOJA = 'Respuestas';
var ENCABEZADOS = [
  'Fecha y hora',
  'Nombre y apellido',
  'Ubicación / localidad',
  'Actividad mixta (extensiva / bajo riego)',
  '1. ¿Por qué hago lo que hago?',
  '2. Dificultades para producir',
  '3. Motivos de productividad estancada',
  '4. ¿La tecnología mejorará su productividad?',
  '5. Mesa de trabajo en la que participaría',
  '6. Otro tema a plantear',
  '7. Una cuestión a resolver en 2026/27',
  'Observaciones'
];

// Sirve el formulario (la URL /exec es la que se comparte por WhatsApp)
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Formulario')
    .setTitle('Cuestionario Ganadero – CCIA SR')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Solo para cuando el HTML se aloja fuera de Apps Script
function doPost(e) {
  guardarRespuesta(e.parameter);
  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}

function guardarRespuesta(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.openById(ID_PLANILLA);
    var hoja = ss.getSheetByName(NOMBRE_HOJA) || ss.insertSheet(NOMBRE_HOJA);
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(ENCABEZADOS);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([
      new Date(),
      limpiar(d.nombre, 120),
      limpiar(d.ubicacion, 200),
      limpiar(d.mixta, 10),
      limpiar(d.p1, 500),
      limpiar(d.p2, 500),
      limpiar(d.p3, 500),
      limpiar(d.p4, 500),
      limpiar(d.mesa, 500),
      limpiar(d.tema, 1000),
      limpiar(d.p5, 1000),
      limpiar(d.obs, 2000)
    ]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

// Recorta el texto y evita que una respuesta se interprete como fórmula
function limpiar(v, max) {
  var s = String(v || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
