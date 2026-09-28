// Pegar en Extensiones > Apps Script de la hoja donde se guardan los registros.
// Implementar > Administrar implementaciones > Editar > Nueva versión (si ya estaba publicado)
//   Ejecutar como: Yo   |   Quién tiene acceso: Cualquier persona

const HOJA = "Registros";
const PERSONAL_ID = "18ij-dLKtajitTqGTWqrPDzly8P6uDA5e4YLjS5WY9xs"; // hoja con los datos de personal
const COLUMNAS = ["Fecha de registro","Fecha","ID personal","Técnico","Programa","Laboratorio","Ambiente","NRC","Curso","Docente(s)","Opción"];

// Busca un ID (columna B) en la primera pestaña de la hoja de personal.
// C = apellido paterno, D = apellido materno, E = nombres.
function doGet(e) {
  const id = String((e.parameter && e.parameter.id) || "").trim();
  let out = { ok: false };
  if (id) {
    const sh = SpreadsheetApp.openById(PERSONAL_ID).getSheets()[0];
    const hit = sh.getRange("B:B").createTextFinder(id).matchEntireCell(true).findNext();
    if (hit) {
      const v = sh.getRange(hit.getRow(), 3, 1, 3).getValues()[0];
      const nombre = (String(v[0]) + " " + String(v[1])).replace(/\s+/g, " ").trim() + ", " + String(v[2]).trim();
      out = { ok: true, nombre: nombre.toUpperCase() };
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
  if (sh.getLastRow() === 0) sh.appendRow(COLUMNAS);
  const ahora = new Date();
  JSON.parse(e.postData.contents).forEach(r =>
    sh.appendRow([ahora, r.fecha, r.idPersonal, r.tecnico, r.programa, r.laboratorio, r.ambiente, r.nrc, r.curso, r.docentes, r.opcion]));
  return ContentService.createTextOutput("ok");
}
