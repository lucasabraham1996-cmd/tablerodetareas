const SYNC_VERSION = "2026-09-29.3";
const SPREADSHEET_ID = "1cl-A3qyT9llxfibZmqXkO__mO-DpnTBdFponeajseSM";
const TASK_SHEET = "Tareas";
const SYNC_SHEET = "Sincronización";
const HEADERS = [
  "Tarea","Categoría","Tipo de acción","Responsable","Fecha límite","Estado",
  "Subestado","Prioridad","Progreso subtareas","Subtareas","Periodicidad","Descripción"
];

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const rows = Array.isArray(payload.rows) ? payload.rows : [];

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(TASK_SHEET);
    if (!sheet) throw new Error("No existe la hoja Tareas");

    const neededRows = rows.length + 1;
    if (neededRows > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), neededRows - sheet.getMaxRows());
    }

    const existingRows = Math.max(sheet.getLastRow() - 1, 0);
    if (existingRows > 0) {
      sheet.getRange(2, 1, existingRows, HEADERS.length).clearContent();
    }

    if (rows.length) {
      const values = rows.map(row => HEADERS.map(header => row[header] ?? ""));
      sheet.getRange(2, 1, values.length, HEADERS.length).setValues(values);
    }

    const syncSheet = ss.getSheetByName(SYNC_SHEET);
    if (syncSheet) {
      syncSheet.getRange("B7").setValue(new Date());
      syncSheet.getRange("B8").setValue(rows.length);
      syncSheet.getRange("B9").setValue("OK · v" + SYNC_VERSION);
    }

    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({ok:true, count:rows.length}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const syncSheet = ss.getSheetByName(SYNC_SHEET);
    if (syncSheet && e && e.parameter && e.parameter.ping === "1") {
      syncSheet.getRange("B9").setValue("ENDPOINT ACTIVO · v" + SYNC_VERSION);
      syncSheet.getRange("B7").setValue(new Date());
      SpreadsheetApp.flush();
    }
  } catch (err) {}

  return ContentService
    .createTextOutput("Sincronización activa · v" + SYNC_VERSION + ". Si ves este mensaje, la URL /exec funciona.")
    .setMimeType(ContentService.MimeType.TEXT);
}
