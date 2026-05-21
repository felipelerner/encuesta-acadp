// ══════════════════════════════════════════════════════════════
//  ACADP · Relevamiento de Áreas — Google Apps Script Backend
//  Pegá este código completo en script.google.com
// ══════════════════════════════════════════════════════════════

// ── Configuración ──────────────────────────────────────────────
const SHEET_NAME = 'Respuestas';
const SECRET_KEY = 'acadp2025';   // ← Cambiá esta clave (usala en el panel admin)
// ──────────────────────────────────────────────────────────────

const HEADERS = [
  'Timestamp', 'Área', 'Rol', 'Responsable', 'Anónimo',
  'Carga', 'Tareas (JSON)', 'Tareas manuales',
  'Dependencias', 'Mejoras dependencias',
  'Trabas', 'Mejoras trabas',
  'Picos de carga', 'Cambios propuestos', 'Cómo implementarlo',
];

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // Formato de encabezados
    const header = sheet.getRange(1, 1, 1, HEADERS.length);
    header.setFontWeight('bold');
    header.setBackground('#8B1A5E');
    header.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    [2,3,4].forEach(c => sheet.setColumnWidth(c, 140));
    [7,8,9,10,11,12,13,14,15].forEach(c => sheet.setColumnWidth(c, 260));
  }
  return sheet;
}

// ── Recibe respuestas del formulario ──────────────────────────
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      data.timestamp      || new Date().toISOString(),
      data.area           || '',
      data.rol            || '',
      data.responsable    || '',
      data.anonimo        || false,
      data.carga          || '',
      data.tareas         || '',
      data.tareasManuales || '',
      data.dependencias   || '',
      data.mejorasDep     || '',
      data.trabas         || '',
      data.mejorasTrabas  || '',
      data.picos          || '',
      data.cambios        || '',
      data.implementacion || '',
    ]);

    return respond({ ok: true });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

// ── Devuelve todas las respuestas al panel admin ───────────────
function doGet(e) {
  if (!e.parameter.key || e.parameter.key !== SECRET_KEY) {
    return respond({ error: 'unauthorized' });
  }

  try {
    const sheet  = getOrCreateSheet();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return respond({ rows: [] });

    const headers = values[0];
    const fieldMap = {
      'Timestamp':           'timestamp',
      'Área':                'area',
      'Rol':                 'rol',
      'Responsable':         'responsable',
      'Anónimo':             'anonimo',
      'Carga':               'carga',
      'Tareas (JSON)':       'tareas',
      'Tareas manuales':     'tareasManuales',
      'Dependencias':        'dependencias',
      'Mejoras dependencias':'mejorasDep',
      'Trabas':              'trabas',
      'Mejoras trabas':      'mejorasTrabas',
      'Picos de carga':      'picos',
      'Cambios propuestos':  'cambios',
      'Cómo implementarlo':  'implementacion',
    };

    const rows = values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const key = fieldMap[h] || h;
        obj[key] = row[i];
      });
      return obj;
    });

    return respond({ rows });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
