// ============================================================================
// APPS SCRIPT — Hub de Recursos Humanos
// Reemplaza todo el contenido de tu archivo Code.gs con esto
// ============================================================================

// Nombres de sheets (cambia si usas otros)
const SHEET_NAMES = {
  HC: 'HC',
  BAJAS: 'Bajas',
  CONFIG: 'AppConfig',
  BITACORA: 'Bitacora',
  PENDIENTES: 'Pendientes',
  EVENTOS: 'Eventos'
};

// Campos canónicos que maneja la app (deben coincidir con FIELD_SYNONYMS del front)
const CANONICAL_FIELDS = ['id', 'nombre', 'fechaIngreso', 'puesto', 'imss', 'salarioDiario'];

// ============================================================================
// Helpers
// ============================================================================
function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    hc:     getOrCreateSheet(ss, SHEET_NAMES.HC,     CANONICAL_FIELDS),
    bajas:  getOrCreateSheet(ss, SHEET_NAMES.BAJAS,  ['id', 'nombre', 'fechaBaja', 'puesto', 'fechaIngreso']),
    config: getOrCreateSheet(ss, SHEET_NAMES.CONFIG, ['key', 'value'])
  };
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// GET — lectura de datos desde cualquier dispositivo
// ============================================================================
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';

    if (action === 'getHC')    return jsonResponse(getHCData());
    if (action === 'getBajas') return jsonResponse(getBajasData());
    if (action === 'ping')     return jsonResponse({ success: true, message: 'OK', time: new Date() });

    // ---- Módulo Bitácora y Agenda ----
    if (action === 'getBitacora')   return jsonResponse(getBitacoraData(e.parameter.fecha));
    if (action === 'getPendientes') return jsonResponse(getPendientesData());
    if (action === 'getEventos')    return jsonResponse(getEventosData());
    if (action === 'getAgenda') {
      return jsonResponse({
        success: true,
        bitacora: getBitacoraData(e.parameter.fecha).data,
        pendientes: getPendientesData().data,
        eventos: getEventosData().data
      });
    }

    return jsonResponse({ success: false, error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function getHCData() {
  const sheets = getSheets();
  const hcSheet = sheets.hc;
  const lastRow = hcSheet.getLastRow();

  if (lastRow < 2) {
    return { success: true, data: [], lastUpdate: getConfigValue('HC_LastUpdate') };
  }

  const range   = hcSheet.getRange(1, 1, lastRow, CANONICAL_FIELDS.length);
  const values  = range.getValues();
  const headers = values[0];
  const data    = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  return { success: true, data: data, lastUpdate: getConfigValue('HC_LastUpdate') };
}

function getBajasData() {
  const sheets  = getSheets();
  const bSheet  = sheets.bajas;
  const lastRow = bSheet.getLastRow();

  if (lastRow < 2) return { success: true, data: [] };

  const values  = bSheet.getRange(1, 1, lastRow, bSheet.getLastColumn()).getValues();
  const headers = values[0];
  const data    = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  return { success: true, data: data };
}

function getConfigValue(key) {
  const sheets = getSheets();
  const data   = sheets.config.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function setConfigValue(key, value) {
  const sheets = getSheets();
  const data   = sheets.config.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheets.config.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheets.config.appendRow([key, value]);
}

// ============================================================================
// POST — enrutador de acciones
// ============================================================================
function doPost(e) {
  try {
    let postData = {};
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      postData = e.parameter;
    }

    const action = postData.action;
    let response;

    if (action === 'uploadHC') {
      response = uploadHCData(postData);
    } else if (action === 'uploadBajas') {
      response = uploadBajasData(postData);

    // ---- Módulo Bitácora y Agenda ----
    } else if (action === 'addBitacora') {
      response = addBitacoraEntry(postData.payload || {});
    } else if (action === 'deleteBitacora') {
      response = deleteAgendaRow(SHEET_NAMES.BITACORA, (postData.payload || {}).id);
    } else if (action === 'addPendiente') {
      response = addPendienteEntry(postData.payload || {});
    } else if (action === 'togglePendiente') {
      response = togglePendienteEntry((postData.payload || {}).id, (postData.payload || {}).completado);
    } else if (action === 'deletePendiente') {
      response = deleteAgendaRow(SHEET_NAMES.PENDIENTES, (postData.payload || {}).id);
    } else if (action === 'addEvento') {
      response = addEventoEntry(postData.payload || {});
    } else if (action === 'deleteEvento') {
      response = deleteAgendaRow(SHEET_NAMES.EVENTOS, (postData.payload || {}).id);

    } else {
      response = { success: false, error: 'Acción no reconocida: ' + action };
    }

    return jsonResponse(response);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================================
// Cargar HC nuevo + detectar bajas automáticas
// ============================================================================
function uploadHCData(postData) {
  try {
    const hcData = postData.data || [];

    if (!Array.isArray(hcData) || hcData.length === 0) {
      return { success: false, error: 'No hay datos para cargar' };
    }

    const sheets  = getSheets();
    const hcSheet = sheets.hc;

    // 1. Leer HC anterior para detectar bajas
    const oldData    = hcSheet.getDataRange().getValues();
    const oldHeaders = oldData[0] || CANONICAL_FIELDS;
    const idIdx      = oldHeaders.indexOf('id');
    const oldRecords = oldData.slice(1).map(row => {
      const obj = {};
      oldHeaders.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
    const oldIds = oldRecords.map(r => String(r.id || '').trim()).filter(Boolean);

    // 2. Limpiar sheet (conservar headers)
    if (hcSheet.getLastRow() > 1) {
      hcSheet.getRange(2, 1, hcSheet.getLastRow() - 1, hcSheet.getLastColumn()).clearContent();
    }

    // 3. Escribir nuevos datos en orden canónico
    const rows = hcData.map(item => CANONICAL_FIELDS.map(f => item[f] !== undefined ? item[f] : ''));
    if (rows.length > 0) {
      hcSheet.getRange(2, 1, rows.length, CANONICAL_FIELDS.length).setValues(rows);
    }

    // 4. Detectar bajas (IDs que estaban antes y ya no están)
    const newIds   = hcData.map(item => String(item.id || '').trim());
    const bajasIds = oldIds.filter(id => id && newIds.indexOf(id) === -1);

    if (bajasIds.length > 0) {
      const bajasSheet = sheets.bajas;
      const fechaBaja  = new Date();
      bajasIds.forEach(id => {
        const oldRec = oldRecords.find(r => String(r.id || '').trim() === id);
        if (oldRec) {
          bajasSheet.appendRow([
            oldRec.id          || '',
            oldRec.nombre      || '',
            fechaBaja,
            oldRec.puesto      || '',
            oldRec.fechaIngreso || ''
          ]);
        }
      });
    }

    // 5. Actualizar config
    setConfigValue('HC_LastUpdate',   new Date());
    setConfigValue('HC_RecordCount',  hcData.length);

    return {
      success:            true,
      message:            'HC actualizado correctamente',
      bajasDetectadas:    bajasIds.length,
      registrosCargados:  hcData.length
    };
  } catch (err) {
    return { success: false, error: 'Error al cargar HC: ' + err.toString() };
  }
}

// ============================================================================
// Cargar Bajas desde el frontend (reemplaza el contenido de la sheet Bajas)
// ============================================================================
function uploadBajasData(postData) {
  try {
    const bajasData   = postData.data || [];
    const BAJAS_FIELDS = ['id', 'nombre', 'fechaBaja', 'puesto', 'fechaIngreso'];

    if (!Array.isArray(bajasData) || bajasData.length === 0) {
      return { success: false, error: 'No hay datos de bajas para cargar' };
    }

    const sheets = getSheets();
    const bSheet = sheets.bajas;

    // Asegurar headers
    if (bSheet.getLastRow() === 0) {
      bSheet.getRange(1, 1, 1, BAJAS_FIELDS.length).setValues([BAJAS_FIELDS]).setFontWeight('bold');
      bSheet.setFrozenRows(1);
    }

    // Limpiar datos anteriores (conservar headers)
    if (bSheet.getLastRow() > 1) {
      bSheet.getRange(2, 1, bSheet.getLastRow() - 1, BAJAS_FIELDS.length).clearContent();
    }

    // Escribir nuevos datos en orden canónico
    const rows = bajasData.map(item => BAJAS_FIELDS.map(f => item[f] !== undefined ? item[f] : ''));
    if (rows.length > 0) {
      bSheet.getRange(2, 1, rows.length, BAJAS_FIELDS.length).setValues(rows);
    }

    setConfigValue('Bajas_LastUpdate',  new Date());
    setConfigValue('Bajas_RecordCount', bajasData.length);

    return {
      success:           true,
      message:           'Bajas actualizadas correctamente',
      registrosCargados: bajasData.length
    };
  } catch (err) {
    return { success: false, error: 'Error al cargar bajas: ' + err.toString() };
  }
}

// ============================================================================
// Función de prueba manual (ejecutar desde el editor de Apps Script)
// ============================================================================
function testSetup() {
  const sheets = getSheets();
  Logger.log('HC sheet: '     + sheets.hc.getName());
  Logger.log('Bajas sheet: '  + sheets.bajas.getName());
  Logger.log('Config sheet: ' + sheets.config.getName());
  Logger.log('Sheets creados/verificados correctamente');
}

// ============================================================================
// MÓDULO BITÁCORA Y AGENDA
// ============================================================================

function getAgendaSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    bitacora:   getOrCreateSheet(ss, SHEET_NAMES.BITACORA,   ['id', 'fecha', 'hora', 'categoria', 'descripcion', 'nota']),
    pendientes: getOrCreateSheet(ss, SHEET_NAMES.PENDIENTES, ['id', 'texto', 'fecha', 'completado', 'fecha_creacion']),
    eventos:    getOrCreateSheet(ss, SHEET_NAMES.EVENTOS,    ['id', 'fecha', 'titulo', 'tipo', 'notas'])
  };
}

function agendaSheetToObjects(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values  = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function formatAgendaDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function agendaFindRow(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // +2: 1-indexed + header row
  }
  return -1;
}

// ---- Bitácora ----
function getBitacoraData(fecha) {
  const sheets = getAgendaSheets();
  let rows = agendaSheetToObjects(sheets.bitacora).map(r => {
    r.fecha = formatAgendaDate(r.fecha);
    return r;
  });
  if (fecha) rows = rows.filter(r => r.fecha === fecha);
  rows.sort((a, b) => (a.hora || '').localeCompare(b.hora || ''));
  return { success: true, data: rows };
}

function addBitacoraEntry(p) {
  try {
    const sheets = getAgendaSheets();
    const id = Utilities.getUuid();
    sheets.bitacora.appendRow([id, p.fecha, p.hora, p.categoria, p.descripcion, p.nota || '']);
    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: 'Error al guardar bitácora: ' + err.toString() };
  }
}

// ---- Pendientes ----
function getPendientesData() {
  const sheets = getAgendaSheets();
  const rows = agendaSheetToObjects(sheets.pendientes).map(r => {
    r.fecha = formatAgendaDate(r.fecha);
    r.fecha_creacion = formatAgendaDate(r.fecha_creacion);
    r.completado = (r.completado === true || r.completado === 'true' || r.completado === 'TRUE');
    return r;
  });
  return { success: true, data: rows };
}

function addPendienteEntry(p) {
  try {
    const sheets = getAgendaSheets();
    const id = Utilities.getUuid();
    sheets.pendientes.appendRow([id, p.texto, p.fecha || '', false, formatAgendaDate(new Date())]);
    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: 'Error al guardar pendiente: ' + err.toString() };
  }
}

function togglePendienteEntry(id, completado) {
  try {
    const sheets = getAgendaSheets();
    const row = agendaFindRow(sheets.pendientes, id);
    if (row === -1) return { success: false, error: 'Pendiente no encontrado: ' + id };
    const headers = sheets.pendientes.getRange(1, 1, 1, sheets.pendientes.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf('completado');
    sheets.pendientes.getRange(row, colIndex + 1).setValue(completado);
    return { success: true, id: id, completado: completado };
  } catch (err) {
    return { success: false, error: 'Error al actualizar pendiente: ' + err.toString() };
  }
}

// ---- Eventos ----
function getEventosData() {
  const sheets = getAgendaSheets();
  const rows = agendaSheetToObjects(sheets.eventos).map(r => {
    r.fecha = formatAgendaDate(r.fecha);
    return r;
  });
  return { success: true, data: rows };
}

function addEventoEntry(p) {
  try {
    const sheets = getAgendaSheets();
    const id = Utilities.getUuid();
    sheets.eventos.appendRow([id, p.fecha, p.titulo, p.tipo || 'junta', p.notas || '']);
    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: 'Error al guardar evento: ' + err.toString() };
  }
}

// ---- Borrado genérico (bitácora / pendientes / eventos) ----
function deleteAgendaRow(sheetName, id) {
  try {
    if (!id) return { success: false, error: 'Falta id' };
    const sheets = getAgendaSheets();
    const map = {
      [SHEET_NAMES.BITACORA]: sheets.bitacora,
      [SHEET_NAMES.PENDIENTES]: sheets.pendientes,
      [SHEET_NAMES.EVENTOS]: sheets.eventos
    };
    const sheet = map[sheetName];
    if (!sheet) return { success: false, error: 'Hoja no reconocida: ' + sheetName };
    const row = agendaFindRow(sheet, id);
    if (row === -1) return { success: false, error: 'No encontrado: ' + id };
    sheet.deleteRow(row);
    return { success: true, id: id, deleted: true };
  } catch (err) {
    return { success: false, error: 'Error al eliminar: ' + err.toString() };
  }
}
