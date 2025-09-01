/***** CONFIG *****/
const TIMINGS_SPREADSHEET_ID = 'D3oZFOyBD2qWdBq108h0eYK8NtagGkiMUvFQl2O1E';
const ACTIVIDAD_OBJETIVO   = 'SEGUIMIENTO DEL PROYECTO';
const CABECERA_FILA        = 1;

const COL_FECHA            = 1; // A
const COL_OT               = 4; // D
const COL_ACTIVIDAD        = 5; // E
const COL_ACT_TIMING       = 6; // F
const COL_CLIENTE          = 8; // H

// === NUEVAS columnas usadas por Reporte Final (no afectan comportamiento previo) ===
const COL_NOMBRE     = 2; // B
const COL_TIEMPO_EST = 3; // C  (TIEMPO ESTIMADO (HORAS))
const COL_DESC_ACT   = 7; // G  (DESCRIPCIÓN DE LA ACTIVIDAD)

// Validaciones de fecha
const REQUERIR_MES    = true;  // Debe estar en el mes actual

// Hojas internas fijas (ocultas)
const OTS_SHEET_NAME            = '__ots';                 // columnas: nombre_OT | Client
const ACTIVIDADES_SHEET_NAME    = '__actividades_timing';  // columnas: una por OT; debajo, las actividades
const FECHAS_SHEET_NAME         = '__fechas_timing';       // columnas: nombre_OT | Task_Name | Start_Date | Finish_Date

// Encabezados en TIMING
const TIMING_HEADER_ROW     = 1;
const TIMING_TASK_HEADER    = 'Task_Name';
const TIMING_START_HEADER   = 'Start_Date';
const TIMING_FINISH_HEADER  = 'Finish_Date';
const TIMING_CLIENT_HEADER  = 'Client';

// Colores
const COLOR_EN_CURSO   = '#51e425';
const COLOR_VENCIDA    = '#ff7a5c';
const COLOR_NO_INICIA  = '#77bbff';
const COLOR_SIN_COLOR  = null;

/***** CACHÉS POR EJECUCIÓN (mejoran velocidad sin cambiar comportamiento) *****/
const _CACHE = {
  ssActive: null,
  ssTimings: null,
  sheets: {},
  otsRange: null,
  otsList: null,             // array con nombres de OT
  otToClient: null,          // Map nombre_OT -> cliente
  actividadesHeaderIndex: {},// Map OT -> {colIdx, lastRow}
  fechasList: null,          // array de filas de __fechas_timing
  fechasMap: null,           // Map `${ot}||${task}` -> {start, finish}
  tz: null,                  // zona horaria cacheada
  _protBySheetId: new Map(), // protecciones cacheadas por hoja
  _lastColorKey: new Map()   // debounce de coloración por fila
};

const NORM_ACTIVIDAD_OBJETIVO = normalize_(ACTIVIDAD_OBJETIVO);

/***** CACHÉ PERSISTENTE (para onEdit ultrarrápido) *****/
const _DOC_PROPS = PropertiesService.getDocumentProperties();
const _CACHE_SVC = CacheService.getDocumentCache();
const KEY_OT_LIST   = 'OT_LIST_JSON';
const KEY_OT_TASKS  = 'OT_TASKS_JSON';

function _persistCaches_(otList, otTasksMap) {
  try {
    const listTxt = JSON.stringify(otList || []);
    const mapTxt  = JSON.stringify(otTasksMap || {});
    _DOC_PROPS.setProperty(KEY_OT_LIST, listTxt);
    _DOC_PROPS.setProperty(KEY_OT_TASKS, mapTxt);
    _CACHE_SVC.put(KEY_OT_LIST, listTxt, 21600);  // 6h
    _CACHE_SVC.put(KEY_OT_TASKS, mapTxt, 21600);
  } catch(_) {}
}
function _readOTListFast_() {
  let txt = _CACHE_SVC.get(KEY_OT_LIST) || _DOC_PROPS.getProperty(KEY_OT_LIST);
  if (!txt) return [];
  try { return JSON.parse(txt) || []; } catch(_) { return []; }
}
function _readOTTasksMapFastAll_() {
  let txt = _CACHE_SVC.get(KEY_OT_TASKS) || _DOC_PROPS.getProperty(KEY_OT_TASKS);
  if (!txt) return {};
  try { return JSON.parse(txt) || {}; } catch(_) { return {}; }
}
function _readTasksForOTFast_(ot) {
  if (!ot) return [];
  const map = _readOTTasksMapFastAll_();
  const arr = map[ot] || [];
  return Array.isArray(arr) ? arr.filter(Boolean) : [];
}

/***** DV (DataValidation) cache por ejecución *****/
const _DV_CACHE = { otList: null, tasksByOT: new Map() };
function getDVForOTList_(list) {
  if (_DV_CACHE.otList) return _DV_CACHE.otList;
  _DV_CACHE.otList = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  return _DV_CACHE.otList;
}
function getDVForTasks_(ot, tasks) {
  if (_DV_CACHE.tasksByOT.has(ot)) return _DV_CACHE.tasksByOT.get(ot);
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(tasks, true)
    .setAllowInvalid(false)
    .build();
  _DV_CACHE.tasksByOT.set(ot, dv);
  return dv;
}

/***** MENÚ + MODAL *****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // === Modal al abrir (bloqueante) ===
  ui.alert(
    'Formato del reporte',
    '⚠️ AVISO\n\n' +
    'Si en "TIEMPO ESTIMADO (HORAS)" o "ACTIVIDAD GENERAL / TIMING" aparece un valor con advertencia roja (no coincide con su menú desplegable) o NO llenan la columna de NOMBRE, esa fila será EXCLUIDA y NO contará en el Reporte Final.\n\n' +
    'Cualquier duda, aclaración o error de formato, repórtalo con RRHH ANTES de que termine el mes en curso.',
    ui.ButtonSet.OK
  );

  // === Menú 1: Actualizar Timings (renombrado) ===
  ui.createMenu('Actualizar Timings')
    .addItem('Actualizar todos los timings', 'refreshEverything_')
    .addToUi();

  // === Menú 2: Reporte (nuevo) ===
  ui.createMenu('Reporte')
    .addItem('Generar Reporte Final', 'generateFinalReport_')
    .addToUi();
}

function refreshEverything_() {
  refreshOTs_();
  refreshActividades_();
  refreshFechas_();
}

/***** EDITOR PRINCIPAL *****/
function onEdit(e) {
  if (!e || !e.range) return;

  const sh  = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= CABECERA_FILA) return;

  if (!_CACHE.tz) _CACHE.tz = _getActiveSS_().getSpreadsheetTimeZone();

  // Evita trabajo si el usuario dejó el mismo valor
  if (typeof e.value !== 'undefined' && typeof e.oldValue !== 'undefined' && e.value === e.oldValue) return;

  // Prefetch de la fila (una lectura)
  const maxCol = Math.max(COL_CLIENTE, COL_ACT_TIMING, COL_ACTIVIDAD, COL_OT, COL_FECHA);
  const rowVals = sh.getRange(row, 1, 1, maxCol).getValues()[0];

  /***** VALIDACIÓN DE FECHA *****/
  if (col === COL_FECHA) {
    const val = rowVals[COL_FECHA - 1];
    if (val instanceof Date) {
      const hoy   = toMidnight_(new Date(), _CACHE.tz);
      const fecha = toMidnight_(val, _CACHE.tz);
      let ok = true;
      if (REQUERIR_MES) ok = ok && (fecha.getFullYear() === hoy.getFullYear() && fecha.getMonth() === hoy.getMonth());
      if (!ok) {
        e.range.clearContent();
        SpreadsheetApp.getActive().toast('Fecha inválida: solo mes actual.', 'Validación', 3);
        return;
      } else {
        e.range.setBorder(false, false, false, false, false, false, null, null);
      }
    }
  }

  /***** REACCIONES DE CAMPOS *****/
  if (col === COL_ACTIVIDAD) {
    // Limpia OT + ACT_TIMING + CLIENTE en bloque
    sh.getRangeList([colToA1_(COL_OT, row), colToA1_(COL_ACT_TIMING, row), colToA1_(COL_CLIENTE, row)]).clearContent();
    unlockClientCell_(sh, row);
  }

  if (col === COL_OT) {
    sh.getRangeList([colToA1_(COL_ACT_TIMING, row), colToA1_(COL_CLIENTE, row)]).clearContent();
    unlockClientCell_(sh, row);
  }

  // Solo recalcula validaciones cuando cambian ACT o OT
  if (col === COL_ACTIVIDAD || col === COL_OT) {
    applyRowValidations_(sh, row, e);
  }

  // Recolorea si cambian ACT, OT o ACT_TIMING
  if (col === COL_ACTIVIDAD || col === COL_OT || col === COL_ACT_TIMING) {
    applyTimingColorForRow_(sh, row);
  }

  // Autollenado + lock de CLIENTE cuando es SEGUIMIENTO y hay OT
  const actividadGeneral = String(rowVals[COL_ACTIVIDAD - 1] || '').trim();
  const esSeg = normalize_(actividadGeneral) === NORM_ACTIVIDAD_OBJETIVO;
  const otEdit = (typeof e.value !== 'undefined' && col === COL_OT) ? e.value : rowVals[COL_OT - 1];
  const ot     = String(otEdit || '').trim();

  if (esSeg && ot) {
    const cliente = getClientFromLocalOTs_(ot);
    const cCli = sh.getRange(row, COL_CLIENTE);
    if (cliente) setIfChanged_(cCli, cliente); else cCli.clearContent();
    lockClientCell_(sh, row);
  } else {
    unlockClientCell_(sh, row);
  }
}


/***** VALIDACIÓN POR FILA (OT/Actividad Timing) *****/
function applyRowValidations_(sh, row, eOpt) {
  const cAct = sh.getRange(row, COL_ACTIVIDAD);
  const cOT  = sh.getRange(row, COL_OT);
  const cAT  = sh.getRange(row, COL_ACT_TIMING);

  const esSeg = normalize_(cAct.getValue()) === NORM_ACTIVIDAD_OBJETIVO;
  if (!esSeg) {
    // Si no es seguimiento, sin validaciones
    cOT.setDataValidation(null);
    cAT.setDataValidation(null);
    return;
  }

  // 1) OTs desde caché (rapidísimo); fallback a rango si no hay cache
  const otList = _readOTListFast_();
  if (otList.length) {
    const dvOT = getDVForOTList_(otList);
    // solo aplica si cambió para evitar I/O
    const cur = cOT.getDataValidation();
    if (!cur || String(cur) !== String(dvOT)) cOT.setDataValidation(dvOT);
  } else {
    const otRange = getLocalOTRange_();
    const ruleOT  = SpreadsheetApp.newDataValidation().requireValueInRange(otRange, true).setAllowInvalid(false).build();
    const cur = cOT.getDataValidation();
    if (!cur || String(cur) !== String(ruleOT)) cOT.setDataValidation(ruleOT);
  }

  // 2) Actividades de la OT elegida
  const valorOT = (eOpt && typeof eOpt.value !== 'undefined' && eOpt.range.getColumn() === COL_OT)
    ? String(eOpt.value || '').trim()
    : String(cOT.getValue() || '').trim();

  if (!valorOT || (otList.length && !otList.includes(valorOT))) {
    cAT.setDataValidation(null);
    cAT.clearContent();
    return;
  }

  const tasks = _readTasksForOTFast_(valorOT);
  if (tasks.length) {
    const dvAT = getDVForTasks_(valorOT, tasks);
    const cur = cAT.getDataValidation();
    if (!cur || String(cur) !== String(dvAT)) cAT.setDataValidation(dvAT);
  } else {
    // Fallback a rango oculto si aún no hay cache de tareas
    const tasksRange = getTasksRangeFromLocal_(valorOT);
    if (tasksRange) {
      const ruleAT = SpreadsheetApp.newDataValidation().requireValueInRange(tasksRange, true).setAllowInvalid(false).build();
      const cur = cAT.getDataValidation();
      if (!cur || String(cur) !== String(ruleAT)) cAT.setDataValidation(ruleAT);
    } else {
      cAT.setDataValidation(null);
      cAT.clearContent();
    }
  }
}

/***** COLOR POR FECHAS (solo __fechas_timing, con caché + debounce) *****/
function applyTimingColorForRow_(sh, row) {
  const vals = sh.getRange(row, 1, 1, Math.max(COL_ACT_TIMING, COL_OT, COL_ACTIVIDAD)).getValues()[0];
  const actividadGeneral = String(vals[COL_ACTIVIDAD - 1] || '').trim();
  const cell  = sh.getRange(row, COL_ACT_TIMING);

  if (normalize_(actividadGeneral) !== NORM_ACTIVIDAD_OBJETIVO) {
    setBgIfChanged_(cell, COLOR_SIN_COLOR);
    setNoteIfChanged_(cell, '');
    return;
  }

  const ot    = String(vals[COL_OT - 1] || '').trim();
  const tarea = String(vals[COL_ACT_TIMING - 1] || '').trim();
  if (!ot || !tarea) {
    setBgIfChanged_(cell, COLOR_SIN_COLOR);
    setNoteIfChanged_(cell, '');
    return;
  }

  // Debounce simple por ejecución para evitar doble trabajo si nada cambió
  const key = `${sh.getSheetId()}:${row}:${ot}:${tarea}`;
  if (_CACHE._lastColorKey.get(row) === key) return;
  _CACHE._lastColorKey.set(row, key);

  const fechas = getDatesFromLocal_(ot, tarea); // usa caché interna
  if (!fechas) {
    setBgIfChanged_(cell, COLOR_SIN_COLOR);
    setNoteIfChanged_(cell, 'No hay fechas en __fechas_timing para "'+tarea+'" (OT: "'+ot+'"). Usa “Actualizar fechas”.');
    return;
  }

  if (!_CACHE.tz) _CACHE.tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const hoy    = stripTime_(new Date(), _CACHE.tz);
  const start  = stripTime_(fechas.start, _CACHE.tz);
  const finish = stripTime_(fechas.finish, _CACHE.tz);

  let color = COLOR_SIN_COLOR;
  if (hoy >= start && hoy <= finish) color = COLOR_EN_CURSO;
  else if (hoy > finish)             color = COLOR_VENCIDA;
  else if (hoy < start)              color = COLOR_NO_INICIA;

  const noteTxt = 'Inicio: ' + formatDMY_(start, _CACHE.tz) + '\nFin: ' + formatDMY_(finish, _CACHE.tz);
  setBgIfChanged_(cell, color);
  setNoteIfChanged_(cell, noteTxt);
}

/***** BOTONES: RECONSTRUCCIÓN DE ARCHIVOS INTERNOS *****/
function refreshOTs_() {
  const sh = ensureOTsSheet_();
  sh.clear();
  sh.getRange(1,1,1,2).setValues([['nombre_OT','Client']]);

  const timing = _getTimingsSS_();
  const sheets = timing.getSheets();

  const rows = [];
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    const otName = s.getName();
    const client = getFirstClientFromTimingSheet_(s);
    rows.push([otName, client]);
  }
  if (rows.length) sh.getRange(2,1,rows.length,2).setValues(rows);
  if (!sh.isSheetHidden()) sh.hideSheet();

  // invalidar cachés en memoria
  _CACHE.otsRange = null;
  _CACHE.otsList = null;
  _CACHE.otToClient = null;

  // --- Actualiza cache persistente de OTs (mantén tareas previas si existen)
  const otList = rows.map(r => String(r[0]).trim()).filter(Boolean);
  const existingTasksMap = _readOTTasksMapFastAll_(); // conserva lo que ya tenías
  _persistCaches_(otList, existingTasksMap);

  SpreadsheetApp.getActive().toast(`OTs actualizadas: ${rows.length}`, 'OK', 3);
}

function refreshActividades_() {
  const sh = ensureActividadesSheet_();
  sh.clear();

  const ots = getAllOTsFromLocal_();
  if (!ots.length) {
    SpreadsheetApp.getActive().toast('No hay OTs en __ots. Primero usa "Actualizar OTs".', 'Atención', 4);
    // deja caches limpias
    _persistCaches_([], {});
    return;
  }

  const timing = _getTimingsSS_();
  const dataCols = [];
  let maxLen = 0;

  // Para cache persistente
  const otTasksMap = {};
  for (let i = 0; i < ots.length; i++) {
    const ot = ots[i];
    const tasks = readTasksFromTimingByName_(ot, timing); // pasa SS para no abrirlo de nuevo
    dataCols.push([ot, ...tasks]);
    otTasksMap[ot] = tasks;
    if (tasks.length > maxLen) maxLen = tasks.length;
  }

  if (dataCols.length) {
    for (let i = 0; i < dataCols.length; i++) while (dataCols[i].length < maxLen + 1) dataCols[i].push('');
    const rows = transpose_(dataCols);
    sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  }
  if (!sh.isSheetHidden()) sh.hideSheet();

  // invalidar caché de cabeceras por OT
  _CACHE.actividadesHeaderIndex = {};

  // También refresca cache persistente de OTs y tareas
  _persistCaches_(ots.slice(), otTasksMap);

  SpreadsheetApp.getActive().toast(`Actividades actualizadas para ${ots.length} OTs.`, 'OK', 3);
}

function refreshFechas_() {
  const sh = ensureFechasSheet_();
  sh.clear();
  sh.getRange(1,1,1,4).setValues([['nombre_OT','Task_Name','Start_Date','Finish_Date']]);

  const ots = getAllOTsFromLocal_();
  if (!ots.length) {
    SpreadsheetApp.getActive().toast('No hay OTs en __ots. Primero usa "Actualizar OTs".', 'Atención', 4);
    return;
  }

  const timing = _getTimingsSS_();
  const rows = [];

  for (let i = 0; i < ots.length; i++) {
    const ot = ots[i];
    const s = timing.getSheetByName(ot);
    if (!s) continue;
    const lastRow = s.getLastRow();
    const lastCol = s.getLastColumn();
    if (lastRow < TIMING_HEADER_ROW + 1) continue;

    const headers = s.getRange(TIMING_HEADER_ROW, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
    const iTask   = headers.indexOf(TIMING_TASK_HEADER)   + 1;
    const iStart  = headers.indexOf(TIMING_START_HEADER)  + 1;
    const iFinish = headers.indexOf(TIMING_FINISH_HEADER) + 1;
    if (iTask <= 0 || iStart <= 0 || iFinish <= 0) continue;

    // Lee solo las 3 columnas necesarias (reduce IO)
    const n = lastRow - TIMING_HEADER_ROW;
    const tasks  = s.getRange(TIMING_HEADER_ROW + 1, iTask,   n, 1).getValues();
    const starts = s.getRange(TIMING_HEADER_ROW + 1, iStart,  n, 1).getValues();
    const ends   = s.getRange(TIMING_HEADER_ROW + 1, iFinish, n, 1).getValues();

    for (let r = 0; r < n; r++) {
      const task   = String(tasks[r][0]).trim();
      if (!task) continue;
      const start  = toDateSafe_(starts[r][0]);
      const finish = toDateSafe_(ends[r][0]);
      rows.push([ot, task, start || '', finish || '']);
    }
  }

  if (rows.length) sh.getRange(2,1,rows.length,4).setValues(rows);
  if (!sh.isSheetHidden()) sh.hideSheet();

  // invalidar y precargar caché de fechas
  _CACHE.fechasList = rows;
  _CACHE.fechasMap = null;

  SpreadsheetApp.getActive().toast(`Fechas actualizadas: ${rows.length} filas.`, 'OK', 3);
}

/***** LECTURAS DESDE LOS ARCHIVOS INTERNOS (con caché) *****/
function getLocalOTRange_() {
  if (_CACHE.otsRange) return _CACHE.otsRange;
  const sh = ensureOTsSheet_();
  const last = Math.max(sh.getLastRow(), 1);
  _CACHE.otsRange = (last < 2) ? sh.getRange(2,1,1,1) : sh.getRange(2,1,last-1,1);
  return _CACHE.otsRange;
}

function getTasksRangeFromLocal_(otName) {
  const sh = ensureActividadesSheet_();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  if (lastCol < 1) return null;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x).trim());
  const colIdx = headers.indexOf(otName) + 1;
  if (colIdx <= 0) return null;

  const lastRow = getColumnLastRow_(sh, colIdx, 2);
  if (lastRow < 2) return null;
  return sh.getRange(2, colIdx, lastRow-1, 1);
}

function _buildFechasMapIfNeeded_() {
  if (_CACHE.fechasMap) return;
  const sh = ensureFechasSheet_();
  const lastRow = sh.getLastRow();
  let vals;
  if (_CACHE.fechasList) {
    vals = _CACHE.fechasList;
  } else {
    if (lastRow < 2) {
      _CACHE.fechasMap = new Map();
      return;
    }
    vals = sh.getRange(2,1,lastRow-1,4).getValues(); // nombre_OT | Task_Name | Start | Finish
  }
  const map = new Map();
  for (let i = 0; i < vals.length; i++) {
    const [ot, task, s, f] = vals[i];
    if (!ot || !task) continue;
    const start  = toDateSafe_(s);
    const finish = toDateSafe_(f);
    if (!start || !finish) continue;
    map.set(`${String(ot).trim()}||${String(task).trim()}`, { start, finish });
  }
  _CACHE.fechasMap = map;
}

function getDatesFromLocal_(otName, taskName) {
  _buildFechasMapIfNeeded_();
  if (!_CACHE.fechasMap) return null;
  return _CACHE.fechasMap.get(`${otName}||${taskName}`) || null;
}

function getAllOTsFromLocal_() {
  if (_CACHE.otsList) return _CACHE.otsList.slice();
  const sh = ensureOTsSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    _CACHE.otsList = [];
    return [];
  }
  const list = sh.getRange(2,1,last-1,1).getValues().flat().map(v => String(v).trim()).filter(Boolean);
  _CACHE.otsList = list;
  return list.slice();
}

/***** CLIENTE (desde __ots) *****/
function getClientFromLocalOTs_(otName) {
  if (_CACHE.otToClient) {
    return _CACHE.otToClient.get(otName) || '';
  }
  const sh = ensureOTsSheet_();
  const lastRow = sh.getLastRow();
  const map = new Map();
  if (lastRow >= 2) {
    const vals = sh.getRange(2,1,lastRow-1,2).getValues(); // nombre_OT | Client
    for (let i=0;i<vals.length;i++){
      const k = String(vals[i][0]).trim();
      const v = String(vals[i][1] || '').trim();
      if (k) map.set(k, v);
    }
  }
  _CACHE.otToClient = map;
  return map.get(otName) || '';
}

/***** BLOQUEO / DESBLOQUEO DE CLIENTE (con protecciones cacheadas) *****/
function lockClientCell_(sh, row) {
  removeClientProtectionIfExists_(sh, row);
  const prot = sh.getRange(row, COL_CLIENTE).protect();
  prot.setDescription(makeClientLockDesc_(sh, row));
  prot.setWarningOnly(false);
  const editors = prot.getEditors();
  if (editors && editors.length) prot.removeEditors(editors);
  if (prot.canDomainEdit()) prot.setDomainEdit(false);
  _pushProtCache_(sh, prot);
}

function unlockClientCell_(sh, row) {
  removeClientProtectionIfExists_(sh, row);
}

function removeClientProtectionIfExists_(sh, row) {
  const sheetId = sh.getSheetId();
  let list = _CACHE._protBySheetId.get(sheetId);
  if (!list) {
    list = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    _CACHE._protBySheetId.set(sheetId, list);
  }
  const rangeA1 = sh.getRange(row, COL_CLIENTE).getA1Notation();
  const desc    = makeClientLockDesc_(sh, row);
  const kept = [];
  for (const p of list) {
    try {
      const r = p.getRange();
      if (r && (r.getA1Notation() === rangeA1 || p.getDescription() === desc)) {
        p.remove();
      } else {
        kept.push(p);
      }
    } catch(_) {}
  }
  _CACHE._protBySheetId.set(sheetId, kept);
}

function _pushProtCache_(sh, prot) {
  const sheetId = sh.getSheetId();
  let list = _CACHE._protBySheetId.get(sheetId);
  if (!list) list = [];
  list.push(prot);
  _CACHE._protBySheetId.set(sheetId, list);
}

function makeClientLockDesc_(sh, row) {
  return `__lock_cliente_${sh.getSheetId()}_${row}`;
}

/***** UTILIDADES TIMING *****/
function readTasksFromTimingByName_(otSheetName, timingSSOpt) {
  const timing = timingSSOpt || _getTimingsSS_();
  const sh = timing.getSheetByName(otSheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < TIMING_HEADER_ROW + 1) return [];

  const headers = sh.getRange(TIMING_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const colTask = headers.findIndex(h => String(h).trim() === TIMING_TASK_HEADER) + 1;
  if (colTask <= 0) return [];

  const vals = sh.getRange(TIMING_HEADER_ROW + 1, colTask, lastRow - TIMING_HEADER_ROW, 1)
    .getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v);

  const seen = new Set();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i];
    if (!seen.has(v)) { seen.add(v); out.push(v); }
  }
  return out;
}

function getFirstClientFromTimingSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < TIMING_HEADER_ROW + 1) return '';
  const headers = sheet.getRange(TIMING_HEADER_ROW, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const iClient = headers.indexOf(TIMING_CLIENT_HEADER) + 1;
  if (iClient <= 0) return '';
  const colVals = sheet.getRange(TIMING_HEADER_ROW + 1, iClient, lastRow - TIMING_HEADER_ROW, 1).getValues();
  for (let i=0;i<colVals.length;i++){
    const v = String(colVals[i][0]).trim();
    if (v) return v;
  }
  return '';
}

/***** HELPERS GENERALES *****/
function setSelectedTask(task, row, sheetName) {
  const ss = _getActiveSS_();
  const sh = ss.getSheetByName(sheetName);
  sh.getRange(row, COL_ACT_TIMING).setValue(task);
}

// Escrituras “solo si cambia”
function setIfChanged_(range, value) {
  const cur = range.getValue();
  if (cur !== value) range.setValue(value);
}
function setBgIfChanged_(range, color) {
  const cur = range.getBackground();
  if (cur !== color) range.setBackground(color);
}
function setNoteIfChanged_(range, note) {
  const cur = range.getNote();
  if (cur !== note) range.setNote(note);
}

// Miscelánea
function toDateSafe_(v) {
  if (v instanceof Date) return v;
  const s = String(v).trim();
  if (!s) return null;
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/); // d/m/Y
  if (m1) return new Date(parseInt(m1[3],10), parseInt(m1[2],10)-1, parseInt(m1[1],10));
  const m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/); // Y-m-d
  if (m2) return new Date(parseInt(m2[1],10), parseInt(m2[2],10)-1, parseInt(m2[3],10));
  return null;
}
function stripTime_(d, tz) { return new Date(Utilities.formatDate(d, tz, 'yyyy/MM/dd')); }
function formatDMY_(d, tz) { return Utilities.formatDate(d, tz, 'dd/MM/yyyy'); }
function normalize_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}
function transpose_(arrOfCols) {
  const rows = Math.max(...arrOfCols.map(c => c.length));
  const cols = arrOfCols.length;
  const out = Array.from({length: rows}, () => Array(cols).fill(''));
  for (let c=0;c<cols;c++){
    const colArr = arrOfCols[c];
    for (let r=0;r<colArr.length;r++){
      out[r][c] = colArr[r];
    }
  }
  return out;
}
function getColumnLastRow_(sheet, col, startRow) {
  const last = sheet.getLastRow();
  const n = Math.max(last - startRow + 1, 1);
  const vals = sheet.getRange(startRow, col, n, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return startRow + i;
  }
  return startRow - 1;
}
function toMidnight_(date, tz) {
  return new Date(Utilities.formatDate(date, tz, "yyyy-MM-dd'T'00:00:00"));
}
function colToA1_(col, row) {
  // Convierte índice de columna (1-based) a letra(s) + fila (e.g., 4,10 -> D10)
  let c = col, s = '';
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s + row;
}

/***** ENSURE SHEETS (ocultas) *****/
function ensureOTsSheet_() {
  if (_CACHE.sheets[OTS_SHEET_NAME]) return _CACHE.sheets[OTS_SHEET_NAME];
  const ss = _getActiveSS_();
  let sh = ss.getSheetByName(OTS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(OTS_SHEET_NAME);
  if (!sh.isSheetHidden()) sh.hideSheet();
  _CACHE.sheets[OTS_SHEET_NAME] = sh;
  return sh;
}
function ensureActividadesSheet_() {
  if (_CACHE.sheets[ACTIVIDADES_SHEET_NAME]) return _CACHE.sheets[ACTIVIDADES_SHEET_NAME];
  const ss = _getActiveSS_();
  let sh = ss.getSheetByName(ACTIVIDADES_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(ACTIVIDADES_SHEET_NAME);
  if (!sh.isSheetHidden()) sh.hideSheet();
  _CACHE.sheets[ACTIVIDADES_SHEET_NAME] = sh;
  return sh;
}
function ensureFechasSheet_() {
  if (_CACHE.sheets[FECHAS_SHEET_NAME]) return _CACHE.sheets[FECHAS_SHEET_NAME];
  const ss = _getActiveSS_();
  let sh = ss.getSheetByName(FECHAS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(FECHAS_SHEET_NAME);
  if (!sh.isSheetHidden()) sh.hideSheet();
  _CACHE.sheets[FECHAS_SHEET_NAME] = sh;
  return sh;
}

/***** APERTURAS MEMOIZADAS *****/
function _getActiveSS_() {
  if (_CACHE.ssActive) return _CACHE.ssActive;
  _CACHE.ssActive = SpreadsheetApp.getActive();
  return _CACHE.ssActive;
}
function _getTimingsSS_() {
  if (_CACHE.ssTimings) return _CACHE.ssTimings;
  _CACHE.ssTimings = SpreadsheetApp.openById(TIMINGS_SPREADSHEET_ID);
  return _CACHE.ssTimings;
}

/***** === NUEVO: Generar Reporte Final === *****/
function generateFinalReport_() {
  const ss = _getActiveSS_();
  const src = ss.getActiveSheet();
  const srcName = src.getName();

  // Copia de la pestaña
  const copy = src.copyTo(ss);
  const newName = `REPORTE FINAL-${srcName}`;
  try { copy.setName(newName); }
  catch (e) {
    const t = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmmss');
    copy.setName(`${newName}_${t}`);
  }
  ss.setActiveSheet(copy);
  ss.moveActiveSheet(ss.getNumSheets());

  const sh = copy;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= CABECERA_FILA) {
    SpreadsheetApp.getUi().alert('No hay filas de datos para procesar.');
    return;
  }

  // === Helpers ===
  function getFirstDVRuleInColumn_(sheet, col, startRow) {
    const maxR = sheet.getMaxRows();
    const rng = sheet.getRange(startRow, col, Math.max(1, maxR - startRow + 1), 1);
    const dvMat = rng.getDataValidations();
    for (let i = 0; i < dvMat.length; i++) {
      const rule = dvMat[i][0];
      if (rule) return rule;
    }
    return null;
  }
  function inferAllowedList_(sheet, col, startRow) {
    const rule = getFirstDVRuleInColumn_(sheet, col, startRow);
    if (!rule || !rule.getCriteriaType) return null;
    const type = rule.getCriteriaType();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const arr = (rule.getCriteriaValues()[0] || []).map(v => String(v).trim()).filter(Boolean);
      return arr.length ? arr : null;
    }
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const rng = rule.getCriteriaValues()[0];
      if (!rng || !rng.getDisplayValues) return null;
      const vals = rng.getDisplayValues().flat().map(v => String(v).trim()).filter(Boolean);
      return vals.length ? vals : null;
    }
    return null;
  }
  function toNormSet_(arr) { return new Set((arr || []).map(x => normalize_(x))); }

  const dataStart = CABECERA_FILA + 1;

  // === 1) Listas permitidas
  const allowedTiempoList =
    inferAllowedList_(src, COL_TIEMPO_EST, dataStart) ||
    inferAllowedList_(sh,  COL_TIEMPO_EST, dataStart);

  // Fallback para ACTIVIDAD GENERAL (por si no detectamos DV en la columna)
  const FALLBACK_ACT_LIST = [
    'SEGUIMIENTO DEL PROYECTO',
    'COMIDA',
    'VACACIONES',
    'RRHH',
    'CAPACITACIÓN',   // con acento
    'CAPACITACION',   // sin acento (por si tu lista está así)
    'JUNTA CLIENTE',
    'JUNTA CON CLIENTE',
    'JUNTA INTERNA',
    'OT INTERNA'
  ];
  const allowedActList =
    inferAllowedList_(src, COL_ACTIVIDAD, dataStart) ||
    inferAllowedList_(sh,  COL_ACTIVIDAD, dataStart) ||
    FALLBACK_ACT_LIST.slice(); // <- aquí garantizamos lista

  const allowedActSetNorm = toNormSet_(allowedActList);
  const tiempoRegex = /^([0-1]?\d|2[0-3]):[0-5]\d$/; // fallback estricto hh:mm

  // === 2) Eliminar filas que NO cumplan
  const rowsToDelete = [];
  for (let r = dataStart; r <= lastRow; r++) {
    const vNombre = String(sh.getRange(r, COL_NOMBRE).getDisplayValue() || '').trim();
    if (!vNombre) { rowsToDelete.push(r); continue; }

    const vTiempo = String(sh.getRange(r, COL_TIEMPO_EST).getDisplayValue() || '').trim();
    const vAct    = String(sh.getRange(r, COL_ACTIVIDAD).getDisplayValue() || '').trim();

    let okTiempo = true;
    if (allowedTiempoList) {
      okTiempo = (vTiempo === '') ? true : allowedTiempoList.includes(vTiempo);
    } else {
      okTiempo = (vTiempo === '') ? true : tiempoRegex.test(vTiempo);
    }

    // ACTIVIDAD GENERAL: comparación por texto normalizado (con fallback de lista)
    const actNorm = normalize_(vAct);
    const okActividad = (vAct === '') ? true : allowedActSetNorm.has(actNorm);

    if (!okTiempo || !okActividad) rowsToDelete.push(r);
  }
  for (let i = rowsToDelete.length - 1; i >= 0; i--) sh.deleteRow(rowsToDelete[i]);

  const lastRow2 = sh.getLastRow();
  if (lastRow2 <= CABECERA_FILA) {
    SpreadsheetApp.getActive().toast('Todas las filas fueron eliminadas por validación.', 'Reporte Final', 5);
    return;
  }

  // === 3) Quitar TODAS las validaciones
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setDataValidation(null);

  // === 4) Normalizar a texto
  const rngData = sh.getRange(dataStart, 1, lastRow2 - CABECERA_FILA, Math.max(lastCol, COL_CLIENTE));
  let vals = rngData.getDisplayValues().map(row => row.map(v => (v == null ? '' : String(v))));

  // === 5) Reglas por actividad + fecha (igual que ya tenías, con limpieza de descripción en COMIDA/VACACIONES)
  for (let i = 0; i < vals.length; i++) {
    const r = i + dataStart;

    let vFecha     = vals[i][COL_FECHA-1];
    let vNombre    = vals[i][COL_NOMBRE-1];
    let vTiempo    = vals[i][COL_TIEMPO_EST-1];
    let vOT        = vals[i][COL_OT-1];
    let vActGral   = vals[i][COL_ACTIVIDAD-1];
    let vActTiming = vals[i][COL_ACT_TIMING-1];
    let vDesc      = vals[i][COL_DESC_ACT-1];
    let vCliente   = vals[i][COL_CLIENTE-1];

    const act = (vActGral || '').toString().trim().toUpperCase();
    const isCOMIDA = act === 'COMIDA';
    const isVAC    = act === 'VACACIONES';

    let fechaObj = toDateSafe_(vFecha);
    if (!vFecha || !String(vFecha).trim()) {
      vFecha = 'EL USUARIO NO LLENO LOS DATOS';
    } else if (fechaObj instanceof Date) {
      const day = fechaObj.getDay();
      if (day === 0 || day === 6) sh.getRange(r, COL_FECHA).setBackground(COLOR_NO_INICIA);
      else sh.getRange(r, COL_FECHA).setBackground(null);
    }

    if (isCOMIDA) {
      vTiempo = '1:00';
      vOT = ''; vActTiming = ''; vDesc = ''; vCliente = '';
    } else if (isVAC) {
      vTiempo = '10:00';
      vOT = ''; vActTiming = ''; vDesc = ''; vCliente = '';
    } else if (act === 'RRHH' || act === 'CAPACITACIÓN' || act === 'CAPACITACION') {
      vOT = ''; vActTiming = ''; vCliente = '';
    } else if (act === 'JUNTA CLIENTE' || act === 'JUNTA CON CLIENTE') {
      vActTiming = '';
      if (!vFecha || !String(vFecha).trim()) vFecha = 'USUARIO NO LLENO LOS DATOS';
      if (!vTiempo || !String(vTiempo).trim()) vTiempo = 'USUARIO NO LLENO LOS DATOS';
      if (!vDesc || !String(vDesc).trim()) vDesc = 'USUARIO NO LLENO LOS DATOS';
      if (!vCliente || !String(vCliente).trim()) vCliente = 'USUARIO NO LLENO LOS DATOS';
      if (!vOT || !String(vOT).trim()) vOT = 'EL USUARIO NO LLENO LOS DATOS O EL CLIENTE NO TIENE OT';
    } else if (act === 'JUNTA INTERNA') {
      vActTiming = '';
      if (!vFecha || !String(vFecha).trim()) vFecha = 'USUARIO NO LLENO LOS DATOS';
      if (!vTiempo || !String(vTiempo).trim()) vTiempo = 'USUARIO NO LLENO LOS DATOS';
      if (!vDesc || !String(vDesc).trim()) vDesc = 'USUARIO NO LLENO LOS DATOS';
      if (!vCliente || !String(vCliente).trim()) vCliente = 'USUARIO NO LLENO LOS DATOS O NO ES JUNTA SOBRE UN CLIENTE/OT)';
    } else if (act === 'OT INTERNA') {
      vActTiming = ''; vCliente = '';
      if (!vFecha || !String(vFecha).trim()) vFecha = 'USUARIO NO LLENO LOS DATOS';
      if (!vTiempo || !String(vTiempo).trim()) vTiempo = 'USUARIO NO LLENO LOS DATOS';
      if (!vDesc || !String(vDesc).trim()) vDesc = 'USUARIO NO LLENO LOS DATOS';
      if (!vOT || !String(vOT).trim()) vOT = 'USUARIO NO LLENO LOS DATOS';
    } else if (normalize_(act) === NORM_ACTIVIDAD_OBJETIVO) {
      if (!vFecha || !String(vFecha).trim()) vFecha = 'USUARIO NO LLENO LOS DATOS';
      if (!vTiempo || !String(vTiempo).trim()) vTiempo = 'USUARIO NO LLENO LOS DATOS';
      if (!vDesc || !String(vDesc).trim()) vDesc = 'USUARIO NO LLENO LOS DATOS';
      if (!vOT || !String(vOT).trim()) vOT = 'USUARIO NO LLENO LOS DATOS';
      if (!vActTiming || !String(vActTiming).trim()) vActTiming = 'USUARIO NO LLENO LOS DATOS';
      if (!vCliente || !String(vCliente).trim()) vCliente = 'USUARIO NO LLENO LOS DATOS';
    }

    if (!vFecha || !String(vFecha).trim()) vFecha = 'USUARIO NO LLENO LOS DATOS';
    if (!vTiempo || !String(vTiempo).trim()) vTiempo = 'USUARIO NO LLENO LOS DATOS';
    if (!(isCOMIDA || isVAC)) {
      if (!vDesc || !String(vDesc).trim()) vDesc = 'USUARIO NO LLENO LOS DATOS';
    } else {
      vDesc = '';
    }

    const toTxt = x => (x === '' ? '' : "'" + String(x));
    vals[i][COL_FECHA-1]      = toTxt(vFecha);
    vals[i][COL_NOMBRE-1]     = toTxt(vNombre);
    vals[i][COL_TIEMPO_EST-1] = toTxt(vTiempo);
    vals[i][COL_OT-1]         = toTxt(vOT);
    vals[i][COL_ACTIVIDAD-1]  = toTxt(vActGral);
    vals[i][COL_ACT_TIMING-1] = toTxt(vActTiming);
    vals[i][COL_DESC_ACT-1]   = toTxt(vDesc);
    vals[i][COL_CLIENTE-1]    = toTxt(vCliente);
  }

  // === 6) Escribir resultados
  sh.getRange(dataStart, 1, vals.length, Math.max(vals[0].length, lastCol)).setValues(vals);

  SpreadsheetApp.getActive().toast(`Reporte final generado en la pestaña "${sh.getName()}".`, 'Reporte Final', 4);
}




