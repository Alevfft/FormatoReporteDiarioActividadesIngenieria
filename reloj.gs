/***** CONFIGURACIÓN *****/
const SHEET_ASISTENCIA = 'Asistencia';
const SHEET_USUARIOS = 'Config_Usuarios';
const SHEET_MENSAJES = 'Config_Mensajes';

// Columnas (1-indexed) en Asistencia
const COL_APELLIDO = 3;  // C
const COL_FECHA = 4;  // D
const COL_ENTRADA = 7;  // G
const COL_SALIDA = 8;  // H
const COL_COMIDA_INI = 9;  // I
const COL_COMIDA_FIN = 10; // J
const COL_TIEMPO_DESCANSO = 11; // K (opcional)
const COL_COMENTARIOS = 12; // L
const COL_FIRMA = 13; // M

// Claves esperadas en Config_Mensajes
const KEY_INASISTENCIA = 'INASISTENCIA';
const KEY_RETARDO = 'RETARDO';
const KEY_NO_SALIDA = 'NO_SALIDA';
const KEY_SALIDA_ANT = 'SALIDA_ANTICIPADA';
const KEY_NO_COMIDA_I = 'NO_COMIDA_I';
const KEY_NO_COMIDA_F = 'NO_COMIDA_F';
const KEY_NO_COMIDA_IF = 'NO_COMIDA_IF';
const KEY_COMIDA_TARDE = 'COMIDA_TARDE';

// Defaults para nuevos usuarios
const DEF_H_ENTRADA = '08:00';
const DEF_H_SALIDA = '18:00';
const DEF_TOL_MIN = 15;   // minutos de tolerancia
const DEF_COMIDA_MIN = 60;   // minutos
const DEF_EXENTA = false;

/***** MENÚ *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Asistencia')
    .addItem('Inicializar páginas y celdas', 'initSetup')
    .addItem('Sincronizar usuarios nuevos (desde Asistencia)', 'syncUsuariosDesdeAsistencia')
    .addSeparator()
    .addItem('Evaluar Asistencia', 'evaluarAsistencia')
    .addToUi();
}

/***** INICIALIZACIÓN Y SINCRONIZACIÓN *****/
function initSetup() {
  const ss = SpreadsheetApp.getActive();

  // 1) Asegurar hojas y encabezados
  const shAsis = ensureSheetWithHeaders_(
    ss,
    SHEET_ASISTENCIA,
    [
      'ID de Empleado', // A
      'Nombre',         // B
      'Apellido',       // C
      'Fecha',          // D
      'Día de Semana',  // E
      'Excepción',      // F (no se usa, pero mantenemos)
      'Entrada',        // G
      'Salida',         // H
      'Inicio Descanso',// I
      'Fin Descanso',   // J
      'Tiempo Descanso',// K
      'Comentarios',    // L
      'FIRMA'           // M
    ]
  );

  // Garantizar que existan Comentarios y FIRMA aun si la hoja ya existía
  ensureColumnsAsistencia_(shAsis);

  const shUsers = ensureSheetWithHeaders_(
    ss,
    SHEET_USUARIOS,
    [
      'Apellido',      // A
      'Nombre',        // B
      'Hora_Entrada',  // C HH:MM
      'Hora_Salida',   // D HH:MM
      'ToleranciaMin', // E
      'ComidaMin',     // F
      'ExentaInasistencia' // G TRUE/FALSE
    ]
  );

  const shMsgs = ensureSheetWithHeaders_(
    ss,
    SHEET_MENSAJES,
    ['Clave', 'Mensaje', 'ColorHex']
  );
  seedMensajesPorDefecto_(shMsgs);

  // 2) Sincronizar usuarios desde Asistencia a Config_Usuarios
  syncUsuariosDesdeAsistencia();

  // 3) Reevaluar todo y ordenar
  evaluarAsistencia();

  SpreadsheetApp.getActive().toast('Inicialización/verificación completa.');
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  // Poner encabezados si faltan o están incompletos
  const neededCols = headers.length;
  const firstRow = sh.getRange(1, 1, 1, neededCols).getValues()[0];
  let mustWrite = false;
  for (let i = 0; i < neededCols; i++) {
    if (!firstRow[i] || firstRow[i].toString().trim() === '') {
      mustWrite = true; break;
    }
  }
  if (mustWrite) {
    sh.getRange(1, 1, 1, neededCols).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureColumnsAsistencia_(sh) {
  const lastCol = sh.getLastColumn();
  // Si la hoja existe pero tiene menos de 13 columnas, ampliamos
  if (lastCol < COL_FIRMA) {
    sh.insertColumnsAfter(lastCol, COL_FIRMA - lastCol);
  }
  // Garantizar encabezados puntuales
  const headers = sh.getRange(1, 1, 1, COL_FIRMA).getValues()[0];
  if (!headers[COL_COMENTARIOS - 1]) sh.getRange(1, COL_COMENTARIOS).setValue('Comentarios');
  if (!headers[COL_FIRMA - 1]) sh.getRange(1, COL_FIRMA).setValue('FIRMA');
  sh.setFrozenRows(1);
}

function syncUsuariosDesdeAsistencia() {
  const ss = SpreadsheetApp.getActive();
  const shAsis = ss.getSheetByName(SHEET_ASISTENCIA);
  const shUsers = ss.getSheetByName(SHEET_USUARIOS);
  if (!shAsis || !shUsers) throw new Error('Faltan hojas para sincronizar.');

  const lrA = shAsis.getLastRow();
  if (lrA < 2) { SpreadsheetApp.getActive().toast('No hay filas en Asistencia para sincronizar.'); return; }

  // Mapa existente de usuarios en Config_Usuarios
  const lrU = shUsers.getLastRow();
  const existing = {};
  if (lrU >= 2) {
    const uVals = shUsers.getRange(2, 1, lrU - 1, 2).getValues(); // A:B
    for (const r of uVals) {
      const ap = (r[0] || '').toString().trim().toUpperCase();
      const no = (r[1] || '').toString().trim().toUpperCase();
      if (ap && no) existing[`${ap}|${no}`] = true;
    }
  }

  // Recolectar únicos desde Asistencia
  const aVals = shAsis.getRange(2, 1, lrA - 1, Math.max(COL_APELLIDO, 2)).getValues();
  const toAppend = [];
  for (const r of aVals) {
    const nombre = (r[1] || '').toString().trim().toUpperCase(); // B
    const apellido = (r[2] || '').toString().trim().toUpperCase(); // C
    if (!apellido || !nombre) continue;
    const key = `${apellido}|${nombre}`;
    if (!existing[key]) {
      toAppend.push([
        apellido,
        nombre,
        DEF_H_ENTRADA,
        DEF_H_SALIDA,
        DEF_TOL_MIN,
        DEF_COMIDA_MIN,
        DEF_EXENTA
      ]);
      existing[key] = true;
    }
  }
  if (toAppend.length) {
    shUsers.getRange(shUsers.getLastRow() + 1, 1, toAppend.length, 7).setValues(toAppend);
  }

  SpreadsheetApp.getActive().toast(`Sincronización de usuarios completada. Nuevos: ${toAppend.length}.`);
}

/***** ACCIONES PRINCIPALES (evaluación y orden) *****/
function evaluarAsistencia() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ASISTENCIA);
  if (!sh) throw new Error('No existe la hoja Asistencia');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  evaluarFilas_(sh, 2, lastRow);
  ordenarPorApellido_(sh);
}

// Automático al editar
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_ASISTENCIA) return;

    const col = e.range.getColumn();
    const colsRelevantes = [COL_APELLIDO, COL_FECHA, COL_ENTRADA, COL_SALIDA, COL_COMIDA_INI, COL_COMIDA_FIN];
    if (colsRelevantes.indexOf(col) === -1) return;

    evaluarFilas_(sh, e.range.getRow(), e.range.getRow());
    ordenarPorApellido_(sh);
  } catch (err) {
    console.error(err);
  }
}

/***** LÓGICA DE REGLAS *****/
function evaluarFilas_(sh, rStart, rEnd) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'America/Mexico_City';

  const usuarios = cargarUsuarios_();   // key: "APELLIDO|NOMBRE"
  const mensajes = cargarMensajes_();   // key: CLAVE -> {texto, color}

  const rng = sh.getRange(rStart, 1, rEnd - rStart + 1, sh.getLastColumn());
  const values = rng.getValues();
  const backgrounds = rng.getBackgrounds();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const idx = i;

    const apellido = (row[COL_APELLIDO - 1] || '').toString().trim().toUpperCase();
    const nombre = (row[1] || '').toString().trim().toUpperCase();
    const fecha = row[COL_FECHA - 1];
    const excepcionRaw = row[5];
    const esDescanso = esDiaDescanso_(excepcionRaw);
    row[COL_COMENTARIOS - 1] = '';
    setRowBg_(backgrounds, idx, null);

    if (!apellido || !fecha) { values[i] = row; continue; }

    const uKey = `${apellido}|${nombre}`;
    const uCfg = usuarios[uKey];

    const entrada = coerceToDateTime_(fecha, row[COL_ENTRADA - 1], tz);
    const salida = coerceToDateTime_(fecha, row[COL_SALIDA - 1], tz);
    const ci = coerceToDateTime_(fecha, row[COL_COMIDA_INI - 1], tz);
    const cf = coerceToDateTime_(fecha, row[COL_COMIDA_FIN - 1], tz);

    const comentarios = [];
    if (esDescanso) {
      values[i] = row;
      continue;
    }
    // INASISTENCIA: sin entrada y sin salida (si no está exento)
    if (!entrada && !salida) {
      const exenta = (uCfg && uCfg.exenta) ? true : false;
      if (!exenta && mensajes[KEY_INASISTENCIA]) {
        comentarios.push(mensajes[KEY_INASISTENCIA].texto);
        setRowBg_(backgrounds, idx, mensajes[KEY_INASISTENCIA].color);
      }
      row[COL_COMENTARIOS - 1] = comentarios.join(' | ');
      values[i] = row;
      continue;
    }

    // RETARDO
    if (uCfg && entrada && uCfg.hEntrada) {
      const limite = addMinutes_(uCfg.hEntrada, uCfg.toleranciaMin || 0);
      if (entrada.getTime() > limite.getTime()) {
        const minutos = Math.round((entrada - limite) / 60000);
        if (mensajes[KEY_RETARDO]) {
          const texto = mensajes[KEY_RETARDO].texto.replace(/\{X\}/g, minutos.toString());
          comentarios.push(texto);
          setRowBgIfEmpty_(backgrounds, idx, mensajes[KEY_RETARDO].color);
        }
      }
    }

    // NO SALIDA
    if (entrada && !salida && mensajes[KEY_NO_SALIDA]) {
      comentarios.push(mensajes[KEY_NO_SALIDA].texto);
      setRowBgIfEmpty_(backgrounds, idx, mensajes[KEY_NO_SALIDA].color);
    }

    // SALIDA ANTICIPADA
    if (uCfg && salida && uCfg.hSalida && salida.getTime() < uCfg.hSalida.getTime()) {
      if (mensajes[KEY_SALIDA_ANT]) {
        comentarios.push(mensajes[KEY_SALIDA_ANT].texto);
        setRowBgIfEmpty_(backgrounds, idx, mensajes[KEY_SALIDA_ANT].color);
      }
    }

    // COMIDA
    if (!ci && !cf && mensajes[KEY_NO_COMIDA_IF]) {
      comentarios.push(mensajes[KEY_NO_COMIDA_IF].texto);
      setRowBgIfEmpty_(backgrounds, idx, mensajes[KEY_NO_COMIDA_IF].color);
    } else if (ci && !cf && mensajes[KEY_NO_COMIDA_F]) {
      comentarios.push(mensajes[KEY_NO_COMIDA_F].texto);
      setRowBgIfEmpty_(backgrounds, idx, mensajes[KEY_NO_COMIDA_F].color);
    } else if (!ci && cf && mensajes[KEY_NO_COMIDA_I]) {
      comentarios.push(mensajes[KEY_NO_COMIDA_I].texto);
      setRowBgIfEmpty_(backgrounds, idx, mensajes[KEY_NO_COMIDA_I].color);
    } else if (ci && cf && uCfg && uCfg.comidaMin > 0) {
      const durMin = Math.round((cf - ci) / 60000);
      if (durMin > uCfg.comidaMin && mensajes[KEY_COMIDA_TARDE]) {
        const tarde = durMin - uCfg.comidaMin;
        const texto = mensajes[KEY_COMIDA_TARDE].texto.replace(/\{X\}/g, tarde.toString());
        comentarios.push(texto);
        setRowBgIfEmpty_(backgrounds, idx, mensajes[KEY_COMIDA_TARDE].color);
      }
    }

    row[COL_COMENTARIOS - 1] = comentarios.join(' | ');
    values[i] = row;
  }

  rng.setValues(values);
  rng.setBackgrounds(backgrounds);
}

function ordenarPorApellido_(sh) {
  const lr = sh.getLastRow();
  if (lr < 2) return;
  sh.getRange(2, 1, lr - 1, sh.getLastColumn()).sort([{ column: COL_APELLIDO, ascending: true }]);
}

/***** CARGA CONFIG *****/
function cargarUsuarios_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USUARIOS);
  if (!sh) return {};
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'America/Mexico_City';
  const vals = sh.getRange(2, 1, lastRow - 1, 7).getValues(); // A:G
  const map = {};
  for (const r of vals) {
    const ap = (r[0] || '').toString().trim().toUpperCase();
    const nom = (r[1] || '').toString().trim().toUpperCase();
    const he = r[2]; // HH:MM
    const hs = r[3]; // HH:MM
    const tol = Number(r[4] || 0);
    const com = Number(r[5] || 60);
    const ex = !!r[6];

    const hEntrada = he ? parseClockToday_(he, tz) : null;
    const hSalida = hs ? parseClockToday_(hs, tz) : null;

    map[`${ap}|${nom}`] = {
      hEntrada, hSalida,
      toleranciaMin: tol,
      comidaMin: com,
      exenta: ex
    };
  }
  return map;
}

function cargarMensajes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MENSAJES);
  if (!sh) return {};
  const lr = sh.getLastRow();
  if (lr < 2) return {};
  const vals = sh.getRange(2, 1, lr - 1, 3).getValues(); // A:C
  const map = {};
  for (const r of vals) {
    const key = (r[0] || '').toString().trim().toUpperCase();
    if (!key) continue;
    map[key] = { texto: (r[1] || '').toString(), color: (r[2] || '').toString() || null };
  }
  // Defaults si faltan
  map[KEY_INASISTENCIA] = map[KEY_INASISTENCIA] || { texto: 'INASISTENCIA A REVISAR', color: '#FFE0E0' };
  map[KEY_RETARDO] = map[KEY_RETARDO] || { texto: 'LLEGÓ {X} MINUTOS TARDE', color: '#FFF3CD' };
  map[KEY_NO_SALIDA] = map[KEY_NO_SALIDA] || { texto: 'NO CHECO HORARIO DE SALIDA', color: '#E0E0FF' };
  map[KEY_SALIDA_ANT] = map[KEY_SALIDA_ANT] || { texto: 'CHECO ANTES DE SU HORARIO DE SALIDA', color: '#D6EAF8' };
  map[KEY_NO_COMIDA_I] = map[KEY_NO_COMIDA_I] || { texto: 'NO CHECO LA HORA DE COMIDA INICIO', color: '#FADBD8' };
  map[KEY_NO_COMIDA_F] = map[KEY_NO_COMIDA_F] || { texto: 'NO CHECO LA HORA DE COMIDA FIN', color: '#FADBD8' };
  map[KEY_NO_COMIDA_IF] = map[KEY_NO_COMIDA_IF] || { texto: 'NO CHECO LA HORA DE COMIDA INICIO/FIN', color: '#FADBD8' };
  map[KEY_COMIDA_TARDE] = map[KEY_COMIDA_TARDE] || { texto: 'CHECO {X} MINUTOS TARDE EN COMIDA', color: '#F9E79F' };
  return map;
}

function seedMensajesPorDefecto_(shMsgs) {
  const lr = shMsgs.getLastRow();
  if (lr > 1) return; // ya tiene algo
  const rows = [
    [KEY_INASISTENCIA, 'INASISTENCIA A REVISAR', '#FFE0E0'],
    [KEY_RETARDO, 'LLEGÓ {X} MINUTOS TARDE', '#FFF3CD'],
    [KEY_NO_SALIDA, 'NO CHECO HORARIO DE SALIDA', '#E0E0FF'],
    [KEY_SALIDA_ANT, 'CHECO ANTES DE SU HORARIO DE SALIDA', '#D6EAF8'],
    [KEY_NO_COMIDA_I, 'NO CHECO LA HORA DE COMIDA INICIO', '#FADBD8'],
    [KEY_NO_COMIDA_F, 'NO CHECO LA HORA DE COMIDA FIN', '#FADBD8'],
    [KEY_NO_COMIDA_IF, 'NO CHECO LA HORA DE COMIDA INICIO/FIN', '#FADBD8'],
    [KEY_COMIDA_TARDE, 'CHECO {X} MINUTOS TARDE EN COMIDA', '#F9E79F']
  ];
  shMsgs.getRange(2, 1, rows.length, 3).setValues(rows);
}

/***** HELPERS DE TIEMPO Y FORMATO *****/
function coerceToDateTime_(baseDate, timeCell, tz) {
  if (!timeCell && timeCell !== 0) return null;
  if (timeCell instanceof Date) {
    const d = new Date(baseDate);
    d.setHours(timeCell.getHours(), timeCell.getMinutes(), 0, 0);
    return d;
  }
  const s = timeCell.toString().trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const d = new Date(baseDate);
  d.setHours(parseInt(m[1], 10), parseInt(m[2], 10), 0, 0);
  return d;
}

function esDiaDescanso_(val) {
  if (val == null) return false;
  // Normaliza: quita acentos y pasa a MAYÚSCULAS
  const s = val.toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().trim();
  // Coincide con cualquier variante razonable
  return s.includes('DESCANSO'); // cubre "DIA DE DESCANSO", "DESCANSO", etc.
}


function parseClockToday_(hhmm, tz) {
  const now = new Date();
  const d = new Date(now);
  const m = hhmm.toString().trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  d.setHours(parseInt(m[1], 10), parseInt(m[2], 10), 0, 0);
  return d;
}

function addMinutes_(dateObj, minutes) {
  const d = new Date(dateObj);
  d.setMinutes(d.getMinutes() + minutes);
  return d;
}

function setRowBg_(bgs, idx, hexOrNull) {
  for (let c = 0; c < bgs[0].length; c++) bgs[idx][c] = hexOrNull || null;
}

function setRowBgIfEmpty_(bgs, idx, hex) {
  if (!bgs[idx][0]) setRowBg_(bgs, idx, hex || null);
}
