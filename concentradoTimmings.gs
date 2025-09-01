/***** CONFIG *****/
const APPLY_TO_ALL_SHEETS = true;
const DATE_HEADERS = ['Start_Date', 'Finish_Date'];
const HEADER_ROW = 1;
const DATE_FORMAT = 'dd/MM/yyyy';

/***** MENÚ *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TIMING')
    .addItem('Limpiar fechas (sin hora)', 'cleanTimingDates')
    .addToUi();
}

/***** ACCIÓN PRINCIPAL *****/
function cleanTimingDates() {
  const ss = SpreadsheetApp.getActive();
  const sheets = APPLY_TO_ALL_SHEETS ? ss.getSheets() : [ss.getActiveSheet()];
  sheets.forEach(sh => cleanDatesInSheet_(sh));
}

/***** IMPLEMENTACIÓN *****/
function cleanDatesInSheet_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROW + 1) return;

  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const targetCols = DATE_HEADERS
    .map(h => headers.findIndex(x => String(x).trim() === h) + 1)
    .filter(idx => idx > 0);
  if (!targetCols.length) return;

  targetCols.forEach(col => {
    const numRows = lastRow - HEADER_ROW;
    const rng     = sh.getRange(HEADER_ROW + 1, col, numRows, 1);

    // Leemos ambos: crudo (tipo real) y lo que se muestra en pantalla
    const raw  = rng.getValues();
    const disp = rng.getDisplayValues();

    for (let i = 0; i < numRows; i++) {
      const vRaw  = raw[i][0];
      const vDisp = disp[i][0];

      // 1) Si ya es Date, solo quitar la hora
      if (vRaw instanceof Date && !isNaN(vRaw)) {
        raw[i][0] = new Date(vRaw.getFullYear(), vRaw.getMonth(), vRaw.getDate());
        continue;
      }

      // 2) Parsear lo visible (maneja 'p. m.' / 'a. m.' y variados)
      const dFromDisp = toDateOnly_(vDisp);
      if (dFromDisp) { raw[i][0] = dFromDisp; continue; }

      // 3) Si es número serial de Sheets, normalizar
      if (typeof vRaw === 'number' && isFinite(vRaw)) {
        const ms = Math.floor(vRaw * 24 * 60 * 60 * 1000); // quita fracción (hora)
        const d  = new Date(ms);
        raw[i][0] = new Date(d.getFullYear(), d.getMonth(), d.getDate());
        continue;
      }

      // Si no se pudo convertir, dejar como está
    }

    rng.setValues(raw);
    rng.setNumberFormat(DATE_FORMAT); // ejemplo: 'dd/MM/yyyy'
  });
}

/***** HELPERS *****/
// Convierte texto/fecha a Date sin hora (00:00). Tolera:
//  - "27/08/25 12:00 p. m." / "27/08/2025 06:00 PM"
//  - "27/08/2025", "27-08-2025", "27/08/25"
//  - "2025-08-27 12:00", "2025/08/27"
function toDateOnly_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  if (v == null) return null;

  let s = String(v).trim();
  if (!s) return null;

  // Normaliza "a. m." / "p. m." y espacios no estándar
  s = s
    .replace(/\ba\.?\s*m\.?\b/gi, 'AM')   // a. m. → AM
    .replace(/\bp\.?\s*m\.?\b/gi, 'PM')   // p. m. → PM
    .replace(/\u00A0/g, ' ')              // NBSP → espacio normal
    .replace(/\s+/g, ' ')
    .trim();

  // Caso 1: dd/MM/(yy|yyyy) al inicio (con o sin hora)
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2}|\d{4})/);
  if (m) {
    let d = parseInt(m[1], 10);
    let M = parseInt(m[2], 10) - 1;
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000; // 27/08/25 → 2025
    const out = new Date(y, M, d);
    return isNaN(out) ? null : out;
  }

  // Caso 2: yyyy-MM-dd (o '/')
  m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const y = parseInt(m[1], 10);
    const M = parseInt(m[2], 10) - 1;
    const d = parseInt(m[3], 10);
    const out = new Date(y, M, d);
    return isNaN(out) ? null : out;
  }

  // Intento genérico (último recurso)
  const dGeneric = new Date(s);
  if (!isNaN(dGeneric)) {
    return new Date(dGeneric.getFullYear(), dGeneric.getMonth(), dGeneric.getDate());
  }
  return null;
}
