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
    const rng = sh.getRange(HEADER_ROW + 1, col, numRows, 1);
    const vals = rng.getValues();

    for (let i = 0; i < vals.length; i++) {
      const v = vals[i][0];
      const d = toDateOnly_(v);   // ← convierte texto/fecha a Date sin hora (00:00)
      if (d) vals[i][0] = d;      // si no pudo convertir, deja el valor tal cual
    }

    rng.setValues(vals);
    rng.setNumberFormat(DATE_FORMAT); // mostrar sin hora
  });
}

/***** HELPERS *****/
// Convierte a Date sin hora. Acepta Date, número serial o textos como "25 July 2023 09:00 a. m."
function toDateOnly_(v) {
  if (v instanceof Date) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  if (typeof v === 'number') {
    // número serial de Sheets: quitar fracción (hora)
    return new Date(Math.floor(v * 24 * 60 * 60 * 1000));
  }
  if (typeof v === 'string') {
    let s = v.trim();
    if (!s) return null;
    // normalizar "a. m." / "p. m." → AM/PM y quitar puntos
    s = s.replace(/\ba\.?\s*m\.?\b/gi, 'AM')
         .replace(/\bp\.?\s*m\.?\b/gi, 'PM')
         .replace(/\s+/g, ' ')
         .replace(/\.\b/g, '');

    // Intento directo
    let d = new Date(s);
    if (!isNaN(d)) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }

    // Intento extra: "25 July 2023" (sin hora)
    const m = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
    if (m) {
      d = new Date(`${m[2]} ${m[1]}, ${m[3]}`);
      if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
  }
  return null; // no convertible
}
