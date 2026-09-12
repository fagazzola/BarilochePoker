const { readRange, writeRange } = require("./graph");

// Definición de cada hoja del Excel: nombre exacto de la pestaña y cantidad
// de columnas de datos. Los encabezados (fila 1) los escribís vos a mano una
// sola vez al crear el archivo — ver README.md para el texto exacto.
//
// "entregaFichas" es la hoja de auditoría de la pantalla "Entrega de
// fichas" (botón "Guardar en Excel ahora"): a diferencia de las demás hojas
// (que siempre se reescriben enteras con el último estado), esta se usa
// solo para AGREGAR filas (ver appendRows) — cada click del botón suma una
// nueva tanda de filas con su propio timestamp, así queda un historial para
// auditorías futuras en vez de perder los pushes anteriores. Por eso tiene
// muchas más filas reservadas (maxRows) que el resto.
const SHEETS = {
  roster: { name: "Jugadores", cols: 7 },
  games: { name: "Partidas", cols: 8 },
  resultados: { name: "Resultados", cols: 12 },
  meta: { name: "Meta", cols: 2 },
  entregaFichas: { name: "EntregaFichas", cols: 15, maxRows: 5000 },
};

// Cantidad de filas de datos reservadas por hoja, por defecto (de sobra para
// un grupo de amigos jugando durante años). Si algún día lo superan, subir
// este número o el "maxRows" propio de la hoja en SHEETS.
const MAX_ROWS = 500;

function colLetter(n) {
  let s = "";
  let num = n;
  while (num > 0) {
    const m = (num - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

function rangeFor(cols, maxRows) {
  return `A2:${colLetter(cols)}${(maxRows || MAX_ROWS) + 1}`;
}

async function getRows(key) {
  const cfg = SHEETS[key];
  if (!cfg) throw new Error(`Hoja desconocida: ${key}`);
  const maxRows = cfg.maxRows || MAX_ROWS;
  const values = await readRange(cfg.name, rangeFor(cfg.cols, maxRows));
  // Filtramos filas completamente vacías (el rango siempre trae maxRows filas)
  return values.filter((row) => row.some((c) => c !== "" && c !== null && c !== undefined));
}

async function setRows(key, rows) {
  const cfg = SHEETS[key];
  if (!cfg) throw new Error(`Hoja desconocida: ${key}`);
  const maxRows = cfg.maxRows || MAX_ROWS;
  const padded = rows.map((r) => {
    const row = r.slice(0, cfg.cols).map((v) => (v === undefined || v === null ? "" : v));
    while (row.length < cfg.cols) row.push("");
    return row;
  });
  while (padded.length < maxRows) padded.push(new Array(cfg.cols).fill(""));
  await writeRange(cfg.name, rangeFor(cfg.cols, maxRows), padded.slice(0, maxRows));
}

// Agrega filas nuevas al final de lo que ya haya escrito en la hoja, sin
// tocar lo anterior (a diferencia de setRows, que siempre reescribe la hoja
// entera). Pensada para hojas de auditoría/historial como "EntregaFichas".
async function appendRows(key, newRows) {
  const cfg = SHEETS[key];
  if (!cfg) throw new Error(`Hoja desconocida: ${key}`);
  if (!newRows || newRows.length === 0) return;
  const maxRows = cfg.maxRows || MAX_ROWS;
  const existing = await getRows(key);
  const startRow = existing.length + 2; // fila 1 = encabezados
  const endRow = startRow + newRows.length - 1;
  if (endRow > maxRows + 1) {
    throw new Error(
      `La hoja "${cfg.name}" llegó a su límite de ${maxRows} filas — hay que archivar filas viejas o ampliar el rango en el código.`
    );
  }
  const padded = newRows.map((r) => {
    const row = r.slice(0, cfg.cols).map((v) => (v === undefined || v === null ? "" : v));
    while (row.length < cfg.cols) row.push("");
    return row;
  });
  const address = `A${startRow}:${colLetter(cfg.cols)}${endRow}`;
  await writeRange(cfg.name, address, padded);
}

module.exports = { SHEETS, MAX_ROWS, getRows, setRows, appendRows, colLetter };
