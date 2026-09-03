const { readRange, writeRange } = require("./graph");

// Definición de cada hoja del Excel: nombre exacto de la pestaña y cantidad
// de columnas de datos. Los encabezados (fila 1) los escribís vos a mano una
// sola vez al crear el archivo — ver README.md para el texto exacto.
const SHEETS = {
  roster: { name: "Jugadores", cols: 7 },
  games: { name: "Partidas", cols: 8 },
  resultados: { name: "Resultados", cols: 12 },
  meta: { name: "Meta", cols: 2 },
};

// Cantidad de filas de datos reservadas por hoja (de sobra para un grupo de
// amigos jugando durante años). Si algún día lo superan, subir este número.
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

function rangeFor(cols) {
  return `A2:${colLetter(cols)}${MAX_ROWS + 1}`;
}

async function getRows(key) {
  const cfg = SHEETS[key];
  if (!cfg) throw new Error(`Hoja desconocida: ${key}`);
  const values = await readRange(cfg.name, rangeFor(cfg.cols));
  // Filtramos filas completamente vacías (el rango siempre trae MAX_ROWS filas)
  return values.filter((row) => row.some((c) => c !== "" && c !== null && c !== undefined));
}

async function setRows(key, rows) {
  const cfg = SHEETS[key];
  if (!cfg) throw new Error(`Hoja desconocida: ${key}`);
  const padded = rows.map((r) => {
    const row = r.slice(0, cfg.cols).map((v) => (v === undefined || v === null ? "" : v));
    while (row.length < cfg.cols) row.push("");
    return row;
  });
  while (padded.length < MAX_ROWS) padded.push(new Array(cfg.cols).fill(""));
  await writeRange(cfg.name, rangeFor(cfg.cols), padded.slice(0, MAX_ROWS));
}

module.exports = { SHEETS, MAX_ROWS, getRows, setRows, colLetter };
