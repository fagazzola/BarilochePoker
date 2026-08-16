const { computeSettlement } = require("./settlement");

// Límite de seguridad por celda de Excel (~32,767 caracteres). Las fotos de
// avatar en base64 grandes se descartan al sincronizar para no romper el
// guardado; los íconos (emoji) no tienen este problema.
const MAX_CELL_CHARS = 30000;

// Excel a veces "reconoce" un texto tipo 2026-08-14 como una fecha real y lo
// guarda internamente como su número de serie (días desde el 30/12/1899).
// Si eso pasa, Graph API nos devuelve ese número en vez del texto. Esta
// función lo detecta y lo reconstruye de vuelta a YYYY-MM-DD.
function normalizeDateCell(val) {
  if (typeof val === "number" && isFinite(val)) {
    const utcDays = Math.floor(val - 25569); // 25569 = días entre 1900-01-01 y 1970-01-01 (con el bug de Excel)
    const ms = utcDays * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  return String(val || "");
}

/* ---------------- Jugadores ---------------- */
function rosterToRows(roster) {
  return (roster || []).map((p) => {
    let avatarType = p.avatar?.type || "";
    let avatarValue = p.avatar?.value || "";
    if (avatarValue.length > MAX_CELL_CHARS) {
      // Foto demasiado grande para una celda de Excel: se omite (el jugador
      // se sincroniza igual, solo pierde el avatar en esta copia).
      avatarType = "";
      avatarValue = "";
    }
    return [
      p.id,
      p.name,
      p.active ? "SI" : "NO",
      avatarType,
      avatarValue,
      p.createdAt ? new Date(p.createdAt).toISOString() : "",
    ];
  });
}
function rowsToRoster(rows) {
  return rows
    .map((r) => ({
      id: String(r[0] || ""),
      name: String(r[1] || ""),
      active: String(r[2] || "").toUpperCase() === "SI",
      avatar: r[3] ? { type: String(r[3]), value: String(r[4] || "") } : null,
      createdAt: r[5] ? (Date.parse(r[5]) || Date.now()) : Date.now(),
    }))
    .filter((p) => p.id && p.name);
}

/* ---------------- Partidas ---------------- */
function gamesToRows(games) {
  return (games || []).map((g) => [
    g.id,
    g.date,
    g.loteValue,
    g.rake,
    g.hostId,
    JSON.stringify(g.playerIds || []),
    g.finished ? "SI" : "NO",
    JSON.stringify({
      purchases: g.purchases || [],
      dinner: g.dinner || { total: 0, waiter: 0, alcoholFee: 0, alcohol: {}, paid: {}, paymentMethod: {} },
      finalChips: g.finalChips || {},
    }),
  ]);
}
function rowsToGames(rows) {
  return rows
    .map((r) => {
      let detalle = {};
      try { detalle = JSON.parse(r[7] || "{}"); } catch { detalle = {}; }
      let playerIds = [];
      try { playerIds = JSON.parse(r[5] || "[]"); } catch { playerIds = []; }
      return {
        id: String(r[0] || ""),
        date: normalizeDateCell(r[1]),
        loteValue: Number(r[2]) || 0,
        rake: Number(r[3]) || 0,
        hostId: String(r[4] || ""),
        playerIds,
        finished: String(r[6] || "").toUpperCase() === "SI",
        purchases: detalle.purchases || [],
        dinner: detalle.dinner || { total: 0, waiter: 0, alcoholFee: 0, alcohol: {}, paid: {}, paymentMethod: {} },
        finalChips: detalle.finalChips || {},
        results: null,
      };
    })
    .filter((g) => g.id);
}

/* ---------------- Meta (partida activa + contraseña de administrador) ---------------- */
// IMPORTANTE: setRows reescribe la hoja entera, así que siempre hay que
// incluir la fila de la contraseña al guardar "active", o se perdería.
function metaToRows(activeGame, adminPassword) {
  return [
    ["active", activeGame ? JSON.stringify(activeGame) : ""],
    ["admin_password", adminPassword || ""],
  ];
}
function rowsToActiveGame(rows) {
  const row = rows.find((r) => String(r[0]) === "active");
  if (!row || !row[1]) return null;
  try {
    return JSON.parse(row[1]);
  } catch {
    return null;
  }
}
function rowsToAdminPassword(rows) {
  const row = rows.find((r) => String(r[0]) === "admin_password");
  return row ? String(row[1] || "") : "";
}

/* ---------------- Resultados (espejo legible, se regenera entero) ---------------- */
function buildResultadosRows(games, roster) {
  const rows = [];
  (games || [])
    .filter((g) => g.finished)
    .forEach((g) => {
      const r = computeSettlement(g, roster);
      r.players.forEach((p) => {
        const cashLotes = (g.purchases || [])
          .filter((pu) => pu.playerId === p.playerId && pu.type === "cash")
          .reduce((s, pu) => s + pu.lotes, 0);
        const virtualLotes = (g.purchases || [])
          .filter((pu) => pu.playerId === p.playerId && pu.type === "virtual")
          .reduce((s, pu) => s + pu.lotes, 0);
        rows.push([
          g.id,
          g.date,
          p.name,
          cashLotes,
          p.cashAmount,
          virtualLotes,
          p.virtualAmount,
          p.totalBuyIn,
          p.cashOut,
          p.pagoCash,
          p.pagoTransfer,
          p.balance,
        ]);
      });
    });
  return rows;
}

// Convierte las filas crudas de la hoja "Resultados" a objetos legibles, para
// que el dashboard de estadísticas los consuma directo sin tener que
// recalcular nada.
function rowsToResultados(rows) {
  return rows
    .map((r) => ({
      gameId: String(r[0] || ""),
      fecha: normalizeDateCell(r[1]),
      jugador: String(r[2] || ""),
      lotesCash: Number(r[3]) || 0,
      buyInCash: Number(r[4]) || 0,
      lotesVirtual: Number(r[5]) || 0,
      buyInVirtual: Number(r[6]) || 0,
      totalBuyIn: Number(r[7]) || 0,
      cashOut: Number(r[8]) || 0,
      pagoCash: Number(r[9]) || 0,
      pagoTransfer: Number(r[10]) || 0,
      balance: Number(r[11]) || 0,
    }))
    .filter((row) => row.gameId && row.jugador);
}

module.exports = {
  rosterToRows, rowsToRoster,
  gamesToRows, rowsToGames,
  metaToRows, rowsToActiveGame, rowsToAdminPassword,
  buildResultadosRows, rowsToResultados,
};
