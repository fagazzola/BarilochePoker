const { computeSettlement } = require("./settlement");

// Límite de seguridad por celda de Excel (~32,767 caracteres). Las fotos de
// avatar en base64 grandes se descartan al sincronizar para no romper el
// guardado; los íconos (emoji) no tienen este problema.
const MAX_CELL_CHARS = 30000;

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
        date: String(r[1] || ""),
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

/* ---------------- Meta (partida activa) ---------------- */
function metaToRows(activeGame) {
  return [["active", activeGame ? JSON.stringify(activeGame) : ""]];
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

module.exports = {
  rosterToRows, rowsToRoster,
  gamesToRows, rowsToGames,
  metaToRows, rowsToActiveGame,
  buildResultadosRows,
};
