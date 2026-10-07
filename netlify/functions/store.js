const { getRows, setRows, appendRows } = require("./lib/sheets");
const {
  rosterToRows, rowsToRoster,
  gamesToRows, rowsToGames,
  rowsToAdminPassword,
  buildResultadosRows, rowsToResultados,
  entregaFichasToRows,
  logPetLotesToRows,
} = require("./lib/mapping");
const { getActiveGame, setActiveGame } = require("./lib/blobs");

const HEADERS = { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" };

// La app (src/App.jsx) manda claves como "poker-roster", "poker-games",
// "poker-active-game" y "poker-admin-password" (definidas en la constante
// KEYS). Acá se normalizan a los nombres internos que usa este archivo.
// También se aceptan los nombres cortos para poder probar a mano desde el navegador.
const KEY_MAP = {
  "poker-roster": "roster",
  "poker-games": "games",
  "poker-active-game": "active",
  "poker-admin-password": "adminPassword",
  "poker-resultados": "resultados",
  "poker-entrega-fichas": "entregaFichas",
  "poker-log-pet-lotes": "logPetLotes",
  "roster": "roster",
  "games": "games",
  "active": "active",
  "adminPassword": "adminPassword",
  "resultados": "resultados",
  "entregaFichas": "entregaFichas",
  "logPetLotes": "logPetLotes",
};

exports.handler = async (event) => {
  const rawKey = event.queryStringParameters && event.queryStringParameters.key;
  const key = KEY_MAP[rawKey];

  try {
    if (event.httpMethod === "GET") {
      if (key === "roster") {
        const rows = await getRows("roster");
        return respond(200, { value: rowsToRoster(rows) });
      }
      if (key === "games") {
        const rows = await getRows("games");
        return respond(200, { value: rowsToGames(rows) });
      }
      if (key === "active") {
        // La partida EN CURSO vive en un blob de Netlify, no en Excel — ver
        // lib/blobs.js. Mucho más rápido y sin la intermitencia de Microsoft
        // Graph; el Excel recién se escribe una sola vez, al cierre formal.
        const value = await getActiveGame();
        return respond(200, { value });
      }
      if (key === "adminPassword") {
        // Contraseña de administrador: se lee de la hoja "Meta" (fila
        // key=admin_password), la carga el dueño del Excel a mano — la app
        // nunca la escribe, solo la lee para comparar.
        const rows = await getRows("meta");
        return respond(200, { value: rowsToAdminPassword(rows) });
      }
      if (key === "resultados") {
        // Solo lectura: para el dashboard de estadísticas. Nunca se escribe
        // acá — la hoja "Resultados" siempre se regenera desde "games".
        const rows = await getRows("resultados");
        return respond(200, { value: rowsToResultados(rows) });
      }
      return respond(400, { error: `key inválida: ${rawKey}` });
    }

    if (event.httpMethod === "POST") {
      const body = JSON.parse(event.body || "{}");
      const value = body.value;

      if (key === "roster") {
        await setRows("roster", rosterToRows(value || []));
        return respond(200, { ok: true });
      }
      if (key === "games") {
        await setRows("games", gamesToRows(value || []));
        // Regenerar la hoja "Resultados" (espejo legible) a partir de las
        // partidas finalizadas, para que siempre quede al día en el Excel.
        try {
          const rosterRows = await getRows("roster");
          const roster = rowsToRoster(rosterRows);
          await setRows("resultados", buildResultadosRows(value || [], roster));
        } catch (e) {
          // No queremos que un fallo al regenerar "Resultados" tumbe el guardado
          // de las partidas en sí — se loguea y se sigue.
          console.error("No se pudo regenerar la hoja Resultados:", e);
        }
        return respond(200, { ok: true });
      }
      if (key === "active") {
        // Partida en curso: se guarda en el blob de Netlify, nunca en Excel.
        await setActiveGame(value === undefined ? null : value);
        return respond(200, { ok: true });
      }
      if (key === "entregaFichas") {
        // Se dispara UNA sola vez, en bloque, justo al cerrar formalmente la
        // partida (ya no en cada click de un botón "guardar" durante el
        // juego): agrega el snapshot final a la hoja "EntregaFichas" (nunca
        // sobreescribe lo ya guardado), para dejar un historial auditable.
        const { gameId, gameDate, rows } = body;
        if (!gameId || !Array.isArray(rows)) {
          return respond(400, { error: "faltan gameId o rows para entregaFichas" });
        }
        await appendRows("entregaFichas", entregaFichasToRows(gameId, gameDate, rows));
        return respond(200, { ok: true });
      }
      if (key === "logPetLotes") {
        // Igual que "entregaFichas": se dispara una sola vez al cerrar la
        // partida, con el log completo de toda la noche (ya no evento por
        // evento mientras se juega) — se agrega como filas nuevas a
        // "LogPetLotes", nunca se sobreescribe lo ya guardado.
        const { gameId, gameDate, rows } = body;
        if (!gameId || !Array.isArray(rows)) {
          return respond(400, { error: "faltan gameId o rows para logPetLotes" });
        }
        await appendRows("logPetLotes", logPetLotesToRows(gameId, gameDate, rows));
        return respond(200, { ok: true });
      }
      // No se expone un POST para "adminPassword" ni "resultados": la primera
      // se define a mano en el Excel, la segunda se regenera sola desde "games".
      return respond(400, { error: `key inválida: ${rawKey}` });
    }

    return respond(405, { error: "Método no soportado" });
  } catch (e) {
    console.error(e);
    return respond(500, { error: String((e && e.message) || e) });
  }
};

function respond(statusCode, obj) {
  return { statusCode, headers: HEADERS, body: JSON.stringify(obj) };
}
