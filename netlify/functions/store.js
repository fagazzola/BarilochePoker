const { getRows, setRows } = require("./lib/sheets");
const {
  rosterToRows, rowsToRoster,
  gamesToRows, rowsToGames,
  metaToRows, rowsToActiveGame,
  buildResultadosRows,
} = require("./lib/mapping");

const HEADERS = { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" };

// La app (src/App.jsx) manda claves como "poker-roster", "poker-games" y
// "poker-active-game" (definidas en la constante KEYS). Acá se normalizan a
// los nombres internos que usa este archivo. También se aceptan los nombres
// cortos ("roster","games","active") para poder probar a mano desde el navegador.
const KEY_MAP = {
  "poker-roster": "roster",
  "poker-games": "games",
  "poker-active-game": "active",
  "roster": "roster",
  "games": "games",
  "active": "active",
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
        const rows = await getRows("meta");
        return respond(200, { value: rowsToActiveGame(rows) });
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
          console.error("No se pudo regenerar la hoja Resultados:", e);
        }
        return respond(200, { ok: true });
      }
      if (key === "active") {
        await setRows("meta", metaToRows(value));
        return respond(200, { ok: true });
      }
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
