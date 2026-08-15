const { getRows, setRows } = require("./lib/sheets");
const {
  rosterToRows, rowsToRoster,
  gamesToRows, rowsToGames,
  metaToRows, rowsToActiveGame,
  buildResultadosRows,
} = require("./lib/mapping");

const HEADERS = { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" };

exports.handler = async (event) => {
  const key = event.queryStringParameters && event.queryStringParameters.key;

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
      return respond(400, { error: `key inválida: ${key}` });
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
        await setRows("meta", metaToRows(value));
        return respond(200, { ok: true });
      }
      return respond(400, { error: `key inválida: ${key}` });
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
