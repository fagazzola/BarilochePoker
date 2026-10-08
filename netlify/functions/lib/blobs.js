const { getStore } = require("@netlify/blobs");

// Blobs de Netlify: almacenamiento nativo del propio sitio (sin OAuth, sin
// Microsoft Graph, sin límites de frecuencia ni refresh tokens que vencen).
// Se usa exclusivamente para el estado de la PARTIDA EN CURSO — todo lo que
// cambia todo el tiempo mientras se juega (compras, cena, lote/rake, log de
// compras, borrador de "Entrega de fichas", ajustes, etc. — en la práctica,
// el objeto `game` completo que antes se mandaba a cada rato a la hoja
// "Meta" del Excel). Al escribir en cada click directo a Microsoft Graph es
// donde aparecía la intermitencia (un guardado que "pega" y después
// desaparece): Graph tiene latencia variable y un flujo de OAuth con
// refresh token de por medio, nada ideal para guardar en cada tecla.
//
// El Excel sigue siendo el registro DEFINITIVO de cada partida, pero ahora
// se escribe ahí una sola vez: al cerrar formalmente la partida (ver
// `pushExcelAudit` en src/App.jsx, y las keys "games"/"entregaFichas"/
// "logPetLotes" de store.js, todas disparadas juntas en ese momento).
const STORE_NAME = "bariloche-active";
const ACTIVE_KEY = "active-game";

// "strong": siempre lee el último valor escrito (consistencia "read your
// writes"), al costo de un poquito más de latencia por request. Por default,
// Netlify Blobs usa consistencia "eventual" (más rápida, pero una lectura
// justo después de escribir puede traer todavía el valor viejo por unos
// segundos) — eso, combinado con el polling cada 4s de la app, podía producir
// una lectura vieja que pisara un cambio recién guardado. Acá no hay margen
// para esa demora: la partida en curso necesita que cada lectura refleje
// siempre la última escritura.
// En un sitio desplegado normalmente por Netlify, getStore() recibe solo
// (sin hacer nada extra) el contexto necesario para hablar con Blobs —
// Netlify se lo inyecta solo a la función en tiempo de ejecución. En
// algunos sitios/despliegues esa inyección automática no llega, y
// getStore() tira: "The environment has not been configured to use
// Netlify Blobs. To use it manually, supply the following properties
// when creating a store: siteID, token". Para esos casos se puede pasar
// el Site ID y un token a mano, vía 2 variables de entorno opcionales
// (ver README.md, sección "Blobs: configuración manual"):
//   NETLIFY_SITE_ID    — Site settings → General → Site details → Site ID
//   NETLIFY_BLOBS_TOKEN — User settings → Applications → Personal access tokens → New access token
// Si no están cargadas, se sigue intentando la configuración automática
// de siempre (no rompe nada en los sitios donde sí funciona sola).
function activeStore() {
  const opts = { name: STORE_NAME, consistency: "strong" };
  if (process.env.NETLIFY_SITE_ID && process.env.NETLIFY_BLOBS_TOKEN) {
    opts.siteID = process.env.NETLIFY_SITE_ID;
    opts.token = process.env.NETLIFY_BLOBS_TOKEN;
  }
  return getStore(opts);
}

async function getActiveGame() {
  const value = await activeStore().get(ACTIVE_KEY, { type: "json" });
  return value === null || value === undefined ? null : value;
}

async function setActiveGame(value) {
  const s = activeStore();
  if (value === null || value === undefined) {
    // Partida cerrada/cancelada: no tiene sentido dejar un blob viejo
    // dando vueltas — se borra en vez de guardar un "null".
    await s.delete(ACTIVE_KEY);
    return;
  }
  await s.setJSON(ACTIVE_KEY, value);
}

module.exports = { getActiveGame, setActiveGame };
