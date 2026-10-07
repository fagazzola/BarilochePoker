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

function activeStore() {
  return getStore(STORE_NAME);
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
