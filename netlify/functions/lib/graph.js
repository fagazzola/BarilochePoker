const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

// Intercambia el refresh token guardado en las variables de entorno de Netlify
// por un access token de corta duración. Se pide uno nuevo en cada invocación
// (las funciones son stateless) — para un grupo de amigos el volumen de
// llamadas es bajísimo, así que no hace falta cachear.
async function getAccessToken() {
  const required = ["MS_CLIENT_ID", "MS_CLIENT_SECRET", "MS_REFRESH_TOKEN"];
  for (const key of required) {
    if (!process.env[key]) {
      throw new Error(`Falta configurar la variable de entorno ${key} en Netlify.`);
    }
  }
  const params = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID,
    client_secret: process.env.MS_CLIENT_SECRET,
    grant_type: "refresh_token",
    refresh_token: process.env.MS_REFRESH_TOKEN,
    scope: "offline_access Files.ReadWrite",
  });
  const res = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params,
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(
      "No se pudo renovar el token de Microsoft. Es probable que el refresh token haya expirado " +
      "(pasa si la app no se usa por 90+ días) — hay que rehacer la autorización una vez. Detalle: " +
      JSON.stringify(data)
    );
  }
  // Nota: si en el futuro Microsoft rota el refresh_token, lo ideal es reemplazar
  // MS_REFRESH_TOKEN en Netlify con data.refresh_token. Por ahora se ignora
  // porque para uso personal el token entregado suele mantenerse válido.
  return data.access_token;
}

function workbookPath(subpath) {
  const filePath = process.env.MS_EXCEL_PATH || "/BarilochePoker.xlsx";
  const normalized = filePath.startsWith("/") ? filePath : `/${filePath}`;
  return `${GRAPH_BASE}/me/drive/root:${normalized}:/workbook${subpath}`;
}

async function graphFetch(subpath, options = {}) {
  const token = await getAccessToken();
  const res = await fetch(workbookPath(subpath), {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Error de Microsoft Graph (${res.status}): ${text}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function readRange(sheet, address) {
  const data = await graphFetch(`/worksheets('${sheet}')/range(address='${address}')`);
  return (data && data.values) || [];
}

async function writeRange(sheet, address, values) {
  return graphFetch(`/worksheets('${sheet}')/range(address='${address}')`, {
    method: "PATCH",
    body: JSON.stringify({ values }),
  });
}

module.exports = { getAccessToken, readRange, writeRange, graphFetch };
