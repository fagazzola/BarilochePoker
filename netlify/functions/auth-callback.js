// Esta función solo se usa UNA VEZ, para el setup inicial. Microsoft redirige
// acá después de que autorizás el acceso, con un "code" en la URL. Esta
// función lo cambia por un refresh_token, que es lo que hay que copiar a
// la variable de entorno MS_REFRESH_TOKEN en Netlify.
exports.handler = async (event) => {
  const code = event.queryStringParameters && event.queryStringParameters.code;
  const errorParam = event.queryStringParameters && event.queryStringParameters.error;

  if (errorParam) {
    return html(400, `<h2>Microsoft devolvió un error</h2><pre>${escapeHtml(JSON.stringify(event.queryStringParameters, null, 2))}</pre>`);
  }
  if (!code) {
    return html(400, "<h2>Falta el parámetro 'code' en la URL.</h2><p>Este link solo funciona si te redirige Microsoft después de iniciar sesión — no lo abras directamente.</p>");
  }

  const required = ["MS_CLIENT_ID", "MS_CLIENT_SECRET", "MS_REDIRECT_URI"];
  for (const k of required) {
    if (!process.env[k]) {
      return html(500, `<h2>Falta configurar ${k} en las variables de entorno de Netlify.</h2>`);
    }
  }

  const params = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID,
    client_secret: process.env.MS_CLIENT_SECRET,
    grant_type: "authorization_code",
    code,
    redirect_uri: process.env.MS_REDIRECT_URI,
    scope: "offline_access Files.ReadWrite",
  });

  const res = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params,
  });
  const data = await res.json();

  if (!res.ok) {
    return html(500, `<h2>Error al intercambiar el código</h2><pre>${escapeHtml(JSON.stringify(data, null, 2))}</pre>`);
  }

  return html(200, `
    <h2>Autorización completada ✅</h2>
    <p>Copiá este valor y guardalo como variable de entorno <code>MS_REFRESH_TOKEN</code>
    en Netlify (Site settings → Environment variables), después redeployá el sitio.</p>
    <textarea style="width:100%;height:140px;font-family:monospace;">${escapeHtml(data.refresh_token || "")}</textarea>
    <p style="color:#a33;">Este valor es sensible — no lo compartas ni lo dejes en capturas de pantalla.</p>
  `);
};

function html(statusCode, bodyHtml) {
  return {
    statusCode,
    headers: { "Content-Type": "text/html; charset=utf-8" },
    body: `<!doctype html><html><body style="font-family:sans-serif;max-width:640px;margin:40px auto;padding:0 16px;">${bodyHtml}</body></html>`,
  };
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
