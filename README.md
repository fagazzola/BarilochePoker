# BARILOCHE — Registro de Poker (con Excel como base de datos)

El Excel en tu OneDrive (vía Microsoft Graph) es el registro **definitivo**
de cada partida, pero ya no se toca mientras se juega. Mientras una partida
está en curso, todo (compras, cena, lote/rake, log de compras, entrega de
fichas) vive en un **blob de Netlify** — almacenamiento propio del sitio,
sin OAuth ni límites de Microsoft Graph, mucho más rápido y sin la
intermitencia de "se guarda y después desaparece". Recién al cerrar
formalmente la partida se escribe todo de una vez al Excel: la fila de
"Partidas"/"Resultados" y, en bloque, las hojas de auditoría "EntregaFichas"
y "LogPetLotes". Cualquiera con el link del sitio ve y edita los mismos
datos.

## 1. Crear el archivo Excel

En tu OneDrive, creá un archivo llamado **`BarilochePoker.xlsx`** (en la raíz,
o ajustá `MS_EXCEL_PATH` si lo ponés en una carpeta). Adentro, creá 4 hojas
con estos nombres **exactos** y estos encabezados en la fila 1:

**Hoja "Jugadores"**
| id | nombre | activo | avatarTipo | avatarValor | fechaAlta | pin |
|----|--------|--------|------------|-------------|-----------|-----|

**Hoja "Partidas"**
| id | fecha | loteValue | rake | hostId | playerIds | finished | detalleJson |
|----|-------|-----------|------|--------|-----------|----------|--------------|

**Hoja "Resultados"** (se regenera sola, es solo para que la mires — no hace falta tocarla)
| gameId | fecha | jugador | lotesCash | buyInCash | lotesVirtual | buyInVirtual | totalBuyIn | cashOut | pagoCash | pagoTransfer | balance |
|--------|-------|---------|-----------|-----------|--------------|--------------|------------|---------|----------|--------------|---------|

**Hoja "Meta"**
| key | value |
|-----|-------|

**Hoja "EntregaFichas"** (auditoría — se le agrega un snapshot una sola vez, al cerrar formalmente cada partida; nunca se sobreescribe)
| timestamp | gameId | fecha | jugadorId | jugador | debeVirtual | pagaVirtual | fichasRemanentes | ajusteManual | fichasTotales | rake | cashDisponible | virtualPendiente | totalA | totalB |
|-----------|--------|-------|-----------|---------|-------------|-------------|-------------------|---------------|----------------|------|-----------------|-------------------|--------|--------|

**Hoja "LogPetLotes"** (auditoría — se le agrega el log completo de la noche una sola vez, al cerrar formalmente cada partida, un renglón por cada evento de compra/solicitud de lotes, con hora exacta; nunca se sobreescribe)
| timestamp | gameId | fecha | jugadorId | jugador | tipo | accion | origen | lotes | monto | valorLote |
|-----------|--------|-------|-----------|---------|------|--------|--------|-------|-------|-----------|

`tipo` es `cash` o `virtual`. `accion` es una de: `compra directa` (el host le suma un lote a mano), `solicitud enviada`, `solicitud aprobada`, `solicitud rechazada` (pedido de un jugador desde su celular) o `cancelación` (el host deshace la última compra de ese tipo). `origen` es `host` o `jugador`. `lotes`/`monto` van en negativo en una `cancelación`; en una `solicitud enviada` o `solicitud rechazada` llevan el monto que se pidió aunque no haya generado ninguna compra real (la compra real solo queda contabilizada en `compra directa` y `solicitud aprobada`). Por eso esta hoja es un registro narrativo para auditar la actividad — para totales de buy-in siempre hay que usar la hoja "Partidas" (columna `detalleJson.purchases`) o "Resultados", nunca sumar esta.

Si ese empuje en bloque falla al cerrar una partida (por ejemplo, sin
conexión un instante), la app lo avisa con un botón "Reintentar" en la
pantalla de resultados y en el histórico — no se pierde nada, el detalle ya
quedó guardado dentro de la propia partida.

No hace falta escribir nada más — la app llena las filas de datos sola.

## Partida en curso: blob de Netlify, no Excel

Mientras una partida está en curso (botón "active" en `src/App.jsx`), la app
no toca el Excel en absoluto: lee y escribe un blob de Netlify (ver
`netlify/functions/lib/blobs.js`, paquete `@netlify/blobs`). En la gran
mayoría de los sitios esto funciona solo, sin ninguna variable de entorno —
Netlify le inyecta el contexto necesario a la función automáticamente. La
hoja "Meta" del Excel sigue existiendo solo para la contraseña de
administrador (fila `admin_password`, la cargás a mano); la fila vieja
`active` que pudiera haber quedado ahí de versiones anteriores de la app ya
no se lee ni se escribe — es basura inofensiva, se puede borrar a mano o
dejar como está.

### Blobs: configuración manual (solo si ves "problemas de conexión")

Si la app te muestra un error como:

> poker-active-game: The environment has not been configured to use Netlify
> Blobs. To use it manually, supply the following properties when creating
> a store: siteID, token

significa que en tu sitio esa inyección automática no está llegando, y hay
que indicarle a mano a qué sitio conectarse. Cargá estas 2 variables en
**Site settings → Environment variables** (igual que las `MS_*`):

- **`NETLIFY_SITE_ID`** — en el panel de Netlify, andá a **Site settings →
  General → Site details**, copiá el **"Site ID"** (es un código largo tipo
  `a1b2c3d4-...`).
- **`NETLIFY_BLOBS_TOKEN`** — hace falta un token personal de Netlify.
  Arriba a la derecha, tu avatar de usuario → **User settings** →
  **Applications** → sección "Personal access tokens" → **New access
  token**. Le ponés cualquier nombre (por ej. `bariloche-blobs`) y Netlify
  te muestra el token **una sola vez**: copialo entero y pegalo en esta
  variable.

Redeployá el sitio después de cargar las 2 variables (Netlify no las toma
hasta el próximo deploy). Con eso `getStore()` ya no depende de la
inyección automática y el error desaparece.

## 2. Desplegar en Netlify

1. Subí esta carpeta a un repo de GitHub y conectalo en Netlify (o usá
   `netlify deploy` con la CLI desde acá). Netlify va a correr `npm run build`
   y desplegar `dist/` + las funciones de `netlify/functions`.
2. Una vez que el sitio tenga una URL (`https://TU-SITIO.netlify.app`), andá a
   **Site settings → Environment variables** y cargá:
   - `MS_CLIENT_ID` — el que copiaste de Azure
   - `MS_CLIENT_SECRET` — el que generaste en Azure
   - `MS_REDIRECT_URI` — `https://TU-SITIO.netlify.app/api/auth-callback`
   - `MS_EXCEL_PATH` — `/BarilochePoker.xlsx` (o la ruta donde lo creaste)
   - `MS_REFRESH_TOKEN` — lo vamos a completar en el paso 4
3. Redeployá el sitio para que tome las variables.

## 3. Terminar el registro en Azure

Volvé a la app registration que creaste (`portal.azure.com` → Entra ID → App
registrations → tu app) → **Authentication** → **Add a platform** → **Web**,
y pegá ahí el mismo valor de `MS_REDIRECT_URI` de arriba.

## 4. Autorización única (para obtener el refresh token)

Con el `MS_CLIENT_ID` y el `MS_REDIRECT_URI` armá esta URL (reemplazando esos
dos valores) y abrila en el navegador, logueada con tu cuenta de Microsoft:

```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=TU_CLIENT_ID&response_type=code&redirect_uri=TU_REDIRECT_URI&response_mode=query&scope=offline_access%20Files.ReadWrite
```

Te va a pedir iniciar sesión y aceptar permisos. Al aceptar, te redirige a
`/api/auth-callback`, que te muestra un `refresh_token` en pantalla. Copiá ese
valor completo y pegalo en la variable `MS_REFRESH_TOKEN` en Netlify. Redeployá
una última vez.

Después de esto, la app ya lee y escribe directo en tu Excel. Si en unos meses
ves errores de "token expirado", es porque el refresh token venció por
inactividad — se repite solo este paso 4.

## Notas

- Las fotos de avatar muy grandes no se guardan en Excel (hay un límite de
  tamaño por celda) — para el roster funciona mejor usar los íconos.
- La hoja "Resultados" se reescribe entera cada vez que se cierra una partida,
  no hace falta editarla a mano.
