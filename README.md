# BARILOCHE — Registro de Poker (con Excel como base de datos)

Esta versión guarda todo en un archivo de Excel en tu OneDrive, a través de
funciones de Netlify que hablan con Microsoft Graph. Cualquiera con el link
del sitio ve y edita los mismos datos.

## 1. Crear el archivo Excel

En tu OneDrive, creá un archivo llamado **`BarilochePoker.xlsx`** (en la raíz,
o ajustá `MS_EXCEL_PATH` si lo ponés en una carpeta). Adentro, creá 4 hojas
con estos nombres **exactos** y estos encabezados en la fila 1:

**Hoja "Jugadores"**
| id | nombre | activo | avatarTipo | avatarValor | fechaAlta |
|----|--------|--------|------------|-------------|-----------|

**Hoja "Partidas"**
| id | fecha | loteValue | rake | hostId | playerIds | finished | detalleJson |
|----|-------|-----------|------|--------|-----------|----------|--------------|

**Hoja "Resultados"** (se regenera sola, es solo para que la mires — no hace falta tocarla)
| gameId | fecha | jugador | lotesCash | buyInCash | lotesVirtual | buyInVirtual | totalBuyIn | cashOut | pagoCash | pagoTransfer | balance |
|--------|-------|---------|-----------|-----------|--------------|--------------|------------|---------|----------|--------------|---------|

**Hoja "Meta"**
| key | value |
|-----|-------|

No hace falta escribir nada más — la app llena las filas de datos sola.

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
