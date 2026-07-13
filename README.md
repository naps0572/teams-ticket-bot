# Bot de tickets para Teams (OpenRouter + ITSM)

Chatbot para Microsoft Teams que conversa con el usuario, entiende su problema
usando un modelo de **OpenRouter**, y cuando tiene la información suficiente crea
un ticket en tu **ITSM** vía API. Construido con el **Teams SDK v2**
(`@microsoft/teams.*`) en Node.js / TypeScript.

## Cómo funciona

```
Usuario en Teams → Azure Bot Service → este backend (Teams SDK)
                                          ├─ OpenRouter  (conversa y extrae los campos del ticket)
                                          └─ API ITSM    (crea el ticket)
                                          ↓
                       Tarjeta (Adaptive Card) con el número de ticket
```

El backend mantiene el estado de cada conversación, le pide al modelo que
devuelva un JSON con `{ respuesta, listo, ticket }`, y cuando `listo` es `true`
llama al conector del ITSM y responde con una tarjeta de confirmación.

## Estructura

```
teams-ticket-bot/
├── src/
│   ├── index.ts         # Arranque de la app + manejador de mensajes
│   ├── openrouter.ts    # Cliente de OpenRouter (conversación + extracción)
│   ├── itsm.ts          # Conector del ITSM  ← AQUÍ va tu API
│   ├── conversation.ts  # Estado de conversación (en memoria)
│   ├── cards.ts         # Tarjeta de confirmación del ticket
│   └── types.ts         # Tipos compartidos
├── .env.example
├── package.json
└── tsconfig.json
```

## Requisitos

- Node.js 20 o superior.

## 1) Probar en local (sin Teams, en 2 minutos)

El Teams SDK trae **DevTools**, una interfaz web para chatear con el bot sin
desplegarlo. En este modo, si no configuras el ITSM, el bot corre en **dry-run**
y devuelve tickets simulados (`LOCAL-xxxx`), así pruebas toda la conversación.

```bash
npm install
cp .env.example .env      # y pon tu OPENROUTER_API_KEY
npm run dev
```

Abre **http://localhost:3979/devtools** y escríbele, por ejemplo:
*"No me carga el correo en Outlook desde esta mañana"*. El bot te hará las
preguntas que falten y al final "creará" el ticket.

## 2) Conectar tu ITSM

Toda la integración vive en `src/itsm.ts`, en la función `createTicket`. Está
marcado con `TODO` lo que hay que ajustar:

- **endpoint**: la URL y ruta de creación de tickets.
- **headers**: el tipo de autenticación (API key, Bearer, etc.).
- **body**: el mapeo de campos (`asunto`, `descripcion`, `categoria`,
  `prioridad`, correo del solicitante) a los nombres que espera tu API.
- **respuesta**: de dónde sacar el número/ID y la URL del ticket creado.

Luego define en `.env`:

```
ITSM_BASE_URL=https://tu-itsm.ejemplo.com
ITSM_API_KEY=...
```

En cuanto `ITSM_BASE_URL` esté definido, el bot deja el dry-run y llama a tu API.

## 3) Llevarlo a Teams

Para que viva dentro de Teams necesitas dos cosas:

1. **Un recurso Azure Bot** (capa gratuita F0 disponible). Registra la app en
   Entra ID, crea el recurso Azure Bot, activa el canal **Microsoft Teams** y
   apunta el *messaging endpoint* a `https://TU-DOMINIO/api/messages`. De ahí
   salen `CLIENT_ID`, `CLIENT_SECRET` y `TENANT_ID` para tu `.env`.
2. **Un paquete de app de Teams** (`manifest.json` + iconos) para instalar el
   bot. La forma más rápida de generarlo y depurar es el **Microsoft 365 Agents
   Toolkit** (extensión de VS Code) o el CLI del SDK:

   ```bash
   npm install -g @microsoft/teams.cli
   teams new typescript scaffold --template echo
   ```

   Eso crea un `appPackage/` válido; copia ese `appPackage/` a este proyecto y
   reemplaza el `src/` generado por el de aquí.

> Para exponer tu `localhost` a Teams durante pruebas puedes usar un túnel
> (por ejemplo el *dev tunnel* del propio toolkit o ngrok) y poner esa URL
> pública como messaging endpoint del Azure Bot.

## Notas para producción

- **Quita DevTools**: solo debe usarse en desarrollo (no tiene autenticación).
  El código ya lo activa únicamente cuando `NODE_ENV` no es `production`.
- **Estado persistente**: `conversation.ts` guarda el estado en memoria. Para
  varias instancias o reinicios sin perder contexto, cámbialo por Redis, Cosmos
  DB o una tabla SQL manteniendo la misma interfaz.
- **Credenciales**: configura `CLIENT_ID` / `CLIENT_SECRET` / `TENANT_ID` para
  que el endpoint `/api/messages` valide que las peticiones vienen de Teams.

## Comandos

| Comando          | Qué hace                                        |
|------------------|-------------------------------------------------|
| `npm run dev`    | Desarrollo con recarga (nodemon + DevTools)     |
| `npm run build`  | Compila TypeScript a `dist/`                    |
| `npm start`      | Ejecuta la versión compilada                    |
| `npm run typecheck` | Verifica tipos sin compilar                  |

Dentro del chat, escribe `reiniciar` (o `reset` / `cancelar`) para empezar un
ticket nuevo.
