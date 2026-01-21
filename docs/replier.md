# Replier - Sistema de Sincronización del Recinto

## Descripción General

**Replier** es un servicio Node.js que mantiene el estado del Recinto (Hemiciclo) en memoria y notifica los cambios a webhooks suscriptos. Funciona como el hub central de sincronización de estado de las bancas legislativas.

- **Versión:** 1.1.0
- **Puerto por defecto:** 17001
- **Protocolo:** HTTPS + WebSocket

## Arquitectura

```
┌─────────────────────────────────────────────────────────────┐
│                     REPLIER SERVICE                          │
├─────────────────────────────────────────────────────────────┤
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────────┐  │
│  │   Recinto   │───▶│ EventEmitter│───▶│     Replier     │  │
│  │  (Estado)   │    │  (Eventos)  │    │    (Daemon)     │  │
│  └─────────────┘    └─────────────┘    └────────┬────────┘  │
│         │                                        │           │
│         ▼                                        ▼           │
│  ┌─────────────┐                       ┌─────────────────┐  │
│  │  WebSocket  │                       │ WebhookPublisher│  │
│  │ (Broadcast) │                       │   (Salida)      │  │
│  └─────────────┘                       └─────────────────┘  │
├─────────────────────────────────────────────────────────────┤
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────────┐  │
│  │  Votacion   │    │  Commander  │    │  API REST v1/v2 │  │
│  │  (Dominio)  │    │ (Emulador)  │    │  (Endpoints)    │  │
│  └─────────────┘    └─────────────┘    └─────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

## Estructura del Proyecto

```
fuente/replier/
├── src/
│   ├── _core/                    # Módulos core reutilizables
│   │   ├── EventEmitter.js       # Sistema de eventos (RxJS)
│   │   ├── WebhookPublisher.js   # Publicador de webhooks
│   │   ├── endpoints.js          # Middleware de autenticación
│   │   ├── errors.js             # Clases de error
│   │   ├── http.js               # Cliente HTTP (Axios)
│   │   ├── websocket.js          # Servidor WebSocket
│   │   ├── id.js                 # Generador de IDs (ULID)
│   │   └── swagger.js            # Documentación OpenAPI
│   ├── config.js                 # Configuración centralizada
│   ├── index.js                  # Punto de entrada
│   ├── controllers/
│   │   ├── api/v1/               # Endpoints API v1
│   │   │   ├── BancasController.js
│   │   │   ├── DebugController.js
│   │   │   ├── RecintoController.js
│   │   │   ├── ReplierController.js
│   │   │   ├── VotacionesController.js
│   │   │   └── WebhookController.js
│   │   ├── api/v2/               # Endpoints API v2
│   │   │   └── WebhookController.js
│   │   ├── HealthcheckController.js
│   │   └── HemicicloController.js
│   ├── daemons/
│   │   ├── Replier.js            # Daemon de sincronización
│   │   └── Commander.js          # Emulador (modo debug)
│   ├── domain/
│   │   ├── Banca.js              # Modelo de asiento
│   │   ├── Recinto.js            # Modelo del hemiciclo
│   │   └── Votacion.js           # Modelo de votación
│   ├── routes/
│   │   ├── api.routes.js         # Rutas de API
│   │   └── app.routes.js         # Rutas de aplicación
│   ├── views/
│   │   └── hemiciclo.twig        # Vista del hemiciclo
│   └── public/                   # Assets estáticos
└── package.json
```

## Configuración

Variables de entorno principales (sample.env):

| Variable | Default | Descripción |
|----------|---------|-------------|
| `ENV` | desarrollo | Entorno de ejecución |
| `DEBUG` | true | Modo debug activo |
| `SERVICE_HOST` | localhost | Host del servicio |
| `SERVICE_PORT` | 17001 | Puerto HTTPS |
| `SERVICE_API_KEY` | - | API key para autenticación |
| `WEBHOOK_API_KEY` | - | API key para webhooks entrantes |
| `RECINTO_NRO_BANCAS` | 257 | Número de asientos |
| `RECINTO_PRESIDENTE_POR_DEFECTO` | - | CUIL del presidente |
| `RECINTO_WEBHOOK_URL` | - | URL webhook de salida |
| `RECINTO_WEBHOOK_API_KEY` | - | API key del webhook |
| `RECINTO_WEBHOOK_REPLY_MS` | 0 | Intervalo de réplica (0=ante cambios) |
| `VOTACIONES_WEBHOOK_URL` | - | URL webhook de votaciones |

## Dominio

### Banca

Representa un asiento del Recinto:

```javascript
{
  numero: 0-256,              // Número de banca
  bascula: boolean,           // ¿Ocupado?
  identificacion: null|cuil   // CUIL del diputado (si está identificado)
}
```

**Reglas de negocio:**
- No se puede identificar sin bascula activa
- Al desactivar bascula, se limpia la identificación

### Recinto

Mantiene el estado completo del hemiciclo (257 bancas):

```javascript
{
  nroBancas: 257,
  nroIdentificados: number,
  nroBasculasActivas: number,
  bancas: Array<Banca>,
  bancaPorCuil: {cuil => nroBanca},
  presidentePorDefecto: cuil,
  emulacion: boolean
}
```

**Eventos emitidos:**
- `banca.bascula.activa` - Cuando alguien se sienta
- `banca.bascula.inactiva` - Cuando alguien se para
- `banca.identificacion` - Cuando alguien se identifica
- `banca.desidentificacion` - Cuando se limpia identificación

**Restricciones:**
- Banca 0 siempre activa con el presidente
- El presidente solo puede estar en una banca a la vez
- Banca 0 no se puede desactivar

### Votación

Representa una votación activa:

```javascript
{
  idVotacion: number,
  estado: 'CREADA'|'INICIADA'|'CERRADA'|'CANCELADA',
  duracionEnSegundos: number,
  timestampInicio: timestamp,
  timestampFin: timestamp,
  votoPorCuil: {cuil => voto}
}
```

**Tipos de voto:**
| Código | Significado |
|--------|-------------|
| 0 | Afirmativo |
| 1 | Negativo |
| 2 | Abstención |
| 3 | Ausente (interno) |
| 4 | Presente sin votar (interno) |

## Daemons

### Replier (src/daemons/Replier.js)

Daemon principal que sincroniza el estado del Recinto a webhooks remotos.

**Funcionamiento:**
1. Se suscribe a eventos del Recinto (EventEmitter)
2. Serializa el estado con formato optimizado
3. Envía POST al webhook configurado
4. Aplica throttle de 1 segundo entre envíos

**Modos de operación:**
- `RECINTO_WEBHOOK_REPLY_MS = 0`: Replica ante cada cambio
- `RECINTO_WEBHOOK_REPLY_MS > 0`: Replica periódicamente (mínimo 1000ms)

**Payload enviado:**
```json
{
  "type": "RECINTO_ESTADO",
  "payload": [
    {"b": 1, "i": 1234567890},
    {"b": 0, "i": 0}
  ]
}
```
Donde: `b` = bascula (1/0), `i` = identificación (CUIL o 0)

### Commander (src/daemons/Commander.js)

Emulador para modo DEBUG. Simula acciones de diputados.

**Modos:**
- `offline`: Inactivo
- `actioner`: Emula sentarse/identificarse/pararse
- `voter`: Emula votaciones

## API REST

### Endpoints Principales

#### Recinto

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| GET | `/v1/recinto` | Obtener estado completo |
| GET | `/v1/recinto:sync` | Obtener estado serializado |
| POST | `/v1/recinto:configurar` | Activar/desactivar emulación |
| POST | `/v1/recinto:limpiar-identificaciones` | Limpiar todas las identificaciones |

#### Bancas

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| GET | `/api/v1/bancas` | Listar todas las bancas |
| GET | `/api/v1/bancas/{banca}` | Obtener una banca |
| PUT | `/api/v1/bancas/{banca}` | Actualizar banca **(solo modo emulación)** |

**PUT /api/v1/bancas/{banca} - Modo Debug/Emulación**

Este endpoint permite actualizar el estado de una banca individual, pero **solo funciona cuando el Recinto está en modo emulación** (`recinto.emulacion === true`).

Request:
```json
{
  "bascula": true,
  "identificacion": 20123456789
}
```

Si el modo emulación está desactivado, responde con error 417:
```json
{
  "error": {
    "code": 417,
    "message": "La emulación de recinto no se encuentra activa"
  }
}
```

Para activar el modo emulación:
```
POST /api/v1/recinto:configurar
{"emulacion": true}
```

> **Nota:** Este endpoint NO es utilizado por Black Sender en producción. Black Sender envía los datos via `POST /api/v1/webhook` con vectores completos. El PUT a bancas individuales es para testing y emulación manual.

#### Votaciones

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| GET | `/v1/votaciones` | Listar votaciones activas |
| DELETE | `/v1/votaciones` | Limpiar votaciones no iniciadas |
| GET | `/v1/votaciones/{id}` | Obtener votación |
| POST | `/v1/votaciones/{id}:iniciar` | Iniciar votación |
| POST | `/v1/votaciones/{id}:cerrar` | Cerrar votación |
| POST | `/v1/votaciones/{id}:cancelar` | Cancelar votación |
| GET | `/v1/votaciones/{id}/votos/{cuil}` | Obtener voto |
| PUT | `/v1/votaciones/{id}/votos/{cuil}` | Registrar voto |
| DELETE | `/v1/votaciones/{id}/votos/{cuil}` | Eliminar voto |

#### Webhooks (Entrada)

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| POST | `/v1/webhook` | Webhook desde BlackProxy/hardware |
| POST | `/v1/webhook/debug` | Debug webhook de salida |
| POST | `/v2/webhook` | Webhook v2 desde controlador |

#### Replier Control

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| POST | `/v1/replier:pause` | Pausar réplicas |
| POST | `/v1/replier:unpause` | Reanudar réplicas |

### Autenticación

Dos tipos de API Key via header `X-API-KEY`:

1. **SERVICE_API_KEY**: Para endpoints internos
2. **WEBHOOK_API_KEY**: Para webhooks entrantes

## WebSocket

El servicio expone un servidor WebSocket que:
- Broadcast del estado del Recinto cada 1 segundo (throttled)
- Envía: `{bancas, totalBasculasActivas, totalIdentificaciones}`
- Soporta ping/pong para keepalive

## Flujos de Operación

### Flujo de Actualización en Tiempo Real

```
1. BlackProxy/Hardware envía POST /v1/webhook
   ↓
2. WebhookController parsea datos de basculas/identificaciones
   ↓
3. Recinto actualiza estado de bancas
   ↓
4. Recinto emite eventos (banca.bascula.activa, etc)
   ↓
5. EventEmitter notifica a suscriptores
   ↓
6. Replier escucha eventos y serializa estado
   ↓
7. Replier envía POST a webhook configurado
   ↓
8. WebSocket broadcast a clientes conectados
```

### Flujo de Votación

```
1. API: POST /v1/votaciones/{id}:iniciar
   ↓
2. Votacion inicia timer de cierre automático
   ↓
3. Diputados votan via PUT /v1/votaciones/{id}/votos/{cuil}
   ↓
4. Timer expira o POST /v1/votaciones/{id}:cerrar
   ↓
5. Replier pausa temporalmente
   ↓
6. votacionesPublisher envía resultado a webhook
   ↓
7. Replier reanuda automáticamente (7 segundos)
```

## Dependencias Principales

- **express**: Framework web
- **helmet**: Seguridad HTTP
- **rxjs**: Sistema de eventos
- **axios**: Cliente HTTP
- **ws**: WebSocket server
- **twig**: Templates
- **swagger-jsdoc/swagger-ui-express**: Documentación API
- **async-mutex**: Control de concurrencia
- **underscore**: Utilidades (throttle, shuffle)
