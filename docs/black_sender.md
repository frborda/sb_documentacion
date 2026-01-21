# Black Sender - Consumidor y Distribuidor de Presencia

## Descripción General

**Black Sender** es un servicio Node.js que actúa como puente entre la base de datos de hardware del Recinto y el sistema Replier. Realiza polling continuo sobre una tabla SQL Server que contiene vectores de presencia e identificación, y distribuye los cambios a webhooks configurados.

- **Versión:** 1.1.0
- **Puerto por defecto:** 17000
- **Protocolo:** HTTPS

## Arquitectura

```
┌─────────────────────────────────────────────────────────────┐
│                   BLACK SENDER SERVICE                       │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────────────────────────────────────────────┐    │
│  │           BancasConsumerReplier (Daemon)             │    │
│  │                                                      │    │
│  │   ┌──────────────┐         ┌──────────────────┐     │    │
│  │   │   Polling    │────────▶│  Detección de    │     │    │
│  │   │   (1 seg)    │         │    Cambios       │     │    │
│  │   └──────────────┘         └────────┬─────────┘     │    │
│  │                                      │               │    │
│  │   ┌──────────────┐                  ▼               │    │
│  │   │  Mapa CUIL   │         ┌──────────────────┐     │    │
│  │   │  (cuilPorId) │────────▶│   Traducción     │     │    │
│  │   └──────────────┘         │   ID → CUIL      │     │    │
│  │                             └────────┬─────────┘     │    │
│  │                                      │               │    │
│  │                                      ▼               │    │
│  │                             ┌──────────────────┐     │    │
│  │                             │  Webhook POST    │     │    │
│  │                             │  (throttled)     │     │    │
│  │                             └──────────────────┘     │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
│  ┌──────────────────┐    ┌──────────────────────────────┐   │
│  │   API REST       │    │      SQL Server (MSSQL)      │   │
│  │   /v1/daemon     │    │   - vector                   │   │
│  └──────────────────┘    │   - DiputadosCuil            │   │
│                          │   - legisladores_activos     │   │
│                          └──────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

## Estructura del Proyecto

```
fuente/black_sender/
├── src/
│   ├── _core/                    # Módulos core reutilizables
│   │   ├── Database.js           # Conexión a SQL Server
│   │   ├── EventEmitter.js       # Sistema de eventos (RxJS)
│   │   ├── endpoints.js          # Middleware de autenticación
│   │   ├── errors.js             # Clases de error
│   │   ├── http.js               # Cliente HTTP (Axios)
│   │   ├── swagger.js            # Documentación OpenAPI
│   │   └── index.js              # Exporta módulos core
│   ├── config.js                 # Configuración centralizada
│   ├── index.js                  # Punto de entrada
│   ├── controllers/
│   │   ├── HealthcheckController.js
│   │   └── api/v1/
│   │       ├── DaemonController.js
│   │       └── index.js
│   ├── daemons/
│   │   └── BancasConsumerReplier.js  # Daemon principal
│   └── routes/
│       └── api.routes.js
└── package.json
```

## Configuración

Variables de entorno principales (sample.env):

| Variable | Default | Descripción |
|----------|---------|-------------|
| `ENV` | desarrollo | Entorno de ejecución |
| `DEBUG` | true | Modo debug activo |
| `SERVICE_HOST` | localhost | Host del servicio |
| `SERVICE_PORT` | 17000 | Puerto HTTPS |
| `SERVICE_API_KEY` | - | API key para autenticación |
| `SERVICE_SSL_KEY` | - | Ruta al certificado SSL key |
| `SERVICE_SSL_CERT` | - | Ruta al certificado SSL cert |
| `SENDER_WEBHOOK_URL` | - | URL del webhook destino (Replier) |
| `SENDER_WEBHOOK_API_KEY` | - | API key del webhook |
| `DATABASE_HOST` | - | Host SQL Server |
| `DATABASE_PORT` | 1433 | Puerto SQL Server |
| `DATABASE_NAME` | - | Nombre de la base de datos |
| `DATABASE_USER` | - | Usuario de base de datos |
| `DATABASE_PASS` | - | Contraseña de base de datos |

## Módulos Core

### Database.js

Abstracción para conectarse a SQL Server (MSSQL):

```javascript
// Métodos disponibles
db.conectar()   // Establece conexión y valida con ping
db.ping()       // Verifica conectividad
db.query(q)     // Ejecuta query SQL, retorna recordsets
```

### Otros módulos

Los módulos `EventEmitter`, `endpoints`, `errors`, `http` y `swagger` son compartidos con el proyecto Replier y funcionan de manera idéntica.

## Daemon: BancasConsumerReplier

El daemon principal que realiza el polling y distribución de datos.

### Estado Interno

```javascript
{
  db: Database,                     // Conexión a MSSQL
  webhooks: [],                     // URLs destino para replicar
  pollingMilliseconds: 1000,        // Intervalo de polling (1 segundo)

  cuilPorId: {},                    // Mapa {diputado_id: cuil}
  vectorPresenciaAnterior: '',      // Cache del vector anterior
  vectorIdentificacionAnterior: '', // Cache del vector anterior

  cacheBasculas: undefined,         // Último vector_presencia procesado
  cacheIdentificaciones: undefined, // Último vector_identificacion procesado

  keepReplying: true                // Siempre replicar incluso sin cambios
}
```

### Flujo de Ejecución

#### 1. Inicialización (`run()`)

```
run()
  │
  ├─▶ conectar a DB
  │
  ├─▶ generarMapaCuilPorId()
  │     │
  │     ├─▶ chequearCuilesDiputados()  [valida que todos tengan CUIL]
  │     │
  │     └─▶ query DiputadosCuil → llena cuilPorId
  │
  └─▶ consume()  [inicia polling infinito]
```

#### 2. Polling (`consume()`)

Cada 1 segundo:

```
consume()
  │
  ├─▶ Query tabla "vector"
  │     - Obtiene vector_presencia
  │     - Obtiene vector_identificacion
  │
  ├─▶ Fuerza banca 0 a '1' (presidente siempre presente)
  │
  ├─▶ Detecta cambios vs versión anterior
  │
  └─▶ Si hay cambios O keepReplying=true:
        │
        ├─▶ Parsea identificaciones (IDs → CUILs)
        │
        └─▶ reply() → Envía webhook
```

#### 3. Replicación (`reply()`)

```javascript
// Payload enviado
{
  "basculas": "11010101...",                    // Vector de 257 caracteres (0/1)
  "identificaciones": "23045123456;0;23045678912;..."  // CUILs separados por ;
}
```

- Usa `_.throttle()` para limitar frecuencia
- POST a cada webhook registrado
- Headers: `X-API-KEY`, `Content-Type: application/json`
- Timeout: 5 segundos

### Métodos Públicos

| Método | Descripción |
|--------|-------------|
| `run()` | Inicia el daemon (conecta DB + inicia polling) |
| `to(url, apiKey)` | Registra webhook destino |
| `regenerarCuiles()` | Recarga mapa de CUILs desde DB |
| `toJSON()` | Serializa estado para debugging |

## Base de Datos

### Tablas Utilizadas

#### vector
Contiene los vectores actuales de presencia e identificación del hardware.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `vector_presencia` | string | 257 caracteres (0/1) indicando ocupación |
| `vector_identificacion` | string | IDs de diputados separados por ; |

#### DiputadosCuil
Mapeo de diputados a su CUIL.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `id` | int | ID del diputado |
| `cuil` | bigint | CUIL del diputado |

#### legisladores_activos
Lista de legisladores activos para validación.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `id` | int | ID del diputado |
| `nombre` | string | Nombre |
| `apellido` | string | Apellido |

## API REST

### Endpoints

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| GET | `/healthcheck` | Estado del servicio |
| GET | `/v1/daemon` | Obtener estado del daemon |
| POST | `/v1/daemon:regenerar-cuiles` | Regenerar mapa de CUILs |

### GET /healthcheck

Retorna estado del sistema:

```json
{
  "status": "ok",
  "env": "desarrollo",
  "debug": true,
  "replier": {
    "webhook": "https://..."
  },
  "database": {
    "host": "...",
    "port": 1433,
    "name": "...",
    "user": "..."
  }
}
```

### GET /v1/daemon

Retorna estado del daemon:

```json
{
  "pollingMilliseconds": 1000,
  "vectorBasculas": "11010101...",
  "vectorIdentificaciones": "23045123456;0;...",
  "cuilPorId": {
    "1": 20123456789,
    "2": 27234567890
  }
}
```

### POST /v1/daemon:regenerar-cuiles

Recarga el mapa de CUILs desde la base de datos. Útil cuando se agregan nuevos legisladores.

### Autenticación

Todos los endpoints requieren header `X-API-KEY` con el valor de `SERVICE_API_KEY`.

## Dependencias Principales

- **express**: Framework web
- **helmet**: Seguridad HTTP headers
- **morgan**: Logging de requests
- **mssql**: Cliente SQL Server
- **axios**: Cliente HTTP para webhooks
- **rxjs**: Sistema de eventos
- **swagger-jsdoc/swagger-ui-express**: Documentación API
- **underscore**: Utilidades (throttle)

## Flujo Completo de Datos

```
┌──────────────────┐
│   Hardware del   │
│     Recinto      │
│   (Basculas +    │
│  Identificación) │
└────────┬─────────┘
         │
         ▼ (Escribe en DB)
┌──────────────────┐
│   SQL Server     │
│   Tabla: vector  │
└────────┬─────────┘
         │
         ▼ (Polling cada 1s)
┌──────────────────┐
│  Black Sender    │
│  (Este servicio) │
└────────┬─────────┘
         │
         ▼ (POST webhook)
┌──────────────────┐
│     Replier      │
│   (Destino)      │
└──────────────────┘
```

## Casos de Uso

1. **Sincronización de presencia**: El hardware detecta cuando un diputado se sienta/levanta, escribe en la tabla `vector`, Black Sender lo detecta y notifica a Replier.

2. **Identificación de diputados**: Cuando un diputado se identifica con su huella/tarjeta, el hardware actualiza `vector_identificacion`, Black Sender traduce el ID a CUIL y notifica.

3. **Regeneración de CUILs**: Cuando se agregan nuevos legisladores, se puede invocar el endpoint para recargar el mapeo sin reiniciar el servicio.
