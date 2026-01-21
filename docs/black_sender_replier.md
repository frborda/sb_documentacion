# Intercomunicación Black Sender - Replier

## Descripción General

Este documento describe la comunicación entre los sistemas **Black Sender** y **Replier**, que en conjunto forman el sistema de sincronización de estado del Recinto (Hemiciclo) legislativo.

## Arquitectura de Integración

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                        FLUJO COMPLETO DE DATOS                               │
└─────────────────────────────────────────────────────────────────────────────┘

    ┌───────────────┐         ┌───────────────┐         ┌───────────────┐
    │   HARDWARE    │         │  BLACK SENDER │         │    REPLIER    │
    │   (Basculas   │         │   :17000      │         │    :17001     │
    │    + Huella)  │         │               │         │               │
    └───────┬───────┘         └───────┬───────┘         └───────┬───────┘
            │                         │                         │
            │  Escribe               │  Polling               │  Mantiene
            │  vectores              │  cada 1s               │  estado en
            │                         │                         │  memoria
            ▼                         ▼                         ▼
    ┌───────────────┐         ┌───────────────┐         ┌───────────────┐
    │  SQL Server   │◀───────│    Daemon     │────────▶│    Recinto    │
    │  (tabla       │  Query  │ BancasConsumer│  POST   │   (257 bancas)│
    │   vector)     │         │   Replier     │ webhook │               │
    └───────────────┘         └───────────────┘         └───────┬───────┘
                                                                 │
                                                                 ▼
                                                        ┌───────────────┐
                                                        │   Webhooks    │
                                                        │  de Salida    │
                                                        │ + WebSocket   │
                                                        └───────────────┘
```

## Protocolo de Comunicación

### Dirección del Flujo

```
Black Sender ──────────▶ Replier
             POST /v1/webhook
```

Black Sender actúa como **productor** y Replier como **consumidor** de los datos de presencia e identificación.

### Endpoint de Comunicación

| Componente | URL | Método |
|------------|-----|--------|
| Origen | Black Sender | POST |
| Destino | `https://{REPLIER_HOST}:17001/v1/webhook` | - |

### Autenticación

La comunicación utiliza autenticación via API Key:

```http
POST /v1/webhook HTTP/1.1
Host: replier:17001
Content-Type: application/json
X-API-KEY: {WEBHOOK_API_KEY}
```

**Configuración en Black Sender:**
- `SENDER_WEBHOOK_URL`: URL completa del webhook de Replier
- `SENDER_WEBHOOK_API_KEY`: API Key configurada en Replier

**Configuración en Replier:**
- `WEBHOOK_API_KEY`: API Key que valida las peticiones entrantes

## Formato de Datos

### Payload enviado por Black Sender

```json
{
  "basculas": "1101010100000...",
  "identificaciones": "20123456789;0;27234567890;0;..."
}
```

| Campo | Tipo | Descripción |
|-------|------|-------------|
| `basculas` | string | Vector de 257 caracteres. Cada posición representa una banca: `1` = ocupada, `0` = vacía |
| `identificaciones` | string | CUILs separados por `;`. Cada posición corresponde a una banca. `0` = no identificado |

### Ejemplo de Payload

```json
{
  "basculas": "110100...",
  "identificaciones": "20123456789;27234567890;0;20345678901;..."
}
```

Interpretación:
- Banca 0: ocupada (`1`), identificado con CUIL `20123456789`
- Banca 1: ocupada (`1`), identificado con CUIL `27234567890`
- Banca 2: vacía (`0`), no identificado (`0`)
- Banca 3: ocupada (`1`), identificado con CUIL `20345678901`

## Procesamiento en Replier

### WebhookController (src/controllers/api/v1/WebhookController.js)

Cuando Replier recibe el webhook:

```javascript
// 1. Parsea los vectores
const basculas = req.body.basculas.split('')
const identificaciones = req.body.identificaciones.split(';')

// 2. Para cada banca (0-256):
for (let i = 0; i < 257; i++) {
  const basculaActiva = basculas[i] === '1'
  const cuil = parseInt(identificaciones[i]) || null

  // 3. Actualiza el estado del Recinto
  if (basculaActiva) {
    recinto.activarBascula(i)
    if (cuil) {
      recinto.identificar(i, cuil)
    }
  } else {
    recinto.desactivarBascula(i)
  }
}
```

### Respuesta de Replier

```json
{
  "totalProcesados": 257,
  "totalExitosos": 255,
  "totalFallidos": 2,
  "fallidos": [
    { "banca": 45, "error": "CUIL duplicado" },
    { "banca": 123, "error": "Banca inválida" }
  ]
}
```

## Flujo de Sincronización

### Diagrama de Secuencia

```
┌─────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────┐
│Hardware │     │ SQL Server  │     │Black Sender │     │ Replier │
└────┬────┘     └──────┬──────┘     └──────┬──────┘     └────┬────┘
     │                 │                   │                  │
     │ Detecta cambio  │                   │                  │
     │────────────────▶│                   │                  │
     │                 │                   │                  │
     │                 │  Polling (1s)     │                  │
     │                 │◀──────────────────│                  │
     │                 │                   │                  │
     │                 │  Retorna vector   │                  │
     │                 │──────────────────▶│                  │
     │                 │                   │                  │
     │                 │                   │ Detecta cambio   │
     │                 │                   │ vs anterior      │
     │                 │                   │                  │
     │                 │                   │ POST /v1/webhook │
     │                 │                   │─────────────────▶│
     │                 │                   │                  │
     │                 │                   │                  │ Actualiza
     │                 │                   │                  │ Recinto
     │                 │                   │                  │
     │                 │                   │  200 OK          │
     │                 │                   │◀─────────────────│
     │                 │                   │                  │
     │                 │                   │                  │ Emite eventos
     │                 │                   │                  │ a suscriptores
     │                 │                   │                  │
```

### Tiempos y Throttling

| Componente | Intervalo | Descripción |
|------------|-----------|-------------|
| Black Sender Polling | 1000ms | Lee la tabla `vector` cada segundo |
| Black Sender Reply | ~1000ms | Throttle para evitar sobrecarga |
| Replier WebSocket | 1000ms | Broadcast del estado a clientes |

## Transformación de Datos

### En Black Sender

```
Tabla vector (SQL Server)           Payload HTTP
┌─────────────────────────┐        ┌────────────────────────┐
│ vector_presencia:       │        │ basculas: "110100..."  │
│ "110100..."             │───────▶│                        │
│                         │        │                        │
│ vector_identificacion:  │        │ identificaciones:      │
│ "1;2;0;4;..."           │──┬────▶│ "20123456789;          │
│ (IDs numéricos)         │  │     │  27234567890;0;..."    │
└─────────────────────────┘  │     │ (CUILs)                │
                             │     └────────────────────────┘
    ┌────────────────────┐   │
    │ Mapa cuilPorId     │───┘
    │ {1: 20123456789,   │  Traduce ID → CUIL
    │  2: 27234567890,   │
    │  4: 20345678901}   │
    └────────────────────┘
```

### En Replier

```
Payload HTTP                         Estado del Recinto
┌────────────────────────┐          ┌─────────────────────────┐
│ basculas: "110100..."  │──────────│ bancas: [               │
│                        │          │   {b:1, i:20123456789}, │
│ identificaciones:      │          │   {b:1, i:27234567890}, │
│ "20123456789;          │──────────│   {b:0, i:0},           │
│  27234567890;0;..."    │          │   {b:1, i:0},           │
└────────────────────────┘          │   ...                   │
                                    │ ]                       │
                                    └─────────────────────────┘
```

## Manejo de Errores

### En Black Sender

```javascript
// Si el webhook falla, loguea pero continúa el polling
try {
  await axios.post(webhook.url, payload, { timeout: 5000 })
} catch (error) {
  console.error('Webhook falló:', error.message)
  // No detiene el servicio, continúa en siguiente iteración
}
```

### En Replier

```javascript
// Procesa cada banca individualmente, errores no detienen el proceso
for (let i = 0; i < 257; i++) {
  try {
    // Procesar banca i
  } catch (error) {
    fallidos.push({ banca: i, error: error.message })
  }
}
// Retorna resultado parcial aunque haya errores
```

## Reglas de Negocio Compartidas

### Banca 0 (Presidente)

Ambos sistemas mantienen la regla de que la banca 0 siempre está activa:

**Black Sender:**
```javascript
// Fuerza banca 0 a '1' antes de procesar
vector.vector_presencia = '1' + vector.vector_presencia.substring(1)
```

**Replier:**
```javascript
// Banca 0 siempre tiene presidente identificado
// No se puede desactivar bascula de banca 0
```

### Modo Emulación

Cuando Replier está en modo emulación, **ignora los webhooks entrantes**:

```javascript
// WebhookController.js
if (recinto.emulacion) {
  return res.json({ message: 'Ignorado: modo emulación activo' })
}
```

Esto permite hacer pruebas sin que los datos reales interfieran.

## Configuración de Integración

### Archivo .env de Black Sender

```env
# Webhook destino (Replier)
SENDER_WEBHOOK_URL=https://localhost:17001/v1/webhook
SENDER_WEBHOOK_API_KEY=tu-api-key-secreta

# Base de datos (origen de datos)
DATABASE_HOST=servidor-sql
DATABASE_PORT=1433
DATABASE_NAME=recinto_db
DATABASE_USER=usuario
DATABASE_PASS=password
```

### Archivo .env de Replier

```env
# API Key para webhooks entrantes
WEBHOOK_API_KEY=tu-api-key-secreta

# Configuración del recinto
RECINTO_NRO_BANCAS=257
RECINTO_PRESIDENTE_POR_DEFECTO=20123456789
```

## Monitoreo y Debug

### Verificar estado de Black Sender

```bash
curl -H "X-API-KEY: tu-api-key" https://localhost:17000/v1/daemon
```

Respuesta:
```json
{
  "pollingMilliseconds": 1000,
  "vectorBasculas": "110100...",
  "vectorIdentificaciones": "20123456789;27234567890;...",
  "cuilPorId": { "1": 20123456789, ... }
}
```

### Verificar estado de Replier

```bash
curl -H "X-API-KEY: tu-api-key" https://localhost:17001/v1/recinto:sync
```

Respuesta:
```json
[
  {"b": 1, "i": 20123456789},
  {"b": 1, "i": 27234567890},
  {"b": 0, "i": 0},
  ...
]
```

### Webhook de Debug en Replier

```bash
# Ver qué está enviando Replier a sus webhooks de salida
curl -X POST -H "X-API-KEY: tu-api-key" \
  https://localhost:17001/v1/webhook/debug \
  -d '{"type":"RECINTO_ESTADO","payload":[...]}'
```

## Diagrama de Puertos y Conexiones

```
┌──────────────────────────────────────────────────────────────┐
│                     INFRAESTRUCTURA                           │
├──────────────────────────────────────────────────────────────┤
│                                                               │
│   ┌─────────────────┐                                        │
│   │  SQL Server     │                                        │
│   │  :1433          │◀───────────┐                           │
│   └─────────────────┘            │                           │
│                                  │ Query                     │
│   ┌─────────────────┐            │                           │
│   │  Black Sender   │────────────┘                           │
│   │  :17000 HTTPS   │                                        │
│   │                 │─────────────┐                          │
│   └─────────────────┘             │ POST /v1/webhook         │
│                                   │                          │
│   ┌─────────────────┐             │                          │
│   │  Replier        │◀────────────┘                          │
│   │  :17001 HTTPS   │                                        │
│   │  :17001 WSS     │─────────────┐                          │
│   └─────────────────┘             │                          │
│                                   │ WebSocket                │
│   ┌─────────────────┐             │                          │
│   │  Clientes Web   │◀────────────┘                          │
│   │  (Hemiciclo UI) │                                        │
│   └─────────────────┘                                        │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

## Resumen de Responsabilidades

| Sistema | Responsabilidad |
|---------|-----------------|
| **Hardware** | Detectar presencia física e identificación de diputados |
| **SQL Server** | Almacenar vectores de estado actuales |
| **Black Sender** | Polling de DB, traducción ID→CUIL, distribución via webhook |
| **Replier** | Mantener estado en memoria, validar reglas de negocio, distribuir a clientes |

## Troubleshooting

### Black Sender no envía datos

1. Verificar conexión a base de datos: `GET /healthcheck`
2. Verificar que el webhook está configurado: revisar `.env`
3. Verificar logs de error en consola

### Replier no recibe datos

1. Verificar que `WEBHOOK_API_KEY` coincide con `SENDER_WEBHOOK_API_KEY`
2. Verificar que Replier no está en modo emulación
3. Verificar logs del webhook: `POST /v1/webhook/debug`

### Datos inconsistentes

1. Verificar mapa de CUILs: `GET /v1/daemon` (Black Sender)
2. Regenerar CUILs si es necesario: `POST /v1/daemon:regenerar-cuiles`
3. Comparar vectores en DB vs estado en Replier
