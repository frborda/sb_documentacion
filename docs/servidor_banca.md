# Servidor de Bancas (servidor_banca)

## Descripción General

**Servidor de Bancas** es el componente central de comunicación desarrollado en Visual Basic 6 que actúa como **puente bidireccional entre el sistema SQV (Sistema de Quorum y Votación) y las terminales/bancas físicas**. Gestiona todas las comunicaciones TCP/IP con las bancas, maneja una cola de mensajes, y traduce los comandos entre ambos sistemas.

## Arquitectura del Sistema

```
┌─────────────────────┐                              ┌──────────────────────┐
│                     │      SQL (Tabla mensajes)    │                      │
│     SQV_Server      │◀─────────────────────────────│   Servidor de        │
│                     │                              │      Bancas          │
│   (Consola/Cartel)  │─────────────────────────────▶│                      │
│                     │      SQL (Tabla mensajes)    │   (Este aplicativo)  │
└─────────────────────┘                              └──────────┬───────────┘
                                                               │
                                           ┌───────────────────┼───────────────────┐
                                           │                   │                   │
                                      TCP/7000            TCP/7000           TCP/7000
                                           │                   │                   │
                                           ▼                   ▼                   ▼
                                    ┌──────────┐        ┌──────────┐        ┌──────────┐
                                    │ Banca 1  │        │ Banca 2  │   ...  │ Banca N  │
                                    │ Terminal │        │ Terminal │        │ Terminal │
                                    └──────────┘        └──────────┘        └──────────┘
```

## Componentes Principales

### 1. FormMain.frm (Formulario Principal)
- **WSocket(0..N)**: Array de controles WinSock para conexiones múltiples
- **Timer**: Temporizador principal para polling de cola de mensajes
- **txtLog**: Log de actividad
- **prgBAR**: Barra de progreso para envío de huellas

### 2. ModuloGeneral.bas (Lógica Principal)
- Conexión a SQL Server
- Interpretación de mensajes SQV → Bancas
- Interpretación de respuestas Bancas → SQV
- Gestión de legisladores y huellas

### 3. ManejodeCola.bas (Sistema de Colas)
- Cola de mensajes en memoria (RecordSet desconectado)
- Control de reintentos
- Timeouts y secuenciación

### 4. Tipos.bas (Definición de Estructuras)
- Types para bancas, sockets, mensajes

## Estructuras de Datos Principales

### BancaIP
```vb
Type BancaIP
    Banca              As Integer   ' Número de banca
    IP                 As String    ' Dirección IP
    Puerto             As String    ' Puerto (7000)
    Estado             As Boolean   ' True = Conectada
    tEstado            As String    ' Estado actual
    tVersion           As String    ' Versión de datos
    tBancaSecuencia    As String    ' Últimas identificaciones
End Type
```

### MensajeSQV
```vb
Type MensajeSQV
    sTipo        As String  ' mset, mget, mevt
    sComponente  As String  ' term, term.auth, term.keyb, etc.
    sObjeto      As String  ' Número de banca o "brc" (broadcast)
    sAtributo    As String  ' action, state, etc.
    sValor       As String  ' Valor del comando
    sComentario  As String  ' Información adicional
End Type
```

### Conversores Socket ↔ Banca
```vb
Type BancaSkt   ' Banca → Socket
     Socket As Integer
     Estado As Boolean
End Type

Type SktBanca   ' Socket → Banca
     Banca As Integer
     Estado As Boolean
End Type
```

## Protocolo de Comunicación

### Comandos SQV → Bancas (InterpretaDatosSQV)

| Tipo | Componente | Atributo | Valor | Comando Banca |
|------|------------|----------|-------|---------------|
| mset | term.auth | action | auth_start | SIDRXH (Identificación huella) |
| mset | term.auth | action | auth_key_start | SIDRNX (Identificación teclado) |
| mset | term.auth | action | auth_restart | SCANCL + SIDRXH |
| mset | term.auth | action | auth_cancel | SCANCL |
| mset | term.auth | action | auth_test | SIDRXH (modo mantenimiento) |
| mset | term.keyb | state | onvotnum | SVOTNU 03 |
| mset | term.keyb | state | offvotnum | SFINNU |
| mset | term.keyb | state | onvotnom | SVOTAR 03 |
| mset | term.keyb | state | offvotnom | SFINVT |
| mset | term.led1 | state | on | SACKID |
| mset | term.ledk1/k2 | state | off* | SLIMVT |
| mset | term.mon | action | reset | Cerrar socket |
| mset | term.mon | action | sync | EnviarHuellasHCDN |
| mget | term | state | - | STATUS |

### Respuestas Bancas → SQV (InterpretaDatosSkt)

| Comando | Descripción | Mensaje a SQV |
|---------|-------------|---------------|
| TIDVAL | Identificación válida | term.auth/result = [ID_HEX] |
| TIDINV | Identificación inválida | term.auth/result = negative |
| TIDOUT | Timeout de identificación | term.auth/result = timeout |
| TVOTOX | Voto emitido (S/N/A) | term.keyb.si/no/ab/state = on |
| TESTAD | Estado del terminal | term/state = ok, term.seat/switch |
| TPRESE | Switch cerrado (sentado) | term.seat/switch = closed |
| TAUSEN | Switch abierto (ausente) | term.seat/switch = open |
| TACKNL | Acknowledge | (procesamiento interno) |
| TNACKN | Negative Acknowledge | (reintentos o cierre) |
| TLEVER | Versión de la banca | Actualiza BancasIP |
| TVERRE | Huellas recibidas OK | (fin de sincronización) |

## Máquina de Estados de Conexión

```
                         ┌──────────────────┐
                         │    INICIALIZADO  │
                         └────────┬─────────┘
                                  │ CargarBancas()
                                  ▼
┌──────────┐  WSocket_Connect  ┌─────────────┐
│          │◀──────────────────│  CONECTANDO │
│          │                   └──────┬──────┘
│          │                          │ (Éxito)
│  CERRADO │                          ▼
│          │   WSocket_Close   ┌─────────────┐
│          │◀──────────────────│  CONECTADO  │
│          │                   │   (sck=7)   │
└──────────┘                   └──────┬──────┘
      ▲                               │
      │     ReconectarBancas()        │ Timer (cada 30s)
      │                               ▼
      │                        ┌─────────────┐
      └────────────────────────│  VERIFICAR  │
                               │   STATUS    │
                               └─────────────┘
```

## Sistema de Cola de Mensajes

### Estructura de la Cola (RecordSet)
```vb
RsCola.Fields:
    - Secuencia   (Char 1)     ' ASCII 125-250
    - Socket      (Integer)    ' Índice del socket
    - Mensaje     (VarChar)    ' Comando a enviar
    - Tick        (Double)     ' Timestamp del envío
    - Reintentos  (Integer)    ' Contador de reintentos
```

### Flujo de Envío
```
                    ┌───────────────────┐
                    │ EnviarSktxCola()  │
                    └─────────┬─────────┘
                              │
          ┌───────────────────┼───────────────────┐
          │                   │                   │
          ▼                   ▼                   ▼
    ┌──────────┐       ┌──────────┐       ┌──────────┐
    │   brc    │       │  1;0;1   │       │   123    │
    │(broadcast)│      │ (vector) │       │ (única)  │
    └────┬─────┘       └────┬─────┘       └────┬─────┘
         │                  │                  │
         ▼                  ▼                  ▼
    ┌─────────────────────────────────────────────┐
    │              CargarCola()                   │
    │   RsCola.AddNew → Secuencia, Socket, Msg    │
    └─────────────────────────────────────────────┘
                              │
                              ▼
    ┌─────────────────────────────────────────────┐
    │          EnviarDatosSkt() [Timer]           │
    │   Por cada socket con mensajes pendientes   │
    │   → EnviarxSkt(secuencia, socket, mensaje)  │
    └─────────────────────────────────────────────┘
```

### Timeouts y Reintentos
```
TimeOutMensaje    = 6000 ms   ' Timeout máximo
TimeOutReintentos = 3000 ms   ' Tiempo entre reintentos
ReintentosMensajes = 3        ' Máximo de reintentos
Máximo reintentos  = 10       ' Cierra socket si se excede
```

## Proceso de Envío de Huellas (EnviarHuellasHCDN)

### Secuencia Completa
```
1. Preparación
   ├── Leer version_datos_sqv de config
   ├── Contar legisladores con huellas
   └── Marcar bancas como "ERROR_PROCESANDO"

2. Por cada banca destino
   ├── EstablecerEnviandoHuellas(fObjeto)
   ├── Enviar SNUVER (versión + cantidad)
   └── Por cada legislador:
       └── Enviar SRLEGI (template hexadecimal)

3. Finalización
   ├── Recibir TVERRE de cada banca
   ├── Enviar SLEVER (solicitar versión)
   └── Actualizar BancasIP con nueva versión
```

### Formato SRLEGI (HCDN 2011)
```
SRLEGI [TEMPLATE_HEXADECIMAL]
       │
       └── Template de huella dactilar en formato hexadecimal
           Tamaño variable según el sensor biométrico
```

## Interacción con sqv_server

### Tablas de Comunicación SQL Server

```sql
-- Mensajes de SQV → Servidor de Bancas
CREATE TABLE sqv_sb_mensajes (
    id          INT IDENTITY,
    Tipo        VARCHAR(10),
    Componente  VARCHAR(50),
    Objeto      VARCHAR(10),
    Atributo    VARCHAR(30),
    Valor       VARCHAR(100),
    Comentario  VARCHAR(50),
    Fecha       DATE,
    Hora        TIME
)

-- Mensajes de Servidor de Bancas → SQV
CREATE TABLE sb_sqv_mensajes (
    -- Estructura similar
)
```

### Procedimiento de Inserción
```sql
EXEC insert_sb_sqv_mensajes
    @Tipo, @Componente, @Objeto,
    @Atributo, @Valor, @Comentario,
    @Fecha, @Hora
```

## Ciclo Principal (Timer)

```vb
Private Sub Timer_Timer()
    CountTimer = CountTimer + 1

    ' Cada ciclo (1ms)
    If FlagTimerLeoCola = True Then
        FlagTimerLeoCola = False
        Call EnviarDatosSkt        ' Procesar cola de envío
        Call LeerMensajesSQV       ' Leer mensajes de SQV
        FlagTimerLeoCola = True
    End If

    ' Cada 300 ciclos (~30s)
    If CountTimer = 300 Then
        CountTimer = 0
        If Not EstadoEnviandoHuellas Then
            Call EnviarSktxCola("brc", "STATUS")  ' Heartbeat
            Call ReconectarBancas                  ' Reconectar caídas
        End If
    End If
End Sub
```

## Gestión de Estados de Bancas

### Arrays de Control
```vb
UltimoEstadodeBanca()     ' "ok", "off"
UltimoEstadodePresencia() ' "closed", "open"
BanderaReset()            ' True si necesita reset
FueConfigurada()          ' True si recibió SCONFG
EnviandoHuellas()         ' True durante sincronización
BancasDeshabilitadas()    ' True si está deshabilitada
```

### Secuencia de Configuración Inicial
```
WSocket_Connect(Index)
        │
        ├── EnviarSktxCola("SCANCL")
        │
        ├── Enviar mensaje term/state=ok a SQV
        │
        └── (Al recibir TESTAD)
                │
                └── Configurar(Index)
                        ├── SCONFG 9010061
                        ├── STATUS
                        └── SLEVER
```

## Diagrama de Secuencia de Votación

```
sqv_server          servidor_banca         Banca           Usuario
    │                     │                  │                │
    │ term.keyb/state=    │                  │                │
    │    onvotnom         │                  │                │
    │────────────────────▶│                  │                │
    │                     │   SVOTAR 03      │                │
    │                     │─────────────────▶│                │
    │                     │                  │  TACKNL        │
    │                     │◀─────────────────│                │
    │                     │                  │                │
    │                     │                  │   [Presiona]   │
    │                     │                  │◀───────────────│
    │                     │    TVOTOX S      │                │
    │                     │◀─────────────────│                │
    │                     │                  │                │
    │ term.keyb.si/       │                  │                │
    │    state=on         │                  │                │
    │◀────────────────────│                  │                │
    │                     │                  │                │
    │ term.ledk1/state=on │                  │                │
    │────────────────────▶│                  │                │
    │                     │   SACKVT S       │                │
    │                     │─────────────────▶│  [LED verde]   │
    │                     │                  │────────────────▶
```

## Constantes Importantes

```vb
MAX_BANCA = 256              ' Máximo número de bancas
MODOLIGHT = True             ' Reduce logging
cNIVEL_LOG = 3               ' 3=todos, 2=algunos, 1=incidentes
Puerto por defecto = 7000    ' Puerto TCP de las bancas
```

## Manejo de Errores

### Error 40006 (Socket no conectado)
```vb
If Err.Number = 40006 Then
    Call WSocketClose(fSocket, "ERROR SOCKET")
    ' Reconexión automática en siguiente ciclo
End If
```

### Máximo de Reintentos Excedido
```vb
If RsCola!reintentos > 10 Then
    Call WSocketClose(RsCola!Socket, "MAX REINTENTOS")
End If
```

## Logs y Diagnóstico

- **txtLog**: Errores y eventos principales
- **txtEnviando**: Estado de envío de huellas por banca (0/1)
- **txtLogEnvio**: Estado de completitud de envíos (@=OK, A-Z=demora, *=timeout)
- **Log_Banca()**: Log individual por banca (si MODOLIGHT=False)
- **Log_DEBUG()**: Log detallado de operaciones
