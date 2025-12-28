# EnvioHuellasSQV (enviar_huella)

## Descripción General

**EnvioHuellasSQV** es una aplicación utilitaria desarrollada en Visual Basic 6 cuya función principal es **enviar templates de huellas dactilares a una banca específica** de forma manual. Se utiliza para sincronización individual de datos biométricos cuando es necesario actualizar una terminal específica.

## Arquitectura

```
┌─────────────────────┐          TCP/7000           ┌──────────────────┐
│  EnvioHuellasSQV    │ ────────────────────────────▶│  Terminal/Banca  │
│  (Cliente)          │                              │  (Servidor)      │
└─────────────────────┘                              └──────────────────┘
         │
         │ ADO/SQL
         ▼
┌─────────────────────┐
│   SQL Server        │
│   - SQV_Config      │
│   - Base Vigente    │
└─────────────────────┘
```

## Componentes Principales

### Formulario Principal (`frmMain.frm`)
- **txtBanca**: Campo de entrada para el número de banca destino
- **cmdEnviarHuellas**: Botón que inicia el envío de huellas
- **prgBar**: Barra de progreso del envío
- **WinSock**: Control de comunicación TCP (puerto 7000)

### Módulo de Tipos (`Module1.bas`)
Define la estructura `TLegislador`:
```vb
Type TLegislador
     sId          As String      ' ID del legislador
     sNombre      As String      ' Nombre
     sApellido    As String      ' Apellido
     sDNI         As String      ' Documento de identidad
     sClase       As String      ' Clase (S=Legislador, V=Mantenimiento)
     sIcono       As String      ' Icono asociado
     sTemplate    As String      ' Template de huella (texto)
     sTemplate11() As Byte       ' Template de huella (binario)
End Type
```

## Flujo de Operación

### 1. Inicialización
```
Form_Load()
    └── AbrirConexionSQLServer()
            ├── Conectar a SQV_Config
            ├── Leer variable 'base_vigente'
            └── Reconectar a base de datos vigente
```

### 2. Proceso de Envío de Huellas

```
Usuario ingresa número de banca
         │
         ▼
cmdEnviarHuellas_Click()
         │
         ├── Validar confirmación del usuario (MsgBox)
         │
         ├── Consultar IP de la banca (tabla BancasIp)
         │
         └── WinSock.Connect(IP, Puerto 7000)
                  │
                  ▼ (Evento WinSock_Connect)
                  │
         EnviarHuellasHCDN(ip)
```

### 3. Protocolo de Comunicación con la Banca

```
┌──────────────────────────────────────────────────────────────────┐
│                     SECUENCIA DE COMANDOS                        │
├──────────────────────────────────────────────────────────────────┤
│  1. SCANCL     →    Cancelar operaciones pendientes              │
│     ← TACKNL        Acknowledgement                              │
│                                                                  │
│  2. SNUVER     →    Enviar versión (8 bytes version + 4 cant.)   │
│     ← TACKNL        Acknowledgement                              │
│                                                                  │
│  3. SNUVER     →    (Se envía dos veces para confirmación)       │
│     ← TACKNL        Acknowledgement                              │
│                                                                  │
│  4. Por cada legislador:                                         │
│     SRLEGI    →     Template de huella en hexadecimal            │
│     ← TACKNL        Acknowledgement                              │
│                                                                  │
│  5. Al finalizar:                                                │
│     ← TVERRE        Terminal Verification Ready                  │
└──────────────────────────────────────────────────────────────────┘
```

## Formato de Mensajes

### SRLEGI (Envío de Legislador)
```
fSRLEGI [TEMPLATE_HEX]
│└──────────────────── Template biométrico en hexadecimal
└── Prefijo 'f' para mensajes rápidos (sin cola)
```

### SNUVER (Nueva Versión)
```
SNUVER [VERSION_8][CANTIDAD_4]
        │           └── Cantidad de personas en decimal (4 bytes)
        └── Versión de datos (8 caracteres: DDMMHHMM)
```

## Base de Datos

### Consultas Principales

1. **Obtener IP de banca**:
```sql
SELECT Ip FROM BancasIp WHERE BancaNumero = @BancaNumero
```

2. **Obtener legisladores activos con huellas**:
```sql
SELECT sb.*
FROM (
    SELECT a.id FROM legisladores_sb a
    INNER JOIN legisladores_activos ON legisladores_activos.id = a.id
    WHERE a.tipo = 1 AND descripcion <> 'Activo sin incorporar'
    UNION
    SELECT b.id FROM legisladores_sb b WHERE b.tipo = 0
) U
INNER JOIN legisladores_sb sb ON U.id = sb.id
ORDER BY sb.tipo DESC, U.id
```

## Estados del Socket

| Estado | Descripción |
|--------|-------------|
| 0 | Cerrado |
| 7 | Conectado |

## Manejo de Respuestas

```vb
Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)
    ' Respuestas esperadas:
    ' - TACKNL: Operación exitosa
    ' - TESTAD: Estado (se ignora durante envío)
    ' - TVERRE: Proceso completado exitosamente
End Sub
```

## Interacción con Otros Componentes

### Relación con servidor_banca
- **EnvioHuellasSQV** se conecta directamente a las bancas
- Es una alternativa manual al envío masivo que realiza **servidor_banca**
- Útil para sincronizar bancas individuales sin afectar al resto

### Relación con sqv_server
- Ambos leen de la misma base de datos de legisladores
- sqv_server coordina las operaciones de votación
- EnvioHuellasSQV solo se encarga de la sincronización biométrica

## Diagrama de Estados

```
                    ┌─────────────┐
                    │   INACTIVO  │
                    └──────┬──────┘
                           │ Click en "Enviar"
                           ▼
                    ┌─────────────┐
                    │ CONECTANDO  │
                    └──────┬──────┘
                           │ WinSock_Connect
                           ▼
                    ┌─────────────┐
                    │  ENVIANDO   │◀─┐
                    │   SCANCL    │  │
                    └──────┬──────┘  │
                           │         │
                           ▼         │
                    ┌─────────────┐  │
                    │  ENVIANDO   │  │ Por cada
                    │   SNUVER    │  │ legislador
                    └──────┬──────┘  │
                           │         │
                           ▼         │
                    ┌─────────────┐  │
                    │  ENVIANDO   │──┘
                    │   SRLEGI    │
                    └──────┬──────┘
                           │ TVERRE recibido
                           ▼
                    ┌─────────────┐
                    │ COMPLETADO  │
                    └─────────────┘
```

## Funciones Auxiliares

| Función | Descripción |
|---------|-------------|
| `BinAHex()` | Convierte array de bytes a string hexadecimal |
| `CerosIzquierda()` | Rellena string con ceros a la izquierda |
| `strString()` | Formatea string a longitud fija |
| `NullCadena()` | Maneja valores nulos de base de datos |
| `rtaOk()` | Verifica si la respuesta contiene TACKNL |

## Consideraciones de Uso

1. **Una sola banca a la vez**: La aplicación solo puede enviar a una banca por ejecución
2. **Confirmación requerida**: Se solicita confirmación antes de iniciar
3. **Validación de banca**: Verifica que la banca exista en la tabla BancasIp
4. **Progreso visual**: La barra de progreso muestra el avance del envío
5. **Bloqueo durante envío**: El botón se deshabilita mientras se envían las huellas
