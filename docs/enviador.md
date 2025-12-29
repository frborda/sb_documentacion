# EnviadorPro (enviador)

## Descripción General

**EnviadorPro** es una aplicación utilitaria desarrollada en Visual Basic 6 cuya función principal es **enviar templates de huellas dactilares a una banca específica** de forma manual. Es una versión mejorada del envío de huellas con capacidades de reintento automático y carga previa de datos en memoria para mayor eficiencia.

## Arquitectura

```
┌─────────────────────┐          TCP/7000           ┌──────────────────┐
│   EnviadorPro       │ ────────────────────────────▶│  Terminal/Banca  │
│   (Cliente)         │                              │  (Servidor)      │
└─────────────────────┘                              └──────────────────┘
         │
         │ ADO/SQL
         ▼
┌─────────────────────┐
│   SQL Server        │
│   - SQV_Config      │
│   - Base Vigente    │
└─────────────────────┘
         │
         │ Logging
         ▼
┌─────────────────────┐
│  C:\logBanca.txt    │
└─────────────────────┘
```

## Componentes Principales

### Formulario Principal (`frmMain.frm`)
- **txtBanca**: Campo de entrada para el número de banca destino
- **cmdEnviar**: Botón que inicia el proceso de envío
- **prgBar**: Barra de progreso del envío (MSComctlLib.ProgressBar)
- **WinSock**: Control de comunicación TCP (puerto 7000)

### Módulo de Tipos y Funciones (`Module1.bas`)
Define la estructura `TLegislador` y funciones auxiliares:
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

## Variables Globales del Formulario

| Variable | Tipo | Descripción |
|----------|------|-------------|
| `ipBanca` | String | IP de la banca destino |
| `version` | String | Versión de datos (8+4 bytes) |
| `huellas()` | String | Array con todas las huellas a enviar |
| `cantHuellas` | Integer | Cantidad total de huellas |
| `sLegisla` | TLegislador | Estructura del legislador actual |
| `huellaActual` | Integer | Índice de huella en proceso |
| `firstNuver` | Boolean | Flag para control de doble SNUVER |
| `enviado` | String | Último comando enviado |

## Flujo de Operación

### 1. Inicialización
```
Form_Load()
    ├── Inicializar variable enviado = ""
    ├── AbrirConexionSQLServer()
    │       ├── Conectar a SQV_Config
    │       ├── Leer variable 'base_vigente'
    │       └── Reconectar a base de datos vigente
    └── Crear archivo de log (C:\logBanca.txt)
```

### 2. Proceso de Obtención de Datos
```
Usuario ingresa número de banca
         │
         ▼
cmdEnviar_Click()
         │
         └── ObtieneDatosEnviar()
                  │
                  ├── Consultar IP de la banca (tabla BancasIp)
                  │
                  ├── Deshabilitar controles UI
                  │
                  └── EnviarHuellasHCDN(ip)
                           │
                           ├── Obtener version_datos_sqv de config
                           │
                           ├── Contar legisladores con huella
                           │
                           └── Cargar array huellas[] en memoria
                                    │
                                    └── Procesar()
```

### 3. Proceso de Conexión y Envío
```
Procesar()
    ├── Inicializar huellaActual = -1
    ├── Configurar barra de progreso
    └── WinSock.Connect(ipBanca, 7000)
             │
             ▼ (Evento WinSock_Connect)
             │
         Enviar("SCANCL")
```

### 4. Protocolo de Comunicación con la Banca

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
│     ← TACKNL        Acknowledgement (éxito)                      │
│     ← TNACK         Negative Ack (reintento automático)          │
│                                                                  │
│  5. Al finalizar:                                                │
│     ← TVERRE/ERRE   Terminal Verification Ready                  │
└──────────────────────────────────────────────────────────────────┘
```

## Formato de Mensajes

### Prefijo de Envío
Todos los mensajes se envían con prefijo `f` para mensajes rápidos (sin cola):
```vb
Private Sub Enviar(pDato As String)
    WinSock.SendData "f" & pDato
    enviado = pDato
End Sub
```

### SRLEGI (Envío de Legislador)
```
fSRLEGI [TEMPLATE_HEX]
│└──────────────────── Template biométrico en hexadecimal
└── Prefijo 'f' para mensajes rápidos
```

### SNUVER (Nueva Versión)
```
SNUVER [VERSION_8][CANTIDAD_4]
        │           └── Cantidad de personas en decimal (4 bytes)
        └── Versión de datos (8 caracteres: DDMMHHMM)
```

La versión se construye a partir de `version_datos_sqv`:
```vb
strPrefijoVersionBanca = Replace(rsBanca.Fields(0).Value, "_", "")
strPrefijoVersionBanca = Left(strPrefijoVersionBanca, 4) & Mid(strPrefijoVersionBanca, 9, 4)
version = strPrefijoVersionBanca & Format(nCantLegisladoresConHuella, "0000")
```

## Base de Datos

### Conexión Inicial (SQV_Config)
```
Server: 10.1.1.5
Database: SQV_Config
User: SQV
Password: hcdn11
```

### Consultas Principales

1. **Obtener base vigente**:
```sql
SELECT valor FROM configuracion WHERE variable = 'base_vigente'
```

2. **Obtener versión de datos**:
```sql
SELECT TOP 1 version_datos_sqv FROM config
```

3. **Obtener IP de banca**:
```sql
SELECT Ip FROM BancasIp WHERE BancaNumero = @BancaNumero
```

4. **Contar legisladores con huellas**:
```sql
SELECT COUNT(*)
FROM (
    SELECT a.id FROM legisladores_sb a
    INNER JOIN legisladores_activos ON legisladores_activos.id = a.id
    WHERE a.tipo = 1 AND descripcion <> 'Activo sin incorporar'
    UNION
    SELECT b.id FROM legisladores_sb b WHERE b.tipo = 0
) U
INNER JOIN legisladores_sb sb ON U.id = sb.id
```

5. **Obtener legisladores activos con huellas**:
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

## Manejo de Respuestas

```vb
Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)
    ' Respuestas manejadas:
    ' - TACKNL: Operación exitosa, continuar con siguiente paso
    ' - TNACK:  Error, reintentar envío de huella actual
    ' - ERRE:   Proceso completado exitosamente
    ' - Otro:   Mensaje desconocido, mostrar error y terminar
End Sub
```

### Máquina de Estados de Respuesta

```
TACKNL recibido
    │
    ├── enviado = "SCANCL"  → Enviar SNUVER
    │
    ├── enviado = "SNUVER"
    │       │
    │       ├── firstNuver = False → Enviar SNUVER (segunda vez)
    │       │
    │       └── firstNuver = True  → Enviar primera huella
    │
    └── enviado = "SRLEGI" → Enviar siguiente huella

TNACK recibido
    └── enviado = "SRLEGI" → Reintentar misma huella

ERRE recibido
    └── Proceso completado, mostrar éxito, cerrar conexión
```

## Diagrama de Estados

```
                    ┌─────────────┐
                    │   INACTIVO  │
                    └──────┬──────┘
                           │ Click en "Enviar"
                           ▼
                    ┌─────────────┐
                    │  CARGANDO   │ ← Carga huellas en memoria
                    │   HUELLAS   │
                    └──────┬──────┘
                           │
                           ▼
                    ┌─────────────┐
                    │ CONECTANDO  │
                    └──────┬──────┘
                           │ WinSock_Connect
                           ▼
                    ┌─────────────┐
                    │  ENVIANDO   │
                    │   SCANCL    │
                    └──────┬──────┘
                           │ TACKNL
                           ▼
                    ┌─────────────┐
                    │  ENVIANDO   │
                    │  SNUVER(1)  │
                    └──────┬──────┘
                           │ TACKNL
                           ▼
                    ┌─────────────┐
                    │  ENVIANDO   │
                    │  SNUVER(2)  │
                    └──────┬──────┘
                           │ TACKNL
                           ▼
                    ┌─────────────┐
                    │  ENVIANDO   │◀──┐
                    │   SRLEGI    │   │ TNACK (reintento)
                    └──────┬──────┘───┘
                           │
                    ┌──────┴──────┐
                    │             │
              TACKNL│             │ERRE
          (más huellas)     (fin)
                    │             │
                    ▼             ▼
              ┌─────────┐  ┌─────────────┐
              │ Siguiente│  │ COMPLETADO  │
              │  huella  │  └─────────────┘
              └────┬────┘
                   │
                   └──────▶ (volver a ENVIANDO SRLEGI)
```

## Manejo de Errores

### Error de Conexión
```vb
Private Sub WinSock_Error(...)
    ' - Muestra mensaje de error
    ' - Rehabilita controles UI
    ' - Cierra conexión WinSock
End Sub
```

### Mensaje Desconocido
Si se recibe un mensaje no reconocido, la aplicación muestra el mensaje y termina.

## Sistema de Logging

La aplicación mantiene un log en `C:\logBanca.txt`:
```vb
' Al iniciar (Form_Load)
Open "C:\logBanca.txt" For Output As #1    ' Crea/sobrescribe
Print #1, Now()
Close #1

' Durante operación (Log)
Open "C:\logBanca.txt" For Append As #1    ' Añade
Print #1, texto & Now()
Close #1
```

## Funciones Auxiliares

| Función | Descripción |
|---------|-------------|
| `AbrirConexionSQLServer()` | Conecta a SQL Server, lee base vigente y reconecta |
| `SetearRsW()` | Configura y ejecuta consulta en recordset ADO |
| `BinAHex()` | Convierte array de bytes a string hexadecimal |
| `CerosIzquierda()` | Rellena string con ceros a la izquierda |
| `strString()` | Formatea string a longitud fija (izq/der) |
| `NullCadena()` | Maneja valores nulos de base de datos |
| `Log()` | Escribe en archivo de log con timestamp |

## Diferencias con EnvioHuellasSQV (enviar_huella)

| Característica | EnviadorPro | EnvioHuellasSQV |
|----------------|-------------|-----------------|
| Nombre interno | EnviadorPro | EnvioHuellasSQV |
| Carga de huellas | Previa en memoria | Durante envío |
| Reintento automático | Sí (TNACK) | No |
| Sistema de logging | Sí (C:\logBanca.txt) | No |
| Detección fin | "ERRE" | "TVERRE" |
| Control de estado | Variable `enviado` | Implícito |

## Consideraciones de Uso

1. **Carga previa**: Las huellas se cargan en memoria antes de iniciar la conexión
2. **Reintentos automáticos**: Si recibe TNACK, reintenta automáticamente la misma huella
3. **Una sola banca a la vez**: Solo puede enviar a una banca por ejecución
4. **Log persistente**: El archivo de log se sobrescribe en cada ejecución
5. **Validación de banca**: Verifica que la banca exista en la tabla BancasIp
6. **Progreso visual**: La barra de progreso muestra el avance del envío
7. **Bloqueo durante envío**: Los controles se deshabilitan mientras se envían las huellas
8. **Modo prueba**: Muestra advertencia si la base está en modo prueba

## Interacción con Otros Componentes

### Relación con servidor_banca
- **EnviadorPro** se conecta directamente a las bancas
- Es una alternativa manual al envío masivo que realiza **servidor_banca**
- Útil para sincronizar bancas individuales sin afectar al resto

### Relación con sqv_server
- Ambos leen de la misma base de datos de legisladores
- sqv_server coordina las operaciones de votación
- EnviadorPro solo se encarga de la sincronización biométrica

### Relación con enviar_huella
- Ambos cumplen la misma función básica
- EnviadorPro es una versión mejorada con reintentos y logging
- Comparten el mismo protocolo de comunicación
