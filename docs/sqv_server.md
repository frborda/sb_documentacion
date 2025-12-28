# SQV Server (sqv_server)

## Descripción General

**SQV Server** (Sistema de Quorum y Votación) es la aplicación principal desarrollada en Visual Basic 6 que gestiona todo el **sistema de votación legislativa**. Coordina el quorum, votaciones nominales/numéricas, identificación biométrica de legisladores, y presenta los resultados en carteles visuales. Se comunica con las bancas físicas a través del **servidor_banca**.

## Arquitectura del Sistema

```
┌─────────────────────────────────────────────────────────────────────────┐
│                            SQV Server                                    │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  ┌──────────────┐   ┌──────────────┐   ┌──────────────┐                │
│  │   frmMain    │   │ frmCartel2011│   │  Datos.bas   │                │
│  │  (Control)   │   │  (Display)   │   │  (Estado)    │                │
│  └──────┬───────┘   └──────────────┘   └──────────────┘                │
│         │                                                               │
│         │ SQL Server                                                    │
│         ▼                                                               │
│  ┌────────────────────────────────────────────────────────────────┐    │
│  │                    Tablas de Mensajes                           │    │
│  │  sqv_sb_mensajes (SQV → Servidor Bancas)                       │    │
│  │  sb_sqv_mensajes (Servidor Bancas → SQV)                       │    │
│  └────────────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────────────┘
                                    │
                                    │ SQL
                                    ▼
                        ┌───────────────────────┐
                        │    Servidor de        │
                        │       Bancas          │
                        └───────────┬───────────┘
                                    │
                    ┌───────────────┼───────────────┐
                    ▼               ▼               ▼
              ┌──────────┐   ┌──────────┐   ┌──────────┐
              │ Banca 1  │   │ Banca 2  │   │ Banca N  │
              └──────────┘   └──────────┘   └──────────┘
```

## Componentes Principales

### Formularios

| Formulario | Descripción |
|------------|-------------|
| `frmMain.frm` | Control principal y lógica de votación |
| `frmCartel2011.frm` | Cartel visual del recinto (168KB) |
| `frmMainAEB.frm` | Versión alternativa del cartel (722KB) |
| `frmConfig.frm` | Configuración de conexión |
| `frmAsamblea.frm` | Vista de asamblea |
| `frmSexto.frm` | Vista sexto intermedio |
| `frmNoFrame.frm` | Vista sin marco |

### Módulos

| Módulo | Descripción |
|--------|-------------|
| `Datos.bas` | Estructuras de datos y estado global |
| `VL.bas` | Votación legislativa |
| `Pendientes.bas` | Gestión de pendientes |
| `VotoRemoto.bas` | Soporte voto remoto |
| `Encriptador.cls` | Encriptación de configuración |

### Clases

| Clase | Descripción |
|-------|-------------|
| `clsLegislador.cls` | Datos de legislador |
| `colLegisladores.cls` | Colección de legisladores |
| `clsSvData.cls` | Datos del servidor |

## Estructuras de Datos Principales

### EstadoServer (Estado Global del Sistema)
```vb
Type EstadoServer
    ' Vectores de estado por banca (0 a UltimaBanca)
    VectorPresencia()           ' "1"=Presente, "0"=Ausente, "X"=Inhabilitada
    VectorIdentificacion()      ' ID del legislador o "0"=No identificado
    VectorColor()               ' Color visual de la banca
    VectorResultados()          ' " "=Abstención, "s"=Sí, "n"=No, "a"=Abstención autorizada
    VectorError()               ' " "=Sin error, "W"=Error IOC
    VTipoIdentificacion()       ' " "=Huella, "T"=Teclado

    ' Estado de operación
    TipoDeOperacion             ' "votnom", "votnum", "quorum", etc.
    EstadoVotacion_y_PasList    ' "votando", "larga", "finalizada", "empate"

    ' Contadores
    Presentes                   As Long
    Ausentes                    As Long
    OcupadosNoIdentificados     As Long

    ' Sesión
    Sesion                      As Long
    NroActa                     As Long
    PeriodoLegislativo          As String

    ' Mayoría y quorum
    BaseMayoria                 As String
    TipoMayoria                 As String
    TipoMayoriaQuorum           As String

    ' Presidente
    ModoVotaPresidente          As Boolean
    ResultadoVotoPresidente     As String
    EsperarVotoPresidente       As Boolean

    ' Tiempos
    TiempoParaVotacion          As Long
    FechaVotacion               As Date
End Type
```

### DatosCartel (Datos para Visualización)
```vb
Type DatosCartel
    Presentes                   As Long
    Ausentes                    As Long
    Resultado                   As String
    Afirmativos                 As Long
    Negativos                   As Long
    Abstenciones                As Long
    MinimoVotosParaAfirmativo   As Long
    LeyendaQuorum               As String
    LeyendaTiempo               As String
    LeyendaTipoOperacion        As String
End Type
```

### MensajeSistema (Comunicación con Servidor de Bancas)
```vb
Type MensajeSistema
    sTipo        As String   ' mset, mget, mevt
    sComponente  As String   ' term, term.auth, term.keyb
    sObjeto      As String   ' Número de banca o "brc"
    sAtributo    As String   ' action, state, result
    sValor       As String   ' Valor del mensaje
    sComentario  As String   ' Información adicional
End Type
```

## Constantes de Estado

```vb
' Presencia
PRESENTE           = "1"
AUSENTE            = "0"
BANCA_INHABILITADA = "X"

' Identificación
NO_IDENTIFICADO = "0"

' Votación
ABSTENCION            = " "
AFIRMATIVO            = "s"
NEGATIVO              = "n"
ABSTENCION_AUTORIZADA = "a"

' Tipo de identificación
TIPO_IDENTIFICACION_HUELLA  = " "
TIPO_IDENTIFICACION_TECLADO = "T"

' Colores
cGRIS    = 0   ' Votó (oculto)
cBLANCO  = 1   ' Ausente
cAMARILLO = 2  ' Presente no identificado
cROJO    = 3   ' Voto negativo
cCELESTE = 4   ' Identificado
cNARANJA = 5   ' -
cVERDE   = 6   ' Voto afirmativo
cNEGRO   = 7   ' Abstención autorizada
cOLIVA   = 8   ' -
cAZUL    = 9   ' Error IOC (switch)
cMARRON  = 10  ' Banca inhabilitada
```

## Máquina de Estados de Operación

### Estados Principales
```
                    ┌──────────────┐
                    │   INACTIVO   │
                    │   (quorum)   │
                    └──────┬───────┘
                           │
         ┌─────────────────┼─────────────────┐
         │                 │                 │
         ▼                 ▼                 ▼
┌──────────────┐  ┌──────────────┐  ┌──────────────┐
│   VOTACION   │  │    PASE      │  │ IDENTIFICAR  │
│   NOMINAL    │  │  DE LISTA    │  │   (prueba)   │
│   (votnom)   │  │   (paslis)   │  │              │
└──────┬───────┘  └──────────────┘  └──────────────┘
       │
       ├── votando (habilita teclados)
       │
       ├── larga (tiempo extendido)
       │
       ├── finalizada
       │
       └── empate (espera voto presidente)
```

### Flujo de Votación Nominal

```
        ┌─────────────┐
        │   INICIO    │
        │  (votnom)   │
        └──────┬──────┘
               │ term.keyb/state=onvotnom
               ▼
        ┌─────────────┐
        │  VOTANDO    │ ◀─── Legisladores emiten votos
        └──────┬──────┘      term.keyb.si/no/state=on
               │
               │ Tiempo agotado o todos votaron
               ▼
        ┌─────────────┐
        │   LARGA     │ (Tiempo extendido opcional)
        └──────┬──────┘
               │
               │ term.keyb/state=offvotnom
               ▼
        ┌─────────────┐
        │  EVALUANDO  │
        │  RESULTADO  │
        └──────┬──────┘
               │
       ┌───────┴───────┐
       │               │
       ▼               ▼
┌─────────────┐ ┌─────────────┐
│ FINALIZADA  │ │   EMPATE    │
│ (Afirm/Neg) │ │ (Pdte vota) │
└─────────────┘ └──────┬──────┘
                       │
                       ▼
               ┌─────────────┐
               │ DESEMPATE   │
               └─────────────┘
```

## Interacción con Servidor de Bancas

### Envío de Mensajes (SQV → Servidor de Bancas)
```vb
' Insertar mensaje en tabla sqv_sb_mensajes
Sub EjecutarSQL(strSQL As String)
    Cn.Execute(strSQL)
End Sub

' Ejemplo: Iniciar votación nominal
INSERT INTO sqv_sb_mensajes
    (Tipo, Componente, Objeto, Atributo, Valor)
VALUES
    ('mset', 'term.keyb', 'brc', 'state', 'onvotnom')
```

### Recepción de Mensajes (Servidor de Bancas → SQV)
```vb
' Leer mensajes de tabla sb_sqv_mensajes
Sub LeerMensajesDelSB()
    strSQL = "SELECT * FROM sb_sqv_mensajes WHERE id > " & ultimoID
    ' Procesar cada mensaje
    Select Case sComponente
        Case "term"
            ' Estado de terminal (ok/off)
        Case "term.auth"
            ' Resultado de identificación
        Case "term.keyb.si", "term.keyb.no", "term.keyb.ab"
            ' Voto recibido
        Case "term.seat"
            ' Estado de presencia (switch)
    End Select
End Sub
```

### Mensajes Principales

#### De SQV a Bancas (mset)
| Componente | Atributo | Valor | Descripción |
|------------|----------|-------|-------------|
| term.auth | action | auth_start | Iniciar identificación por huella |
| term.auth | action | auth_cancel | Cancelar identificación |
| term.keyb | state | onvotnom | Iniciar votación nominal |
| term.keyb | state | offvotnom | Finalizar votación nominal |
| term.keyb | state | onvotnum | Iniciar votación numérica |
| term.keyb | state | offvotnum | Finalizar votación numérica |
| term.led1 | state | on | Confirmar identificación |
| term.ledk1 | state | on | Confirmar voto SI |
| term.ledk2 | state | on | Confirmar voto NO |
| term.mon | action | sync | Sincronizar huellas |
| term.mon | action | reset | Reiniciar banca |

#### De Bancas a SQV (mevt)
| Componente | Atributo | Valor | Descripción |
|------------|----------|-------|-------------|
| term | state | ok/off | Estado de conexión |
| term.auth | result | [ID_HEX] | Identificación exitosa |
| term.auth | result | negative | Identificación fallida |
| term.auth | result | timeout | Timeout de identificación |
| term.seat | switch | closed | Legislador sentado |
| term.seat | switch | open | Legislador ausente |
| term.keyb.si | state | on | Voto afirmativo |
| term.keyb.no | state | on | Voto negativo |
| term.keyb.ab | state | on | Abstención |

## Lógica de Colores (AsignarColor)

```vb
Function AsignarColor(xBancaColor As Long) As String
    ' Error de IOC (switch)
    If VectorError = ERROR_IOC Then
        Return cAZUL

    ' Durante presentación de resultados
    ElseIf TipoDeOperacion In ("votnom", "votnum") And
           EstadoVotacion = "finalizada" Then
        If VectorPresenciaCong = BANCA_INHABILITADA Then
            Return cMARRON
        ElseIf VectorPresenciaCong = AUSENTE Then
            Return cBLANCO
        ElseIf VectorPresenciaCong = PRESENTE Then
            If VectorResultados = AFIRMATIVO Then
                Return cVERDE
            ElseIf VectorResultados = NEGATIVO Then
                Return cROJO
            ElseIf VectorResultados = ABSTENCION_AUTORIZADA Then
                Return cNEGRO
            End If
        End If

    ' Durante operación normal
    ElseIf VectorPresencia = BANCA_INHABILITADA Then
        Return cMARRON
    ElseIf VectorPresencia = AUSENTE Then
        Return cBLANCO
    ElseIf VectorPresencia = PRESENTE Then
        If VectorIdentificacion <> NO_IDENTIFICADO Then
            Return cCELESTE
        Else
            Return cAMARILLO
        End If
    End If
End Function
```

## Cartel Visual (frmCartel2011)

### Elementos del Cartel
```
┌────────────────────────────────────────────────────────────────┐
│ PRESENTES: [XX]              AUSENTES: [XX]                    │
├────────────────────────────────────────────────────────────────┤
│ [Fecha]         [Hora]              [Leyenda Quorum]           │
├────────────────────────────────────────────────────────────────┤
│                     [Período y Sesión]                         │
│                     [Título del Acta]                          │
│                     [Mayoría Requerida]                        │
├────────────────────────────────────────────────────────────────┤
│ [Tipo de Operación]                    TIEMPO: [Estado]        │
├────────────────────────────────────────────────────────────────┤
│                                                                │
│                    ┌─────────────────┐                         │
│                    │   HEMICICLO     │                         │
│                    │   (Bancas 1-N)  │                         │
│                    │    [ctrBanca]   │                         │
│                    └─────────────────┘                         │
│                                                                │
├────────────────────────────────────────────────────────────────┤
│                       [RESULTADO]                              │
├────────────────────────────────────────────────────────────────┤
│ AFIRMATIVOS:[XX]   NEGATIVOS:[XX]   ABSTENCIONES:[XX]         │
└────────────────────────────────────────────────────────────────┘
```

### Control ctrBanca
- Representa visualmente cada banca en el hemiciclo
- Cambia de color según el estado
- Muestra el número de banca

## Flujo de Inicialización

```
Main() [Datos.bas]
    │
    ├── Verificar instancia previa
    │
    ├── Encripta.Password = "ClaveInvulnerable350"
    │
    ├── DeterminarStringConexion()
    │       ├── Leer sqv.dat (encriptado)
    │       ├── Conectar a SQV_Config
    │       └── Obtener base_vigente
    │
    ├── LeerConfig()
    │       ├── Cantidad de legisladores
    │       ├── Tiempos de votación
    │       └── Cantidad de bancas
    │
    ├── InicializarValores()
    │       ├── Dimensionar vectores
    │       └── Estados iniciales
    │
    └── frmCartel2011.Show
```

## Tablas de Base de Datos

### Principales
| Tabla | Descripción |
|-------|-------------|
| config | Configuración general del sistema |
| legisladores | Datos de todos los legisladores |
| legisladores_activos | Legisladores habilitados |
| legisladores_sb | Datos para servidor de bancas |
| sesion | Sesiones legislativas |
| BancasIp | IPs y estado de bancas |
| sqv_sb_mensajes | Mensajes SQV → Servidor Bancas |
| sb_sqv_mensajes | Mensajes Servidor Bancas → SQV |
| ComunicacionRapida | Flags de comunicación |
| tipmay | Tipos de mayoría |
| basemay | Bases de mayoría |

### Campos de Config
```sql
SELECT
    Cantidad_de_Legisladores,
    cantidad_de_bancas,
    Segundos_de_inicio_operacion,
    Segundos_de_fin_operacion,
    Tiempo_espera_Pase_de_Lista,
    Sensib_scan_neg,
    version_datos_sqv
FROM config
```

## Gestión del Presidente

```vb
' El presidente (banca 0) tiene tratamiento especial
Public Sub ResetearPresidente()
    With EstadoActual
        .VectorPresencia(0) = AUSENTE
        .VectorIdentificacion(0) = 0
        .VectorResultados(0) = ABSTENCION
    End With
End Sub

' En caso de empate, el presidente puede votar
If xHuboEmpate And ModoVotaPresidente Then
    EsperarVotoPresidente = True
    ' Habilitar votación solo para banca 0
End If
```

## Comunicación Rápida

La tabla `ComunicacionRapida` se usa para flags de control:

```sql
-- Indicar que se imprimió el acta
UPDATE ComunicacionRapida SET ImprimirActa = 1

-- Indicar que el tiempo transcurrió
UPDATE ComunicacionRapida SET TiempoTranscurrido = 1
```

## Diagrama de Secuencia: Votación Completa

```
Operador         SQV_Server        servidor_banca      Bancas
    │                 │                   │              │
    │ Iniciar votación│                   │              │
    │────────────────▶│                   │              │
    │                 │ term.keyb/onvotnom│              │
    │                 │──────────────────▶│              │
    │                 │                   │  SVOTAR 03   │
    │                 │                   │─────────────▶│
    │                 │                   │              │
    │                 │                   │   TVOTOX     │
    │                 │◀──────────────────│◀─────────────│
    │                 │   (actualiza      │              │
    │                 │    vectores)      │              │
    │                 │                   │              │
    │ Ver cartel      │                   │              │
    │◀────────────────│                   │              │
    │                 │                   │              │
    │ Cerrar votación │                   │              │
    │────────────────▶│                   │              │
    │                 │term.keyb/offvotnom│              │
    │                 │──────────────────▶│              │
    │                 │                   │   SFINVT     │
    │                 │                   │─────────────▶│
    │                 │                   │              │
    │                 │  (calcular        │              │
    │                 │   resultado)      │              │
    │                 │                   │              │
    │ Ver resultado   │                   │              │
    │◀────────────────│                   │              │
```

## Consideraciones de Implementación

1. **Singleton**: Solo una instancia puede ejecutarse (`App.PrevInstance`)
2. **Base de datos vigente**: Se selecciona dinámicamente (producción/prueba)
3. **Encriptación**: La configuración se almacena encriptada
4. **Vectores congelados**: Se mantienen copias de estado para presentación
5. **Presidente especial**: La banca 0 tiene lógica diferenciada
6. **Broadcast (brc)**: Permite enviar comandos a todas las bancas
