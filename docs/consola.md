# Consola de Operación SQV

## Descripción General

**Consola de Operación SQV** es la aplicación principal de gestión del Sistema de Votación (SQV) desarrollada en Visual Basic 6 para la **Honorable Cámara de Diputados de la Nación Argentina**. Permite controlar sesiones legislativas, gestionar votaciones en tiempo real, administrar legisladores, y generar actas de votación.

## Arquitectura

```
┌─────────────────────────────────────────────────────────────────────────┐
│                        CONSOLA DE OPERACIÓN SQV                          │
├─────────────────────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐    │
│  │   Login     │  │    Menú     │  │  Consola    │  │  Consultas  │    │
│  │  frmLogin   │──│  frmMenu    │──│  Operación  │  │   Actas     │    │
│  └─────────────┘  └─────────────┘  └─────────────┘  └─────────────┘    │
└─────────────────────────────────────────────────────────────────────────┘
         │                                    │
         │ ADO/SQL                           │ TCP/WinSock
         ▼                                    ▼
┌─────────────────────┐              ┌──────────────────┐
│   SQL Server        │              │   sqv_server     │
│   - SQV_Config      │              │   (Puerto 8001)  │
│   - Base Vigente    │              └──────────────────┘
│   - Base Prueba     │                       │
└─────────────────────┘                       ▼
                                     ┌──────────────────┐
                                     │  Bancas (1-257)  │
                                     │  (Puerto 7000)   │
                                     └──────────────────┘
```

## Componentes Principales

### Módulo de Autenticación

#### `frmLogin.frm` - Inicio de Sesión
- **Autenticación**: Usuario y contraseña contra tabla `UsuarioConsola`
- **Modo Prueba**: Checkbox para alternar entre base de producción y prueba
- **Permisos**: Carga permisos del usuario en estructura `PermisosTotales`

```
Estructura de Permisos:
├── ABMLegisladores
├── ABMUsuarios
├── ConsultaActas
├── DefinirOrdenPresidente
├── ExportaActas
├── ModificaActas
├── UsuarioMantenimiento
├── ImprimeActas
├── ConsultaABMLegislador
├── HabilitaBotonesConsola
└── ActualizaaSB
```

### Menú Principal

#### `frmMenu.frm` - Menú Principal
Opciones principales del sistema:

| Botón | Función | Formulario |
|-------|---------|------------|
| Consola de Operación | Acceso a consola de votación | `frmConsolaOperacion` |
| Consultas | Consulta de actas | `frmConsultas` |
| Configuraciones | Configuración del sistema | `FrmConfigurarConsola` |
| Orden de Selección de Presidente | Definir orden de presidencia | `frmOrdenSeleccionPresidente` |
| Períodos Legislativos | Gestión de períodos | `frmAltaPeriodo` |
| Modificación de Datos | ABM de datos | `frmPreABM` |
| Log de Identificaciones | Ver log biométrico | `frmLogIdentificaciones` |
| Estadísticas | Estadísticas del sistema | `frmPreEstadisticas` |
| Salir | Cerrar sesión | - |

#### Tipos de Usuario y Permisos

```
Tipo 0 - Administrador:       Acceso total
Tipo 1 - Administrador Bancas: Solo configuración de bancas
Tipo 2 - Operador Avanzado:   Operación sin configuración
Tipo 3 - Operador Básico:     Operación sin configuración
Tipo 4 - Operador Consulta:   Solo consultas de actas
```

### Consola de Operación

#### `FRMCONSOLAOPERACIONAEB.FRM` - Consola Principal
Formulario principal de operación que incluye:

**Controles de Sesión:**
- `cmdPeriodoLegislativo`: Seleccionar período legislativo
- `cmdNuevaSesion`: Crear nueva sesión
- `cmdCambiarSesion`: Cambiar a otra sesión

**Controles de Votación:**
- `cmdVotacion`: Iniciar votación
- `cmdCancelarVotacion`: Cancelar votación en curso
- `cmdAbstenciones`: Selector de abstenciones
- `cmdReconsiderar`: Reconsiderar votación

**Controles de Configuración:**
- `dcTipoMayoria`: Tipo de mayoría requerida
- `dcBaseMayoria`: Base para calcular mayoría
- `dcTipoOperacion`: Tipo de operación
- `dcTipoQuorum`: Tipo de quórum
- `dcAbstencion`: Configuración de abstenciones

**Indicadores en Tiempo Real:**
- `txtPresentes`: Legisladores presentes
- `txtAusentes`: Legisladores ausentes
- `txtSi`: Votos afirmativos
- `txtNo`: Votos negativos
- `txtAbs`: Abstenciones
- `txtTiempo`: Tiempo de votación
- `txtResultado`: Resultado de la votación
- `TxtQuorum`: Estado del quórum

**Visualización del Recinto:**
- Array `shpBanca(1-257)`: Representación visual de cada banca
- Array `ctrBanca(1-257)`: Número/estado de cada banca
- Colores según estado: presente, ausente, votó sí, votó no, abstención

**Comunicación:**
- `Ws`: Control WinSock para comunicación con sqv_server
- `Timer`: Timer para actualización periódica (500ms)

### Gestión de Sesiones

#### `frmCrearSesion.frm` - Crear Nueva Sesión
- **txtSesion**: Número de sesión (autocalculado)
- **dtFecha**: Fecha de inicio
- **txtProximo**: Próximo número de acta
- **chkProrroga**: Indicador de prórroga

```sql
-- Inserción de nueva sesión
INSERT INTO sesion (Período_Legislativo, Sesión, Fecha_de_inicio,
                    Próximo_Acta, Estado_sesión, Prorroga)
VALUES (@periodo, @sesion, @fecha, @proximoActa, 'nueva', @prorroga)
```

### Gestión de Votos

#### `frmDefinirVoto.frm` - Cambiar Voto de Banca
Permite modificar manualmente el voto de una banca:
- `cmdSi`: Asignar voto afirmativo
- `cmdNo`: Asignar voto negativo
- `cmdAbstener`: Registrar abstención
- `cmdVolver`: Cancelar sin cambios

```vb
' Comunicación con sqv_server para cambio de voto
MensajesSQV.cambioVoto Str(mBanca), "S"  ' Voto Sí
MensajesSQV.cambioVoto Str(mBanca), "N"  ' Voto No
MensajesSQV.cambioVoto Str(mBanca), "A"  ' Abstención
```

### Configuración del Sistema

#### `frmConfig.frm` - Configuración de Conexión
Gestiona la conexión a la base de datos:
- **txtServer**: Servidor SQL Server
- **txtBase**: Base de datos
- **txtUsuario**: Usuario SQL
- **txtPassword**: Contraseña
- Guarda configuración encriptada en `Consola.dat`

#### `frmConfigurarUnidadBanca.frm` - Configuración de Bancas
Grilla con información de todas las bancas:

| Columna | Descripción |
|---------|-------------|
| Banca | Número de banca (1-257) |
| IP | Dirección IP de la banca |
| Puerto | Puerto de comunicación |
| Comentario | Observaciones |
| Versión última sinc. | Última sincronización |
| Versión Datos Banca | Versión en la banca |
| Versión Datos SQV | Versión en el servidor |

## Base de Datos

### Tablas Principales

#### `UsuarioConsola`
```sql
-- Usuarios del sistema de consola
Login, Clave, ABMLegisladores, ABMUsuarios, ConsultasActas,
OrdenPresidente, ImportaTitulos, ModificaActas, ControlaMantenimiento,
ImprimeActas, ConsultaLegisladores, HabilitaBotonesConsola, EnvioDatosSB
```

#### `sesion`
```sql
-- Sesiones legislativas
Período_Legislativo, Sesión, Fecha_de_inicio, Próximo_Acta,
Estado_sesión, Prorroga
```

#### `BancasIP`
```sql
-- Configuración de bancas
BancaNumero, Ip, Puerto, Comentario, IdString, Version,
version_datos_banca, version_datos_sqv
```

#### `perparl`
```sql
-- Períodos parlamentarios
Período_Legislativo, Nro_de_Sesion_actual
```

#### `config`
```sql
-- Configuración general
version_datos_sqv, directorio_enrolamiento, archivo_enrolamiento
```

#### `Configuracion`
```sql
-- Variables de configuración
Variable, Valor
-- Ejemplo: 'base_vigente' -> connection string
```

### Consultas Principales

```sql
-- Obtener base vigente
SELECT valor FROM configuracion WHERE variable = 'base_vigente'

-- Validar usuario
SELECT * FROM UsuarioConsola
WHERE Login = @usuario AND Clave = @clave

-- Validar versiones de bancas
SELECT count(BancaNumero) as cuantos FROM bancasip
WHERE version_datos_sqv <> @version

-- Nueva sesión
SELECT max(Sesión) as maximo FROM sesion WHERE Sesión <> 9999

-- Datos de bancas
SELECT * FROM BancasIP ORDER BY BancaNumero
```

## Comunicación con Otros Componentes

### Comunicación con sqv_server
La consola se comunica con `sqv_server` mediante el módulo `MensajesSQV`:

```vb
' Ejemplos de comandos
MensajesSQV.cambiosesion txtSesion.Text
MensajesSQV.cambioVoto Str(mBanca), "S"
MensajesSQV.SincronizarBancas "brc"
```

### Comunicación con Bancas
Gestión de IPs y sincronización:

```vb
' Actualizar IPs en servidor de bancas
Datos.GrabarMensaje "actualizarips", "", "", True
```

## Flujo de Operación

### 1. Inicio de Sesión
```
Usuario inicia aplicación
         │
         ▼
    frmLogin
         │
         ├── Selecciona modo (Producción/Prueba)
         │
         ├── Ingresa credenciales
         │
         ├── Valida contra UsuarioConsola
         │
         └── Carga permisos → frmMenu
```

### 2. Operación de Votación
```
Operador en frmConsolaOperacion
         │
         ├── Selecciona Período Legislativo
         │
         ├── Crea/Selecciona Sesión
         │
         ├── Configura Tipo de Mayoría y Quórum
         │
         ├── Ingresa Título del Orden del Día
         │
         ├── Inicia Votación (cmdVotacion)
         │       │
         │       ├── Legisladores votan en bancas
         │       │
         │       ├── Timer actualiza indicadores
         │       │
         │       └── Se reciben votos via WinSock
         │
         ├── Finaliza Votación
         │
         └── Genera Acta
```

### 3. Validación de Integridad de Bancas
```
Al cargar frmMenu
         │
         └── ValidarVersionBancas()
                  │
                  ├── Compara version_datos_sqv
                  │
                  ├── Compara version_datos_banca vs version
                  │
                  └── Si hay errores → Muestra frmConfigurarUnidadBanca
```

## Lista de Formularios

### Formularios de Operación
| Formulario | Descripción |
|------------|-------------|
| `frmConsolaOperacion` | Consola principal de votación |
| `frmCrearSesion` | Crear nueva sesión |
| `frmCambiarSesion` | Cambiar sesión activa |
| `frmDefinirVoto` | Modificar voto de banca |
| `frmTiempoVotacion` | Configurar tiempo de votación |
| `frmTituloActa` | Ingresar título de acta |
| `frmElegirPresidente` | Seleccionar presidente |

### Formularios de Consulta
| Formulario | Descripción |
|------------|-------------|
| `frmConsultas` | Consultas de actas |
| `frmConsultarActa` | Ver detalle de acta |
| `frmListarActas` | Listar actas |
| `frmMostrarActas` | Mostrar actas |
| `frmHistorico` | Histórico de datos |
| `frmLogIdentificaciones` | Log de identificaciones biométricas |
| `frmEstadisticas` | Estadísticas del sistema |

### Formularios de ABM
| Formulario | Descripción |
|------------|-------------|
| `frmABMLegisladores` | ABM de legisladores |
| `frmABMBloques` | ABM de bloques políticos |
| `frmABMPartidos` | ABM de partidos políticos |
| `frmABMDistritos` | ABM de distritos |
| `frmABMMandatos` | ABM de mandatos |
| `frmABMSecciones` | ABM de secciones |
| `frmAdministradorLegisladores` | Administrador de legisladores |
| `frmGestionarUsuarios` | Gestión de usuarios |

### Formularios de Configuración
| Formulario | Descripción |
|------------|-------------|
| `frmConfig` | Configuración de conexión |
| `frmConfigurarConsola` | Configuración de consola |
| `frmConfigurarUnidadBanca` | Configuración de bancas |
| `frmEditarDatosUnidadBanca` | Editar datos de banca |
| `frmSetearConfig` | Setear configuración |
| `frmControlSistema` | Control del sistema |
| `frmAltaPeriodo` | Alta de período legislativo |
| `frmOrdenSeleccionPresidente` | Orden de selección de presidente |

### Formularios Auxiliares
| Formulario | Descripción |
|------------|-------------|
| `frmLogin` | Inicio de sesión |
| `frmMenu` | Menú principal |
| `frmCargando` | Pantalla de carga |
| `frmConectando` | Indicador de conexión |
| `frmMessageBox` | Cuadro de mensaje personalizado |
| `frmImpresion` | Gestión de impresión |
| `frmVisor` | Visor de documentos |
| `frmFlotante` | Ventana flotante |

## Funciones Auxiliares

### Módulo `Datos`
- `AbrirDB()`: Abre conexión a base de datos
- `SetearRs()`: Configura recordset
- `SenteciaSQl()`: Ejecuta sentencia SQL
- `GrabarMensaje()`: Graba mensaje para sqv_server
- `establecerIP()`: Establece IP de la consola

### Módulo `Funciones`
- `validarCaracter()`: Valida caracteres de entrada
- `validarNumero()`: Valida entrada numérica
- `seleccionadoTxt()`: Selecciona texto en control

### Módulo `Encripta`
- `EncryptString()`: Encripta cadena de conexión

### Módulo `MensajesSQV`
- `cambiosesion()`: Notifica cambio de sesión
- `cambioVoto()`: Notifica cambio de voto
- `SincronizarBancas()`: Sincroniza bancas

## Consideraciones de Uso

1. **Instancia única**: La aplicación verifica `App.PrevInstance` para evitar múltiples instancias
2. **Modo Prueba**: Permite operar en base de datos de prueba sin afectar producción
3. **Permisos**: Cada función verifica permisos del usuario antes de ejecutar
4. **Validación de bancas**: Al iniciar, valida integridad de versiones de datos en bancas
5. **Capturas automáticas**: Opción para habilitar capturas de pantalla automáticas
6. **Tecla ESC**: En la mayoría de formularios, ESC cierra el formulario
7. **Confirmaciones**: Operaciones críticas requieren confirmación del usuario

## Interacción con Otros Componentes

### Relación con sqv_server
- La consola envía comandos de votación y configuración
- sqv_server coordina con las bancas individuales
- Comunicación bidireccional via WinSock

### Relación con servidor_banca
- La consola gestiona configuración de IPs de bancas
- servidor_banca maneja comunicación directa con hardware

### Relación con EnviadorPro/EnvioHuellasSQV
- Comparten la misma base de datos de legisladores
- La consola gestiona datos, los enviadores sincronizan huellas

## Módulo de Estadísticas

### `frmEstadisticas.frm` - Estadísticas Individuales
Genera informes de votaciones nominales por legislador:

**Funcionalidades:**
- Selección de diputados (activos o todos)
- Filtro por rango de fechas
- Lista de impresión múltiple
- Exportación a formato XML para OpenCMS
- Generación de PDF individual por legislador

**Datos Incluidos:**
- Votos afirmativos, negativos, ausencias
- Desempates como presidente
- Detalle de cada acta votada
- Información de bloque y provincia

```vb
' Estructura de datos de estadística
- Desempates_Negativos
- Desempates_Afirmativos
- CantAfirm (votos afirmativos)
- CantNeg (votos negativos)
- CantAus (ausencias)
- Fechas y períodos de cada votación
```

**Consulta Principal de Estadísticas:**
```sql
SELECT Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado,
       detalleactas.bloque_político, distritos.distrito AS Provincia,
       (SELECT COUNT(Resultado) FROM detalleactas
        WHERE Resultado = 'AFIRMATIVO' AND Legislador_asignado = @id) AS CantAfirm,
       (SELECT COUNT(Resultado) FROM detalleactas
        WHERE Resultado = 'NEGATIVO' AND Legislador_asignado = @id) AS CantNeg,
       (SELECT COUNT(Resultado) FROM detalleactas
        WHERE Resultado = 'AUSENTE' AND Legislador_asignado = @id) AS CantAus
FROM detalleactas
INNER JOIN Legisladores ON Legisladores.id = detalleactas.Legislador_asignado
INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito
WHERE detalleactas.Versión_Acta = 0 AND detalleactas.Legislador_asignado = @id
```

## Módulo de Consultas

### `frmConsultas.frm` - Menú de Consultas
Opciones de consulta disponibles:

| Botón | Función | Descripción |
|-------|---------|-------------|
| Listados de Legisladores | `frmListados` | Consultar legisladores |
| Listado datos de Recinto | `frmListadoDatosRecinto` | Bancas probables y huellas |
| Consultar y modificar actas | `frmMostrarPeriodos` | Gestión de actas de sesión |
| Volver al Menú | - | Regresar al menú principal |

## Módulo de Impresión

### `frmImpresion.frm` - Control de Impresión
Gestiona la espera y emisión de actas:

- Timer de verificación cada 100ms (`tmCheck`)
- Timer de animación cada 250ms (`tmPuntos`)
- Espera almacenamiento del acta en BD
- Impresión automática si está habilitada
- Opción de cancelar espera (no recomendado)

```vb
' Variables de control
Ultimo_Periodo    ' Período de la última votación
Ultima_Sesion     ' Sesión de la última votación
Ultimo_Acta       ' Número de la última acta
ImpresionAutomaticaActivada  ' Flag de auto-impresión
```

**Verificación de Acta:**
```sql
SELECT * FROM actas
WHERE Período_Legislativo = @periodo
  AND Sesión = @sesion
  AND Número_de_Acta = @acta
```

## ABM de Legisladores

### `frmABMLegisladores.frm` - Mantenimiento de Legisladores
Gestión completa de datos de legisladores:

**Datos del Legislador:**
| Campo | Descripción |
|-------|-------------|
| ID | Identificador único |
| Nombre | Nombre del legislador |
| Apellido | Apellido del legislador |
| Sexo | Género |
| Fecha Nacimiento | Fecha de nacimiento |
| Mandato | Tipo de mandato |
| Bloque | Bloque político |
| Distrito | Provincia/Distrito |
| Fotografía | Imagen del legislador |
| Personal Mantenimiento | Flag si es personal técnico |

**Gestión de Bancas:**
- `lblBanca`: Número de banca asignado
- `cmdAsignar`: Asignar número de orden
- `cmdDesasignarBanca`: Quitar banca asignada
- `btnSincronizarEnrolador`: Sincronizar con enrolador
- `cmdAbrirArchivoHuella`: Cargar huella dactilar

**Integración con Enrolador:**
- Conexión a base Access (`base_enrolador`)
- Importación/Exportación de datos biométricos
- Sincronización con sistema de huellas

## Tipos de Operaciones

### Operaciones de Votación
| Código | Descripción |
|--------|-------------|
| `votnom` | Votación Nominal |
| `paslis` | Pase de Lista |

### Estados de Sesión
| Estado | Descripción |
|--------|-------------|
| `nueva` | Sesión recién creada |
| `activa` | Sesión en curso |
| `cerrada` | Sesión finalizada |

### Resultados de Votación
| Resultado | Descripción |
|-----------|-------------|
| `AFIRMATIVO` | Voto a favor |
| `NEGATIVO` | Voto en contra |
| `AUSENTE` | No votó |
| `S/VOTO` | Presidente sin voto |
| `Pte.` | Actuó como presidente |

### Tipos de Período Legislativo
El código de período tiene formato `XXXYZ`:
- `XXX`: Número de período (ej: 132)
- `Y`: Tipo de período (O=Ordinario, E=Extraordinario, P=Prórroga, L=Legislativo)
- `Z`: Tipo de sesión (T=Tablas, H=Homenajes, E=Especial, P=Preparatoria, I=Informativa)

## Archivos de Configuración

| Archivo | Descripción |
|---------|-------------|
| `Consola.dat` | Cadena de conexión encriptada |
| `SQV.DAT` | Configuración legacy |
| `bdExportEnrolamiento.mdb` | Base Access para sincronización |

## Variables Globales Importantes

```vb
' Conexiones
strconexion         ' Cadena de conexión activa
strConexionConfig   ' Conexión a SQV_Config
strBaseProduccion   ' Cadena de producción
strBasePrueba       ' Cadena de prueba

' Estado
FlagBasePrueba      ' Indica si está en modo prueba
gLoginSucceeded     ' Login exitoso
gTipoUsuario        ' Tipo de usuario (0-4)
PermisosTotales     ' Estructura con permisos

' Operación
EntroAMenu          ' Flag de entrada a menú
EntroAConsola       ' Flag de entrada a consola
AutoCaptura         ' Capturas automáticas habilitadas
Modo_Prueba_Seleccionado ' Modo prueba activo
```
