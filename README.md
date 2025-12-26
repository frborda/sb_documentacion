# Protocolo de Comunicación - Sistema de Votación Legislativo

## Servidor de Bancas - Documentación Técnica

**Versión:** 5.02a  
**Última actualización:** Febrero 2011

---

## 1. Descripción General

El Servidor de Bancas es el componente central que gestiona la comunicación entre el Sistema de Quorum y Votación (SQV) y las terminales de votación instaladas en el recinto legislativo. El sistema soporta hasta 256 terminales (bancas) conectadas simultáneamente.

### 1.1 Arquitectura

```
┌─────────────────┐                     ┌─────────────────┐
│   Sistema SQV   │                     │ Servidor Bancas │
│    (Consola)    │◄───────────────────►│                 │
└─────────────────┘   Base de Datos     └────────┬────────┘
                      SQL Server                 │
                                                 │ TCP/IP
                                                 │ Puerto 7000
                                        ┌────────┴────────┐
                                        ▼                 ▼
                                   ┌─────────┐       ┌─────────┐
                                   │ Banca 1 │  ...  │Banca 256│
                                   └─────────┘       └─────────┘
```

### 1.2 Características Principales

- Comunicación bidireccional vía TCP/IP
- Sistema de cola de mensajes con reintentos automáticos
- Soporte para broadcast y direccionamiento selectivo
- Gestión de identificación biométrica por huella dactilar
- Control de presencia mediante switch de asiento
- Soporte para votación nominal y numérica

---

## 2. Protocolo de Comunicación

### 2.1 Parámetros de Conexión

| Parámetro | Valor |
|-----------|-------|
| Protocolo | TCP/IP |
| Puerto | 7000 |
| Encoding | ASCII |
| Terminador | CR+LF (`\r\n`) |

### 2.2 Formato de Mensajes

Todos los mensajes siguen el formato:

```
[SECUENCIA][COMANDO] [PARÁMETROS]\r\n
```

| Campo | Longitud | Descripción |
|-------|----------|-------------|
| SECUENCIA | 1 byte | Carácter de control de flujo |
| COMANDO | 6 bytes | Identificador del comando en mayúsculas |
| PARÁMETROS | Variable | Datos adicionales (opcional) |

### 2.3 Caracteres de Secuencia

| Secuencia | Descripción |
|-----------|-------------|
| `f` | Mensaje prioritario, sin encolamiento |
| `X` | Comando de control/status |
| `*` | Comando de configuración |
| `A-Z` | Respuesta con timeout desde terminal |
| `}`...`ú` (125-250) | Secuencia de cola del servidor |

**Nota:** Los mensajes con secuencia `f` o `A-Z` no pasan por el sistema de cola y no requieren confirmación de entrega.

---

## 3. Sistema de Cola de Mensajes

### 3.1 Parámetros de Configuración

| Parámetro | Valor | Descripción |
|-----------|-------|-------------|
| Timeout de reintento | 3000 ms | Tiempo entre reintentos |
| Máximo de reintentos | 10 | Intentos antes de cerrar conexión |
| Timeout de mensaje | 6000 ms | Tiempo máximo de espera |

### 3.2 Modos de Direccionamiento

El sistema soporta tres modos de direccionamiento:

#### Broadcast
Envía el mensaje a todas las bancas conectadas.
```
Destino: "brc"
```

#### Vector Binario
Selección múltiple mediante cadena de bits separados por punto y coma.
```
Destino: "1;0;1;1;0;..."
         │ │ │ │ └─ Banca 4: No
         │ │ │ └─── Banca 3: Sí
         │ │ └───── Banca 2: Sí
         │ └─────── Banca 1: No
         └───────── Banca 0: Sí
```

#### Banca Individual
Número de banca como cadena.
```
Destino: "45"
```

### 3.3 Control de Duplicidad

El sistema evita el envío de mensajes duplicados:
- Se registra el último mensaje enviado a cada banca
- Los mensajes idénticos se descartan si se envían dentro de 3 segundos
- Los mensajes diferentes siempre se procesan

### 3.4 Gestión de Reintentos

1. El mensaje se encola con secuencia única (125-250)
2. Se envía y se registra el timestamp
3. Si no hay respuesta en 3 segundos, se reintenta
4. Después de 10 reintentos fallidos, se cierra la conexión
5. Al recibir ACK válido, se elimina de la cola

---

## 4. Comandos del Servidor

### 4.1 Control de Estado

| Comando | Parámetros | Descripción |
|---------|------------|-------------|
| `STATUS` | - | Solicita estado completo de la terminal |
| `SRESET` | - | Reinicia la terminal a estado inicial |
| `SCANCL` | - | Cancela la operación en curso |
| `SLEVER` | - | Solicita versión de datos almacenados |

### 4.2 Configuración

| Comando | Parámetros | Descripción |
|---------|------------|-------------|
| `SCONFG` | 7 dígitos | Configura el modo de operación |

#### Formato de SCONFG

```
SCONFG ABCDEFG
       │││││││
       ││││││└─ Reservado
       │││││└── Reservado
       ││││└─── Modo de identificación
       │││└──── Configuración de timeout
       ││└───── Modo de presencia (0=inactivo, 1=activo)
       │└────── Modo de votación
       └─────── Modo general (2=presencia, 9=completo)
```

Configuraciones comunes:
- `2010010` - Modo presencia únicamente
- `9010061` - Modo completo con identificación por huella

### 4.3 Identificación

| Comando | Parámetros | Descripción |
|---------|------------|-------------|
| `SIDRXH` | [^rango] | Inicia identificación por huella dactilar |
| `SIDRNX` | [^rango] | Inicia identificación por teclado numérico |
| `SAUTOD` | ID (10 dígitos) | Asigna identificación manual |
| `SACKID` | - | Confirma identificación exitosa |
| `SACKNL` | - | Acknowledge genérico |

El parámetro opcional `^rango` permite limitar la búsqueda a un subconjunto de huellas (ej: solo personal de mantenimiento).

### 4.4 Votación

| Comando | Parámetros | Descripción |
|---------|------------|-------------|
| `SVOTAR` | timeout (2 dígitos) | Inicia votación nominal (Sí/No/Abstención) |
| `SVOTNU` | timeout (2 dígitos) | Inicia votación numérica |
| `SFINVT` | - | Finaliza votación nominal |
| `SFINNU` | - | Finaliza votación numérica |
| `SLIMVT` | - | Limpia indicadores de voto en pantalla |
| `SACKVT` | S/N/A | Confirma recepción de voto |

El parámetro de timeout indica segundos (ej: `03` = 3 segundos).

### 4.5 Gestión de Huellas Dactilares

| Comando | Parámetros | Descripción |
|---------|------------|-------------|
| `SNUVER` | versión (12 chars) | Inicia nueva versión, borra huellas existentes |
| `SRLEGI` | datos de huella | Envía template de huella dactilar |

#### Formato de SNUVER
```
SNUVER YYYYMMDDNNNN
       │       │
       │       └─── Cantidad de legisladores (4 dígitos decimales)
       └─────────── Fecha en formato compacto
```

#### Formato de SRLEGI

**Huella sin datos personales (tipo 0):**
```
SRLEGI 0[ID_HEX 4][TEMPLATE 2048 bytes hex]
```

**Huella con datos personales (tipo 1):**
```
SRLEGI 1[ID_HEX 4][APELLIDO 44][NOMBRE 44][ID 10][CLASE 1][ICONO 1][TEMPLATE 2048 bytes hex]
```

| Campo | Longitud | Descripción |
|-------|----------|-------------|
| ID_HEX | 4 bytes | Identificador en hexadecimal |
| APELLIDO | 44 bytes | Apellido del legislador |
| NOMBRE | 44 bytes | Nombre del legislador |
| ID | 10 bytes | Número de identificación |
| CLASE | 1 byte | "S" = Legislador, "V" = Mantenimiento |
| ICONO | 1 byte | Código de icono |
| TEMPLATE | 2048 bytes | Template de huella en hexadecimal |

### 4.6 Display

| Comando | Parámetros | Descripción |
|---------|------------|-------------|
| `SINFOR` | texto (40 chars) | Muestra mensaje en pantalla |

---

## 5. Respuestas de la Terminal

### 5.1 Estado

| Respuesta | Formato | Descripción |
|-----------|---------|-------------|
| `TESTAD` | `TESTAD [ESTADO][SWITCH]` | Estado completo |
| `TLEVER` | `TLEVER [VERSION]` | Versión de datos |
| `TACKNL` | `TACKNL [ESTADO][SWITCH]` | Acknowledge positivo |
| `TNACKN` | `TNACKN [COMANDO]` | Acknowledge negativo |

#### Códigos de Estado

| Código | Descripción |
|--------|-------------|
| `EINACT` | Terminal inactiva (sin identificar) |
| `EIDACP` | Identificación aceptada |
| `EIDRXH` | Esperando huella dactilar |
| `EIDRNX` | Esperando entrada por teclado |
| `ELISVT` | En proceso de votación |

#### Códigos de Switch

| Código | Descripción |
|--------|-------------|
| `P` | Presente (switch cerrado, legislador sentado) |
| `A` | Ausente (switch abierto, legislador no presente) |

### 5.2 Identificación

| Respuesta | Formato | Descripción |
|-----------|---------|-------------|
| `TIDVAL` | `TIDVAL [ID 16 hex]` | Identificación exitosa |
| `TIDINV` | `TIDINV` | Huella no reconocida |
| `TIDOUT` | `TIDOUT` | Timeout de identificación |

El campo ID contiene 16 caracteres hexadecimales que identifican al legislador.

### 5.3 Presencia

| Respuesta | Descripción |
|-----------|-------------|
| `TPRESE` | Legislador se sentó (switch cerrado) |
| `TAUSEN` | Legislador se levantó (switch abierto) |

Estas respuestas se generan automáticamente cuando cambia el estado del switch de presencia.

### 5.4 Votación

| Respuesta | Formato | Descripción |
|-----------|---------|-------------|
| `TVOTOX` | `TVOTOX [VOTO]` | Voto emitido |

| Código de Voto | Significado |
|----------------|-------------|
| `S` | Afirmativo (Sí) |
| `N` | Negativo (No) |
| `A` | Abstención |

### 5.5 Carga de Huellas

| Respuesta | Descripción |
|-----------|-------------|
| `TVERRE` | Todas las huellas fueron procesadas correctamente |

---

## 6. Flujos de Operación

### 6.1 Conexión Inicial

```
1. Terminal establece conexión TCP al puerto 7000
2. Servidor envía: SCANCL (cancelar estado previo)
3. Servidor envía: STATUS (solicitar estado)
4. Terminal responde: TESTAD [estado][switch]
5. Servidor envía: SCONFG [configuración]
6. Servidor envía: SLEVER (solicitar versión)
7. Terminal responde: TLEVER [versión]
```

### 6.2 Identificación por Huella

```
1. Servidor envía: SCANCL
2. Servidor envía: SIDRXH
3. Terminal responde: TACKNL EIDRXHP (esperando huella)

[Legislador coloca dedo en lector]

4a. Si huella válida:
    Terminal responde: TIDVAL [ID hexadecimal]
    Servidor envía: SACKID (confirmar en pantalla)

4b. Si huella no reconocida:
    Terminal responde: TIDINV
    Servidor envía: SACKNL (reintentar)

4c. Si timeout:
    Terminal responde: TIDOUT
    Servidor envía: SACKNL (reintentar)
```

### 6.3 Identificación Manual

```
1. Servidor envía: SCANCL
2. Servidor envía: SAUTOD [ID 10 dígitos]
3. Terminal responde: TACKNL EIDACPP
```

### 6.4 Votación Nominal

```
1. Servidor envía: SVOTAR 03 (timeout 3 segundos)
2. Terminal responde: TACKNL ELISVTP (en votación)

[Legislador presiona botón]

3. Terminal envía: TVOTOX S (votó Sí)
4. Servidor envía: SACKVT S (confirmar voto)

[Fin de votación]

5. Servidor envía: SFINVT
6. Servidor envía: SLIMVT (limpiar pantalla)
```

### 6.5 Votación Numérica

```
1. Servidor envía: SVOTNU 03 (timeout 3 segundos)
2. Terminal responde: TACKNL ELISVTP

[Legislador ingresa número]

3. Terminal envía: TVOTOX [valor]
4. Servidor envía: SACKVT [valor]

[Fin de votación]

5. Servidor envía: SFINNU
6. Servidor envía: SLIMVT
```

### 6.6 Carga de Huellas Dactilares

```
1. Servidor envía: SNUVER [versión] (borra huellas existentes)
2. Terminal responde: TACKNL

[Por cada legislador]
3. Servidor envía: SRLEGI [datos de huella]
4. Terminal responde: TACKNL

[Al finalizar todas las huellas]
5. Terminal envía: TVERRE (carga completa)
6. Servidor envía: SLEVER (verificar versión)
7. Terminal responde: TLEVER [versión]
```

**Nota:** Durante la carga de huellas, el servidor suspende el envío de comandos STATUS para evitar interferencias.

### 6.7 Gestión de Presencia

```
[Legislador se sienta]
1. Terminal envía: TPRESE
2. Servidor envía: SACKNL

[Legislador se levanta]
3. Terminal envía: TAUSEN
4. Servidor envía: SACKNL
```

---

## 7. Estados de la Terminal

### 7.1 Diagrama de Estados

```
                    ┌─────────────┐
         ┌─────────►│   EINACT    │◄─────────┐
         │          │ (Inactivo)  │          │
         │          └──────┬──────┘          │
         │                 │                 │
         │    ┌────────────┼────────────┐    │
         │    ▼            ▼            ▼    │
    ┌────┴─────┐    ┌──────────┐    ┌───────┴───┐
    │ EIDRXH   │    │ EIDRNX   │    │  ELISVT   │
    │(Esp.Hue.)│    │(Esp.Tecl)│    │ (Votando) │
    └────┬─────┘    └────┬─────┘    └───────────┘
         │               │
         └───────┬───────┘
                 ▼
         ┌─────────────┐
         │   EIDACP    │
         │(Identificado)│
         └─────────────┘
```

### 7.2 Transiciones

| Estado Origen | Comando | Estado Destino |
|---------------|---------|----------------|
| EINACT | SIDRXH | EIDRXH |
| EINACT | SIDRNX | EIDRNX |
| EINACT | SVOTAR/SVOTNU | ELISVT |
| EIDRXH | TIDVAL + SACKID | EIDACP |
| EIDRXH | SCANCL | EINACT |
| EIDRNX | Identificación OK | EIDACP |
| EIDACP | SCANCL | EINACT |
| ELISVT | SFINVT/SFINNU | EINACT |

---

## 8. Condiciones de Eliminación de Cola

Cada comando en cola se elimina bajo condiciones específicas:

| Comando | Condición de Eliminación |
|---------|--------------------------|
| `SCANCL` | TACKNL EINACTP, TACKNL EINACTA, TACKNL EIDACPP |
| `SIDRXH` | TACKNL EIDRXHP, cualquier TACKNL o TNACKN |
| `SIDRNX` | TACKNL EIDRNXP, TNACKN EIDRNXP, switch Ausente |
| `SVOTAR` | TACKNL ELISVTP, TNACKN ELISVTP, switch Ausente |
| `SVOTNU` | TACKNL ELISVTP, TNACKN ELISVTP, switch Ausente |
| `SLIMVT` | Cualquier respuesta |
| `SNUVER` | Cualquier respuesta |
| `SRLEGI` | Cualquier respuesta |

---

## 9. Polling y Monitoreo

### 9.1 Polling Periódico

El servidor realiza las siguientes acciones cada 30 segundos:

1. Envía `STATUS` en broadcast a todas las bancas
2. Verifica conexiones y reconecta bancas desconectadas
3. Actualiza estado interno de cada terminal

**Excepción:** El polling se suspende durante la carga masiva de huellas.

### 9.2 Detección de Desconexión

Una banca se considera desconectada cuando:
- El socket TCP se cierra
- Se exceden 10 reintentos sin respuesta
- Se detecta error de comunicación

Al detectar desconexión:
1. Se notifica al SQV el estado "off"
2. Se limpia la cola de mensajes de esa banca
3. Se programa reconexión automática

### 9.3 Reconexión Automática

El servidor intenta reconectar automáticamente las bancas desconectadas en cada ciclo de polling. Una vez reconectada:

1. Se envía `SCANCL` para limpiar estado
2. Se solicita `STATUS` para sincronizar
3. Se reenvía configuración con `SCONFG`
4. Se verifica versión de huellas con `SLEVER`

---

## 10. Mantenimiento de Memoria

### 10.1 Limpieza de Cola

La cola de mensajes se reinicia automáticamente cuando:
- No hay mensajes pendientes
- Han pasado más de 15 minutos desde la última limpieza

Esto libera memoria acumulada por registros eliminados.

### 10.2 Cache de Identificación

El servidor mantiene un cache de las últimas identificaciones por banca para optimizar la búsqueda de huellas. Este cache se actualiza automáticamente con cada identificación exitosa.

---

## 11. Tabla de Referencia Rápida

### Comandos del Servidor (S)

| Comando | Cola | Descripción |
|---------|------|-------------|
| STATUS | Sí | Estado de terminal |
| SRESET | Sí | Reiniciar terminal |
| SCANCL | Sí | Cancelar operación |
| SCONFG | No | Configurar modo |
| SLEVER | Sí | Solicitar versión |
| SIDRXH | Sí | Identificación huella |
| SIDRNX | Sí | Identificación teclado |
| SAUTOD | No | ID manual |
| SACKID | Sí | Confirmar ID |
| SACKNL | No | ACK genérico |
| SVOTAR | Sí | Votación nominal |
| SVOTNU | Sí | Votación numérica |
| SFINVT | Sí | Fin votación nominal |
| SFINNU | Sí | Fin votación numérica |
| SLIMVT | Sí | Limpiar votos |
| SACKVT | No | Confirmar voto |
| SNUVER | No | Nueva versión |
| SRLEGI | Sí | Enviar huella |
| SINFOR | No | Mostrar texto |

### Respuestas de Terminal (T)

| Respuesta | Descripción |
|-----------|-------------|
| TESTAD | Estado completo |
| TLEVER | Versión de datos |
| TACKNL | ACK positivo |
| TNACKN | ACK negativo |
| TIDVAL | ID válida |
| TIDINV | ID inválida |
| TIDOUT | Timeout ID |
| TPRESE | Switch cerrado |
| TAUSEN | Switch abierto |
| TVOTOX | Voto emitido |
| TVERRE | Huellas OK |

---

## 12. Apéndice: Ejemplos de Comunicación

### Inicio de Sesión
```
S → B: fSCANCL
B → S: fTACKNL EINACTP
S → B: XSTATUS
B → S: XTESTAD EINACTP
S → B: *SCONFG 9010061
B → S: *TACKNL EINACTP
```

### Identificación Exitosa
```
S → B: }SCANCL
B → S: }TACKNL EINACTP
S → B: ~SIDRXH
B → S: ~TACKNL EIDRXHP
B → S: ATIDVAL 0000000000001234
S → B: SACKID
B → S: ATACKNL EIDACPP
```

### Votación
```
S → B: }SVOTAR 03
B → S: }TACKNL ELISVTP
B → S: ATVOTOX S
S → B: XSACKVT S
S → B: ~SFINVT
B → S: ~TACKNL EINACTP
S → B: SLIMVT
B → S: TACKNL EINACTP
```

---

*Documento generado para uso interno del sistema legislativo.*