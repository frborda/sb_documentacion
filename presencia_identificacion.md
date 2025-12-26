# Presencia e Identificación en Banca

## Sistema de Votación Legislativo - Guía de Operación

**Versión:** 5.02a  
**Módulo:** Control de Presencia e Identificación Biométrica

---

## 1. Descripción General

Este documento describe los mecanismos de control de presencia e identificación por huella dactilar en las terminales de votación. Incluye los comandos del servidor para solicitar identificación y las respuestas automáticas de la terminal ante cambios de estado.

### 1.1 Componentes del Sistema

| Componente | Función |
|------------|---------|
| Switch de asiento | Detecta presencia física del legislador |
| Lector de huellas | Captura y verifica identidad biométrica |
| Display | Muestra estado e información del legislador |

### 1.2 Estados Relevantes

| Estado | Código | Descripción |
|--------|--------|-------------|
| Inactivo | `EINACT` | Terminal sin legislador identificado |
| Esperando huella | `EIDRXH` | Aguardando lectura de huella |
| Esperando teclado | `EIDRNX` | Aguardando entrada numérica |
| Identificado | `EIDACP` | Legislador identificado correctamente |

---

## 2. Control de Presencia

### 2.1 Funcionamiento del Switch

El switch de asiento es un sensor físico que detecta cuando un legislador se sienta o se levanta de su banca. Los cambios de estado se reportan automáticamente al servidor.

### 2.2 Mensajes de Presencia (Terminal → Servidor)

#### TPRESE - Legislador Presente

La terminal envía este mensaje cuando el switch se cierra (legislador se sienta).

| Atributo | Valor |
|----------|-------|
| Mensaje | `TPRESE` |
| Origen | Terminal (automático) |
| Significado | Switch cerrado - Legislador sentado |
| Respuesta del servidor | `SACKNL` |

#### TAUSEN - Legislador Ausente

La terminal envía este mensaje cuando el switch se abre (legislador se levanta).

| Atributo | Valor |
|----------|-------|
| Mensaje | `TAUSEN` |
| Origen | Terminal (automático) |
| Significado | Switch abierto - Legislador ausente |
| Respuesta del servidor | `SACKNL` |

### 2.3 Flujo de Presencia

```
┌──────────────────────────────────────────────────────────┐
│              LEGISLADOR SE SIENTA                        │
└──────────────────────────────────────────────────────────┘
                         │
                         ▼
              ┌─────────────────────┐
              │  Switch se cierra   │
              └──────────┬──────────┘
                         │
                         ▼
              ┌─────────────────────┐
              │ Terminal → Servidor │
              │      TPRESE         │
              └──────────┬──────────┘
                         │
                         ▼
              ┌─────────────────────┐
              │ Servidor → Terminal │
              │      SACKNL         │
              └─────────────────────┘


┌──────────────────────────────────────────────────────────┐
│              LEGISLADOR SE LEVANTA                       │
└──────────────────────────────────────────────────────────┘
                         │
                         ▼
              ┌─────────────────────┐
              │   Switch se abre    │
              └──────────┬──────────┘
                         │
                         ▼
              ┌─────────────────────┐
              │ Terminal → Servidor │
              │      TAUSEN         │
              └──────────┬──────────┘
                         │
                         ▼
              ┌─────────────────────┐
              │ Servidor → Terminal │
              │      SACKNL         │
              └─────────────────────┘
```

### 2.4 Ejemplo de Comunicación

```
# Legislador se sienta
Banca → Servidor: ATPRESE
Servidor → Banca: ASACKNL

# Legislador se levanta
Banca → Servidor: ATAUSEN
Servidor → Banca: ASACKNL
```

### 2.5 Indicador de Presencia en Respuestas

Todas las respuestas de estado incluyen un indicador de presencia:

| Código | Significado |
|--------|-------------|
| `P` | Presente (switch cerrado) |
| `A` | Ausente (switch abierto) |

**Ejemplo en respuesta TESTAD:**
```
TESTAD EINACTP    → Inactivo, Presente
TESTAD EINACTA    → Inactivo, Ausente
TESTAD EIDACPP    → Identificado, Presente
```

---

## 3. Identificación por Huella Dactilar

### 3.1 Comandos del Servidor

#### SIDRXH - Iniciar Identificación por Huella

Solicita a la terminal que inicie el proceso de captura de huella dactilar.

| Atributo | Valor |
|----------|-------|
| Comando | `SIDRXH` |
| Parámetros | `[^rango]` (opcional) |
| Usa cola | Sí |
| Respuesta esperada | `TACKNL EIDRXHP` |

El parámetro opcional `^rango` permite limitar la búsqueda a un subconjunto de huellas (ejemplo: solo personal de mantenimiento).

**Ejemplos:**
```
SIDRXH              → Buscar en todas las huellas
SIDRXH ^0001-0050   → Buscar solo en rango 1 a 50
```

#### SCANCL - Cancelar Operación

Cancela cualquier operación de identificación en curso.

| Atributo | Valor |
|----------|-------|
| Comando | `SCANCL` |
| Parámetros | Ninguno |
| Usa cola | Sí |
| Respuesta esperada | `TACKNL EINACTP` o `TACKNL EINACTA` |

#### SACKID - Confirmar Identificación

Confirma al terminal que la identificación fue aceptada por el sistema.

| Atributo | Valor |
|----------|-------|
| Comando | `SACKID` |
| Parámetros | Ninguno |
| Usa cola | Sí |
| Respuesta esperada | `TACKNL EIDACPP` |

#### SACKNL - Acknowledge Genérico

Confirma recepción de mensaje y permite continuar intentando identificación.

| Atributo | Valor |
|----------|-------|
| Comando | `SACKNL` |
| Parámetros | Ninguno |
| Usa cola | No (envío prioritario) |

#### SAUTOD - Identificación Manual

Asigna una identificación de forma manual sin lectura de huella.

| Atributo | Valor |
|----------|-------|
| Comando | `SAUTOD` |
| Parámetros | ID (10 dígitos) |
| Usa cola | No (envío prioritario) |
| Respuesta esperada | `TACKNL EIDACPP` |

**Ejemplo:**
```
SAUTOD 0012345678    → Asigna ID 0012345678 a la banca
```

---

### 3.2 Respuestas de Identificación (Terminal → Servidor)

#### TIDVAL - Identificación Válida

La terminal encontró una huella coincidente en su base de datos.

| Atributo | Valor |
|----------|-------|
| Respuesta | `TIDVAL [ID]` |
| ID | 16 caracteres hexadecimales |
| Significado | Huella reconocida exitosamente |

**Ejemplo:**
```
TIDVAL 0000000000001234    → Legislador ID 1234 identificado
```

#### TIDINV - Identificación Inválida

La huella capturada no coincide con ninguna en la base de datos.

| Atributo | Valor |
|----------|-------|
| Respuesta | `TIDINV` |
| Significado | Huella no reconocida |
| Acción típica | Servidor envía `SACKNL` para reintentar |

#### TIDOUT - Timeout de Identificación

El tiempo de espera para la captura de huella expiró.

| Atributo | Valor |
|----------|-------|
| Respuesta | `TIDOUT` |
| Significado | No se capturó huella a tiempo |
| Acción típica | Servidor envía `SACKNL` para reintentar |

---

### 3.3 Flujo de Identificación por Huella

```
┌──────────────────────────────────────────────────────────────────┐
│                    INICIO DE IDENTIFICACIÓN                       │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
              ┌───────────────────────────────┐
              │ Servidor → Terminal           │
              │ SCANCL (cancelar estado prev) │
              └───────────────┬───────────────┘
                              │
                              ▼
              ┌───────────────────────────────┐
              │ Terminal → Servidor           │
              │ TACKNL EINACTP               │
              └───────────────┬───────────────┘
                              │
                              ▼
              ┌───────────────────────────────┐
              │ Servidor → Terminal           │
              │ SIDRXH (iniciar lectura)      │
              └───────────────┬───────────────┘
                              │
                              ▼
              ┌───────────────────────────────┐
              │ Terminal → Servidor           │
              │ TACKNL EIDRXHP               │
              │ (esperando huella)            │
              └───────────────┬───────────────┘
                              │
                              ▼
              ┌───────────────────────────────┐
              │   LEGISLADOR COLOCA DEDO      │
              └───────────────┬───────────────┘
                              │
              ┌───────────────┼───────────────┐
              ▼               ▼               ▼
     ┌─────────────┐  ┌─────────────┐  ┌─────────────┐
     │   TIDVAL    │  │   TIDINV    │  │   TIDOUT    │
     │ (ID válido) │  │ (no válido) │  │  (timeout)  │
     └──────┬──────┘  └──────┬──────┘  └──────┬──────┘
            │                │                │
            ▼                └───────┬────────┘
     ┌─────────────┐                 │
     │ Servidor →  │                 ▼
     │ SACKID      │          ┌─────────────┐
     └──────┬──────┘          │ Servidor →  │
            │                 │ SACKNL      │
            ▼                 │ (reintentar)│
     ┌─────────────┐          └──────┬──────┘
     │ TACKNL      │                 │
     │ EIDACPP     │                 ▼
     │(identificado)│         ┌─────────────┐
     └─────────────┘          │ Volver a    │
                              │ EIDRXHP     │
                              └─────────────┘
```

---

### 3.4 Ejemplo de Comunicación Completa

#### Identificación Exitosa

```
# Paso 1: Cancelar estado previo
Servidor → Banca: }SCANCL
Banca → Servidor: }TACKNL EINACTP

# Paso 2: Iniciar lectura de huella
Servidor → Banca: ~SIDRXH
Banca → Servidor: ~TACKNL EIDRXHP

# Paso 3: Legislador coloca dedo - Huella reconocida
Banca → Servidor: ATIDVAL 0000000000001234

# Paso 4: Confirmar identificación
Servidor → Banca: ASACKID
Banca → Servidor: ATACKNL EIDACPP
```

#### Identificación con Reintento

```
# Paso 1-2: Igual que arriba...
Servidor → Banca: }SCANCL
Banca → Servidor: }TACKNL EINACTP
Servidor → Banca: ~SIDRXH
Banca → Servidor: ~TACKNL EIDRXHP

# Paso 3: Huella no reconocida
Banca → Servidor: ATIDINV

# Paso 4: Servidor indica reintentar
Servidor → Banca: fSACKNL

# Paso 5: Legislador reintenta - Ahora sí es válida
Banca → Servidor: ATIDVAL 0000000000001234

# Paso 6: Confirmar identificación
Servidor → Banca: ASACKID
Banca → Servidor: ATACKNL EIDACPP
```

#### Identificación Manual

```
# Cancelar estado previo
Servidor → Banca: fSCANCL
Banca → Servidor: fTACKNL EINACTP

# Asignar identificación directamente
Servidor → Banca: fSAUTOD 0012345678
Banca → Servidor: fTACKNL EIDACPP
```

---

## 4. Consulta de Estado

### 4.1 STATUS - Solicitar Estado Completo

El servidor puede consultar el estado actual de cualquier terminal.

| Atributo | Valor |
|----------|-------|
| Comando | `STATUS` |
| Parámetros | Ninguno |
| Usa cola | Sí |
| Respuesta | `TESTAD [ESTADO][SWITCH]` |

**Ejemplos de respuesta:**
```
TESTAD EINACTP    → Inactivo, Presente
TESTAD EINACTA    → Inactivo, Ausente
TESTAD EIDRXHP    → Esperando huella, Presente
TESTAD EIDACPP    → Identificado, Presente
TESTAD EIDACPA    → Identificado, Ausente (se fue sin desloguearse)
```

### 4.2 Polling Periódico

El servidor realiza polling automático cada 30 segundos:

1. Envía `STATUS` en broadcast a todas las bancas
2. Actualiza estado interno de presencia e identificación
3. Detecta bancas desconectadas

---

## 5. Estados y Transiciones

### 5.1 Diagrama de Estados de Identificación

```
                    ┌─────────────┐
         ┌─────────►│   EINACT    │◄─────────┐
         │          │ (Inactivo)  │          │
         │          └──────┬──────┘          │
         │                 │                 │
         │    SIDRXH       │       SAUTOD    │
         │                 ▼                 │
         │          ┌─────────────┐          │
         │          │   EIDRXH    │          │
    SCANCL          │(Esp. Huella)│          │ SCANCL
         │          └──────┬──────┘          │
         │                 │                 │
         │           TIDVAL│                 │
         │               + │                 │
         │           SACKID│                 │
         │                 ▼                 │
         │          ┌─────────────┐          │
         └──────────│   EIDACP    │──────────┘
                    │(Identificado)│
                    └─────────────┘
```

### 5.2 Tabla de Transiciones

| Estado Origen | Evento/Comando | Estado Destino |
|---------------|----------------|----------------|
| EINACT | SIDRXH | EIDRXH |
| EINACT | SAUTOD | EIDACP |
| EIDRXH | TIDVAL + SACKID | EIDACP |
| EIDRXH | SCANCL | EINACT |
| EIDRXH | TIDINV + SACKNL | EIDRXH (continúa esperando) |
| EIDRXH | TIDOUT + SACKNL | EIDRXH (continúa esperando) |
| EIDACP | SCANCL | EINACT |

---

## 6. Tabla de Referencia Rápida

### 6.1 Comandos del Servidor

| Comando | Parámetros | Descripción | Cola |
|---------|------------|-------------|------|
| `STATUS` | - | Consultar estado | Sí |
| `SCANCL` | - | Cancelar operación | Sí |
| `SIDRXH` | [^rango] | Iniciar ID por huella | Sí |
| `SAUTOD` | ID (10 díg) | ID manual | No |
| `SACKID` | - | Confirmar ID válida | Sí |
| `SACKNL` | - | ACK genérico | No |

### 6.2 Respuestas de la Terminal

| Respuesta | Formato | Descripción |
|-----------|---------|-------------|
| `TESTAD` | `TESTAD [EST][SW]` | Estado completo |
| `TACKNL` | `TACKNL [EST][SW]` | ACK positivo |
| `TPRESE` | `TPRESE` | Switch cerrado (presente) |
| `TAUSEN` | `TAUSEN` | Switch abierto (ausente) |
| `TIDVAL` | `TIDVAL [ID 16 hex]` | ID válida |
| `TIDINV` | `TIDINV` | ID inválida |
| `TIDOUT` | `TIDOUT` | Timeout ID |

### 6.3 Códigos de Estado

| Código | Significado |
|--------|-------------|
| `EINACT` | Inactivo |
| `EIDRXH` | Esperando huella |
| `EIDRNX` | Esperando teclado |
| `EIDACP` | Identificado |

### 6.4 Códigos de Switch

| Código | Significado |
|--------|-------------|
| `P` | Presente |
| `A` | Ausente |

---

## 7. Consideraciones de Implementación

### 7.1 Manejo de Timeouts

- El servidor reintenta comandos automáticamente cada 3 segundos
- Máximo 10 reintentos antes de considerar la banca desconectada
- Los mensajes con secuencia `f` son prioritarios y no pasan por cola

### 7.2 Cache de Identificación

El servidor mantiene un cache de las últimas identificaciones exitosas por banca para optimizar búsquedas posteriores.

### 7.3 Comportamiento ante Ausencia

Si el legislador se levanta (TAUSEN) durante:
- **Identificación en curso**: La operación puede continuar o cancelarse según configuración
- **Ya identificado**: El estado EIDACP se mantiene hasta recibir SCANCL

---

*Documento para uso interno del sistema legislativo.*
