# Despliegue de Huellas Dactilares

## Sistema de Votación Legislativo - Guía de Sincronización

**Versión:** 5.02a  
**Módulo:** Gestión de Templates Biométricos

---

## 1. Descripción General

Este documento describe el proceso de despliegue de huellas dactilares desde el servidor hacia las terminales de votación (bancas). El proceso permite cargar y actualizar los templates biométricos de los legisladores en todas las terminales del recinto.

### 1.1 Requisitos Previos

- Conexión TCP/IP establecida con las terminales (Puerto 7000)
- Templates de huellas en formato hexadecimal (2048 bytes por huella)
- Versión de datos preparada en formato YYYYMMDDNNNN

### 1.2 Consideraciones Importantes

- Durante la carga de huellas, el servidor **suspende el envío de comandos STATUS** para evitar interferencias
- El proceso borra las huellas existentes antes de cargar las nuevas
- Se recomienda realizar el despliegue en horarios de baja actividad

---

## 2. Comandos de Despliegue

### 2.1 SLEVER - Consultar Versión Actual

Solicita la versión de datos de huellas almacenados en la terminal.

| Atributo | Valor |
|----------|-------|
| Comando | `SLEVER` |
| Parámetros | Ninguno |
| Usa cola | Sí |
| Respuesta esperada | `TLEVER [VERSION]` |

**Ejemplo:**
```
Servidor → Banca: SLEVER
Banca → Servidor: TLEVER 202401150089
```

La versión `202401150089` indica: fecha 2024-01-15, 89 legisladores cargados.

---

### 2.2 SNUVER - Iniciar Nueva Versión

Inicia el proceso de carga de una nueva versión de huellas. **Este comando borra todas las huellas existentes en la terminal.**

| Atributo | Valor |
|----------|-------|
| Comando | `SNUVER` |
| Parámetros | Versión (12 caracteres) |
| Usa cola | No (envío prioritario) |
| Respuesta esperada | `TACKNL` |

#### Formato de Versión

```
SNUVER YYYYMMDDNNNN
       │       │
       │       └─── Cantidad de legisladores (4 dígitos decimales)
       └─────────── Fecha en formato compacto (YYYYMMDD)
```

**Ejemplos:**
```
SNUVER 202401150089    → Fecha: 15/01/2024, Legisladores: 89
SNUVER 202312010256    → Fecha: 01/12/2023, Legisladores: 256
```

---

### 2.3 SRLEGI - Enviar Template de Huella

Envía un template de huella dactilar individual a la terminal.

| Atributo | Valor |
|----------|-------|
| Comando | `SRLEGI` |
| Parámetros | Tipo + Datos de huella |
| Usa cola | Sí |
| Respuesta esperada | `TACKNL` |

#### Formato Tipo 0 - Solo Huella

Envía únicamente el template biométrico, sin datos personales.

```
SRLEGI 0[ID_HEX][TEMPLATE]
       │ │      │
       │ │      └─── Template de huella (2048 bytes en hexadecimal = 4096 caracteres)
       │ └────────── Identificador único (4 bytes hexadecimal)
       └──────────── Tipo de registro
```

**Ejemplo:**
```
SRLEGI 0001A[4096 caracteres hexadecimales del template]
```

#### Formato Tipo 1 - Huella con Datos Personales

Envía el template junto con información del legislador para mostrar en pantalla.

```
SRLEGI 1[ID_HEX][APELLIDO][NOMBRE][ID][CLASE][ICONO][TEMPLATE]
```

| Campo | Longitud | Descripción |
|-------|----------|-------------|
| Tipo | 1 byte | `1` para registro completo |
| ID_HEX | 4 bytes | Identificador en hexadecimal |
| APELLIDO | 44 bytes | Apellido (rellenar con espacios) |
| NOMBRE | 44 bytes | Nombre (rellenar con espacios) |
| ID | 10 bytes | Número de documento/identificación |
| CLASE | 1 byte | `S` = Legislador, `V` = Mantenimiento |
| ICONO | 1 byte | Código de icono a mostrar |
| TEMPLATE | 2048 bytes | Template en hexadecimal (4096 chars) |

**Ejemplo:**
```
SRLEGI 1001AGOMEZ PEREZ                                  JUAN CARLOS                                  00123456780100[template]
```

---

### 2.4 TVERRE - Confirmación de Carga Completa

Respuesta que envía la terminal cuando todas las huellas fueron procesadas correctamente.

| Atributo | Valor |
|----------|-------|
| Respuesta | `TVERRE` |
| Origen | Terminal |
| Significado | Todas las huellas almacenadas OK |

---

## 3. Flujo de Despliegue Completo

### 3.1 Secuencia de Operaciones

```
┌─────────────────────────────────────────────────────────────────┐
│                    INICIO DE DESPLIEGUE                         │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  1. Suspender polling de STATUS                                 │
│     (evitar interferencias durante la carga)                    │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  2. Enviar SNUVER [versión]                                     │
│     → Borra huellas existentes                                  │
│     → Terminal responde: TACKNL                                 │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  3. Por cada legislador:                                        │
│     → Enviar SRLEGI [datos de huella]                          │
│     → Esperar TACKNL                                           │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  4. Esperar TVERRE                                              │
│     (confirmación de carga completa)                            │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  5. Enviar SLEVER para verificar                                │
│     → Terminal responde: TLEVER [versión]                       │
│     → Validar que coincida con la versión enviada               │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  6. Reanudar polling de STATUS                                  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    FIN DE DESPLIEGUE                            │
└─────────────────────────────────────────────────────────────────┘
```

### 3.2 Ejemplo de Comunicación Completa

```
# Paso 1: Iniciar nueva versión (borra huellas anteriores)
Servidor → Banca: fSNUVER 202401150003
Banca → Servidor: fTACKNL

# Paso 2: Enviar primera huella
Servidor → Banca: }SRLEGI 10001GARCIA...                    [template1]
Banca → Servidor: }TACKNL

# Paso 3: Enviar segunda huella
Servidor → Banca: ~SRLEGI 10002MARTINEZ...                  [template2]
Banca → Servidor: ~TACKNL

# Paso 4: Enviar tercera huella
Servidor → Banca: SRLEGI 10003LOPEZ...                     [template3]
Banca → Servidor: TACKNL

# Paso 5: Terminal confirma carga completa
Banca → Servidor: TVERRE

# Paso 6: Verificar versión
Servidor → Banca: SLEVER
Banca → Servidor: TLEVER 202401150003
```

---

## 4. Modos de Direccionamiento

### 4.1 Despliegue a Todas las Bancas (Broadcast)

Para enviar huellas a todas las terminales conectadas:

```
Destino: "brc"
```

### 4.2 Despliegue Selectivo

Para enviar a terminales específicas usando vector binario:

```
Destino: "1;0;1;1;0;..."
         │ │ │ │ └─ Banca 4: No recibe
         │ │ │ └─── Banca 3: Sí recibe
         │ │ └───── Banca 2: Sí recibe
         │ └─────── Banca 1: No recibe
         └───────── Banca 0: Sí recibe
```

### 4.3 Despliegue Individual

Para enviar a una única banca:

```
Destino: "45"    → Solo banca 45
```

---

## 5. Manejo de Errores

### 5.1 Timeout y Reintentos

| Parámetro | Valor |
|-----------|-------|
| Timeout entre reintentos | 3000 ms |
| Máximo de reintentos | 10 |
| Acción al exceder reintentos | Cerrar conexión |

### 5.2 Respuestas de Error

| Respuesta | Significado | Acción Recomendada |
|-----------|-------------|-------------------|
| `TNACKN` | Comando rechazado | Verificar formato y reintentar |
| Sin respuesta | Timeout | Reintento automático (hasta 10 veces) |
| Conexión cerrada | Error de comunicación | Reconectar y reiniciar proceso |

### 5.3 Verificación Post-Despliegue

Siempre verificar la versión después del despliegue:

```
Servidor → Banca: SLEVER
Banca → Servidor: TLEVER [versión]
```

Si la versión no coincide con la esperada, reiniciar el proceso para esa terminal.

---

## 6. Tabla de Referencia Rápida

### Comandos

| Comando | Descripción | Cola | Respuesta |
|---------|-------------|------|-----------|
| `SLEVER` | Consultar versión | Sí | `TLEVER [ver]` |
| `SNUVER` | Nueva versión (borra) | No | `TACKNL` |
| `SRLEGI` | Enviar huella | Sí | `TACKNL` |

### Respuestas

| Respuesta | Significado |
|-----------|-------------|
| `TLEVER` | Versión de datos actual |
| `TACKNL` | Comando procesado OK |
| `TNACKN` | Error en comando |
| `TVERRE` | Carga completa exitosa |

---

## 7. Consideraciones de Rendimiento

### 7.1 Tamaño de Datos

- Cada template: 2048 bytes (4096 caracteres hex)
- Con datos personales: ~2150 bytes por legislador
- Para 256 legisladores: ~550 KB total

### 7.2 Tiempos Estimados

| Cantidad | Tiempo Aproximado |
|----------|-------------------|
| 50 legisladores | 2-3 minutos |
| 100 legisladores | 4-6 minutos |
| 256 legisladores | 10-15 minutos |

*Tiempos por terminal. El broadcast permite carga simultánea.*

### 7.3 Recomendaciones

1. Realizar el despliegue en horarios sin sesión
2. Verificar conectividad de todas las bancas antes de iniciar
3. Mantener log de versiones por terminal
4. Implementar verificación automática post-despliegue

---

*Documento para uso interno del sistema legislativo.*
