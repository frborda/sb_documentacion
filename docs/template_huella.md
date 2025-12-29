# Documentación de Formatos Biométricos - Sistema NECAR_HCDN

## Resumen

El sistema de votación legislativo almacena datos biométricos de huellas digitales en la tabla `FingerPrints` con tres campos principales:

| Campo | Propósito | Formato |
|-------|-----------|---------|
| `Imagen` | Template completo + imagen embebida | Propietario (Nitgen) |
| `Minutia` | Template de minutiae para matching | PC2.A1 (Nitgen) |
| `MinutiaFut` | Template alternativo/backup | Formato secundario |

---

## 1. Campo `Imagen` - Template Completo

### 1.1 Estructura General

El campo `Imagen` contiene un template biométrico completo en formato propietario, organizado en secciones delimitadas por marcadores `FF Ax`.

```
Offset  Marcador  Descripción
------  --------  -----------
0x0000  FF A0     Header principal del template
0x0002  FF A2     Parámetros de imagen/sensor
0x00xx  FF A4     Datos de minutiae (puntos característicos)
0x00xx  FF A5     Datos extendidos de minutiae
0x00xx  FF A6     Metadata del dispositivo/captura
0x00xx  FF A3     Imagen comprimida (JPEG embebido)
```

### 1.2 Header Principal (FF A0)

```hex
FF A0 FF A2 00 11 00 FF 01 E0 01 40 02 43 B8 04 34 E8 01 AD 0C
```

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x00 | `FF A0` | Marcador de inicio del template |
| 0x02 | `FF A2` | Inicio de sección de parámetros |
| 0x04 | `00 11` | Longitud de sección (17 bytes) |
| 0x06 | `00 FF` | Versión del formato |
| 0x08 | `01 E0` | Ancho de imagen: 480 px (0x01E0) |
| 0x0A | `01 40` | Alto de imagen: 320 px (0x0140) |
| 0x0C | `02 43` | Resolución: 579 DPI (0x0243) |
| 0x0E | `B8 04` | Flags/configuración del sensor |
| 0x10 | `34 E8` | Calidad de captura |
| 0x12 | `01 AD` | Número de minutiae detectadas |
| 0x14 | `0C` | Bits por píxel / formato de compresión |

### 1.3 Sección de Minutiae (FF A4)

```hex
FF A4 00 3A 09 07 00 00 09 32 D3 26 3C 00 0A E0 F3 1A 84 01...
```

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x00 | `FF A4` | Marcador de sección minutiae |
| 0x02 | `00 3A` | Longitud de sección (58 bytes) |
| 0x04 | `09 07` | Número de minutiae en esta sección |
| 0x06+ | ... | Array de minutiae (6 bytes cada una) |

#### Estructura de cada Minutia (6 bytes):

```
Byte 0-1: Coordenada X (16 bits, big-endian)
Byte 2-3: Coordenada Y (16 bits, big-endian)
Byte 4:   Ángulo (0-255 mapeado a 0°-360°)
Byte 5:   Tipo + Calidad
          - Bits 7-6: Tipo (00=ending, 01=bifurcation, 10=other)
          - Bits 5-0: Calidad (0-63)
```

**Ejemplo de minutia decodificada:**
```hex
09 32 D3 26 3C 00
```
- X = 0x0932 = 2354
- Y = 0xD326 = 54054 (normalizado a dimensiones de imagen)
- Ángulo = 0x3C = 60 → 60 * 360/256 = 84.4°
- Tipo/Calidad = 0x00

### 1.4 Datos Extendidos (FF A5)

```hex
FF A5 01 85 02 00 2C 02 27 63 02 2F 43...
```

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x00 | `FF A5` | Marcador de datos extendidos |
| 0x02 | `01 85` | Longitud (389 bytes) |
| 0x04 | `02 00` | Versión de datos extendidos |
| 0x06+ | ... | Pares de coordenadas adicionales (ridge flow, core/delta) |

Esta sección contiene:
- Información de flujo de crestas (ridge flow)
- Puntos singulares (core, delta)
- Datos de curvatura para matching más preciso

### 1.5 Metadata (FF A6)

```hex
FF A6 00 64 00 00 02 01 02 03 03 06 07 07 0A 0F 0A 0F...
```

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x00 | `FF A6` | Marcador de metadata |
| 0x02 | `00 64` | Longitud (100 bytes) |
| 0x04 | `00 00` | Reservado |
| 0x06 | `02 01` | ID del tipo de sensor |
| 0x08+ | ... | Histograma de calidad, timestamps, etc. |

### 1.6 Imagen Comprimida (FF A3)

```hex
FF A3 00 03 00 F6 F6 F6 FA 7B 7F B7 B7 D3 DB DB DB...
```

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x00 | `FF A3` | Marcador de imagen |
| 0x02 | `00 03` | Tipo de compresión (03 = JPEG modificado) |
| 0x04+ | ... | Datos de imagen comprimida |

La imagen está en un formato JPEG modificado. Los patrones `F6 F6 F6`, `DB DB DB` corresponden a tablas de cuantización y Huffman de JPEG.

**Nota:** Esta imagen NO es directamente extraíble como JPG estándar sin procesamiento adicional.

---

## 2. Campo `Minutia` - Formato PC2.A1 (Nitgen)

### 2.1 Estructura General

```hex
50 43 32 00 41 31 00 05 73 04 28 E7 3D 5A 32 00 3A 08 18 03...
```

Este es el formato nativo de Nitgen para almacenamiento y matching de minutiae.

### 2.2 Header PC2

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x00-0x02 | `50 43 32` | Magic: "PC2" (ASCII) |
| 0x03 | `00` | Null terminator |
| 0x04-0x05 | `41 31` | Versión: "A1" (ASCII) |
| 0x06 | `00` | Null terminator |
| 0x07 | `05` | Tipo de template (05 = fingerprint) |
| 0x08 | `73` | Flags de captura |
| 0x09 | `04` | Número de dedos registrados |

### 2.3 Datos de Configuración

```hex
28 E7 3D 5A 32 00 3A 08 18 03 1B 01 80 07 FC 07 FE 01 FF 02
```

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x0A-0x0D | `28 E7 3D 5A` | Timestamp de registro (Unix time) |
| 0x0E-0x0F | `32 00` | ID del sensor |
| 0x10 | `3A` | Calidad mínima aceptada (58/100) |
| 0x11 | `08` | Resolución del template (8 = 500 DPI) |
| 0x12 | `18` | Número de minutiae (24) |
| 0x13 | `03` | Tipo de dedo (03 = índice derecho) |
| 0x14-0x15 | `1B 01` | Dimensiones normalizadas |
| 0x16-0x19 | `80 07 FC 07` | Bounding box de minutiae |

### 2.4 Mapa de Bits de Características

```hex
FE 01 FF 02 FF E1 FF F8 FF FC FF FE 7F FF BF FF DF FF EF F3...
```

Esta sección contiene un mapa de bits que representa la densidad de minutiae en diferentes regiones de la huella, usado para acelerar el proceso de matching.

### 2.5 Array de Minutiae Comprimido

A partir del offset 0x4D aproximadamente:

```hex
04 D3 20 0C C0 43 FE 26 20 41 3F 82 41 42 7C BE E0...
```

Cada minutia se codifica en 4 bytes (formato comprimido):

```
Bits 31-22: Coordenada X (10 bits, 0-1023)
Bits 21-12: Coordenada Y (10 bits, 0-1023)
Bits 11-4:  Ángulo (8 bits, 0-255)
Bits 3-0:   Tipo + Calidad (4 bits)
```

**Ejemplo de decodificación:**
```hex
04 D3 20 0C
```
En binario: `00000100 11010011 00100000 00001100`

- X = 0b0000010011 = 19
- Y = 0b0100110010 = 306
- Ángulo = 0b00000000 = 0
- Tipo/Calidad = 0b1100 = 12

### 2.6 Datos de Ridge Count

La sección final contiene la matriz de ridge count (conteo de crestas entre minutiae), usada para mejorar la precisión del matching:

```hex
31 00 10 1F 21 00 00 22 0F 0E 00 0F 0F 10 20 20...
```

---

## 3. Campo `MinutiaFut` - Formato Alternativo

### 3.1 Estructura General

```hex
9D 02 02 03 00 58 12 19 41 77 77 77 77 77 77 77 77...
```

Este campo contiene un template en formato diferente, posiblemente para:
- Compatibilidad con otro modelo de lector
- Backup en formato estándar (ISO/IEC 19794-2)
- Migración futura a otro sistema

### 3.2 Header

| Offset | Bytes | Interpretación |
|--------|-------|----------------|
| 0x00 | `9D` | Identificador de formato |
| 0x01 | `02` | Versión mayor |
| 0x02 | `02` | Versión menor |
| 0x03 | `03` | Tipo de biometría (03 = fingerprint) |
| 0x04 | `00` | Reservado |
| 0x05 | `58` | Longitud de datos (88 bytes) |
| 0x06 | `12` | Número de minutiae (18) |
| 0x07 | `19` | Calidad promedio (25/100) |
| 0x08 | `41` | Flags |

### 3.3 Padding

Los bytes `77` repetidos son padding para alinear la estructura:
```hex
77 77 77 77 77 77 77 77 77 77 77 77 77 77 77 77
```

### 3.4 Matriz de Densidad

```hex
0B 0A 0A 08 04 03 01 00 3A 38 36 33 33 77 77 77 77
0F 0D 09 08 07 05 03 01 3B 39 37 36 34 32 30 77 77 77 77
```

Esta es una matriz 16x16 que representa la densidad de crestas en cada región de la huella. Cada byte indica el número de crestas en esa celda.

### 3.5 Datos de Minutiae

A partir del offset 0x120 aproximadamente:

```hex
25 6F 64 47 6A 3A 54 62 22 71 49 51 7E 79 4D 52...
```

Las minutiae están codificadas en 3 bytes cada una:

```
Byte 0: Coordenada X (0-255, normalizada)
Byte 1: Coordenada Y (0-255, normalizada)
Byte 2: Ángulo + Tipo
        - Bits 7-2: Ángulo (0-63, * 5.625° = 0°-360°)
        - Bits 1-0: Tipo (00=ending, 01=bifurcation)
```

### 3.6 Tabla de Índices

La sección final contiene referencias a las minutiae por región:

```hex
00 00 00 00 00 12 16 1B 2B 48 4F 4F 51 55 5E 61 62 68...
```

---

## 4. Comparación de Formatos

| Característica | Imagen | Minutia (PC2) | MinutiaFut |
|----------------|--------|---------------|------------|
| Tamaño típico | ~2.8 KB | ~0.5 KB | ~0.6 KB |
| Contiene imagen | Sí (JPEG mod) | No | No |
| Precisión coords | 16 bits | 10 bits | 8 bits |
| Minutiae máx | ~400 | ~100 | ~100 |
| Uso principal | Visualización + matching | Matching primario | Backup/compatibilidad |

---

## 5. Recomendaciones para Migración

### 5.1 Para autenticación en nuevo sistema

1. **Usar campo `Minutia` (PC2)** como fuente primaria
2. El SDK de Nitgen puede leer este formato directamente
3. Si se migra a otro SDK, convertir a ISO/IEC 19794-2

### 5.2 Para extraer imágenes visuales

1. El campo `Imagen` requiere parsear la estructura completa
2. Extraer sección `FF A3` y reconstruir header JPEG
3. Alternativa: capturar nuevas imágenes con el lector

### 5.3 Código de ejemplo (Go)

```go
// Estructura para parsear header PC2
type PC2Header struct {
    Magic      [4]byte  // "PC2\0"
    Version    [3]byte  // "A1\0"
    Type       uint8    // 0x05 = fingerprint
    Flags      uint8
    FingerCount uint8
}

// Leer template PC2
func ParsePC2Template(data []byte) (*PC2Header, error) {
    if len(data) < 10 {
        return nil, errors.New("data too short")
    }
    if string(data[0:3]) != "PC2" {
        return nil, errors.New("invalid magic")
    }
    
    header := &PC2Header{}
    copy(header.Magic[:], data[0:4])
    copy(header.Version[:], data[4:7])
    header.Type = data[7]
    header.Flags = data[8]
    header.FingerCount = data[9]
    
    return header, nil
}
```

---

## 6. Referencias

- Nitgen SDK Documentation (versión correspondiente al hardware)
- ISO/IEC 19794-2:2011 - Finger minutiae data
- ANSI/INCITS 378-2004 - Finger Minutiae Format for Data Interchange

