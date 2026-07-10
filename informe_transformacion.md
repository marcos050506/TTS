# Informe de Transformación del Dataset TTS

## 1. Objetivo

Transformar un corpus de conversaciones grabadas (audios + transcripciones en .doc) en un dataset estructurado para entrenamiento de TTS, donde cada utterance individual (con audio asociado) se organiza por **intervalos de edad de 5 años** del hablante que lo emitió.

---

## 2. Datos de Entrada

| Elemento | Cantidad | Formato |
|---|---|---|
| Grabaciones (según PDF) | 93 | Códigos GB/GC/GD |
| Transcripciones | 40 | Word .doc (formato Word 97-2003) |
| Audios | 1570 | WAV (nombrados `{turno}.{short_name}.wav`) |

---

## 3. Pipeline de Transformación

### 3.1 Extracción de texto de .doc

Los archivos .doc están en formato Word 97-2003 (OLE2). Se usó `win32com.client` para abrir cada documento con Microsoft Word y extraer el texto completo.

- **Método**: `Word.Application COM automation`
- **Problema encontrado**: El servidor COM no soportaba múltiples invocaciones sucesivas. Solución: abrir y cerrar Word por cada archivo, con `taskkill` para limpiar procesos huérfanos.

### 3.2 Parseo de la transcripción

Cada archivo .doc tiene una estructura de **metadata** (líneas con `@`) seguida del **conversacional** (líneas con intervenciones de participantes).

#### Metadata extraída:
- `@ Título corto: Cam` → nombre corto de la grabación
- `@ Participante: R = Rubén = 18` → mapeo de código de hablante a edad

#### Utterances extraídos:
Se buscaron todos los patrones `((N texto ))N` donde:
- `N` es un número de utterance (1, 2, 3...)
- `texto` es el contenido hablado
- El hablante se determina por la letra al inicio de la línea (`A: ... ((1 ...))1` → hablante A)

**Ejemplo:**
```
Línea: A: ((1 Ese es el típico santiaguero. ))1
→ utterance 1, hablante A, texto: "Ese es el típico santiaguero."
```

### 3.3 Mapeo audio → utterance

Los archivos de audio tienen el formato `{turno}.{short_name}.wav` (ej: `1.Cam.wav`, `2.Cam.wav`).

Se cruza:
- `turno` (ej: 1) ↔ `utterance N` (ej: ((1 ... ))1)
- `short_name` del audio (ej: Cam) ↔ `Título corto` del .doc

### 3.4 Asignación de edad

Para cada utterance se obtiene:
1. Letra del hablante (de la línea de transcripción)
2. Edad del hablante (del header `@ Participante: X = Nombre = Edad`)
3. Intervalo de 5 años: `(edad // 5) * 5` a `((edad // 5) * 5) + 4`

### 3.5 Unificación de nombres

- `Cocina` → `Coc` (unificado con el nombre del PDF)

---

## 4. Output Generado

```
DTset_organizado_v2/
├── {intervalo_edad}/
│   ├── {short_name}_{utterance_N}.wav    ← audio original renombrado
│   └── {short_name}_{utterance_N}.txt    ← texto del utterance
```

### Estadísticas finales

| Métrica | Valor |
|---|---|
| Total utterances en transcripciones | 6,316 |
| Pares audio+texto generados | 1,500 |
| Tasa de matching | 23.7% |
| Intervalos de edad creados | 19 |
| Total archivos WAV copiados | 1,500 |
| Total archivos TXT generados | 1,500 |

### Distribución por intervalo de edad

| Intervalo | Cantidad | Edades representativas |
|-----------|----------|----------------------|
| 15-19 | 110 | 17, 18, 19 |
| 20-24 | 312 | 20, 22, 23 |
| 25-29 | 327 | 26, 27 |
| 30-34 | 171 | 32, 33, 34 |
| 35-39 | 5 | 35, 36, 37, 38, 39 |
| 40-44 | 49 | 40, 43, 44 |
| 45-49 | 1 | 46 |
| 50-54 | 9 | 50, 51, 52, 53, 54 |
| 55-59 | 16 | 55, 56, 57 |
| 70-74 | 27 | 70, 71, 73, 74 |
| 75-79 | 53 | 75, 76, 77, 78, 79 |
| 80-84 | 2 | 80, 82 |
| 85-89 | 46 | 85, 87, 89 |
| 90-94 | 4 | 90, 92, 93 |
| 95-99 | 7 | 95, 96, 97, 98 |
| 100-104 | 229 | 99, 101, 103, 104 |
| 105-109 | 58 | 105, 106, 107, 108 |
| 110-114 | 60 | 110, 111, 112, 113, 114 |
| 120-124 | 14 | 120, 121, 122 |

---

## 5. Decisiones Técnicas

### 5.1 Extracción de .doc

- Se descartó `olefile` porque produce texto con artefactos de codificación (mezcla de UTF-16LE con bytes de formato Word)
- Se optó por `win32com.client` (Word COM) que extrae el texto limpio con la codificación correcta
- Se maneja el COM conflict cerrando y forzando la terminación de WINWORD.EXE entre archivos

### 5.2 Filtro de utterances

- Solo se incluyen utterances con patrón `((N ...))N` (doble paréntesis + número de cierre)
- Se ignoran utterances que no tengan archivo de audio correspondiente
- Se ignoran utterances cuyo hablante no tenga edad registrada en el header

### 5.3 Limpieza de texto

- Se eliminan anotaciones entre corchetes `[N ... ]N`
- Se eliminan marcas de comentario `<N ... >`
- Se colapsan múltiples espacios en uno

---

## 6. Script Generador

El script `organizar_dataset_v2.py` contiene todo el pipeline y puede reutilizarse con el dataset completo cuando esté disponible. Parámetros configurables:

```python
BASE_DIR = r"D:\Marcos\Prog\TTS"    # Ruta base
DTset_DIR = "DTset"                   # Subcarpeta con Audios/ y Transcripcion/
OUT_DIR = "DTset_organizado_v2"       # Subcarpeta de salida
```

---

## 7. Notas para el dataset completo

- El 76.3% de utterances en las transcripciones **no tienen audio asociado** en la muestra (6,316 utterances totales, solo 1,500 con audio). Esto es esperable para una muestra.
- Con el dataset completo, se espera que la tasa de matching mejore significativamente.
- El script asume la misma estructura de nombres de archivo: `{N}.{short_name}.wav`.
- Los intervalos de edad se generan automáticamente según las edades que aparezcan.