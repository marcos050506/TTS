# Dataset TTS - Transformacion de corpus conversacional

Transformacion de un corpus de conversaciones grabadas en un dataset estructurado para entrenamiento de TTS. Cada utterance individual con audio asociado se organiza por intervalos de edad de 5 anos del hablante que lo emitio.

## Datos de entrada

| Elemento | Cantidad | Formato |
|---|---|---|
| Grabaciones (segun PDF) | 93 | Codigos GB/GC/GD |
| Transcripciones | 40 | Word .doc (Word 97-2003) |
| Audios | 1570 | WAV (nombrados `{turno}.{short_name}.wav`) |

## Pipeline de transformacion

### Extraccion de texto de .doc

Los archivos .doc estan en formato Word 97-2003 (OLE2). Se uso win32com.client para abrir cada documento con Microsoft Word y extraer el texto completo. Se descarto olefile porque producia texto con artefactos de codificacion (mezcla de UTF-16LE con bytes de formato Word).

Problema encontrado: el servidor COM no soportaba multiples invocaciones sucesivas. Se soluciono abriendo y cerrando Word por cada archivo, con taskkill para limpiar procesos huerfanos.

### Parseo de la transcripcion

Cada archivo .doc tiene una estructura de metadata (lineas con @) seguida del conversacional (lineas con intervenciones de participantes).

Metadata extraida:
- `@ Titulo corto: Cam` -> nombre corto de la grabacion
- `@ Participante: R = Ruben = 18` -> mapeo de codigo de hablante a edad

Utterances extraidos:
Se buscaron todos los patrones `((N texto ))N` donde N es un numero de utterance y texto es el contenido hablado. El hablante se determina por la letra al inicio de la linea:

```
A: ((1 Ese es el tipico santiaguero. ))1
-> utterance 1, hablante A, texto: "Ese es el tipico santiaguero."
```

### Mapeo audio a utterance

Los archivos de audio tienen el formato `{turno}.{short_name}.wav` (ej: `1.Cam.wav`). Se cruza el turno con el numero de utterance y el short name del audio con el titulo corto del .doc.

### Asignacion de edad

Para cada utterance se obtiene la letra del hablante, su edad desde el header, y se calcula el intervalo de 5 anos: `(edad // 5) * 5` a `((edad // 5) * 5) + 4`.

### Unificacion de nombres

- Cocina -> Coc (unificado con el nombre del PDF)

## Output generado

```
DTset_organizado_v2/
  {intervalo_edad}/
    {short_name}_{utterance_N}.wav
    {short_name}_{utterance_N}.txt
```

## Estadisticas finales

| Metrica | Valor |
|---|---|
| Total utterances en transcripciones | 6,316 |
| Pares audio+texto generados | 1,500 |
| Tasa de matching | 23.7% |
| Intervalos de edad creados | 19 |

Distribucion por intervalo de edad:

| Intervalo | Cantidad | Edades representativas |
|---|---|---|
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

## Decisiones tecnicas

Para la extraccion de .doc se descarto olefile porque producia texto con artefactos de codificacion. Se opto por win32com.client (Word COM) que extrae el texto limpio con la codificacion correcta. Se maneja el conflicto COM cerrando y forzando la terminacion de WINWORD.EXE entre archivos.

Solo se incluyen utterances con patron `((N ...))N` (doble parentesis + numero de cierre). Se ignoran utterances que no tengan archivo de audio correspondiente o cuyo hablante no tenga edad registrada en el header.

El texto se limpia eliminando anotaciones entre corchetes `[N ... ]N`, marcas de comentario `<N ... >`, y colapsando multiples espacios en uno.

## Script generador

El script `organizar_dataset_v2.py` contiene todo el pipeline y esta configurado para reutilizarse con el dataset completo:

```python
BASE_DIR = r"D:\Marcos\Prog\TTS"    # Ruta base
DTset_DIR = "DTset"                   # Subcarpeta con Audios/ y Transcripcion/
OUT_DIR = "DTset_organizado_v2"       # Subcarpeta de salida
```

## Notas para el dataset completo

El 76.3% de utterances en las transcripciones no tienen audio asociado en la muestra (6,316 utterances totales, solo 1,500 con audio). Esto es esperable para una muestra. Con el dataset completo se espera que la tasa de matching mejore significativamente.

El script asume la misma estructura de nombres de archivo: `{N}.{short_name}.wav`. Los intervalos de edad se generan automaticamente segun las edades que aparezcan.