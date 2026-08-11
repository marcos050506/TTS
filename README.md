# Dataset TTS - Transformacion de corpus conversacional

Transformacion de un corpus de conversaciones grabadas en un dataset estructurado para entrenamiento de TTS. Cada utterance individual con audio asociado se organiza por la edad exacta del hablante que lo emitio.

## Instalacion

```bash
pip install pywin32>=306
```

## Uso

```bash
python organizar_dataset.py
```

El script pedira interactivamente:
1. Ruta de la carpeta con los audios (.wav)
2. Ruta de la carpeta con las transcripciones (.doc)

Valida que las rutas existan, contengan archivos del formato esperado, y da 3 intentos antes de salir. El output se crea en `DTset_organizado/` en el mismo directorio donde se ejecuta el script.

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

Para cada utterance se obtiene la letra del hablante y su edad desde el header. Los audios se organizan en carpetas por edad exacta.

### Unificacion de nombres

- Cocina -> Coc (unificado con el nombre del PDF)

### Determinacion de sexo

El sexo puede venir de dos fuentes, en este orden:

1. Del propio doc, si la linea `@ Participante` ya trae el valor al final: `@ Participante: R = Ruben = 18 = H` (formato `H` hombre, `M` mujer, `D` desconocido).
2. De un archivo `sexo_participantes.csv` colocado junto a las transcripciones, con columnas `nombre,sexo,probabilidad,fuente`. El nombre se normaliza quitando acentos y quedandose con el primer token (ej: `Eva Rosa` -> `Eva`). Se usa si el doc no trae el sexo.

Este cache se genera con la API genderize.io. Hay un helper opcional `agregar_sexo_docs.py` que, dada una carpeta de transcripciones, consulta genderize.io para cada participante y reescribe las lineas `@ Participante` embebiendo el sexo dentro del propio .doc. Asi el sexo queda fisicamente en la transcripcion y el script principal lo lee directa.

Los nombres que la API no reconoce o ante fallos de red se marcan como `D` (desconocido).

## Output generado

```
DTset_organizado/
  metadata.csv
  {edad}/
    {short_name}_{utterance_N}.wav
    {short_name}_{utterance_N}.txt
```

El archivo `metadata.csv` contiene todas las entradas con columnas: `short_name`, `utt_num`, `edad`, `sexo`, `texto`, `archivo`. El path en `archivo` es relativo a la raiz del output, permitiendo que cualquier dataloader acceda directamente sin depender de la estructura de carpetas. El valor de `sexo` es `H` (hombre), `M` (mujer) o `D` (desconocido).

## Estadisticas finales

| Metrica | Valor |
|---|---|
| Total utterances en transcripciones | 6,316 |
| Pares audio+texto generados | 1,501 |
| Tasa de matching | 23.8% |
| Edades unicas representadas | 53 |
| Hombres | 590 |
| Mujeres | 911 |
| Sexo desconocido | 0 |

Distribucion por edad exacta:

| Edad | Cantidad | Edad | Cantidad | Edad | Cantidad |
|---|---|---|---|---|---|
| 18 | 110 | 42 | 16 | 100 | 7 |
| 20 | 97 | 43 | 3 | 102 | 3 |
| 22 | 129 | 44 | 16 | 103 | 115 |
| 23 | 73 | 45 | 1 | 104 | 104 |
| 24 | 13 | 53 | 2 | 106 | 39 |
| 25 | 1 | 54 | 7 | 107 | 3 |
| 26 | 206 | 58 | 1 | 109 | 16 |
| 27 | 50 | 59 | 15 | 110 | 23 |
| 28 | 39 | 73 | 6 | 111 | 33 |
| 29 | 31 | 74 | 21 | 112 | 2 |
| 30 | 17 | 75 | 4 | 113 | 2 |
| 31 | 16 | 76 | 1 | 120 | 6 |
| 32 | 26 | 77 | 30 | 121 | 8 |
| 33 | 24 | 78 | 15 | | |
| 34 | 88 | 79 | 3 | | |
| 35 | 3 | 83 | 1 | | |
| 39 | 2 | 84 | 1 | | |
| 40 | 8 | 85 | 14 | | |
| 41 | 7 | 86 | 32 | | |
| | | 93 | 4 | | |
| | | 99 | 7 | | |

## Decisiones tecnicas

Para la extraccion de .doc se descarto olefile porque producia texto con artefactos de codificacion. Se opto por win32com.client (Word COM) que extrae el texto limpio con la codificacion correcta. Se maneja el conflicto COM cerrando y forzando la terminacion de WINWORD.EXE entre archivos.

Solo se incluyen utterances con patron `((N ...))N` (doble parentesis + numero de cierre). Se ignoran utterances que no tengan archivo de audio correspondiente o cuyo hablante no tenga edad registrada en el header.

El texto se limpia eliminando anotaciones entre corchetes `[N ... ]N`, marcas de comentario `<N ... >`, y colapsando multiples espacios en uno.

## Script generador

El script `organizar_dataset.py` contiene todo el pipeline. No tiene rutas fijas: al ejecutarse pide (1) la carpeta de audios y (2) la carpeta de transcripciones, y crea `DTset_organizado/` en el directorio actual. Incluye validacion de formato y hasta 3 reintentos por ruta erronea.

Para no re-abrir Word en cada ejecucion, el texto extraido de cada transcripcion se cachea en `DTset_organizado/_cache_text/`. El cache se regenera solo si el .doc cambia.

## Helper opcional para sexo

`agregar_sexo_docs.py` es opcional y no forma parte del pipeline principal. Su objetivo es reescribir las transcripciones para embeker el sexo en la linea `@ Participante` (consultando la API genderize.io con un timeout y reintentos). Uso:

```
python agregar_sexo_docs.py --consultar <ruta_transcripciones>   # arma el CSV
python agregar_sexo_docs.py --inyectar <ruta_transcripciones>    # embebe sexo en los .doc
```

Guarda copias originales en `backup_originales/` y no vuelve a consultar nombres ya resueltos. Requiere la variable de entorno `GENDERIZE_API_KEY` para evitar el rate-limit de la API.

## Notas para el dataset completo

El 76.3% de utterances en las transcripciones no tienen audio asociado en la muestra (6,316 utterances totales, solo 1,500 con audio). Esto es esperable para una muestra. Con el dataset completo se espera que la tasa de matching mejore significativamente.

El script asume la misma estructura de nombres de archivo: `{N}.{short_name}.wav`. Las carpetas de edad se generan automaticamente segun las edades que aparezcan en los headers de las transcripciones.