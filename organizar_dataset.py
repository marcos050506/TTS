import os, re, shutil, sys, subprocess, csv, unicodedata

import win32com.client

MAX_ATTEMPTS = 3
SHORT_NAME_UNIFY = {"cocina": "Coc"}
SEXO_CACHE_FILENAME = "sexo_participantes.csv"

def normalizar_nombre(nombre):
    nombre = unicodedata.normalize("NFD", nombre)
    nombre = "".join(c for c in nombre if unicodedata.category(c) != "Mn")
    return nombre.split()[0].lower()

def cargar_sexo_cache(transc_dir):
    cache = {}
    path = os.path.join(transc_dir, SEXO_CACHE_FILENAME)
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                cache[row["nombre"]] = row["sexo"]
    return cache

def doc_to_txt_safe(doc_path):
    def _extract():
        word = win32com.client.Dispatch("Word.Application")
        try: word.Visible = False
        except: pass
        try: word.DisplayAlerts = False
        except: pass
        doc = word.Documents.Open(doc_path)
        text = doc.Content.Text
        doc.Close(False)
        return text
    try:
        return _extract()
    except:
        try:
            subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
        except: pass
        import time; time.sleep(2)
        return _extract()

def get_doc_text(doc_path, cache_dir):
    cache_file = os.path.join(cache_dir, os.path.basename(doc_path) + ".txt")
    fresh = os.path.isfile(cache_file) and \
        os.path.getmtime(cache_file) >= os.path.getmtime(doc_path)
    if fresh:
        with open(cache_file, "r", encoding="utf-8") as f:
            return f.read()
    text = doc_to_txt_safe(doc_path)
    os.makedirs(cache_dir, exist_ok=True)
    with open(cache_file, "w", encoding="utf-8") as f:
        f.write(text)
    return text

def parse_transcription(text):
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    short_name = None
    speakers = {}

    for line in lines:
        m = re.match(r"@ T\u00edtulo corto:\s*(\S+)", line)
        if m:
            short_name = m.group(1)
        m = re.match(r"@ Participante:\s*([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1])\s*=\s*(.+?)\s*=\s*(\d+)\s*(?:=\s*([HMD]))?", line)
        if m:
            code = m.group(1)
            age = int(m.group(3))
            name = m.group(2).strip()
            sexo = m.group(4) or "D"
            speakers[code] = {"edad": age, "nombre": name, "sexo": sexo}

    utt_pattern = re.compile(r"\(\((\d+)\s+(.*?)\s*\)\)\1")
    utterances = []

    for line in lines:
        speaker = None
        m_sp = re.match(r"\s*([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1])\s*(?:\([^)]*\))?\s*:", line)
        if m_sp:
            speaker = m_sp.group(1)

        for m in utt_pattern.finditer(line):
            utt_num = int(m.group(1))
            utt_text = m.group(2).strip()
            utt_text = re.sub(r"\[\d+\s*[^\]]*\]\d+", "", utt_text).strip()
            utt_text = re.sub(r"<\d+\s*[^>]*>", "", utt_text).strip()
            utt_text = re.sub(r"\s+", " ", utt_text).strip()
            if utt_text:
                utterances.append({
                    "num": utt_num,
                    "speaker": speaker,
                    "text": utt_text,
                })

    return short_name, speakers, utterances

def scan_audios(audios_dir):
    audio_map = {}
    for fname in os.listdir(audios_dir):
        if not fname.endswith(".wav"):
            continue
        m = re.match(r"(\d+[a-z]?)\.(\w+)\.wav", fname)
        if not m:
            continue
        turn_str = m.group(1)
        short_raw = m.group(2)
        short_name = SHORT_NAME_UNIFY.get(short_raw.lower(), short_raw)
        turn_num = int(re.match(r"(\d+)", turn_str).group(1))
        audio_map[(short_name, turn_num)] = os.path.join(audios_dir, fname)
    return audio_map

def validar_audios_dir(path):
    if not os.path.isdir(path):
        return False, "La ruta no existe o no es una carpeta."
    wavs = [f for f in os.listdir(path) if f.endswith(".wav")]
    if not wavs:
        return False, "La carpeta no contiene archivos .wav."
    validos = sum(1 for f in wavs if re.match(r"\d+[a-z]?\.\w+\.wav", f))
    if validos == 0:
        return False, "Ningun archivo .wav sigue el formato esperado (ej: 1.Cam.wav)."
    return True, f"Ok: {len(wavs)} archivos .wav encontrados ({validos} con formato valido)."

def validar_transc_dir(path):
    if not os.path.isdir(path):
        return False, "La ruta no existe o no es una carpeta."
    docs = [f for f in os.listdir(path) if f.endswith(".doc")]
    if not docs:
        return False, "La carpeta no contiene archivos .doc."
    return True, f"Ok: {len(docs)} archivos .doc encontrados."

def pedir_ruta(mensaje, validador):
    for intento in range(MAX_ATTEMPTS):
        ruta = input(mensaje).strip()
        if not ruta:
            ruta = "."
        valida, msg = validador(ruta)
        if valida:
            print(f"  {msg}")
            return ruta
        restantes = MAX_ATTEMPTS - intento - 1
        print(f"  ERROR: {msg}")
        if restantes > 0:
            print(f"  Intentos restantes: {restantes}")
        else:
            print("  Demasiados intentos fallidos. Saliendo.")
            sys.exit(1)

def organize():
    print("=== ORGANIZADOR DE DATASET TTS ===\n")

    audios_dir = pedir_ruta(
        "Ingrese ruta de la carpeta de AUDIOS (.wav): ",
        validar_audios_dir
    )
    print()

    transc_dir = pedir_ruta(
        "Ingrese ruta de la carpeta de TRANSCRIPCIONES (.doc): ",
        validar_transc_dir
    )
    print()

    out_dir = os.path.join(os.getcwd(), "DTset_organizado")
    os.makedirs(out_dir, exist_ok=True)

    print("Escaneando archivos de audio...", flush=True)
    audio_map = scan_audios(audios_dir)
    print(f"  {len(audio_map)} audios mapeados por (short_name, num)\n", flush=True)

    print("Procesando transcripciones...", flush=True)
    doc_files = sorted([f for f in os.listdir(transc_dir) if f.endswith(".doc")])

    cache_dir = os.path.join(out_dir, "_cache_text")
    docs_data = []
    total_utterances = 0

    for fname in doc_files:
        doc_path = os.path.join(transc_dir, fname)
        print(f"  {fname}...", end=" ", flush=True)

        try:
            text = get_doc_text(doc_path, cache_dir)
        except Exception as e:
            print(f"ERROR: {e}", flush=True)
            continue

        short_name, speakers, utterances = parse_transcription(text)
        if not short_name:
            print("sin titulo corto, se omite", flush=True)
            continue

        docs_data.append((short_name, speakers, utterances))
        total_utterances += len(utterances)
        print(f"{len(utterances)} utterances, {len(speakers)} hablantes", flush=True)

    sexo_cache = cargar_sexo_cache(transc_dir)
    if sexo_cache:
        print(f"  Cache de sexo: {len(sexo_cache)} nombres (fallback para docs sin sexo)", flush=True)
    # resolver sexo del doc; si falta, usar cache por nombre
    for short_name, speakers, utterances in docs_data:
        for info in speakers.values():
            if info.get("sexo", "D") == "D":
                norm = normalizar_nombre(info["nombre"])
                if norm in sexo_cache:
                    info["sexo"] = sexo_cache[norm]

    print("\nOrganizando pares audio+texto...", flush=True)
    matched = 0
    metadata_rows = []

    for short_name, speakers, utterances in docs_data:
        for utt in utterances:
            utt_num = utt["num"]
            speaker = utt["speaker"]
            utt_text = utt["text"]

            audio_path = audio_map.get((short_name, utt_num))
            if not audio_path:
                continue

            info = speakers.get(speaker) if speaker else None
            if not info:
                continue

            age = info["edad"]
            sexo = info.get("sexo", "D")

            age_str = str(age)
            out_folder = os.path.join(out_dir, age_str)
            os.makedirs(out_folder, exist_ok=True)

            new_audio_name = f"{short_name}_{utt_num}.wav"
            shutil.copy2(audio_path, os.path.join(out_folder, new_audio_name))

            txt_path = os.path.join(out_folder, f"{short_name}_{utt_num}.txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(utt_text)

            metadata_rows.append({
                "short_name": short_name,
                "utt_num": utt_num,
                "edad": age,
                "sexo": sexo,
                "texto": utt_text,
                "archivo": f"{age_str}/{new_audio_name}",
            })

            matched += 1

    csv_path = os.path.join(out_dir, "metadata.csv")
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["short_name", "utt_num", "edad", "sexo", "texto", "archivo"])
        writer.writeheader()
        writer.writerows(metadata_rows)
    print(f"  Metadata: {csv_path}", flush=True)

    try:
        subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
    except: pass

    edades_unicas = sorted(set(r["edad"] for r in metadata_rows))
    por_sexo = {}
    for r in metadata_rows:
        por_sexo[r["sexo"]] = por_sexo.get(r["sexo"], 0) + 1
    print(f"\n=== RESUMEN ===", flush=True)
    print(f"  Total utterances encontrados: {total_utterances}", flush=True)
    print(f"  Pares audio+texto generados:  {matched}", flush=True)
    print(f"  Edades unicas representadas:  {len(edades_unicas)}", flush=True)
    print(f"  Por sexo: {por_sexo}", flush=True)
    print(f"  Output: {out_dir}", flush=True)

if __name__ == "__main__":
    organize()
