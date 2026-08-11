import os, re, sys, csv, json, time, shutil, unicodedata
import urllib.request, urllib.parse, urllib.error
import win32com.client

MAX_ATTEMPTS = 3
SEXO_CACHE_FILENAME = "sexo_participantes.csv"
SEXO_API_URL = "https://api.genderize.io/"
BACKUP_DIRNAME = "backup_originales"
PARTICIPANTE_RE = re.compile(
    r"@ Participante:\s*([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1])\s*=\s*(.+?)\s*=\s*(\d+)\s*(?:=\s*([HMD]))?"
)

def normalizar(nombre):
    nombre = unicodedata.normalize("NFD", nombre)
    nombre = "".join(c for c in nombre if unicodedata.category(c) != "Mn")
    return nombre.split()[0].lower()

def cargar_cache(path):
    cache = {}
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                cache[row["nombre"]] = {
                    "sexo": row["sexo"],
                    "probabilidad": row.get("probabilidad", ""),
                    "fuente": row.get("fuente", ""),
                }
    return cache

def guardar_cache(path, cache):
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["nombre", "sexo", "probabilidad", "fuente"])
        writer.writeheader()
        for nombre in sorted(cache):
            info = cache[nombre]
            writer.writerow({
                "nombre": nombre,
                "sexo": info["sexo"],
                "probabilidad": info["probabilidad"],
                "fuente": info["fuente"],
            })

def consultar_api(nombre, max_intentos=5):
    params = {"name": nombre}
    key = os.environ.get("GENDERIZE_API_KEY", "")
    if key:
        params["apikey"] = key
    url = SEXO_API_URL + "?" + urllib.parse.urlencode(params)
    for intento in range(max_intentos):
        try:
            with urllib.request.urlopen(url, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            break
        except urllib.error.HTTPError as e:
            if e.code == 429 or e.code >= 500:
                retry = 30 + intento * 20
                print(f"      limite de API ({e.code}), reintentando en {retry}s...", flush=True)
                time.sleep(retry)
            else:
                return {"sexo": "D", "probabilidad": "", "fuente": "error"}
        except Exception as e:
            print(f"      error de red ({type(e).__name__}), reintentando...", flush=True)
            time.sleep(10)
    else:
        return {"sexo": "D", "probabilidad": "", "fuente": "error"}
    genero = data.get("gender")
    if genero not in ("male", "female"):
        return {"sexo": "D", "probabilidad": "", "fuente": "desconocido"}
    return {"sexo": "H" if genero == "male" else "M",
            "probabilidad": data.get("probability", ""), "fuente": "genderize"}

def resolver_sexos(nombres, cache_path):
    cache = cargar_cache(cache_path)
    pendientes = sorted(
        n for n in nombres
        if n not in cache or cache[n].get("fuente") == "error"
    )
    if pendientes:
        print(f"Consultando sexo de {len(pendientes)} nombres en genderize.io...", flush=True)
        for i, nombre in enumerate(pendientes, 1):
            cache[nombre] = consultar_api(nombre)
            print(f"  [{i}/{len(pendientes)}] {nombre} -> {cache[nombre]['sexo']}", flush=True)
            guardar_cache(cache_path, cache)
            if i < len(pendientes):
                time.sleep(1)
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
        time.sleep(2)
        return _extract()

def verificar_transc_dir(path):
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

def _inyectar(doc_path, sexo_map):
    modificados = []
    word = win32com.client.Dispatch("Word.Application")
    try: word.Visible = False
    except: pass
    try: word.DisplayAlerts = False
    except: pass
    doc = word.Documents.Open(doc_path)
    text = doc.Content.Text

    cambios = []
    for m in PARTICIPANTE_RE.finditer(text):
        norm = normalizar(m.group(2).strip())
        sexo = sexo_map.get(norm, {}).get("sexo", "D")
        actual = m.group(4)
        if actual == sexo:
            continue
        if actual == "D" and sexo == "D":
            continue
        start, end = m.start(), m.end()
        while end > start and text[end - 1] in " \t\r\n":
            end -= 1
        base = text[start:end]
        if actual is not None:
            base = re.sub(r"\s* = [HMD]\s*$", "", base)
        cambios.append((start, end, base + " = " + sexo))
        modificados.append(f"{m.group(1)} {m.group(2).strip()} = {sexo}")

    for start, end, nuevo in reversed(cambios):
        rng = doc.Range(start, end)
        rng.Text = nuevo

    if cambios:
        doc.Save()
    doc.Close(False)
    return modificados

def inyectar_sexo(doc_path, sexo_map):
    try:
        return _inyectar(doc_path, sexo_map)
    except Exception:
        try:
            subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
        except: pass
        time.sleep(2)
        return _inyectar(doc_path, sexo_map)

def main():
    print("=== AGREGAR SEXO A TRANSCRIPCIONES ===\n")

    try:
        subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
    except: pass

    transc_dir = pedir_ruta(
        "Ingrese ruta de la carpeta de TRANSCRIPCIONES (.doc): ",
        verificar_transc_dir
    )
    print()

    backup_dir = os.path.join(transc_dir, BACKUP_DIRNAME)
    doc_files = sorted([f for f in os.listdir(transc_dir) if f.endswith(".doc")])

    args = sys.argv[1:]
    modo_consultar = "--consultar" in args or not args
    modo_inyectar = "--inyectar" in args or not args

    cache_path = os.path.join(transc_dir, SEXO_CACHE_FILENAME)

    if modo_consultar:
        all_names = extraer_nombres(doc_files, transc_dir)
        print(f"  {len(all_names)} nombres unicos", flush=True)
        sexo_map = resolver_sexos(all_names, cache_path)
    else:
        sexo_map = cargar_cache(cache_path)
        print(f"  Usando cache existente: {len(sexo_map)} nombres", flush=True)

    if not modo_inyectar:
        print("  (fase de inyeccion omitida)", flush=True)
        return

    os.makedirs(backup_dir, exist_ok=True)
    print("\nReservando backup de originales...", flush=True)
    count_backup = 0
    for fname in doc_files:
        src = os.path.join(transc_dir, fname)
        dst = os.path.join(backup_dir, fname)
        if not os.path.exists(dst):
            shutil.copy2(src, dst)
            count_backup += 1
    print(f"  {count_backup} archivos en backup_originales/")

    print("\nInyectando sexo en los documentos...", flush=True)
    total_mod = 0
    for fname in doc_files:
        try:
            subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
        except: pass
        time.sleep(1)
        doc_path = os.path.join(transc_dir, fname)
        print(f"  {fname}...", end=" ", flush=True)
        try:
            mods = inyectar_sexo(doc_path, sexo_map)
        except Exception as e:
            print(f"ERROR: {e}", flush=True)
            continue
        print(f"{len(mods)} participantes", flush=True)
        total_mod += len(mods)

    try:
        subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
    except: pass

    print(f"\n=== RESUMEN ===", flush=True)
    print(f"  Participantes con sexo agregado: {total_mod}", flush=True)
    print(f"  Cache de sexos: {cache_path}", flush=True)
    print(f"  Backup: {backup_dir}", flush=True)

def extraer_nombres(doc_files, transc_dir):
    names_path = os.path.join(transc_dir, "nombres_participantes.txt")
    if os.path.isfile(names_path):
        try:
            newest = max(os.path.getmtime(os.path.join(transc_dir, f)) for f in doc_files)
            if os.path.getmtime(names_path) >= newest:
                with open(names_path, "r", encoding="utf-8") as f:
                    return {l.strip() for l in f if l.strip()}
        except Exception:
            pass
    print("Extrayendo nombres de participantes...", flush=True)
    all_names = set()
    for fname in doc_files:
        doc_path = os.path.join(transc_dir, fname)
        try:
            text = doc_to_txt_safe(doc_path)
        except Exception:
            continue
        for line in text.split("\r"):
            m = PARTICIPANTE_RE.match(line.strip())
            if m:
                all_names.add(normalizar(m.group(2).strip()))
    with open(names_path, "w", encoding="utf-8") as f:
        f.write("\n".join(sorted(all_names)))
    return all_names

if __name__ == "__main__":
    main()