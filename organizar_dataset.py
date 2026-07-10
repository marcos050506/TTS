import os, re, shutil, json, sys, subprocess
from collections import defaultdict

import pdfplumber
import win32com.client

BASE_DIR = r"D:\Marcos\Prog\TTS"
DTset_DIR = os.path.join(BASE_DIR, "DTset")
AUDIOS_DIR = os.path.join(DTset_DIR, "Audios")
TRANS_DIR = os.path.join(DTset_DIR, "Transcripcion")
PDF_PATH = os.path.join(BASE_DIR, "Identificac-gral-grabaciones.pdf")
OUT_DIR = os.path.join(BASE_DIR, "DTset_organizado")

SHORT_NAME_UNIFY = {"cocina": "Coc"}

# ─── 1. Parse PDF tables ─────────────────────────────────────────────────────

def grupo_from_code(grab_code):
    if grab_code.startswith("GB"):
        return "primer_grupo"
    if grab_code.startswith("GC"):
        return "segundo_grupo"
    if grab_code.startswith("GD"):
        return "tercer_grupo"
    return "desconocido"

def parse_pdf():
    """Parse PDF tables to get mapping: grab_code -> short_name, transc_code, grupo, participantes"""
    mapping = {}

    with pdfplumber.open(PDF_PATH) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    cells = [c.strip() if c else "" for c in row]
                    col0 = cells[0]

                    if col0 in ("Grab", "Transc", "Participantes", "Grabaci\u00f3n",
                                "Duraci\u00f3n", "TOTAL GENERAL") or "OJO" in col0:
                        continue

                    col1 = cells[1] if len(cells) > 1 else ""

                    # Duration tables: skip
                    if re.match(r"\d{2}:\d{2}:\d{2}", col1):
                        continue

                    # Main mapping table: col0 = "GB005a\\nViv-A", col1 = "TB005a"
                    parts = col0.split("\n")
                    if len(parts) >= 2:
                        grab_code = parts[0].strip()
                        short_name = parts[1].strip()
                        transc_code = None
                        if re.match(r"T[BCD]\d{3}[a-z]?", col1):
                            transc_code = col1

                        if re.match(r"G[BCD]\d{3}[a-z]?", grab_code):
                            mapping[short_name] = {
                                "grab_code": grab_code,
                                "short_name": short_name,
                                "transc_code": transc_code,
                                "grupo": grupo_from_code(grab_code),
                                "participantes": [],
                            }

    return mapping

def extract_participants(mapping):
    """Extract participants from PDF table cells (column 2)"""
    with pdfplumber.open(PDF_PATH) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    cells = [c.strip() if c else "" for c in row]
                    col0 = cells[0]
                    col2 = cells[2] if len(cells) > 2 else ""

                    if not col2 or col2 in ("Participantes",):
                        continue

                    # Find which short name this row belongs to
                    parts = col0.split("\n")
                    if len(parts) >= 2:
                        short_name = parts[1].strip()
                        if short_name in mapping:
                            # Extract participants from col2
                            for line in col2.split("\n"):
                                line = line.strip().strip("~").strip()
                                m = re.match(r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1])\s*=\s*(.+?)\s*=\s*(\d+|[a-j])", line)
                                if m:
                                    code = m.group(1)
                                    name = m.group(2).strip()
                                    age_str = m.group(3)
                                    try:
                                        age = int(age_str)
                                    except ValueError:
                                        age = age_str
                                    mapping[short_name]["participantes"].append({
                                        "code": code, "name": name, "age": age
                                    })
    return mapping

# ─── 2. Scan audio files ─────────────────────────────────────────────────────

def scan_audios():
    audio_groups = defaultdict(list)
    for fname in os.listdir(AUDIOS_DIR):
        if not fname.endswith(".wav"):
            continue
        m = re.match(r"(\d+[a-z]?)\.(\w+)\.wav", fname)
        if not m:
            continue
        turn = m.group(1)
        short_raw = m.group(2)
        short_name = SHORT_NAME_UNIFY.get(short_raw.lower(), short_raw)
        audio_groups[short_name].append({
            "turn": turn,
            "filename": fname,
            "path": os.path.join(AUDIOS_DIR, fname),
        })
    for sn in audio_groups:
        audio_groups[sn].sort(key=lambda x: int(re.match(r"(\d+)", x["turn"]).group(1)))
    return audio_groups

# ─── 3. Convert .doc -> .txt ──────────────────────────────────────────────────

def doc_to_txt_safe(doc_path):
    """Open .doc with Word and extract text. Handles COM conflicts."""
    def _extract():
        word = win32com.client.Dispatch("Word.Application")
        try:
            word.Visible = False
        except:
            pass
        try:
            word.DisplayAlerts = False
        except:
            pass
        doc = word.Documents.Open(doc_path)
        text = doc.Content.Text
        doc.Close(False)
        return text
    try:
        return _extract()
    except:
        # If that fails, kill Word processes and retry once
        try:
            subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"],
                           capture_output=True, timeout=5)
        except:
            pass
        import time
        time.sleep(2)
        return _extract()

def convert_all_docs(transc_files):
    texts = {}
    for code, path in transc_files.items():
        try:
            text = doc_to_txt_safe(path)
            texts[code] = text
            print(f"  {code}: {len(text)} chars", flush=True)
        except Exception as e:
            print(f"  {code}: ERROR - {e}", flush=True)
    try:
        subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"],
                       capture_output=True, timeout=5)
    except:
        pass
    return texts

# ─── 4. Sex inference ────────────────────────────────────────────────────────

MASCULINE_NAMES = {
    "ruben", "agustin", "ernesto", "abel", "roberto", "ivon", "orlando",
    "damian", "pedro juan", "sergio", "carlos", "ramon", "martinez",
    "gomez", "romilio", "euripides", "enmanuel", "anibal", "salomon",
    "perico", "otto", "idalberto", "tony", "eduardo", "ruslan",
    "saldivar", "andres", "cesar", "helio omar", "tulio", "raul",
    "tomas", "pedro", "amilcar", "froilan", "nelson", "hugo", "antonio",
    "gorriaran", "fonseca",
}

FEMININE_NAMES = {
    "ana belen", "ines", "jimena", "ester", "celia", "graciela",
    "aimee", "berta", "alina", "fatima", "helena", "sucel",
    "concepcion", "adelaida", "marlen", "mirta", "domitila",
    "carmen", "dalia", "amalia", "teresa", "jazmin", "idania",
    "emiliana", "rosana", "josefina", "eva rosa", "maritza",
    "regla", "isabel", "sandra", "goar", "rebeca", "liset",
    "juliana", "cacha", "silvia", "mara", "cecilia", "leticia",
    "dania", "odalis", "raquel", "laura", "magali", "liliana",
    "mayelin", "obelia", "caridad", "rosalina", "eduviges",
    "clara", "lisbet", "rosaida", "elena", "clodomira", "liuba",
    "carina", "puchunga", "calixta", "maria", "vilma", "patricia",
    "carmela", "keila", "manuela", "fernanda", "angela", "garcia",
}

def infer_sex(name):
    name_lower = name.lower().strip()
    if name_lower in MASCULINE_NAMES:
        return "M"
    if name_lower in FEMININE_NAMES:
        return "F"
    return "?"

# ─── 5. Organize ──────────────────────────────────────────────────────────────

def organize_dataset():
    print("Step 1: Parsing PDF tables...", flush=True)
    pdf_data = parse_pdf()
    pdf_data = extract_participants(pdf_data)
    print(f"  Found {len(pdf_data)} short names in PDF", flush=True)

    print("Step 2: Scanning audio files...", flush=True)
    audio_groups = scan_audios()
    print(f"  Found {len(audio_groups)} unique short names in audio", flush=True)
    total_audio = sum(len(v) for v in audio_groups.values())
    print(f"  Total audio files: {total_audio}", flush=True)

    print("Step 3: Scanning transcription files...", flush=True)
    transc_files = {}
    for fname in os.listdir(TRANS_DIR):
        m = re.match(r"(T[BCD]\d{3}[a-z]?)-s\.doc", fname)
        if m:
            transc_files[m.group(1)] = os.path.join(TRANS_DIR, fname)
    print(f"  Found {len(transc_files)} transcription files", flush=True)

    # Determine which transc codes are needed (those matching audio short names)
    needed_transc = set()
    for sn in audio_groups:
        info = pdf_data.get(sn)
        if info and info.get("transc_code"):
            needed_transc.add(info["transc_code"])

    transc_to_convert = {k: v for k, v in transc_files.items() if k in needed_transc}
    print(f"  Will convert {len(transc_to_convert)} files (matching audio groups)", flush=True)

    print("Step 4: Converting .doc -> .txt...", flush=True)
    transc_texts = convert_all_docs(transc_to_convert)

    # Build lookup
    transc_to_info = {}
    for sn, info in pdf_data.items():
        if info.get("transc_code"):
            transc_to_info[info["transc_code"]] = info

    print("Step 5: Organizing folders...", flush=True)

    out_con_transc = os.path.join(OUT_DIR, "con_transcripcion")
    out_solo_audio = os.path.join(OUT_DIR, "solo_audios")
    out_solo_transc = os.path.join(OUT_DIR, "solo_transcripciones")

    con_transc_count = 0
    solo_audio_count = 0
    solo_transc_count = 0

    # Group by grupo for all short names
    sn_to_grupo = {}
    for sn, info in pdf_data.items():
        sn_to_grupo[sn] = info.get("grupo", "desconocido")

    # --- con_transcripcion ---
    for short_name in sorted(audio_groups):
        info = pdf_data.get(short_name)
        if not info:
            continue
        transc_code = info.get("transc_code")
        if not transc_code or transc_code not in transc_texts:
            continue

        grupo = info.get("grupo", "desconocido")
        folder = os.path.join(out_con_transc, grupo, short_name)
        os.makedirs(folder, exist_ok=True)

        for audio in audio_groups[short_name]:
            shutil.copy2(audio["path"], os.path.join(folder, audio["filename"]))

        txt_path = os.path.join(folder, f"{transc_code}.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(transc_texts[transc_code])

        meta = {
            "grab_code": info["grab_code"],
            "transc_code": transc_code,
            "short_name": short_name,
            "grupo": grupo,
            "total_audio_files": len(audio_groups[short_name]),
            "participantes": [
                {"code": p["code"], "name": p["name"], "age": p["age"], "sex": infer_sex(p["name"])}
                for p in info["participantes"]
            ],
        }
        with open(os.path.join(folder, "metadata.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        con_transc_count += 1
        print(f"  [con][{grupo}] {short_name}: {len(audio_groups[short_name])} audios + {transc_code}.txt", flush=True)

    # --- solo_audios ---
    for short_name in sorted(audio_groups):
        info = pdf_data.get(short_name)
        if info and info.get("transc_code") and info["transc_code"] in transc_texts:
            continue

        grupo = sn_to_grupo.get(short_name, "desconocido")
        folder = os.path.join(out_solo_audio, grupo, short_name)
        os.makedirs(folder, exist_ok=True)

        for audio in audio_groups[short_name]:
            shutil.copy2(audio["path"], os.path.join(folder, audio["filename"]))

        if info:
            meta = {
                "grab_code": info["grab_code"],
                "transc_code": info.get("transc_code"),
                "short_name": short_name,
                "grupo": grupo,
                "total_audio_files": len(audio_groups[short_name]),
                "participantes": [
                    {"code": p["code"], "name": p["name"], "age": p["age"], "sex": infer_sex(p["name"])}
                    for p in info["participantes"]
                ],
            }
            with open(os.path.join(folder, "metadata.json"), "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False, indent=2)

        solo_audio_count += 1
        print(f"  [audio][{grupo}] {short_name}: {len(audio_groups[short_name])} audios", flush=True)

    # --- solo_transcripciones ---
    for transc_code in sorted(transc_texts):
        info = transc_to_info.get(transc_code)
        sn = info["short_name"] if info and info.get("short_name") else transc_code

        if sn in audio_groups:
            continue

        grupo = sn_to_grupo.get(sn, "desconocido") if info else "desconocido"
        folder = os.path.join(out_solo_transc, grupo, sn)
        os.makedirs(folder, exist_ok=True)

        with open(os.path.join(folder, f"{transc_code}.txt"), "w", encoding="utf-8") as f:
            f.write(transc_texts[transc_code])

        solo_transc_count += 1
        print(f"  [transc][{grupo}] {sn} ({transc_code})", flush=True)

    print(f"\n=== RESUMEN ===", flush=True)
    print(f"  Con audio + transcripcion: {con_transc_count}", flush=True)
    print(f"  Solo audio:               {solo_audio_count}", flush=True)
    print(f"  Solo transcripcion:       {solo_transc_count}", flush=True)
    print(f"\nOutput: {OUT_DIR}", flush=True)

if __name__ == "__main__":
    organize_dataset()
