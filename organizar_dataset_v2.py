import os, re, shutil, json, sys, subprocess
from collections import defaultdict

import win32com.client

BASE_DIR = r"D:\Marcos\Prog\TTS"
DTset_DIR = os.path.join(BASE_DIR, "DTset")
AUDIOS_DIR = os.path.join(DTset_DIR, "Audios")
TRANS_DIR = os.path.join(DTset_DIR, "Transcripcion")
OUT_DIR = os.path.join(BASE_DIR, "DTset_organizado_v2")

SHORT_NAME_UNIFY = {"cocina": "Coc"}

def age_interval(age):
    low = (age // 5) * 5
    high = low + 4
    return f"{low}-{high}"

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

def parse_transcription(text):
    """Parse a transcription .doc text.
    Returns:
        short_name: str
        speaker_ages: dict {letter: age}
        utterances: list of (utt_number, speaker_letter, text)
    """
    lines = text.split("\r")
    short_name = None
    speaker_ages = {}

    # First pass: header metadata
    for line in lines:
        m = re.match(r"@ T\u00edtulo corto:\s*(\S+)", line)
        if m:
            short_name = m.group(1)
        m = re.match(r"@ Participante:\s*([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1])\s*=\s*(.+?)\s*=\s*(\d+)", line)
        if m:
            code = m.group(1)
            age = int(m.group(3))
            speaker_ages[code] = age

    utt_pattern = re.compile(r"\(\((\d+)\s+(.*?)\s*\)\)\1")

    utterances = []

    for line in lines:
        # Detect speaker at line start: "A: ..." or "A (nombre): ..."
        speaker = None
        m_sp = re.match(r"\s*([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1])\s*(?:\([^)]*\))?\s*:", line)
        if m_sp:
            speaker = m_sp.group(1)

        # Find all ((N ...))N in this line
        for m in utt_pattern.finditer(line):
            utt_num = int(m.group(1))
            utt_text = m.group(2).strip()
            # Clean up interlinear annotations [N ... ]N
            utt_text = re.sub(r"\[\d+\s*[^\]]*\]\d+", "", utt_text).strip()
            # Clean up <N ... >N markers
            utt_text = re.sub(r"<\d+\s*[^>]*>", "", utt_text).strip()
            utt_text = re.sub(r"\s+", " ", utt_text).strip()
            if utt_text:
                utterances.append({
                    "num": utt_num,
                    "speaker": speaker,
                    "text": utt_text,
                })

    return short_name, speaker_ages, utterances

def scan_audios():
    audio_map = {}  # (short_name, utt_num) -> filepath
    for fname in os.listdir(AUDIOS_DIR):
        if not fname.endswith(".wav"):
            continue
        m = re.match(r"(\d+[a-z]?)\.(\w+)\.wav", fname)
        if not m:
            continue
        turn_str = m.group(1)
        short_raw = m.group(2)
        short_name = SHORT_NAME_UNIFY.get(short_raw.lower(), short_raw)
        turn_num = int(re.match(r"(\d+)", turn_str).group(1))
        audio_map[(short_name, turn_num)] = os.path.join(AUDIOS_DIR, fname)
    return audio_map

def organize():
    print("Scanning audio files...", flush=True)
    audio_map = scan_audios()
    print(f"  Found {len(audio_map)} audio files mapped by (short_name, num)", flush=True)

    print("Processing transcription files...", flush=True)
    doc_files = [f for f in os.listdir(TRANS_DIR) if f.endswith(".doc")]

    total_utterances = 0
    matched = 0

    for fname in sorted(doc_files):
        doc_path = os.path.join(TRANS_DIR, fname)
        print(f"  {fname}...", end=" ", flush=True)

        try:
            text = doc_to_txt_safe(doc_path)
        except Exception as e:
            print(f"ERROR extracting: {e}", flush=True)
            continue

        short_name, speaker_ages, utterances = parse_transcription(text)
        if not short_name:
            print("no short name found, skipping", flush=True)
            continue

        total_utterances += len(utterances)
        local_matched = 0

        for utt in utterances:
            utt_num = utt["num"]
            speaker = utt["speaker"]
            utt_text = utt["text"]

            # Look up audio
            audio_path = audio_map.get((short_name, utt_num))
            if not audio_path:
                continue

            # Determine age
            age = None
            if speaker and speaker in speaker_ages:
                age = speaker_ages[speaker]

            if age is None:
                continue

            interval = age_interval(age)
            out_folder = os.path.join(OUT_DIR, interval)
            os.makedirs(out_folder, exist_ok=True)

            # Copy audio with new name
            new_audio_name = f"{short_name}_{utt_num}.wav"
            shutil.copy2(audio_path, os.path.join(out_folder, new_audio_name))

            # Save text
            with open(os.path.join(out_folder, f"{short_name}_{utt_num}.txt"), "w", encoding="utf-8") as f:
                f.write(utt_text)

            local_matched += 1

        matched += local_matched
        print(f"{len(utterances)} utterances, {local_matched} matched", flush=True)

    print(f"\n=== RESUMEN ===", flush=True)
    print(f"  Total utterances found: {total_utterances}", flush=True)
    print(f"  Matched audio+text:     {matched}", flush=True)
    print(f"  Output: {OUT_DIR}", flush=True)

    # Kill lingering Word
    try:
        subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
    except: pass

if __name__ == "__main__":
    organize()