import os
import re
import sys
import json
import time
import win32com.client as win32
from google import genai
from google.genai import types
from dotenv import load_dotenv
import llm_utils

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

# --- CONFIGURATION ---
AUDIO_PATH = r"F:\ירושלים\2026\05- מאי\ירושלים\מחוזי\3.5.26\OA-פלדמן - 67104-01-20 מיום- 3.5.26\מקור\260503_2064.mp3"
PDF_PATH = r"F:\ירושלים\2026\05- מאי\ירושלים\מחוזי\3.5.26\OA-פלדמן - 67104-01-20 מיום- 3.5.26\20260504112931010.pdf"
VOICE_LIBRARY_PATH = r"D:\yoel\projects\auto\voice_library"
CASE_NUM = "67104-01-20"
WITNESS_NAME = "שרה דווידי"   # Leave empty to pass via --witness flag; set EXTRACT_FROM_PDF=False to skip PDF entirely
EXTRACT_FROM_PDF = False  # Set True to auto-extract witness from PDF each run
CHUNK_MINUTES = 10
MAX_MINUTES = 10
MODEL_NAME = "gemini-3.1-flash-lite-preview"
HEADER_CACHE_PATH = r"D:\yoel\projects\auto\speaker_headers.json"
WORD_RUBRIC_PATH = r"D:\yoel\projects\auto\voice_library\67104-01-20פרידמן20.4.doc"
TEMPLATE_DOC = r"d:\yoel\projects\auto\-בימש מחוזי תל אביב 67104-01-20-בנימין נתניהו.23.2.26.doc"

# Voice samples for speaker identification
VOICE_SAMPLES = {
    "יהודית תירוש 1.mp3":      'עו"ד יהודית תירוש',
    "עוד יהודית תירוש 2.mp3":  'עו"ד יהודית תירוש',
    "עוד עמית חדד.mp3":        'עו"ד עמית חדד',
    "שופטת ר פרידמן פלדמן.mp3": "כב' השופטת ר' פרידמן-פלדמן - אב\"ד",
    "שופט משה בר עם 1.mp3":    "כב' השופט משה בר עם",
    "שופט משה בר עם 2.mp3":    "כב' השופט משה בר עם",
    "עוד ניצן וולקן.mp3":      'עו"ד ניצן וולקן',
    "זק חן 1.mp3":             'עו"ד ז\'ק חן',
    "זק חן 2.mp3":             'עו"ד ז\'ק חן',
}

HEBREW_MONTHS = {
    1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל",
    5: "מאי", 6: "יוני", 7: "יולי", 8: "אוגוסט",
    9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר",
}

# Permanent protocol header — dynamic fields: {hearing_date}, {witness_name}
PROTOCOL_HEADER_LINES = [
    "בבית משפט המחוזי בירושלים",
    '\t\t\t\tת"פ {case_num}',
    "ישיבת יום {hearing_date}",
    "הדיון מתקיים בבית המשפט המחוזי בירושלים",
    "",
    'בפני:\tהרכב\t\t\tכבוד השופטת רבקה פרידמן-פלדמן – אב"ד',
    "\t\t\t\t\tכבוד השופט משה בר עם",
    "\t\t\t\t\tכבוד השופט עודד שחם",
    "",
    "",
    "בעניין:\t\t\t\tמדינת ישראל",
    "",
    "\t\t\t\t\t\t\t\tהמאשימה",
    "",
    "נ ג ד",
    "",
    "",
    "הנאשמים:\t\t\t1. בנימין נתניהו",
    "\t\t\t\t\t2. שאול אלוביץ' ת\"ז 042089367",
    "\t\t\t\t\t3. איריס אלוביץ' ת\"ז 056381353",
    "\t\t\t\t\t4. ארנון מוזס, ת\"ז 051659159",
    "",
    "\t\t\t\t\t\t\t\tהנאשמים",
    "",
    "",
    "נוכחים:",
    'מטעם המאשימה:\tעו"ד יהודית תירוש, עו"ד ניצן וולקן, עו"ד ספיר יזרלביץ',
    'מטעם הנאשמים:\tב"כ נאשם 1 – עו"ד עמית חדד, עו"ד נועה מילשטיין, עו"ד ישראל וולנרמן',
    '\t\t\t\tב"כ נאשמים 2 ו-3 – עו"ד ז\'ק חן',
    "",
    "",
    "העד:\t{witness_name}",
    "",
    "",
]


def derive_hearing_date(audio_path):
    """Derive hearing date in Hebrew from audio filename (YYMMDD prefix)."""
    m = re.match(r'^(\d{2})(\d{2})(\d{2})', os.path.basename(audio_path))
    if m:
        year = int("20" + m.group(1))
        month = int(m.group(2))
        day = int(m.group(3))
        month_name = HEBREW_MONTHS.get(month, str(month))
        return f"{day} ב{month_name} {year}"
    return ""


def extract_speakers_from_word_doc(doc_path, client):
    """Extract speaker roster from an existing Word protocol document."""
    import pythoncom
    llm_utils.log(f"[WORD] Reading speaker roster from {os.path.basename(doc_path)}...")
    word = None
    try:
        pythoncom.CoInitialize()
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(doc_path)
        header_text = doc.Content.Text[:2000]
        doc.Close(False)

        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=[f"""This is the header of an Israeli court protocol document. Extract all speaker names exactly as written.
Preserve names exactly as written including all Hebrew titles, abbreviations, and punctuation (e.g., כב' השופטת ר' פרידמן-פלדמן - אב"ד). Do NOT simplify or translate.
Return ONLY valid JSON with keys: judges, prosecution_lawyers, defense_lawyers, defendants (arrays of full names with titles).

Document text:
{header_text}"""],
        )
        clean = re.sub(r"```(?:json)?\s*|\s*```", "", response.text).strip()
        result = json.loads(clean)
        llm_utils.log(f"[WORD] Extracted: {result}")
        return result
    except Exception as e:
        llm_utils.log(f"[WORD] Error reading Word doc: {e}")
        return None
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass


def save_header_to_cache(case_num, speakers_data):
    """Save extracted speaker header to cache file."""
    cache = {}
    if os.path.exists(HEADER_CACHE_PATH):
        try:
            with open(HEADER_CACHE_PATH, "r", encoding="utf-8") as f:
                cache = json.load(f)
        except:
            pass

    cache[case_num] = speakers_data
    with open(HEADER_CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    llm_utils.log(f"[CACHE] Saved header for case {case_num}")


def load_header_from_cache(case_num):
    """Load extracted speaker header from cache file."""
    if not os.path.exists(HEADER_CACHE_PATH):
        return None
    try:
        with open(HEADER_CACHE_PATH, "r", encoding="utf-8") as f:
            cache = json.load(f)
        val = cache.get(case_num)
        if isinstance(val, list) and len(val) > 0:
            val = val[0]
        return val
    except:
        return None


def extract_speakers_from_pdf(pdf_path, client):
    """Extract speaker names and roles from the first page of the PDF using LLM."""
    llm_utils.log("[PDF] Extracting speakers from PDF...")
    pdf_file = None
    try:
        pdf_file = client.files.upload(
            file=pdf_path,
            config=types.UploadFileConfig(mime_type="application/pdf"),
        )
        while pdf_file.state.name == "PROCESSING":
            time.sleep(2)
            pdf_file = client.files.get(name=pdf_file.name)
        if pdf_file.state.name != "ACTIVE":
            return None
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=["""Extract all speaker names and formal titles from this court hearing document.
Return ONLY valid JSON with keys: judges, prosecution_lawyers, defense_lawyers, defendants (arrays of full names with titles).""",
                pdf_file],
        )
        clean = re.sub(r"```(?:json)?\s*|\s*```", "", response.text).strip()
        result = json.loads(clean)
        if isinstance(result, list) and len(result) > 0:
            result = result[0]
        llm_utils.log(f"[PDF] Extracted: {result}")
        return result
    except Exception as e:
        llm_utils.log(f"[PDF] Error: {e}")
        return None
    finally:
        if pdf_file:
            try:
                client.files.delete(name=pdf_file.name)
            except Exception:
                pass


def extract_witness_from_pdf(pdf_path, client):
    """Extract the name of the witness currently on the stand from the hearing's PDF."""
    llm_utils.log("[PDF] Extracting witness name from PDF...")
    pdf_file = None
    try:
        pdf_file = client.files.upload(
            file=pdf_path,
            config=types.UploadFileConfig(mime_type="application/pdf"),
        )
        while pdf_file.state.name == "PROCESSING":
            time.sleep(2)
            pdf_file = client.files.get(name=pdf_file.name)
        if pdf_file.state.name != "ACTIVE":
            return None
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=["""This is an Israeli court hearing document.
Find the name of the witness (עד/עדה) who is scheduled to testify or is currently on the stand.
Return ONLY valid JSON with a single key "witness_name" containing their full name in Hebrew, or null if not found.""",
                pdf_file],
        )
        clean = re.sub(r"```(?:json)?\s*|\s*```", "", response.text).strip()
        result = json.loads(clean)
        name = result.get("witness_name")
        if name:
            llm_utils.log(f"[PDF] Witness extracted: {name}")
        return name
    except Exception as e:
        llm_utils.log(f"[PDF] Witness extraction error: {e}")
        return None
    finally:
        if pdf_file:
            try:
                client.files.delete(name=pdf_file.name)
            except Exception:
                pass


def normalize_to_protocol_labels(text, witness_name, speakers_data):
    """Convert full-name LLM output to protocol labels: witness→ת:, Q&A lawyer→ש:"""
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    judges = set(speakers_data.get('judges', []) if speakers_data else [])

    # Step 1: replace witness name with ת:
    normalized = []
    for ln in lines:
        if '\t' in ln:
            label, content = ln.split('\t', 1)
            label = label.strip()
            if witness_name and witness_name in label:
                normalized.append(f'ת:\t{content.strip()}')
            else:
                normalized.append(ln)
        else:
            normalized.append(ln)

    # Step 2: convert lawyer Q&A turns → ש:
    # A lawyer turn is Q&A when the immediately prev or next line is ת:
    result = []
    for i, ln in enumerate(normalized):
        if '\t' not in ln:
            result.append(ln)
            continue
        label = ln.split('\t')[0].strip()
        content = ln.split('\t', 1)[1]
        if label in ('ת:', 'ש:'):
            result.append(ln)
            continue
        # Judges always keep their full name
        if any(j in label for j in judges):
            result.append(ln)
            continue
        # Check adjacency to ת:
        prev_label = normalized[i-1].split('\t')[0].strip() if i > 0 and '\t' in normalized[i-1] else ''
        next_label = normalized[i+1].split('\t')[0].strip() if i < len(normalized)-1 and '\t' in normalized[i+1] else ''
        if prev_label == 'ת:' or next_label == 'ת:':
            result.append(f'ש:\t{content}')
        else:
            result.append(ln)

    return '\n'.join(result)


def transcribe_vertex_final(audio_path=None, case_num=None, pdf_path=None, witness_name=None, extract_header=False):
    if not audio_path:
        audio_path = AUDIO_PATH
    if not case_num:
        case_num = CASE_NUM
    if not pdf_path:
        pdf_path = PDF_PATH
    # witness_name: use arg → config constant → auto-extract from PDF
    if not witness_name:
        witness_name = WITNESS_NAME

    api_key = os.getenv("GOOGLE_API_KEY_2")
    if not api_key:
        llm_utils.log("[ERROR] GOOGLE_API_KEY_2 not found in .env")
        return

    client = genai.Client(api_key=api_key)

    # Auto-extract witness from PDF only when explicitly enabled
    _generic_witness_terms = {"עד", "עדה", "עד/ה", "עד/עדה", "העד", "העדה"}
    if EXTRACT_FROM_PDF and not witness_name and pdf_path and os.path.exists(pdf_path):
        extracted = extract_witness_from_pdf(pdf_path, client) or ""
        if extracted and extracted.strip() not in _generic_witness_terms:
            witness_name = extracted
    if witness_name:
        llm_utils.log(f"[WITNESS] {witness_name}")
    else:
        llm_utils.log("[WITNESS] Unknown — LLM will infer from context")

    # Load speaker roster: cache → Word rubric → PDF → generic
    speakers_data = load_header_from_cache(case_num)
    if speakers_data:
        llm_utils.log(f"[HEADER] Loaded cached roster for case {case_num}")
    elif os.path.exists(WORD_RUBRIC_PATH):
        speakers_data = extract_speakers_from_word_doc(WORD_RUBRIC_PATH, client)
        if speakers_data:
            save_header_to_cache(case_num, speakers_data)
    elif extract_header and os.path.exists(pdf_path):
        speakers_data = extract_speakers_from_pdf(pdf_path, client)
        if speakers_data:
            save_header_to_cache(case_num, speakers_data)

    if not speakers_data:
        llm_utils.log("[WARNING] No speaker roster found, LLM will infer from context")
        speakers_data = {}

    speakers_context = format_speaker_context(speakers_data)

    # Build judge anchor — first judge in roster is the presiding judge who opens the session
    presiding_judge = ""
    if speakers_data and speakers_data.get("judges"):
        presiding_judge = speakers_data["judges"][0]

    if witness_name:
        witness_section = f"WITNESS ON THE STAND: {witness_name}\n- Always label their turns as: {witness_name}"
    else:
        witness_section = "WITNESS ON THE STAND: Identify from context (the person being formally questioned by lawyers)\n- Label them by their name once identified."

    judge_anchor = f"""PRESIDING JUDGE — CRITICAL:
The presiding judge is: {presiding_judge}
HOW TO IDENTIFY: The judge is the authority figure who:
  • Opens with "בוקר טוב לכולם" / "אנחנו בתיק {case_num}" / "היום [date]"
  • Warns the witness: "את חייבת לומר את האמת"
  • Manages the session, addresses lawyers, calls people to the stand
  • Is addressed by others as "כבודה" or "כבודכם"
USE THIS EXACT LABEL for all those turns: {presiding_judge}
NEVER use ת: or ש: for the judge — ALWAYS the full name above.""" if presiding_judge else ""

    system_instruction = f"""You are a verbatim court transcriber for case {case_num}.

═══════════════════════════════════════
SPEAKER IDENTIFICATION — PRIORITY ORDER
═══════════════════════════════════════
Use the following priority order to identify every speaker. Never skip to a lower priority if a higher one is available.

1. EXPLICIT LINGUISTIC CONTEXT (highest priority)
   - Speaker identifies their own role: "בית המשפט שואל...", "כבודה, התביעה מתנגדת..."
   - Speaker is addressed by name or title by another participant
   - Role-specific legal language that unambiguously signals the speaker

2. IMPLICIT CONTEXT — infer from speech patterns
   - Rulings, decisions, directing the proceeding → Judge
   - Questioning a witness in a formal Q&A, raising objections → Attorney (use ש: during Q&A)
   - Testifying, answering under oath → Witness (use ת:)

3. VOICE PROFILE (tiebreaker only)
   - Use voice samples provided ONLY when linguistic context (1 or 2) is insufficient or ambiguous
   - Do NOT override a clear linguistic signal with a voice match

4. UNKNOWN SPEAKER — when none of the above apply

RULE: Only assign a specific identity if you have at least ONE of:
  • The speaker explicitly stated their role or name
  • The speaker was addressed by name or title
  • A matching voice profile was provided
→ When uncertain: use UNKNOWN SPEAKER

═══════════════════════════════════════
SPEAKER ROSTER
═══════════════════════════════════════
{speakers_context}

{judge_anchor}

═══════════════════════════════════════
SPEAKER LABELS
═══════════════════════════════════════
{witness_section}

LABEL RULES — use the FULL NAME for every speaker, every time:
- Judge:    ALWAYS the exact full formal name from the roster (כב', ר', אב"ד, etc.)
- Attorney: FULL NAME from roster exactly as written — always, even during Q&A
- Witness:  FULL NAME (as given above) — always, every turn
- Unknown:  UNKNOWN SPEAKER

NEVER use ש: or ת: — use full names only.
Context clues: "כבודה/כבודו/כבודכם" = addressing a judge. "חברי/חברתי" = lawyer to lawyer.

═══════════════════════════════════════
OUTPUT FORMAT
═══════════════════════════════════════
Each speaker turn on its own line — NO timestamps, NO brackets:
FULL_NAME\tSPOKEN TEXT

Examples of correct format:
כב' השופטת ר' פרידמן-פלדמן - אב"ד\tבוקר טוב לכולם.
עו"ד עמית חדד\tכבודה, יש לי בקשה.
עו"ד ז'ק חן\tמה שמך?
שרה דווידי\tשרה דווידי.

- Every speaker change → new line with new label
- Verbatim. No paraphrasing. No summaries.
- Include every word, hesitation, correction, interruption
- Tab between label and text, nothing else
- If no speech: return EMPTY_SEGMENT"""

    base_dir = os.path.dirname(audio_path)
    txt_filename = f"Verbatim_{os.path.basename(audio_path).replace('.mp3', '.txt')}"
    txt_path = os.path.join(base_dir, txt_filename)

    with open(txt_path, "w", encoding="utf-8") as f:
        pass

    audio_file = None
    voice_files = []
    cache = None
    try:
        # Upload voice reference samples so Gemini can cross-reference speaker identity
        llm_utils.log("[UPLOAD] Uploading voice reference samples...")
        voice_parts = []
        for idx, (filename, speaker_label) in enumerate(VOICE_SAMPLES.items()):
            sample_path = os.path.join(VOICE_LIBRARY_PATH, filename)
            if not os.path.exists(sample_path):
                llm_utils.log(f"[VOICE] Missing sample: {filename} — skipped")
                continue
            try:
                with open(sample_path, "rb") as fp:
                    vf = client.files.upload(
                        file=fp,
                        config=types.UploadFileConfig(mime_type="audio/mpeg", display_name=f"voice_ref_{idx:02d}.mp3"),
                    )
                while vf.state.name == "PROCESSING":
                    time.sleep(2)
                    vf = client.files.get(name=vf.name)
                if vf.state.name == "ACTIVE":
                    voice_files.append(vf)
                    voice_parts.append(f"VOICE REFERENCE — {speaker_label}:")
                    voice_parts.append(vf)
                    llm_utils.log(f"[VOICE] Uploaded: {filename} ({speaker_label})")
                else:
                    llm_utils.log(f"[VOICE] Sample stuck in state {vf.state.name}: {filename}")
            except Exception as e:
                llm_utils.log(f"[VOICE] Failed to upload {filename}: {e}")

        llm_utils.log("[UPLOAD] Uploading audio to Gemini Files API...")
        audio_file = client.files.upload(
            file=audio_path,
            config=types.UploadFileConfig(mime_type="audio/mpeg", display_name=os.path.basename(audio_path)),
        )
        while audio_file.state.name == "PROCESSING":
            time.sleep(5)
            audio_file = client.files.get(name=audio_file.name)
        if audio_file.state.name != "ACTIVE":
            llm_utils.log(f"[ERROR] File stuck in state: {audio_file.state.name}")
            return
        llm_utils.log("[UPLOAD] File ACTIVE.")

        cache_contents = voice_parts + ["MAIN AUDIO TO TRANSCRIBE:", audio_file]

        llm_utils.log("[CACHE] Creating context cache (TTL 30m)...")
        cache = client.caches.create(
            model=MODEL_NAME,
            config=types.CreateCachedContentConfig(
                system_instruction=system_instruction,
                contents=cache_contents,
                ttl="1800s",
            ),
        )
        llm_utils.log(f"[CACHE] Cache created: {cache.name}")

        safety_settings = [
            types.SafetySetting(category="HARM_CATEGORY_HARASSMENT",        threshold="BLOCK_NONE"),
            types.SafetySetting(category="HARM_CATEGORY_HATE_SPEECH",       threshold="BLOCK_NONE"),
            types.SafetySetting(category="HARM_CATEGORY_DANGEROUS_CONTENT", threshold="BLOCK_NONE"),
            types.SafetySetting(category="HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold="BLOCK_NONE"),
        ]

        for start_min in range(0, MAX_MINUTES, CHUNK_MINUTES):
            end_min = start_min + CHUNK_MINUTES
            llm_utils.log(f"[PROCESS] Segment {start_min}:00 - {end_min}:00...")

            start_ts = f"{start_min // 60:02d}:{start_min % 60:02d}:00"
            end_ts = f"{end_min // 60:02d}:{end_min % 60:02d}:00"
            prompt = (
                f"Transcribe the audio segment from {start_ts} to {end_ts}. "
                f"CRITICAL FORMAT: each speaker turn on its own separate line, NO timestamps, NO brackets. "
                f"Format: LABEL\\tSPOKEN TEXT\\n"
                f"Every time the speaker changes, start a NEW LINE with the new label. "
                f"Verbatim. Identify speakers using linguistic context first, then voice profiles. "
                f"If no speech, return EMPTY_SEGMENT."
            )

            success = False
            for attempt in range(3):
                try:
                    response = client.models.generate_content(
                        model=MODEL_NAME,
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            cached_content=cache.name,
                            safety_settings=safety_settings,
                        ),
                    )
                    text = response.text.strip() if response.text else ""
                    llm_utils.log(f"[LLM RAW] Segment {start_min}:\n{text[:500]}")

                    if text and text != "EMPTY_SEGMENT":
                        # Clean up line by line — trust the model's newlines
                        clean = []
                        for ln in text.split('\n'):
                            ln = ln.strip()
                            if not ln:
                                continue
                            clean.append(ln)
                        normalized = normalize_to_protocol_labels(
                            '\n'.join(clean), witness_name, speakers_data
                        )
                        with open(txt_path, "a", encoding="utf-8") as f:
                            f.write(normalized.strip() + "\n")
                        llm_utils.log(f"[OK] Segment {start_min} saved ({len(normalized.splitlines())} lines).")
                        success = True
                        break
                    elif text == "EMPTY_SEGMENT":
                        llm_utils.log(f"[INFO] Segment {start_min} empty.")
                        success = True
                        break
                    time.sleep(5)
                except Exception as e:
                    llm_utils.log(f"[RETRY] Attempt {attempt+1} failed: {e}")
                    time.sleep(5)

            if not success:
                llm_utils.log(f"[WARN] Failed segment {start_min}.")

        hearing_date = derive_hearing_date(audio_path)
        llm_utils.log(f"[SUCCESS] Done. Output: {txt_path}")
        apply_formatting_final(txt_path, case_num=case_num, hearing_date=hearing_date, witness_name=witness_name)

    except Exception as e:
        llm_utils.log(f"[ERROR] {e}")
    finally:
        if cache:
            try:
                client.caches.delete(name=cache.name)
                llm_utils.log("[CACHE] Cache deleted.")
            except Exception:
                pass
        if audio_file:
            try:
                client.files.delete(name=audio_file.name)
                llm_utils.log("[UPLOAD] Audio file deleted.")
            except Exception:
                pass
        for vf in voice_files:
            try:
                client.files.delete(name=vf.name)
            except Exception:
                pass
        if voice_files:
            llm_utils.log(f"[UPLOAD] {len(voice_files)} voice sample(s) deleted.")


def format_speaker_context(speakers_data):
    """Format extracted speaker data into a readable context block."""
    if isinstance(speakers_data, list) and len(speakers_data) > 0:
        speakers_data = speakers_data[0]
    if not speakers_data or not isinstance(speakers_data, dict):
        return "Speakers: Extract from audio context using the labeling rules above."

    context = []

    if speakers_data.get("judges"):
        context.append("JUDGES: " + ", ".join(speakers_data["judges"]))
    if speakers_data.get("prosecution_lawyers"):
        context.append("PROSECUTION: " + ", ".join(speakers_data["prosecution_lawyers"]))
    if speakers_data.get("defense_lawyers"):
        context.append("DEFENSE: " + ", ".join(speakers_data["defense_lawyers"]))
    if speakers_data.get("defendants"):
        context.append("DEFENDANTS: " + ", ".join(speakers_data["defendants"]))

    return "\n".join(context) if context else "Speakers: Extract from audio context."


def apply_formatting_final(txt_path, case_num=CASE_NUM, hearing_date="", witness_name=""):
    """Create Word document from transcription with Hebrew RTL formatting and protocol header."""
    import pythoncom
    llm_utils.log("[WORD] Creating Word document...")
    word = None
    try:
        pythoncom.CoInitialize()

        with open(txt_path, "r", encoding="utf-8") as f:
            raw_lines = [l.rstrip("\n") for l in f.readlines()]

        transcript_lines = [line.strip() for line in raw_lines if line.strip()]

        header_lines = [
            l.format(case_num=case_num, hearing_date=hearing_date, witness_name=witness_name or "")
            for l in PROTOCOL_HEADER_LINES
        ]

        all_lines = header_lines + transcript_lines

        save_path = txt_path.replace(".txt", ".docx")

        # Kill any stale Word instance before starting fresh
        try:
            stale = win32.GetActiveObject("Word.Application")
            stale.Quit(SaveChanges=0)
            time.sleep(1)
        except Exception:
            pass

        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.Documents.Add()
        doc = word.ActiveDocument
        doc.Content.Delete()

        # Step 1 — type all lines using Selection (reliable multi-paragraph approach)
        sel = word.Selection
        sel.HomeKey(Unit=6)  # wdStory
        for i, line_text in enumerate(all_lines):
            sel.Font.Name = "David"
            sel.Font.Size = 12
            sel.TypeText(line_text)
            if i < len(all_lines) - 1:
                sel.TypeParagraph()

        # Step 2 — sweep all paragraphs: set RTL, justify, 1.5 spacing, bold speaker label
        for p in doc.Paragraphs:
            p.Format.ReadingOrder = 1       # RTL
            p.Format.Alignment = 3          # Justify
            p.Format.LineSpacingRule = 5    # Multiple
            p.Format.LineSpacing = 18       # 1.5 lines
            line_text = p.Range.Text.rstrip('\r\n\x0d')
            if ': ' in line_text:
                speaker = line_text.split(': ', 1)[0]
                doc.Range(p.Range.Start, p.Range.Start + len(speaker) + 1).Font.Bold = True

        # Step 3 — final font sweep (catches para 1 default style)
        doc.Content.Font.Name = "David"
        doc.Content.Font.Size = 12

        # Line numbering + footer
        try:
            section = doc.Sections(1)
            section.PageSetup.LineNumbering.Active = True
            section.PageSetup.LineNumbering.StartingNumber = 1
            section.PageSetup.LineNumbering.CountBy = 1
            section.PageSetup.LineNumbering.RestartMode = 1  # wdRestartPage
            footer = section.Footers(1)
            footer.Range.ParagraphFormat.Alignment = 1
            footer.PageNumbers.Add(PageNumberAlignment=1, FirstPage=True)
        except Exception as e:
            llm_utils.log(f"[WORD] Section setup warning: {e}")

        doc.SaveAs(os.path.abspath(save_path), FileFormat=16)
        doc.Close(False)
        llm_utils.log(f"[OK] Word document saved: {save_path}")
    except Exception as e:
        llm_utils.log(f"[ERROR] Word formatting failed: {e}")
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    extract_header = "--extract-header" in sys.argv

    # Parse optional arguments: --audio /path --case 67104-01-20 --pdf /path --witness "שם העד"
    audio_path = AUDIO_PATH
    case_num = CASE_NUM
    pdf_path = PDF_PATH
    witness_name = WITNESS_NAME

    for i, arg in enumerate(sys.argv[1:], 1):
        if arg == "--audio" and i < len(sys.argv) - 1:
            audio_path = sys.argv[i + 1]
        elif arg == "--case" and i < len(sys.argv) - 1:
            case_num = sys.argv[i + 1]
        elif arg == "--pdf" and i < len(sys.argv) - 1:
            pdf_path = sys.argv[i + 1]
        elif arg == "--witness" and i < len(sys.argv) - 1:
            witness_name = sys.argv[i + 1]

    transcribe_vertex_final(
        audio_path=audio_path,
        case_num=case_num,
        pdf_path=pdf_path,
        witness_name=witness_name,
        extract_header=extract_header
    )
