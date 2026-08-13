import os
import json
import datetime
import mimetypes
from dotenv import load_dotenv
load_dotenv() # Load variables from .env

from google import genai
from google.genai import types
from google.cloud import storage

LOG_FILE = os.path.join(os.path.dirname(__file__), "logs", "run_log.txt")

def log(msg):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    formatted_msg = f"{timestamp} - {msg}"
    try:
        print(formatted_msg, flush=True)
    except:
        try:
            print(formatted_msg.encode('ascii', errors='replace').decode('ascii'), flush=True)
        except:
            pass
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(formatted_msg + "\n")
    except:
        pass

# הגדרות Vertex AI
GCP_PROJECT_ID = os.getenv("GCP_PROJECT_ID")
GCS_BUCKET_NAME = os.getenv("GCS_BUCKET_NAME")
VERTEX_LOCATION = os.getenv("VERTEX_LOCATION", "global")

# Vertex AI client (google-genai SDK, REST via global endpoint)
_ai_client = None

def _get_client():
    global _ai_client
    if _ai_client is None and GCP_PROJECT_ID:
        _ai_client = genai.Client(
            vertexai=True,
            project=GCP_PROJECT_ID,
            location=VERTEX_LOCATION,
            http_options=types.HttpOptions(api_version="v1"),
        )
    return _ai_client

def init_vertex():
    """נשמר לתאימות אחורה"""
    return _get_client() is not None

def upload_to_gcs(local_path, bucket_name=None):
    """העלאת קובץ ל-Google Cloud Storage"""
    if not bucket_name:
        bucket_name = GCS_BUCKET_NAME
    
    if not bucket_name:
        log("[ERROR] No GCS bucket name configured.")
        return None
        
    try:
        client = storage.Client(project=GCP_PROJECT_ID)
        bucket = client.bucket(bucket_name)
        blob_name = os.path.basename(local_path)
        blob = bucket.blob(blob_name)
        
        log(f"     [GCS] Uploading {blob_name} to gs://{bucket_name}/...")
        blob.upload_from_filename(local_path)
        return f"gs://{bucket_name}/{blob_name}"
    except Exception as e:
        log(f"[ERROR] GCS Upload failed: {e}")
        return None

def delete_from_gcs(gcs_uri):
    """מחיקת קובץ מ-GCS"""
    try:
        if not gcs_uri.startswith("gs://"): return
        path = gcs_uri.replace("gs://", "")
        bucket_name, blob_name = path.split("/", 1)
        
        client = storage.Client(project=GCP_PROJECT_ID)
        bucket = client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        blob.delete()
        log(f"     [GCS] Deleted {blob_name}")
    except Exception as e:
        log(f"[WARNING] GCS Delete failed: {e}")


def init_client(force_rotate=False):  # noqa: ARG001
    """נשמר לתאימות אחורה — Vertex AI משתמש ב-ADC, אין צורך במפתחות API."""
    return GCP_PROJECT_ID is not None

# שמות מודלים
MODEL_FAST = "gemini-3.1-flash-lite-preview"    # מהיר וחסכוני
MODEL_COMPLEX = "gemini-3.1-flash-lite-preview"  # מודל חזק — לשימוש ב-OCR ומשימות מורכבות

def _call_with_retry(func, *args, **kwargs):
    """מריץ פונקציית AI עם Retry במקרה של 429 / quota exhausted."""
    import time
    if not init_client():
        return None
    for attempt in range(5):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            err = str(e)
            if "429" in err or "Resource exhausted" in err or "RESOURCE_EXHAUSTED" in err:
                wait = 30 * (attempt + 1)
                print(f"[WAIT] Quota exhausted. Waiting {wait}s...")
                time.sleep(wait)
            else:
                print(f"[ERROR] AI Error: {e}")
                break
    return None

def ask_llm(prompt, model_name=MODEL_FAST):
    def _run():
        client = _get_client()
        response = client.models.generate_content(model=model_name, contents=prompt)
        return response.text.strip()
    return _call_with_retry(_run)

def ask_llm_json(prompt, model_name=MODEL_FAST):
    def _run():
        client = _get_client()
        system_instruction = "You are a helpful assistant that outputs ONLY valid JSON."
        response = client.models.generate_content(
            model=model_name,
            contents=f"{system_instruction}\n\n{prompt}",
            config=types.GenerateContentConfig(response_mime_type="application/json"),
        )
        return json.loads(response.text.strip())
    return _call_with_retry(_run)

def ask_llm_json_with_file(prompt, file_path, model_name=MODEL_FAST):
    if not init_client():
        return None

    gcs_uri = None
    try:
        log(f"     [GCS] Uploading file for inference: {os.path.basename(file_path)}...")
        gcs_uri = upload_to_gcs(file_path)
        if not gcs_uri:
            log("[ERROR] GCS upload failed — cannot run inference.")
            return None

        mime_type, _ = mimetypes.guess_type(file_path)
        if not mime_type:
            mime_type = "audio/mpeg"

        def _run():
            client = _get_client()
            system_instruction = "You are a helpful assistant that outputs ONLY valid JSON."
            audio_part = types.Part.from_uri(file_uri=gcs_uri, mime_type=mime_type)
            response = client.models.generate_content(
                model=model_name,
                contents=[system_instruction, prompt, audio_part],
                config=types.GenerateContentConfig(response_mime_type="application/json"),
            )
            return json.loads(response.text.strip())

        return _call_with_retry(_run)
    finally:
        if gcs_uri:
            delete_from_gcs(gcs_uri)

def ask_llm_complex(prompt):
    return ask_llm(prompt, model_name=MODEL_COMPLEX)

def ask_llm_complex_json(prompt):
    return ask_llm_json(prompt, model_name=MODEL_COMPLEX)

def extract_judge_llm(text_content):
    """
    מחלץ את שם השופט מתוך טקסט התמלול.
    """
    if not text_content:
        return None
        
    text_sample = text_content[:3000]
    
    prompt = f"""Analyze the following Hebrew legal court transcription.
    NOTE: The text may be REVERSED (Right-to-Left rendering issue). Read it carefully.
    
    Goal: Identify and extract the Judge's name (שם השופט/ת).
    
    Transcription Header:
    "{text_sample}"
    
    Rules:
    1. Look for phrases like "בפני כבוד השופט/ת", "בפני השופט/ת", "לפני השופט/ת".
    2. Extract only the name itself (e.g. "נמרוד אשכול", "מרים אילני"). 
    3. If the judge is a Registrar (רשם/ת), extract that and the name.
    4. VERY IMPORTANT: If the name appears backwards (e.g., "לכימ טיבש" or "ןמדלפ-ןמדירפ"), you MUST reverse the letters so the final output is in correct Hebrew order (e.g., "מיכל שביט" or "פרידמן-פלדמן").
    5. Return ONLY a JSON object with a single key "judge_name".
    """
    
    result = ask_llm_json(prompt)
    if result and "judge_name" in result:
        return result["judge_name"]
    return None

def detect_district_llm(text_content, filename="", possible_districts=None):
    """
    מזהה את המחוז/בית המשפט מתוך טקסט התמלול באמצעות LLM.
    מחזיר מפתח תואם ל-DISTRICT_CONTACTS (למשל 'מחוזי_ירושלים').
    """
    if not text_content and not filename:
        return None
    
    # שולח רק את ההתחלה (הכותרת)
    text_sample = text_content[:4000] if text_content else ""
    
    districts_str = ", ".join(possible_districts) if possible_districts else "Any valid district"
    
    prompt = f"""Analyze the following Hebrew legal court transcription.
    
    Goal: Identify which court district this file belongs to.
    Required result: One of these specific ID keys: [{districts_str}] or null if unsure.
    
    Filename: "{filename}"
    Transcription Header:
    "{text_sample}"
    
    Rules:
    1. Look for the exact court name at the top (שלום, מחוזי, עליון, משפחה, תעבורה).
    2. Identify the city (ירושלים, באר שבע, אשקלון, בית שמש, וכו').
    3. If it's the Supreme Court (עליון), return "ביהמש_העליון".
    4. If it's Jerusalem District, return "מחוזי_ירושלים".
    5. If it's Jerusalem Shalom, return "שלום_ירושלים".
    6. SPECIAL RULE: Both "בית שמש" (Beit Shemesh) cases and "תעבורה ירושלים" (Jerusalem Traffic) cases must be returned as "משפחה_ירושלים".
    7. SPECIAL RULE: If the name "הרכב עמית" appears, return "ביהמש_העליון".
    8. Return ONLY a JSON object with a single key "district_key".
    """
    
    result = ask_llm_json(prompt)
    if result and "district_key" in result:
        return result["district_key"]
    return None


def validate_transcription_quality(text_content: str, expected_case_num: str = "",
                                    expected_date: str = "", filename: str = "",
                                    expected_judge: str = "") -> dict:
    """
    ולידציה מקיפה לאיכות תמלול לפני שליחת מייל.
    בודק:
      1. התאמת מספר תיק: כותרת  גוף
      2. התאמת תאריך: כותרת  גוף (כולל פורמטים עבריים)
      3. שלמות הדיון (מסתיים ב"הקלטה נגמרה" / "הדיון הסתיים" / END marker)
      4. תגיות ריקות: "ת:", "ש:" ללא טקסט (תמלול שנתקע)
      5. שם שופט: עקביות כותרת  גוף + מול expected_judge (אם סופק)
      6. דוברים: האם קיימות שורות דיאלוג עם תוכן בפועל

    מחזיר dict:
      {
        "valid": True/False,   # True = OK לשלוח (אין בעיות חוסמות)
        "score": 0-100,
        "issues": [...],       # בעיות חוסמות (Hard)
        "soft_issues": [...],  # בעיות שניתן להתעלם מהן (Soft)
        "warnings": [...],     # אזהרות כלליות
        "details": {...}       # פרטים מה-AI
      }
    """
    import re

    issues = []
    soft_issues = []
    completeness_issues = []  # incomplete transcript — reduces score but doesn't hard-block
    warnings = []
    details = {}

    if not text_content or len(text_content.strip()) < 50:
        return {
            "valid": False, "score": 0,
            "issues": ["הקובץ ריק או קצר מדי לבדיקה"],
            "warnings": [], "details": {}
        }

    #  בדיקה 1: שלמות הדיון 
    tail = text_content[-500:].strip()
    end_markers = [
        "הקלטה נגמרה", "הקלטה הסתיימה", "הדיון הסתיים", "הדיון נגמר",
        "נגמרה ההקלטה", "תם הדיון", "סיום הקלטה", "recording ended",
        "end of recording", "נגמר הדיון",
        # judicial closing phrases
        "הישיבה נעולה", "הדיון נעול", "הישיבה הסתיימה", "הדיון הסתיים",
        "הישיבה הסתיימה", "הישיבה ננעלת", "ננעלת הישיבה",
        "להפסיק את ההקלטה", "הפסק את ההקלטה", "תפסיק את ההקלטה",
        "תם הדיון", "נעילת הדיון", "נעילת הישיבה",
    ]
    is_complete = any(marker in tail for marker in end_markers)
    details["is_complete"] = is_complete
    details["tail_preview"] = tail[-120:].replace("\n", " ")

    if not is_complete:
        last_char = tail.rstrip()[-1] if tail.rstrip() else ""
        if last_char in ".?!'\"":
            warnings.append("לא נמצא 'הקלטה נגמרה'  אך הטקסט נגמר בסימן פיסוק (ייתכן שלם)")
        else:
            warnings.append("לא נמצא 'הקלטה נגמרה' בסוף הקובץ  כדאי לבדוק")

    #  בדיקה 2: דפוסי תמלול חשוד 
    #  שורות תגית ריקות (ת: / ש: ללא טקסט)
    # תגית ריקה: ת: או ש: (עם גרש אופציונלי) ללא טקסט אחרי
    _SPEAKER_PREFIX = r"(?:\u05ea|\u05e9|\u05e2\u05d5\"\u05d3|\u05d4\u05ea\u05d5\u05d1\u05e2|\u05d4\u05e0\u05d0\u05e9\u05dd|\u05d4\u05e2\u05d3|\u05d4\u05e1\u05e0\u05d9\u05d2\u05d5\u05e8)"
    empty_label_pattern = re.compile(
        r"^[ \t]*" + _SPEAKER_PREFIX + r"['\u2019\u05f3]?[ \t]*[:.][ \t]*$",
        re.MULTILINE
    )
    empty_labels = empty_label_pattern.findall(text_content)

    #  חזרתיות ברצף  בסוף הפרוטוקול (500 תווים אחרונים): שורה שחוזרת 3+ פעמים ברצף
    tail_lines = [ln.strip() for ln in tail.splitlines() if ln.strip()]
    max_consecutive = 1
    if tail_lines:
        cur_run = 1
        for i in range(1, len(tail_lines)):
            if tail_lines[i] == tail_lines[i - 1]:
                cur_run += 1
                max_consecutive = max(max_consecutive, cur_run)
            else:
                cur_run = 1
    details["max_consecutive_repeated_lines"] = max_consecutive

    #  יחס שורות ריקות  אם >35% מהתגיות הן ריקות
    filled_labels = re.findall(
        r"^[ \t]*(?:\u05ea|\u05e9|\u05e2\u05d5\"\u05d3|\u05d4\u05ea\u05d5\u05d1\u05e2|\u05d4\u05e0\u05d0\u05e9\u05dd|\u05d4\u05e2\u05d3|\u05d4\u05e1\u05e0\u05d9\u05d2\u05d5\u05e8)[\u2018\u2019\u05f3']?[ \t]*[:.][ \t]*.{3,}",
        text_content, re.MULTILINE
    )
    total_labels = len(empty_labels) + len(filled_labels)
    empty_ratio = (len(empty_labels) / total_labels) if total_labels > 0 else 0
    details["empty_label_ratio"] = round(empty_ratio, 2)
    details["empty_labels_count"] = len(empty_labels)

    # -- שיפוט --
    if max_consecutive >= 4:
        issues.append(
            f"[REPEATED] תמלול נראה תקוע  שורה חוזרת {max_consecutive} פעמים ברצף בסוף הקובץ"
        )
    elif max_consecutive >= 3:
        soft_issues.append(f"שורה חוזרת {max_consecutive} פעמים ברצף בסוף הקובץ  ייתכן תקיעות")

    if len(empty_labels) >= 10 and empty_ratio > 0.5:
        issues.append(
            f"[CRITICAL] {len(empty_labels)} תגיות ריקות ({int(empty_ratio*100)}%)  תמלול נראה הרוס"
        )
    elif len(empty_labels) >= 3 and empty_ratio > 0.35:
        soft_issues.append(
            f"[WARN] {len(empty_labels)} תגיות ריקות ({int(empty_ratio*100)}%)  ייתכן תמלול חלקי"
        )
    elif len(empty_labels) >= 1:
        warnings.append(f"{len(empty_labels)} תגיות ריקות (ת:/ש: ללא טקסט)")

    #  בדיקה 3: דוברים עם תוכן 
    # תומך בפורמטים: ת: / ת': / ש: / ש': / עו"ד: וכו'
    # זוהי אזהרה בלבד  פורמט משתנה בין פרוטוקול לפרוטוקול
    has_content_pattern = re.compile(
        r"^[ \t]*(?:ת|ש|עו\"ד|התובע|הנאשם|העד|הסניגור|השופט|הנאמן)['\u2019\u05f3]?[ \t]*[:.][ \t]*.{3,}",
        re.MULTILINE
    )
    content_lines = has_content_pattern.findall(text_content)
    details["speaker_lines_with_content"] = len(content_lines)
    if len(content_lines) < 3:
        warnings.append(
            f"זוהו {len(content_lines)} שורות דיאלוג בפורמט סטנדרטי  ייתכן פורמט שונה או קובץ שאינו תמלול"
        )

    #  בדיקה 3: AI  התאמת מספר תיק ותאריך בין כותרת לגוף 
    text_start = text_content[:3000]
    text_body  = text_content[3000:7000] if len(text_content) > 3000 else ""

    static_complete = details.get("is_complete", False)
    prompt = f"""You are validating a Hebrew legal court transcription.

EXPECTED (from filename/folder/Excel):
- Case number: "{expected_case_num}"
- Date: "{expected_date}"  (ISO format YYYY-MM-DD)
- Judge name: "{expected_judge}"  (empty = unknown, skip judge validation)
- Filename: "{filename}"
- Static end-marker check: {"FOUND — a valid end marker was detected in the tail" if static_complete else "NOT FOUND — no end marker detected yet"}

DOCUMENT HEADER (first ~3000 chars):
'''{text_start}'''

DOCUMENT BODY / TAIL SAMPLE (including end of document):
'''{text_body if text_body else "(No additional body/tail provided)"}'''

DOCUMENT TAIL PREVIEW (last ~120 chars):
'''{details['tail_preview']}'''

NOTE: The text above may contain a "[GAP]" or jump suddenly from header to tail. This is normal. 
Focus on the TAIL for the closing phrase check.


Return ONLY a JSON object with these exact keys:
{{
  "header_case_num": "case number found in header or null",
  "body_case_num": "case number found in body or null",
  "header_date": "date found in header normalized to DD/MM/YYYY, or null",
  "body_date": "date found in body normalized to DD/MM/YYYY, or null",
  "header_judge": "judge name found in header (after 'בפני כבוד השופט/ת') or null",
  "body_judge": "judge name found mentioned in body or null",
  "case_num_consistent": true/false,
  "date_consistent": true/false,
  "judge_consistent": true/false,
  "case_matches_expected": true/false,
  "date_matches_expected": true/false,
  "judge_matches_expected": true/false,
  "is_complete": true/false,
  "confidence": 1-10,
  "issues_found": [
     {{"text": "Hebrew description", "severity": "hard|soft"}}
  ]
}}

IMPORTANT RULES:
1. **MATCHING EXPECTED IS MANDATORY**: "case_matches_expected" and "date_matches_expected" MUST be true if the document belongs to the case. If the case number or judge in the document differs from the filename/expected, it's a SUBSTANTIAL ERROR (HARD).
2. **INTERNAL CONSISTENCY**: "case_num_consistent" can be False if the body refers to a different case (like an appeal), and this is a SOFT issue.
3. **INCOMPLETE CHECK**: Set "is_complete"=true if ANY of the following apply:
   - The static end-marker check above says FOUND
   - The tail contains any of these phrases: "סיום הקלטה", "הקלטה נגמרה", "הקלטה הסתיימה", "הדיון הסתיים", "תם הדיון", "נגמרה ההקלטה", "הפסק את ההקלטה", "להפסיק את ההקלטה", "recording ended", "end of recording", "הישיבה נעולה", "הדיון נעול", "ננעלת הישיבה", "נעילת הדיון", "נעילת הישיבה"
   - A judge or party explicitly says to stop the recording (e.g. "אדוני יכול להפסיק את ההקלטה", "תפסיק את ההקלטה")
   - The last substantive line is a clear judicial closing declaration: "הישיבה נעולה" / "הדיון נעול" / "הדיון הסתיים" are all COMPLETE endings — they mean "the session is closed"
   - The last substantive line is a clear judicial closing remark (e.g. setting a deadline, adjournment, "תודה רבה")
   Set "is_complete"=false ONLY if the text cuts off mid-word or mid-sentence with no closing remark at all.
   Missing "הקלטה נגמרה" alone is NOT a reason to set is_complete=false.
4. **FORMATTING**: Ignore / vs - and year format. Also ignore missing dashes in the year part (e.g., 0621 is the same as 06-21). Always consider the case number a match if all the digits appear in the correct sequence, even if separators are missing or different. For dates, 'DD/MM/YYYY' is exactly the same as 'YYYY-MM-DD' (e.g. 19/05/2026 == 2026-05-19). Ignore the order and separators.
5. **JUDGES**: If expected_judge is EMPTY (""), "judge_matches_expected" is ALWAYS True. Judge references in the body as "כב' השופט/ת", "כב' הש'", or by title alone (without the name) are normal and do NOT indicate a mismatch.
6. **SOFT vs HARD**:
   - HARD: Mismatch between Document and Filename/Expected metadata. Text cut off mid-word/mid-sentence with no closing. Major internal conflicts. Stuck repeated text.
   - SOFT: Internal body references to other cases/dates, minor spelling differences (גרוקופף vs גרוסקופף), judge referenced only by title in body.
7. If an issue is specifically allowed (like empty expected_judge), DO NOT list it in "issues_found".
"""

    def _parse_date_to_string(d_val):
        if not d_val: return None
        s = str(d_val).strip().split(" ")[0]
        s = s.replace('\u200e', '').replace('\u200f', '').replace('-', '/').replace('.', '/')
        parts = s.split('/')
        if len(parts) == 3:
            if len(parts[0]) == 4:
                y, m, d = parts[0], parts[1], parts[2]
            elif len(parts[2]) == 4:
                d, m, y = parts[0], parts[1], parts[2]
            else:
                d, m, y = parts[0], parts[1], parts[2]
                y = "20" + y if len(y) == 2 else y
            try:
                return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
            except:
                pass
        return None

    def _clean_case_num(c):
        return "".join(char for char in str(c) if char.isdigit())

    ai_result = ask_llm_json(prompt)
    if ai_result:
        # Programmatic Date Match override
        h_date = _parse_date_to_string(ai_result.get("header_date"))
        exp_date = _parse_date_to_string(expected_date)
        if h_date and exp_date and h_date == exp_date:
            ai_result["date_matches_expected"] = True

        # Programmatic Case Match override
        h_case = _clean_case_num(ai_result.get("header_case_num"))
        exp_case = _clean_case_num(expected_case_num)
        if h_case and exp_case and h_case == exp_case:
            ai_result["case_matches_expected"] = True

        # פוסט-עיבוד: אם expected_judge ריק, אנחנו מתעלמים מכל מה שה-AI אמר על שופטים
        if not expected_judge:
            ai_result["judge_matches_expected"] = True
            # ניקוי הודעות שגיאה על שופטים מה-AI
            if "issues_found" in ai_result:
                ai_result["issues_found"] = [iss for iss in ai_result["issues_found"] 
                                             if "judge" not in str(iss).lower() and "שופט" not in str(iss)]

        details.update(ai_result)

        if ai_result.get("case_num_consistent") is False:
            # אם יש אי-עקביות פנימית אבל ה-AI סימן אותה כ-soft, נכבד זאת
            soft_found = any(iss.get("severity") == "soft" and "תיק" in iss.get("text", "") 
                             for iss in ai_result.get("issues_found", []) if isinstance(iss, dict))
            if soft_found:
                soft_issues.append(f"אי עקביות פנימית במספר תיק (ככל הנראה אזכור של תיק אחר)")
            else:
                issues.append(
                    f"[ISSUE] מספר תיק לא עקבי: כותרת={ai_result.get('header_case_num')}  גוף={ai_result.get('body_case_num')}"
                )
        
        if expected_case_num and ai_result.get("case_matches_expected") is False:
            # שגיאה מהותית: שם הקובץ לא תואם לתוכן המסמך
            issues.append(
                f"[CRITICAL] מספר תיק בתמלול ({ai_result.get('header_case_num')})  שם הקובץ ({expected_case_num})"
            )

        if ai_result.get("date_consistent") is False:
            soft_found = any(iss.get("severity") == "soft" and "תאריך" in iss.get("text", "") 
                             for iss in ai_result.get("issues_found", []) if isinstance(iss, dict))
            if soft_found:
                soft_issues.append(f"אזכור תאריך אחר בגוף המסמך (תקין)")
            else:
                issues.append(
                    f"[ISSUE] תאריך לא עקבי: כותרת={ai_result.get('header_date')}  גוף={ai_result.get('body_date')}"
                )

        if expected_date and ai_result.get("date_matches_expected") is False:
            # שגיאה מהותית: התאריך במסמך לא תואם לתאריך הדיון המצופה
            issues.append(
                f"[CRITICAL] תאריך בתמלול ({ai_result.get('header_date')})  צפוי ({expected_date})"
            )

        if ai_result.get("is_complete") is False:
            if static_complete:
                soft_issues.append("AI: ייתכן שהתמלול נקטע, אך נמצא סמן סיום תקין בסוף הקובץ")
            else:
                completeness_issues.append("התמלול נראה לא שלם או שנקטע באמצע")
            
        #  שם שופט 
        if ai_result.get("judge_consistent") is False:
            warnings.append(
                f"שם השופט בכותרת ({ai_result.get('header_judge')}) לא מופיע בגוף הפרוטוקול  ייתכן תקין בדיון קצר"
            )
        if expected_judge and ai_result.get("judge_matches_expected") is False:
            issues.append(
                f"[ISSUE] שם שופט לא תואם: בתמלול='{ai_result.get('header_judge')}'  צפוי='{expected_judge}'"
            )
            
        # עיבוד issues_found מה-AI
        # if AI said is_complete=False, all related issues_found are completeness (not hard blockers)
        ai_said_incomplete = (ai_result.get("is_complete") is False)
        _incomplete_keywords = (
            "שלם", "נקטע", "complete", "cut", "abrupt", "ending",
            "מסתיים", "מסתיים באמצע", "נגמר", "סיום", "קטוע", "mid-sentence",
            "mid sentence", "incomplete", "truncat", "ללא הצהרת סיום", "ללא סיום",
        )
        for ai_issue in (ai_result.get("issues_found") or []):
            txt = ai_issue.get("text") if isinstance(ai_issue, dict) else str(ai_issue)
            sev = ai_issue.get("severity", "hard") if isinstance(ai_issue, dict) else "hard"
            is_completeness = ai_said_incomplete or any(kw in txt.lower() for kw in _incomplete_keywords)

            if sev == "soft":
                if txt not in soft_issues: soft_issues.append(f"AI: {txt}")
            elif is_completeness:
                if txt not in completeness_issues: completeness_issues.append(f"AI: {txt}")
            else:
                if txt not in issues: issues.append(f"AI: {txt}")
    else:
        warnings.append("לא ניתן לאמת מול AI (שגיאת רשת/quota)  בדיקת Regex בלבד")

    #  חישוב ציון כולל
    # completeness_issues cost 20 pts each but don't hard-block
    score = max(0, 100 - len(issues) * 25 - len(completeness_issues) * 20
                - len(soft_issues) * 10 - len(warnings) * 5)

    # blocked only by real data mismatches; incomplete transcript alone is not a blocker
    valid = len(issues) == 0

    return {
        "valid": valid,
        "score": score,
        "issues": issues + completeness_issues,  # show in log but completeness didn't block
        "soft_issues": soft_issues,
        "warnings": warnings,
        "details": details,
    }


# --- פונקציות עזר מתקדמות ---

def verify_transcription(text_content, expected_case_num="", expected_date="", expected_judge=""):
    """
    בדיקת שפיות לתמלול: מוודא שהמידע בטקסט תואם למה שציפינו.
    בודק: מספר תיק, תאריך, שם שופט, שלמות הטקסט.
    מחזיר dict עם תוצאות הבדיקה.
    """
    if not text_content:
        return {"valid": False, "issues": ["No text content provided"]}
    
    # שולח רק חלק מהטקסט כדי לחסוך טוקנים
    text_sample = text_content[:2000]
    text_end = text_content[-500:] if len(text_content) > 500 else ""
    
    prompt = f"""You are verifying a Hebrew legal court transcription for accuracy.

EXPECTED metadata:
- Case number: "{expected_case_num}"
- Date: "{expected_date}"
- Judge: "{expected_judge}"

DOCUMENT START (first ~2000 chars):
"{text_sample}"

DOCUMENT END (last ~500 chars):
"{text_end}"

Total document length: {len(text_content)} characters, ~{len(text_content.split())} words.

Check the following and return JSON:
{{
  "case_num_match": true/false (does the document contain a matching case number?),
  "date_match": true/false (does the document date match?),
  "judge_match": true/false (does the judge name appear in the document?),
  "text_complete": true/false (does the text appear complete - has proper ending, not cut off mid-sentence?),
  "found_case_num": "the case number found in the document or null",
  "found_date": "the date found in the document or null",
  "found_judge": "the judge name found in the document or null",
  "issues": ["list of any problems found"],
  "confidence": 1-10
}}"""

    result = ask_llm_complex_json(prompt)
    
    if not result:
        return {"valid": False, "issues": ["LLM verification failed"]}
    
    # סיכום: תקין אם כל הבדיקות עברו
    checks = [
        result.get("case_num_match", False) if expected_case_num else True,
        result.get("date_match", False) if expected_date else True,
        result.get("judge_match", False) if expected_judge else True,
        result.get("text_complete", False),
    ]
    result["valid"] = all(checks)
    
    return result

def extract_court_and_city_llm(raw_text, subject=""):
    """
    חילוץ בית משפט ועיר מטקסט PDF של הזמנות (הזמנות.ipynb).
    הטקסט לרוב הפוך (RTL)  LLM מסוגל לקרוא אותו.
    
    מחזיר dict עם: court_type, city, case_num, dates (list), confidence
    """
    if not raw_text or len(raw_text.strip()) < 20:
        return None
    
    # שולח רק 2000 תווים ראשונים לחיסכון
    text_sample = raw_text[:2000]
    
    prompt = f"""Analyze this Hebrew legal court order PDF text. NOTE: The text may be REVERSED (right-to-left rendering issue) - read it accordingly.

Email subject: "{subject}"

PDF text (may be reversed):
"{text_sample}"

Extract:
1. court_type: Type of court (שלום/מחוזי/תעבורה/משפחה/עניינים מקומיים/מיוחד)
2. city: City name in Hebrew (ירושלים/תל אביב/באר שבע/בית שמש/אשקלון etc.)
3. case_num: Case number (format: XXXXX-XX-XX)
4. dates: List of hearing dates and times found [{{"date": "DD/MM/YYYY", "time": "HH:MM", "city": "city if different per date"}}]

Return JSON with keys: "court_type", "city", "case_num", "dates", "confidence" (1-10)
Use null for fields you cannot determine."""

    return ask_llm_complex_json(prompt)

# --- פונקציות עזר מתקדמות ---

def extract_case_details_llm(filename):
    """
    דוגמה: חילוץ פרטי תיק משם קובץ באמצעות Gemini.
    """
    prompt = f"""
    Analyze the following filename and extract:
    1. Case Number (format like 12345-01-23 or similar)
    2. Date (format YYYY-MM-DD, try to infer from filename)
    3. Court Type (Shalom, District, etc. - in Hebrew)
    
    Filename: "{filename}"
    
    Return JSON with keys: "case_num", "date", "court", "confidence" (1-10)
    If a field is missing, use null.
    """
    
    return ask_llm_json(prompt)

if __name__ == "__main__":
    print("--- בדיקת Gemini LLM ---")
    # בדיקה פשוטה
    # res = ask_llm("מה בירת צרפת?")
    # print(f"תשובה: {res}")
    
    # בדיקת JSON
    complex_file = "פרוטוקול דיון 12-34-56 (נמשך) מתאריך 1.1.25 בבית משפט שלום י-ם.docx"
    print(f"\nבודק קובץ: {complex_file}")
    res_json = extract_case_details_llm(complex_file)
    print(f"תוצאה JSON:\n{json.dumps(res_json, indent=2, ensure_ascii=False)}")
