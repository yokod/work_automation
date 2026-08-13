"""
📧 send_district_emails.py
==========================
שולח מיילים לאנשי קשר לפי מחוז אחרי שקובץ תמלול עובד בהצלחה.

תכונות:
- מיפוי מחוז/בית משפט → איש קשר
- שליחה רק בימי עבודה (א-ה) בין 08:00-17:00
- תור מיילים (pending_emails.json) לשעות מחוץ לשעות עבודה
- שליחת תור אחד כל ~5 דקות (Task Scheduler קורא לזה)
- מניעת כפילויות (sent_emails.json)
- DRY_RUN mode לבדיקה בטוחה
"""

import os
import json
import re
import random
import datetime
import win32com.client as win32

# ייבוא כלי ה-AI
try:
    from llm_utils import detect_district_llm, validate_transcription_quality
except ImportError:
    detect_district_llm = None
    validate_transcription_quality = None

# ייבוא בודק עיצוב Word
try:
    from word_format_checker import check_and_fix_word_format
except ImportError:
    check_and_fix_word_format = None

# רשימת שופטים מקובץ מרכזי
try:
    from judges_config import SUPREME_JUDGES
except ImportError:
    SUPREME_JUDGES = ["עמית", "סולברג", "ברק-ארז", "ברק ארז", "מינץ", "וילנר", "גרוסקופף", "שטיין", "כנפי", "כבוב", "כשר", "רונן", "פוגלמן", "אלרון", "חיות"]

# קובץ לוג לתיקים שנכשלו בולידציה
VALIDATION_FAIL_LOG = os.path.join(os.path.dirname(__file__), "logs", "validation_failures.json")

# ============================================================================
# הגדרות
# ============================================================================

DRY_RUN = False  # ← שנה ל-False בהרצה אמיתית!

SENDER_EMAIL = "your_email@example.com"
LOG_DIR = r"d:\yoel\projects\auto\logs"
SENT_LOG_FILE    = os.path.join(LOG_DIR, "sent_emails.json")
PENDING_LOG_FILE = os.path.join(LOG_DIR, "pending_emails.json")

# Cache to avoid repeated slow Outlook scans in one run
_OUTLOOK_SENT_CACHE = None

os.makedirs(LOG_DIR, exist_ok=True)

# שעות עבודה: ימים א-ה (0=Monday...6=Sunday ב-Python, אבל בישראל א=Sunday=6)
WORK_DAYS = {6, 0, 1, 2, 3}  # Sunday=6, Mon=0, Tue=1, Wed=2, Thu=3
WORK_HOUR_START = 8
WORK_HOUR_END   = 17

# ============================================================================
# 🗺️ מיפוי מחוז ← מייל/קבוצה
# ============================================================================

DISTRICT_CONTACTS = {
    # ירושלים
    "מחוזי_ירושלים":   "district_jerusalem_district@court.gov.il",
    "שלום_ירושלים":    "district_jerusalem_magistrate@court.gov.il",
    "משפחה_ירושלים":   "district_jerusalem_family@court.gov.il",

    # עליון
    "ביהמש_העליון":    "supreme_court@court.gov.il",

    # דרום - באר שבע
    "מחוזי_באר_שבע":   "district_bs_district@court.gov.il",
    "שלום_באר_שבע":    "district_bs_magistrate@court.gov.il",

    # דרום - אחר
    "שלום_אשקלון":     "district_ashkelon@court.gov.il",
    "שלום_קריית_גת":   "district_kiryat_gat@court.gov.il",

    # Pending_Other_Region
    "other_region":    "other_region@example.com",
}

# ============================================================================
# ⏰ בדיקת שעות עבודה
# ============================================================================

def is_business_hours(now: datetime.datetime = None) -> bool:
    """
    מחזיר True אם עכשיו הוא ימי עבודה (א-ה) בין 08:00-17:00.
    """
    if now is None:
        now = datetime.datetime.now()
    weekday = now.weekday()  # 0=Mon, 1=Tue, ..., 6=Sun
    hour = now.hour
    return weekday in WORK_DAYS and WORK_HOUR_START <= hour < WORK_HOUR_END


# ============================================================================
# 📋 תור מיילים (pending)
# ============================================================================

def _load_pending() -> list:
    if not os.path.exists(PENDING_LOG_FILE):
        return []
    try:
        with open(PENDING_LOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_pending(data: list):
    with open(PENDING_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def queue_email(case_num: str, date_str: str, file_path: str, filename: str,
                recipient: str, district: str, word_count: int, judge_name: str = ""):
    """מוסיף מייל לתור לשליחה מאוחרת."""
    pending = _load_pending()

    # בדיקה שלא כבר בתור
    for e in pending:
        if e.get("case_num") == case_num and e.get("date") == date_str:
            print(f"   ℹ️ כבר בתור: {case_num} ({date_str})")
            return

    pending.append({
        "case_num":   case_num,
        "date":       date_str,
        "file_path":  file_path,
        "filename":   filename,
        "recipient":  recipient,
        "district":   district,
        "word_count": word_count,
        "judge_name": judge_name,
        "queued_at":  datetime.datetime.now().isoformat(),
    })
    _save_pending(pending)
    print(f"   [QUEUE] נוסף לתור: {case_num} -> {recipient} (ישלח בשעות עבודה)")


def process_one_queued_email() -> bool:
    """
    שולח מייל אחד מהתור (הכי ותיק) ומוחק אותו מהתור.
    מיועד להרצה כל ~5 דקות ע"י Task Scheduler בשעות העבודה.
    מחזיר True אם היה מה לשלוח.
    """
    if not is_business_hours():
        print("   🕐 מחוץ לשעות עבודה — לא שולח.")
        return False

    pending = _load_pending()
    if not pending:
        print("   V אין מיילים בתור.")
        return False

    # הוסף המתנה רנדומלית קצרה (0-60 שניות) לגיוון תוך ה-5 דקות
    delay = random.randint(0, 60)
    print(f"   Wait {delay} seconds before sending...")
    import time
    time.sleep(delay)

    item = pending.pop(0)  # הכי ותיק קודם
    _save_pending(pending)

    print(f"\n[SENDING] שולח מייל מהתור: {item['case_num']} -> {item['recipient']}")
    _do_send_email(
        file_path=item.get("file_path", ""),
        case_num=item["case_num"],
        date_str=item["date"],
        word_count=item.get("word_count", 0),
        recipient=item["recipient"],
        district=item["district"],
        filename=item["filename"],
        judge_name=item.get("judge_name", ""),
    )
    return True



# ============================================================================
# 📋 לוג ולידציה (תיקים שנחסמו)
# ============================================================================
def _find_transcriber_email(outlook, filename: str, days_back: int = 5):
    """מחפש בתיבת הדואר הנכנס אימייל שהכיל את הקובץ הזה כדי למצוא את המייל של המתמלל"""
    try:
        inbox = outlook.GetDefaultFolder(6) # Inbox
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=days_back)).strftime('%d/%m/%Y %H:%M %p')
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        restricted_items = items.Restrict(f"[ReceivedTime] > '{cutoff}'")
        
        for item in restricted_items:
            if item.Class != 43: continue # Only MailItems
            if item.Attachments.Count > 0:
                for att in item.Attachments:
                    if att.FileName.lower() == filename.lower():
                        sender = item.SenderEmailAddress
                        if sender and ('/O=' in sender.upper() or 'EX' in sender.upper()):
                            try:
                                sender = item.Sender.GetExchangeUser().PrimarySmtpAddress
                            except:
                                pass
                                
                        # הגנת בטיחות: למנוע שליחת התראת שגיאה למזכירויות או משתמשים פנימיים
                        if sender:
                            sender_lower = sender.lower()
                            # חסום דומיינים של בתי משפט או חבר החברות
                            if "court.gov.il" in sender_lower or "hever" in sender_lower or "system" in sender_lower:
                                print(f"   [INFO] מתמלל נפסל למשלוח אוטומטי מטעמי בטיחות דומיין: {sender}")
                                return None, None
                                
                        return item, sender
    except Exception as e:
        print(f"   [WARN] שגיאה בחיפוש מתמלל: {e}")
    return None, None
def send_manual_review_notification(filename: str, case_num: str, date_str: str, reason: str, details: list = None):
    """שולח התראה מיידית ליואל, ואם אפשר - מחזיר משוב ישירות למתמלל."""
    try:
        import win32com.client as win32
        outlook = win32.Dispatch("Outlook.Application")
        
        # נסה למצוא את המייל המקורי של המתמלל לפי שם הקובץ
        orig_mail, transcriber_email = _find_transcriber_email(outlook, filename)
        
        details_html = "".join([f"<li>{d}</li>" for d in (details or [])])
        if not details_html and reason:
            details_html = f"<li>{reason}</li>"
            
        if orig_mail and transcriber_email:
            # שלח תגובה למתמלל
            mail = orig_mail.ReplyAll()
            mail.To = transcriber_email
            mail.CC = "your_email@example.com"
            mail.Subject = f"החזרה לתיקון: התגלו בעיות בקובץ התמלול {filename}"
            
            mail.HTMLBody = f"""
            <div dir="rtl" style="font-family: Arial; text-align: right;">
                <h2 style="color: #d93025;">שלום, הקובץ ששלחת לא עבר את הבדיקה האוטומטית ויש לתקן אותו.</h2>
                <p><b>קובץ:</b> {filename}</p>
                <p><b>תיק זיהוי קוד:</b> {case_num} | <b>תאריך זיהוי:</b> {date_str}</p>
                <p><b>סיבת החזרה:</b> {reason}</p>
                <hr>
                <p>פירוט הבעיות שנמצאו להלן (יש לתקן ולהשיב במייל חוזר עם הקובץ התקין):</p>
                <ul>{details_html}</ul>
                <br>
                <p>תודה,<br>מערכת הבדיקה האוטומטית - דיונים ליואל</p>
            </div>
            """ + mail.HTMLBody # Append original email history
            mail.Send()
            print(f"   [NOTIFY] הערות החזרה נשלחו אוטומטית למתמלל: {transcriber_email}")
            
        else:
            # Fallback לתיבה של יואל בלבד
            mail = outlook.CreateItem(0)
            mail.To = "your_email@example.com"
            mail.Subject = f"[בדיקה ידנית] {filename}"
            
            mail.HTMLBody = f"""
            <div dir="rtl" style="font-family: Arial; text-align: right;">
                <h2 style="color: #d93025;">נדרשת בדיקה ידנית לקובץ (המתמלל לא אותר אוטומטית)</h2>
                <p><b>קובץ:</b> {filename}</p>
                <p><b>תיק:</b> {case_num} | <b>תאריך:</b> {date_str}</p>
                <p><b>סיבה:</b> {reason}</p>
                <hr>
                <ul>{details_html}</ul>
                <p style="font-size: 13px; color: #666;">הקובץ הועבר לתיקיית: Requires_Manual_Review</p>
            </div>
            """
            mail.Send()
            print(f"   [NOTIFY] התראת בדיקה ידנית נשלחה ליואל (מתמלל לא נמצא).")
            
    except Exception as e:
        print(f"   [WARN] שגיאה בשליחת התראת בדיקה ידנית: {e}")

def _log_validation_failure(case_num: str, date_str: str, filename: str, vr: dict):
    """מתעד תיק שנכשל בולידציה לקובץ JSON ושולח התראה ליואל."""
    import json
    log_data = []
    if os.path.exists(VALIDATION_FAIL_LOG):
        try:
            with open(VALIDATION_FAIL_LOG, "r", encoding="utf-8") as f:
                log_data = json.load(f)
        except Exception:
            log_data = []
            
    entry = {
        "case_num":  case_num,
        "date":      date_str,
        "filename":  filename,
        "logged_at": datetime.datetime.now().isoformat(),
        "score":     vr.get("score", 0),
        "issues":    vr.get("issues", []),
        "warnings":  vr.get("warnings", []),
        "tail":      (vr.get("details") or {}).get("tail_preview", ""),
    }
    log_data.append(entry)
    
    try:
        os.makedirs(os.path.dirname(VALIDATION_FAIL_LOG), exist_ok=True)
        with open(VALIDATION_FAIL_LOG, "w", encoding="utf-8") as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"   [ERROR] שגיאה בכתיבה ללוג ולידציה: {e}")

    # ── שליחת התראה למייל [FAILED] ──
    send_manual_review_notification(
        filename=filename,
        case_num=case_num,
        date_str=date_str,
        reason="נכשל בבדיקת איכות תמלול (ולידציה)",
        details=entry["issues"]
    )


# ============================================================================
# 📋 לוג שליחה (מניעת כפילויות)
# ============================================================================

def _load_sent_log() -> list:
    if not os.path.exists(SENT_LOG_FILE):
        return []
    try:
        with open(SENT_LOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_sent_log(log_data: list):
    with open(SENT_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(log_data, f, ensure_ascii=False, indent=2)


def is_already_sent(case_num: str, date_str: str, filename: str = "") -> bool:
    """בודק אם כבר נשלח מייל עבור תיק+תאריך זה ביומן הפנימי או באאוטלוק."""
    log_data = _load_sent_log()
    for entry in log_data:
        # Check case AND date
        if entry.get("case_num") == case_num and str(entry.get("date")).strip() == str(date_str).strip():
            return True
            
    if check_outlook_sent_items(case_num, date_str, filename):
        mark_as_sent(case_num, date_str, "MANUAL_OUTLOOK", "N/A", "N/A")
        return True
    
    return False


def check_outlook_sent_items(case_num: str, date_str: str, filename: str = "") -> bool:
    """
    מתחבר לאאוטלוק (Sent Items) ובודק אם נשלח מייל למספר התיק הזה ובתאריך הזה.
    מחזיר True אם נמצא מייל כזה ב-60 הימים האחרונים.
    """
    try:
        global _OUTLOOK_SENT_CACHE
        if _OUTLOOK_SENT_CACHE is None:
            print("   [OUTLOOK] First call this session: Caching sent items (last 60 days)...")
            _OUTLOOK_SENT_CACHE = []
            
            outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
            sent_folder = outlook.GetDefaultFolder(5) # 5 = Sent Items
            
            # 60 days to be safe
            sixty_days_ago = (datetime.datetime.now() - datetime.timedelta(days=90)).strftime('%d/%m/%Y %H:%M %p')
            items = sent_folder.Items
            items.Sort("[SentOn]", True)
            
            restricted_items = items.Restrict(f"[SentOn] > '{sixty_days_ago}'")
            
            for item in restricted_items:
                subj_lower = str(item.Subject).lower()
                # Exclude internal notification emails (like manual review, errors, etc.)
                IGNORE_TAGS = ["[בדיקה ידנית]", "[error]", "[מייל לא ממופה]", "[warning]", "[sync error]", "failure", "נכשל", "שגיאה"]
                if any(tag in subj_lower for tag in IGNORE_TAGS):
                     continue
                _OUTLOOK_SENT_CACHE.append(subj_lower)
            
            print(f"   [OUTLOOK] Cached {len(_OUTLOOK_SENT_CACHE)} sent subjects.")

        case_clean = str(case_num).strip().lower()
        # תאריך בפורמטים שונים שיופיעו אולי בנושא
        d_raw = date_str.split('-') # Expected YYYY-MM-DD
        d_fmts = []
        if date_str:  # Only add original date string if it is not empty
            d_fmts.append(date_str)
            
        if len(d_raw) == 3:
             y, m, d = d_raw
             
             # Also allow versions without leading zeros:
             m_no_z = str(int(m)) if m.isdigit() else m
             d_no_z = str(int(d)) if d.isdigit() else d
             y_short = y[2:] if len(y)==4 else y

             d_fmts.append(f"{d}.{m}.{y_short}") # 15.02.26
             d_fmts.append(f"{d}/{m}/{y}")   # 15/02/2026
             d_fmts.append(f"{d}/{m}/{y_short}") # 15/02/26

             # No leading zeros
             d_fmts.append(f"{d_no_z}.{m_no_z}.{y_short}")
             d_fmts.append(f"{d_no_z}/{m_no_z}/{y}")
             d_fmts.append(f"{d_no_z}/{m_no_z}/{y_short}")
             
             # Reverse formatting
             d_fmts.append(f"{d_no_z}.{m}.{y_short}")
             d_fmts.append(f"{d}.{m_no_z}.{y_short}")
        
        # Case variation check (support -24 vs -2024)
        case_variations = [case_clean]
        if '-' in case_clean:
            parts = case_clean.split('-')
            if len(parts[-1]) == 2: # 22 -> 2022
                case_variations.append("-".join(parts[:-1]) + "-20" + parts[-1])
            elif len(parts[-1]) == 4: # 2022 -> 22
                case_variations.append("-".join(parts[:-1]) + "-" + parts[-1][2:])

        fname_clean = str(filename or "").lower().strip()
        base_name = os.path.splitext(fname_clean)[0] if fname_clean else ""

        for subject_lower in _OUTLOOK_SENT_CACHE:
            # התעלם מהודעות טסט או מאישורי המערכת הפנימיים (או התראות ליואל)
            ignore_keywords = ['טסט', 'test', 'בדיקה ידנית', 'התראה', 'תקלה', 'failed']
            if any(k in subject_lower for k in ignore_keywords):
                continue

            is_match = any(v in subject_lower for v in case_variations)
            if not is_match and base_name and len(base_name) > 5:
                if base_name in subject_lower:
                    is_match = True

            if is_match:
                 # Verify date if it's a different hearing for the same case
                 if not d_fmts: 
                     print(f"   [OUTLOOK] זוהה שנשלח (ללא אימות תאריך, נושא: {subject_lower[:40]}...)")
                     return True
                     
                 if any(fmt in subject_lower for fmt in d_fmts if fmt):
                     print(f"   [OUTLOOK] זוהה שנשלח (נושא: {subject_lower[:40]}...)")
                     return True
                 else:
                     # Relaxed: if we found the case, but the subject doesn't mention ANY other date, 
                     # we assume it's the one (perhaps the user just didn't type the date).
                     # Only reject if we find a DIFFERENT date pattern in the subject.
                     # To avoid matching the case number itself, we remove case variations from the subject first.
                     sub_for_date = subject_lower
                     for v in case_variations:
                         sub_for_date = sub_for_date.replace(v, "CASE_NUM")
                     
                     # Look for any date pattern remaining in the subject (DD.MM.YY, DD-MM-YY, DD/MM/YY, etc.)
                     other_date = re.search(r'\b\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?\b', sub_for_date)
                     if not other_date:
                         print(f"   [OUTLOOK] זוהה שנשלח (זיהוי מספר תיק בלבד, ללא סתירת תאריך, נושא: {subject_lower[:40]}...)")
                         return True
                     # If there is another date, we do NOT return True. We continue searching other emails.
        return False
    except Exception as e:
        print(f"   [ERROR] Outlook check: {e}")
        return False


def mark_as_sent(case_num: str, date_str: str, recipient: str, filename: str, district: str):
    """מסמן תיק כ'נשלח' בלוג."""
    log_data = _load_sent_log()
    log_data.append({
        "case_num":  case_num,
        "date":      str(date_str),
        "sent_at":   datetime.datetime.now().isoformat(),
        "recipient": recipient,
        "filename":  filename,
        "district":  district,
    })
    _save_sent_log(log_data)
    print(f"   [LOG] נרשם כנשלח: {case_num} ({date_str}) -> {recipient}")


# ============================================================================
# 🔍 זיהוי מחוז
# ============================================================================

def detect_district(text_content: str, filename: str = "") -> str | None:
    """
    מזהה מחוז לפי תוכן הקובץ ו/או שם הקובץ.
    משתמש ב-LLM כשיטה עיקרית, וב-Regex כגיבוי.
    חייב להיות זיהוי ודאי. אם לא זוהה - מחזיר None כדי שיטופל ידנית.
    """
    # הגדלנו ל-4000 כדי לתפוס שמות שעפו רחוק למטה בגלל ריווחים
    combined_text = (text_content or "")[:4000]
    combined_all   = combined_text + " " + (filename or "")

    # ── ערי צפון — בתי משפט שייכים ליובל/יוליה (other_region) ──
    # חשוב: בודקים רק "בית משפט ב[עיר]" — לא כל הופעה של שם עיר!
    # ת"א/רחובות/מרכז לא כלולים — קבצינו לא מגיעים משם, אם העיר מוזכרת זו הפניה בלבד
    NON_HANDLED_CITIES = [
        # צפון — שייך ליובל/יוליה
        "נצרת", "נוף הגליל", "עפולה", "טבריה", "כרמיאל", "עכו", "נהריה",
        "חיפה", "קריית ביאליק", "קריית אתא", "קריית מוצקין",
    ]


    # פטור: קבצי עליון (NS prefix) — הם תמיד עוברים לmarim ולא לסינון לפי עיר
    filename_clean = (filename or "").strip()
    is_supreme_file = filename_clean.upper().startswith("NS-") or "הרכב עמית" in combined_all

    if not is_supreme_file:
        import re as _re
        for city in NON_HANDLED_CITIES:
            court_pattern = _re.compile(
                rf'(בית[\s\-]משפט\s*ב?{_re.escape(city)}'
                rf'|ביהמ["\u05f4]?ש\s+{_re.escape(city)})',
                _re.UNICODE
            )
            if court_pattern.search(combined_all):
                print(f"   [OTHER] בית משפט בעיר לא מטופלת: '{city}' -> other_region (יובל/יוליה)")
                return "other_region"



    # --- ניסיון 1: LLM (הכי חכם) ---
    if detect_district_llm:
        try:
            possible = list(DISTRICT_CONTACTS.keys())
            res = detect_district_llm(combined_text, filename, possible_districts=possible)
            if res and res in DISTRICT_CONTACTS:
                # ── ולידציה גיאוגרפית: ה-AI אמר ירושלים — האם ירושלים באמת מוזכרת? ──
                jerusalem_districts = {"שלום_ירושלים", "מחוזי_ירושלים", "משפחה_ירושלים"}
                south_districts     = {"שלום_באר_שבע", "מחוזי_באר_שבע", "שלום_אשקלון", "שלום_קריית_גת"}
                jeru_keywords = ["ירושלים", "jerusalem", "י-ם", "בית שמש", "beitar"]
                south_keywords = ["באר שבע", "אשקלון", "קריית גת", "קרית גת", "אשדוד", "נתיבות", "דימונה", "אילת"]

                if res in jerusalem_districts:
                    if not any(kw in combined_all.lower() for kw in jeru_keywords):
                        print(f"   [WARN] [AI] החזיר '{res}' אבל ירושלים לא מוזכרת — דוחה ומפנה ל-Review")
                        return None
                elif res in south_districts:
                    if not any(kw in combined_all.lower() for kw in south_keywords):
                        print(f"   [WARN] [AI] החזיר '{res}' אבל עיר דרום לא מוזכרת — דוחה ומפנה ל-Review")
                        return None

                print(f"   [AI] המחוז זוהה בהצלחה: {res}")
                return res
        except Exception as e:
            print(f"   [WARN] שגיאה בשימוש ב-AI לזיהוי מחוז: {e}")

    # --- ניסיון 2: Regex (הגיבוי הקשיח) ---
    combined = combined_all
    combined_lower = combined.lower()

    # ── Pending Other Region (זיהוי לפי תבנית שם קובץ) ──
    if re.search(r'\d{3,}-\d{2}-\d{2,4}_\d{1,2}_\d{1,2}_\d{4}_\d+', filename):
        return "other_region"

    # כדי למנוע זיהוי שגוי (למשל סתם כתובת "קריית גת" בכתב האישום),
    # נחפש הקשר של בית משפט קודם, או צירופים חזקים.

    # ── ביהמ"ש העליון ──
    if "העליון" in combined or "supreme court" in combined_lower:
        return "ביהמש_העליון"
    # גיבוי: "הרכב" + שם שופט עליון מהרשימה המרכזית
    if "הרכב" in combined:
        for judge in SUPREME_JUDGES:
            if judge in combined:
                return "ביהמש_העליון"

    # ── דרום - באר שבע / אשקלון / קריית גת ──
    # נבדוק אם יש אזכור מפורש לבית משפט באזור הדרום
    if "באר שבע" in combined or "beer sheva" in combined_lower or "beersheba" in combined_lower:
        if "מחוזי" in combined:
            return "מחוזי_באר_שבע"
        if "שלום" in combined or "משפחה" in combined:
            return "שלום_באר_שבע"
        
    if "אשקלון" in combined or "ashkelon" in combined_lower:
        if "שלום" in combined or "בית משפט" in combined:
            return "שלום_אשקלון"
        
    if "קריית גת" in combined or "קרית גת" in combined:
        if "בית משפט" in combined or "שלום" in combined:
             return "שלום_קריית_גת"

    if any(kw in combined for kw in ["אשדוד", "נתיבות", "דימונה", "אילת", "ערד"]):
        if "בית משפט" in combined or "שלום" in combined:
            return "שלום_באר_שבע"

    # ── בית המשפט העליון ──
    # זיהוי לפי: שם הרכב ידוע, prefix NS, או אזכור מפורש של "בית המשפט העליון"
    SUPREME_PANELS = [
        "הרכב עמית", "הרכב גרוסקופף", "הרכב הנדל", "הרכב וילנר",
        "הרכב ברק-ארז", "הרכב ברק ארז", "הרכב מזוז", "הרכב סולברג", "הרכב פוגלמן",
        "הרכב קרא", "הרכב שטיין", "הרכב כהן", "הרכב אלרון",
        "הרכב דנציגר", "הרכב חיות", "הרכב רובינשטיין",
    ]
    if (any(panel in combined for panel in SUPREME_PANELS)
            or any(kw in combined for kw in ["בית המשפט העליון", "ביה\"מ העליון", "ביהמ\"ש העליון"])
            or combined.upper().startswith("NS-")):
        return "ביהמש_העליון"

    is_jeru = "ירושלים" in combined or "jerusalem" in combined_lower or "י-ם" in combined
    if is_jeru:
        if "משפחה" in combined or "בית שמש" in combined or "תעבורה" in combined:
            return "משפחה_ירושלים"
        if "מחוזי" in combined:
            return "מחוזי_ירושלים"
        if "שלום" in combined:
            return "שלום_ירושלים"
        # בירושלים אנחנו מחמירים: אם כתוב ירושלים אבל לא כתוב מחוזי/שלום/משפחה - לא מזהים אוטומטית
        # זה מונע מתיק מחוזי שמוזכר בו "ירושלים" בטעות להיחשב כשלום.


    # אם הגענו לפה - לא מצאנו שום עיר ששייכת לבית משפט
    return None


# ============================================================================
# 📬 שליחה בפועל (פונקציה פנימית)
# ============================================================================

def _do_send_email(file_path: str, case_num: str, date_str: str,
                   word_count: int, recipient: str, district: str,
                   filename: str, judge_name: str = ""):
    """שולח מייל בפועל (או מדפיס ב-DRY_RUN)."""
    date_display = date_str
    try:
        d = datetime.date.fromisoformat(date_str)
        date_display = d.strftime("%d/%m/%Y")
    except Exception:
        pass

    if is_already_sent(case_num, date_str):
        print(f"   [SKIP] כבר נשלח: {case_num} ({date_str}). מדלג.")
        return

    # נושא: "12345-01-23 חיים משה מיום 03/02/2026"
    judge_part = f" {judge_name}" if judge_name else ""
    subject = f"{case_num}{judge_part} מיום {date_display}"

    # גוף HTML: RTL ומיושר לימין
    html_body = f"""
    <div dir="rtl" style="font-family: Arial, sans-serif; text-align: right;">
        מספר תיק: {case_num}<br>
        תאריך: {date_display}<br>
        {f'שופט: {judge_name}<br>' if judge_name else ''}
    </div>
    """

    if DRY_RUN:
        print(f"   [DRY RUN] -> To: {recipient}")
        print(f"              נושא: {subject}")
        print(f"              גוף:  {html_body}")
        mark_as_sent(case_num, date_str, recipient, filename, district)
        return

    try:
        # ── בדיקה ותיקון עיצוב קובץ Word לפני שליחה ──
        attach_path = file_path  # ברירת מחדל: הקובץ המקורי
        
        # אם הקובץ לא נמצא בנתיב המקורי, אולי הוא כבר הועבר ל-Pending_Word_Count
        if not os.path.exists(attach_path):
            alt_path = fr"C:\Users\yoel\OneDrive - Hever\Jerusalem\Pending_Word_Count\{filename}"
            if os.path.exists(alt_path):
                attach_path = alt_path
                file_path = alt_path # Update file_path for fixer

        if check_and_fix_word_format and attach_path and os.path.exists(attach_path):
            try:
                # התיקון יבוצע (פונט, מרווחים) אך ללא יישור לשני הצדדים (שבוטל במודול)
                fixed_path, fmt_report = check_and_fix_word_format(
                    file_path=file_path,
                    dry_run=False,
                    check_punctuation=False,
                )
                attach_path = fixed_path
            except Exception as fmt_err:
                print(f"   [WARN] שגיאה בבדיקת עיצוב (ממשיך עם הקובץ המקורי): {fmt_err}")

        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = SENDER_EMAIL
        mail.To      = recipient
        mail.Subject = subject
        mail.HTMLBody = html_body

        if attach_path and os.path.exists(attach_path):
            mail.Attachments.Add(attach_path)
        else:
            print(f"   [ERROR] קובץ לא נמצא לצירוף! השליחה בוטלה: {attach_path}")
            return False

        mail.Send()
        print(f"   [SUCCESS] נשלח בהצלחה ל-{recipient} | {subject}")
        mark_as_sent(case_num, date_str, recipient, filename, district)

        # ── שליחת הודעת אישור ליואל [SENT] ──
        try:
            warning_html = ""
            if "תאריך_תוקן" in filename:
                warning_html = """
                <div style="background-color: #fff3cd; color: #856404; padding: 10px; margin-bottom: 10px; border-radius: 5px; border: 1px solid #ffeeba;">
                    צוות הבוטים זיהה במסמך המקורי טעות תאריך (מרחק קטן או חודש עוקב). <b>התאריך תוקן אוטומטית</b> במסמך ובשם הקובץ לתאריך האמיתי מהאקסל.
                </div>
                """
            
            confirm = outlook.CreateItem(0)
            confirm.To = "your_email@example.com"
            confirm.Subject = f"[נשלח ללקוח] {filename}"
            confirm.HTMLBody = f"""
            <div dir="rtl" style="font-family: Arial; text-align: right;">
                {warning_html}
                <h2 style="color: #1a73e8;">פרוטוקול נשלח בהצלחה</h2>
                <p><b>קובץ:</b> {filename}</p>
                <p><b>תיק:</b> {case_num} | <b>תאריך:</b> {date_str}</p>
                <p><b>נשלח אל:</b> {recipient}</p>
                <p><b>מחוז:</b> {district}</p>
            </div>
            """
            confirm.Send()
        except: pass

        # ── עדכון תאריך שליחה באקסל (P/Q בירושלים, T/V בדרום) ──
        try:
            import gsheets_utils
            JERU_DISTRICTS = {"שלום_ירושלים", "מחוזי_ירושלים", "משפחה_ירושלים", "ביהמש_העליון"}
            region = "jerusalem" if district in JERU_DISTRICTS else "south"
            # מעדכנים עמודת 'שליחה ללקוח' (Q בירושלים / V בדרום)
            gsheets_utils.update_sent_dates(case_num, date_str, sent_to_customer=True, region=region)
        except Exception as ge:
            print(f"   [WARN] שגיאה בעדכון תאריך שליחה באקסל: {ge}")

    except Exception as e:
        print(f"   [ERROR] שגיאה בשליחת מייל: {e}")
        queue_email(case_num, date_str, file_path, filename, recipient, district, word_count)


# ============================================================================
# 📨 Entry Point: שליחה מיידית או הכנסה לתור
# ============================================================================

def send_transcription_email(
    file_path: str,
    case_num: str,
    date_obj,
    word_count: int,
    text_content: str = "",
    filename: str = "",
    forced_district: str = None,
) -> bool:
    """
    Entry Point ראשי: נקרא אחרי עיבוד מוצלח.
    - אם שעות עבודה → שולח מיד (אחד בלבד, שאר לתור)
    - אם מחוץ לשעות → מכניס לתור
    """
    fname = filename or os.path.basename(file_path)
    date_str = str(date_obj) if date_obj else ""

    if is_already_sent(case_num, date_str):
        print(f"   [SKIP] כבר נשלח: {case_num} ({date_str}). מדלג.")
        return True

    district = forced_district or detect_district(text_content, fname)
    if not district:
        msg = f"לא זוהה מחוז עבור הקובץ: {fname}. לא נשלח (מצריך טיפול ידני)."
        print(f"   [WARN] {msg}")
        send_manual_review_notification(fname, case_num, date_str, "לא זוהה מחוז אוטומטית")
        return False

    recipient = DISTRICT_CONTACTS.get(district)
    if not recipient:
        msg = f"אין איש קשר למחוז '{district}'."
        print(f"   [WARN] {msg}")
        send_manual_review_notification(fname, case_num, date_str, f"חסר איש קשר למחוז: {district}")
        return False

    # =============== ולידציה איכות תמלול =====================
    judge_name = ""  # יחולץ מהולידציה אם זמין
    if validate_transcription_quality and text_content:
        print(f"   [INFO] מריץ ולידציה על: {fname}...")
        vr = validate_transcription_quality(
            text_content=text_content,
            expected_case_num=case_num,
            expected_date=date_str,
            filename=fname,
        )
        score_text = "[VALID]" if vr["valid"] else "[INVALID]"
        print(f"   {score_text} ולידציה: ציון={vr['score']}/100 | {len(vr['issues'])} בעיות | {len(vr['warnings'])} אזהרות")
        for issue in vr["issues"]:
            print(f"      [ISSUE] {issue}")
        for warn in vr["warnings"]:
            print(f"      [WARN] {warn}")

        # חילוץ שם שופט מתוצאות הולידציה
        judge_name = vr.get("details", {}).get("header_judge") or ""
        if judge_name:
            print(f"      [JUDGE] שופט: {judge_name}")

        if not vr["valid"]:
            print(f"   [INVALID] לא נשלח — תמלול נכשל בולידציה. מתועד ב-validation_failures.json")
            _log_validation_failure(case_num, date_str, fname, vr)
            return False
    # =========================================================

    # =============== בדיקה באאוטלוק ========================
    now = datetime.datetime.now()
    if check_outlook_sent_items(case_num, date_str):
        print(f"   [DONE] נעצר: תיק {case_num} כבר נשלח מהאאוטלוק שלך. מסמן כטופל.")
        mark_as_sent(case_num, date_str, "Manual/Outlook", fname, district)
        return True
    # =========================================================


    if is_business_hours(now):
        pending = _load_pending()
        if pending:
            print(f"   [QUEUE] יש {len(pending)} בתור — מוסיף לסוף (ישלח בתורו)")
            queue_email(case_num, date_str, file_path, fname, recipient, district, word_count, judge_name)
        else:
            print(f"   [SENDING] שולח מיידית: {case_num} -> {recipient}")
            _do_send_email(file_path, case_num, date_str, word_count, recipient, district, fname, judge_name)
    else:
        day_names = {0:"שני", 1:"שלישי", 2:"רביעי", 3:"חמישי", 4:"שישי", 5:"שבת", 6:"ראשון"}
        print(f"   [LATENCY] מחוץ לשעות עבודה (יום {day_names.get(now.weekday(), '?')} {now.strftime('%H:%M')}) -> נכנס לתור")
        queue_email(case_num, date_str, file_path, fname, recipient, district, word_count, judge_name)

    return True


# ============================================================================
# 📁 סריקת שרת "דיונים ליואל" -> שליחה ליובל דדו + יוליה
# ============================================================================

def _extract_case_from_name(name: str, is_content: bool = False) -> str | None:
    """
    Revised Robust Regex:
    1. 3-part: XXXXX-XX-XX or XXX-XX-XX
    2. 2-part: XXXX-XX (Supreme)
    """
    search_text = name[:1000] if is_content else name
    
    # 1. Supreme Court patterns (Explicit)
    supreme_m = re.search(r'(?:ע"א|בג"ץ|ער"מ|בש"פ|דנ"א|רע"א|רע"פ|עע"מ)\s*(\d{2,5})[/-](\d{2,4})', search_text)
    if supreme_m:
        g1, g2 = supreme_m.groups()
        return f"{g1}-{g2}"

    # Use finditer to check context and prioritize clean matches
    patterns = [
        r'(\d{3,})[-/_](\d{1,2})[-/_](\d{2,4})', # 3-part
        r'(\d{2,5})[/-](\d{2,5})(?![/-]\d)'       # 2-part
    ]
    
    for pat in patterns:
        for match in re.finditer(pat, search_text):
            val = match.group(0)
            start = match.start()
            
            # Exclusion context (ignore if preceded by 'ערעור' in content)
            if is_content:
                context_before = search_text[max(0, start-40):start]
                if "ערעור" in context_before or "פסק דינ" in context_before:
                    continue
            
            parts = re.split(r'[-/_]', val)
            p1 = parts[0]
            # Skip if year
            if len(p1) == 4 and (p1.startswith("19") or p1.startswith("20")):
                continue
                
            if len(parts) == 3:
                return f"{parts[0]}-{int(parts[1]):02d}-{parts[2]}"
            else:
                return f"{parts[0]}-{parts[1]}"
                
    return None


def _extract_date_from_name(name: str) -> datetime.date | None:
    """מחלץ תאריך משם תיקייה. תומך ב: DD.MM.YYYY, DD_MM_YYYY, DD-MM-YYYY."""
    for pat in [
        r'(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{4})',
        r'(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{2})(?!\d)',
    ]:
        m = re.search(pat, name)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if y < 100:
                y += 2000
            try:
                return datetime.date(y, mo, d)
            except ValueError:
                pass
    return None


def scan_and_queue_yoal_server(
    server_path: str = None,
    recipients: str = None,
    lookback_days: int = None,
) -> int:
    """
    סורק את שרת 'דיונים ליואל' (F:\\בתי משפט\\דיונים ליואל),
    מחלץ מספרי תיק מתיקיות/קבצים, ומכניס לתור -> יובל דדו + יוליה.
    
    מבנה גמיש — מסתגל לשלוש שכבות:
      שכבה 1 (שנה/חודש):  2026\\02- פברואר\\  -> ממשיך פנימה
      שכבה 2 (תאריך):     12.01.2026\\         -> ממשיך פנימה
      שכבה 3 (תיק):       12345-01-23 שם\\     -> חולץ מספר תיק
      קבצים:              12345-01-23_...docx  -> חולץ ישירות
    """
    import config_drive_paths as conf

    server_path   = server_path   or conf.YOAL_SERVER_PATH
    recipients    = recipients    or conf.YOAL_RECIPIENTS
    lookback_days = lookback_days or conf.YOAL_LOOKBACK_DAYS
    cutoff        = datetime.date.today() - datetime.timedelta(days=lookback_days)

    if not os.path.exists(server_path):
        print(f"   [WARN] שרת יואל לא נמצא: {server_path}")
        return 0

    print(f"\n[INFO] סורק שרת יואל: {server_path} (חלון {lookback_days} ימים)")
    queued = 0
    processed = set()  # למניעת כפילויות באותה ריצה

    def _try_queue(case_num: str, date_obj: datetime.date | None,
                   fpath: str, fname: str):
        """מנסה להכניס תיק לתור אם עדיין לא נשלח."""
        nonlocal queued
        key = (case_num, str(date_obj))
        if key in processed:
            return
        processed.add(key)
        date_str = str(date_obj) if date_obj else "unknown"
        if is_already_sent(case_num, date_str):
            print(f"   [SKIP] כבר נשלח: {case_num} ({date_str})")
            return
        queue_email(case_num, date_str, fpath, fname, recipients, "other_region", 0)
        queued += 1

    # תיקיות פנימיות שאינן מכילות תמלול — מוחרגות מהסריקה
    EXCLUDED_DIRS = {"רישום", "גיבוי", "מקור", "temp", "_old", "archive", "ארכיון"}

    def _scan_folder(folder: str, inherited_date: datetime.date | None = None, depth: int = 0):
        """רקורסיה מוגבלת (עד עומק 4) — מחפשת תיקים ותיקיות."""
        if depth > 4:
            return
        try:
            items = os.listdir(folder)
        except PermissionError:
            return

        for item in items:
            item_path = os.path.join(folder, item)

            # ── דלג על תיקיות פנימיות מוחרגות ──
            if os.path.isdir(item_path) and item in EXCLUDED_DIRS:
                continue

            # ── קובץ Word/Docx ── נסה ישירות לחלץ מספר תיק ותאריך
            if os.path.isfile(item_path):
                if not item.lower().endswith(('.doc', '.docx', '.odt')) or item.startswith('~$'):
                    continue
                cn = _extract_case_from_name(item)
                if not cn:
                    continue
                dt = _extract_date_from_name(item) or inherited_date
                if dt and dt < cutoff:
                    continue
                _try_queue(cn, dt, item_path, item)
                continue

            # ── תיקייה ──
            if not os.path.isdir(item_path):
                continue

            # בדוק אם שם התיקייה מכיל מספר תיק 
            cn = _extract_case_from_name(item)
            if cn:
                # תיקיית תיק — מחפשים תאריך מהשם או מהתיקייה האב
                dt = _extract_date_from_name(item) or inherited_date
                if dt and dt < cutoff:
                    continue
                # מחפשים קובץ Word בתוך תיקיית התיק
                try:
                    inner_files = [f for f in os.listdir(item_path)
                                   if f.lower().endswith(('.doc', '.docx', '.odt'))
                                   and not f.startswith('~$')]
                    if inner_files:
                        best_file = sorted(inner_files)[-1]  # הכי אחרון
                        _try_queue(cn, dt,
                                   os.path.join(item_path, best_file), best_file)
                    else:
                        # אין קובץ פנימי — שולחים עם נתיב התיקייה
                        _try_queue(cn, dt, item_path, item)
                except Exception:
                    _try_queue(cn, dt, item_path, item)
            else:
                # תיקיית ביניים (שנה / חודש / תאריך) — ממשיכים פנימה
                folder_date = _extract_date_from_name(item) or inherited_date
                # סינון: אם זו תיקיית תאריך ב-cutoff — דלג
                fd = _extract_date_from_name(item)
                if fd and fd < cutoff:
                    continue
                _scan_folder(item_path, folder_date, depth + 1)

    _scan_folder(server_path)

    print(f"   [DONE] שרת יואל: {queued} תיקים נוספו לתור -> {recipients}")
    return queued



# ============================================================================
# 🔎 בדיקת דיונים ישנים שטרם נשלחו
# ============================================================================

def check_unsent_old_hearings() -> list:
    """
    מחזיר רשימת התראות על מיילים בתור שממתינים יותר מיום אחד.
    """
    pending = _load_pending()
    if not pending:
        return []

    warnings = []
    cutoff = datetime.datetime.now() - datetime.timedelta(days=1)
    for item in pending:
        try:
            queued_at = datetime.datetime.fromisoformat(item.get("queued_at", ""))
            if queued_at < cutoff:
                warnings.append(
                    f"[WARN] תיק {item['case_num']} ממתין לשליחה מאז {queued_at.strftime('%d/%m %H:%M')} -> {item['recipient']}"
                )
        except Exception:
            pass

    if warnings:
        print(f"\n[WARN] {len(warnings)} מיילים ממתינים מעל יום:")
        for w in warnings:
            print(f"   {w}")

    return warnings


# ============================================================================
# 🚀 Entry Points חיצוניים
# ============================================================================

def run_for_file(file_path: str, case_num: str, date_obj,
                 word_count: int, text_content: str = "", forced_district: str = None) -> bool:
    """נקרא מ-mail_word_count_prod.py אחרי עדכון מוצלח."""
    return send_transcription_email(
        file_path=file_path,
        case_num=case_num,
        date_obj=date_obj,
        word_count=word_count,
        text_content=text_content,
        filename=os.path.basename(file_path),
        forced_district=forced_district,
    )


def run_pending_other_region(input_folder: str = None):
    """
    Entry point לסריקת דיונים ליואל ושליחה ליובל+יוליה.
    סורק גם את שרת יואל, וגם את תיקיית ההמתנה המקומית ב-OneDrive שאליה fast_email_sender מעביר קבצים.
    """
    import config_drive_paths as conf
    
    # 1. Scan network share (F:\)
    queued_f = scan_and_queue_yoal_server()
    
    # 2. Scan local OneDrive pending folder
    pending_folder = os.path.join(r"C:\Users\yoel\OneDrive - Hever\Jerusalem", "Pending_Other_Region")
    queued_p = 0
    if os.path.exists(pending_folder):
        queued_p = scan_and_queue_yoal_server(
            server_path=pending_folder,
            recipients=conf.YOAL_RECIPIENTS,
            lookback_days=60
        )
        
    # Queue processing to actually send the queued items
    run_queue_processor()
    
    return queued_f + queued_p


def run_yoal_server():
    """Entry point עצמאי לסריקת שרת 'דיונים ליואל' — לשימוש מ-Task Scheduler / nightly job."""
    return scan_and_queue_yoal_server()


def run_queue_processor():
    """
    נקרא מ-Task Scheduler כל 5 דקות בשעות עבודה.
    שולח מייל אחד מהתור.
    """
    print(f"\n[QUEUE] Queue Processor --- {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    pending = _load_pending()
    print(f"   [INFO] {len(pending)} מיילים בתור")
    process_one_queued_email()


# ============================================================================
# 🧪 בדיקת עצמאית
# ============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print(f"DRY_RUN = {DRY_RUN}")
    print(f"שעת הרצה: {datetime.datetime.now().strftime('%A %d/%m/%Y %H:%M')}")
    print(f"שעות עבודה עכשיו: {is_business_hours()}")
    print("=" * 60)

    print("\n[INFO] מיפוי מחוזות:")
    for k, v in DISTRICT_CONTACTS.items():
        print(f"  {k:25s} -> {v}")

    print("\n🔍 בדיקת זיהוי מחוז:")
    tests = [
        ("בית המשפט המחוזי ירושלים", ""),
        ("בית משפט השלום בירושלים", ""),
        ("בית משפט לענייני משפחה ירושלים", ""),
        ("בית המשפט המחוזי באר שבע", ""),
        ("בית משפט השלום באר שבע", ""),
        ("בית משפט השלום אשקלון", ""),
        ("בית משפט השלום קריית גת", ""),
        ("", "24238-07-22_08_2_2026_16553635.docx"),
    ]
    for text, fname in tests:
        d = detect_district(text, fname)
        r = DISTRICT_CONTACTS.get(d, "לא נמצא") if d else "[ERROR] לא זוהה"
        label = (text or fname)[:45]
        print(f"  [{label}] -> {d} -> {r}")

    print("\n📬 סטטוס תור:")
    pending = _load_pending()
    sent = _load_sent_log()
    print(f"  בתור: {len(pending)} | נשלחו: {len(sent)}")

    check_unsent_old_hearings()
