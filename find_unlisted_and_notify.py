"""
[SEARCH] find_unlisted_and_notify.py
==============================
מודול משופר: מוצא תיקים שנמצאים בשרת אך חסרים באקסל,
שולח מייל אישור לפני הכנסה לאקסל.

תכונות מיוחדות:
- חלון סריקה: LOOKBACK_DAYS=7 (כדי לא לפספס)
- חילוץ מספר תיק חכם (לא מבלבל שנה עם סיומת תיק)
- זיהוי תיקיות עליון ללא תיקיית-תיק ← סריקת PDF לחילוץ
- שליחת מייל אישור עם רשימה מסודרת
- פונקציה נפרדת לביצוע אחרי אישור
"""

import os
import re
import datetime
import win32com.client as win32
import json
import config_drive_paths as conf

LOG_FILE = os.path.join(os.path.dirname(__file__), "logs", "unlisted_cases.log")
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log(msg):
    try:
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        safe_msg = str(msg).encode('ascii', errors='replace').decode('ascii')
        formatted_msg = f"{timestamp} - {msg}"
        log(f"{timestamp} - {safe_msg}", flush=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(formatted_msg + "\n")
    except:
        try: log(f"LOG_ERR: {str(msg)[:50]}")
        except: pass


# --- CONFIG ---
LOOKBACK_DAYS = 20  # [SUCCESS] הוגדל ל-20 יום כדי למצוא תיקים ישנים שהתפספסו בעבר
APPROVAL_LOG = r"d:\yoel\projects\auto\logs\pending_approval.txt"
EMAIL_TO = "your_email@example.com"

# ============================================================================
# Case Number Extraction - Smart Version
# ============================================================================
def extract_case_smart(text, context_date=None):
    """
    חילוץ מספר תיק חכם עם "סינון תאריכים" שמרני:
    1. מסיר את תאריך ההקשר (context_date) על כל וריאציותיו.
    2. מסיר תבניות שנראות כמו תאריך עם נקודות (DD.MM.YY) - בתיקים אף פעם אין נקודה.
    3. מחפש מספר תיק ממה שנשאר.
    """
    if not text:
        return None

    s = str(text)
    
    # ─── שלב 1: הסרת תאריך ההקשר (אם ידוע) ───
    if context_date:
        try:
            d, m, y = context_date.day, context_date.month, context_date.year
            yy = str(y)[2:]
            # וריאציות נפוצות: 23.2.26, 23-2-26, 2026-02-23...
            date_variants = [
                f"{d}.{m}.{yy}", f"{d:02d}.{m:02d}.{yy}",
                f"{d}/{m}/{yy}", f"{d:02d}/{m:02d}/{yy}",
                f"{d}-{m}-{yy}", f"{d:02d}-{m:02d}-{yy}",
                f"{d}.{m}.{y}", f"{d:02d}.{m:02d}.{y}",
                f"{y}-{m:02d}-{d:02d}" # YYYY-MM-DD
            ]
            for variant in date_variants:
                s = s.replace(variant, " [CONTEXT_DATE_MASK] ")
        except:
            pass

    # ─── שלב 2: הסרת תאריכים עם נקודות (בטוח, כי בתיק אין נקודה) ───
    s = re.sub(r'\b\d{1,2}\.\d{1,2}\.\d{2,4}\b', ' [DOT_DATE_MASK] ', s)

    # ניקוי מילים מנחות
    clean_text = re.split(r'מיום|בתאריך', s)[0].strip()

    # ─── שלב 3: חילוץ מספר התיק ───
    
    # 1. חיפוש 3 חלקים: digits-digits-digits (שלום/מחוזי)
    m3 = re.search(r'(\d{1,10})[-_](\d{1,2})[-_](\d{1,10})', clean_text)
    if m3:
        p1, p2, p3 = m3.groups()
        if len(p1) <= 2 and len(p3) >= 3:
            return f"{p3}-{int(p2):02d}-{p1}"
        return f"{p1}-{int(p2):02d}-{p3}"

    # 2. חיפוש 2 חלקים: digits-digits (עליון)
    m2 = re.search(r'(\d{1,10})[-_/](\d{1,10})', clean_text)
    if m2:
        p1, p2 = m2.groups()
        if len(p1) <= 2 and len(p2) >= 3:
             return f"{p2}/{p1}"
        elif len(p1) >= 3 and len(p2) <= 2:
             return f"{p1}/{p2}"
        else:
             return f"{p1}/{p2}"

    return None


def normalize_case(c):
    """נרמול: מסיר מקפים לצורך השוואה."""
    return c.replace('-', '').replace('/', '').strip()

def normalize_date(val):
    """ממיר ערך תאריך מאקסל ל-datetime.date."""
    if val is None:
        return None
    if hasattr(val, 'date'):
        return val.date()
    if isinstance(val, datetime.datetime):
        return val.date()
    try:
        s = str(val).strip().split(" ")[0]
        for fmt in ["%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y"]:
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except:
                continue
    except:
        pass
    return None


# ============================================================================
# PDF Text Extraction
# ============================================================================
def extract_text_from_pdf(pdf_path):
    """
    חולץ טקסט מ-PDF.
    מנסה pdfplumber קודם, ואז PyMuPDF.
    מחזיר string ריק אם נכשל.
    """
    text = ""
    # Method 1: pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:  # 3 עמודים ראשונים מספיקים
                t = page.extract_text()
                if t:
                    text += t + "\n"
        if text.strip():
            return text.strip()
    except Exception:
        pass

    # Method 2: PyMuPDF (fitz)
    try:
        import fitz
        doc = fitz.open(pdf_path)
        for page_num in range(min(3, len(doc))):
            t = doc[page_num].get_text()
            if t:
                text += t + "\n"
        doc.close()
        if text.strip():
            return text.strip()
    except Exception:
        pass

    return ""


def extract_judge_from_text(text):
    """
    חולץ שם שופט מטקסט PDF.
    תומך בבתי משפט שלום/מחוזי (לפי תבניות) וגם בעליון (לפי רשימה).
    """
    if not text:
        return ""

    # ─── ניסיון 1: תבניות פורמליות ('לפני: השופט X' / 'כבוד השופט/ת X') ───
    patterns = [
        r'לפני[:\s]+(?:כבוד\s+)?(?:ה)?שופט(?:/ת|ת)?[:\s]+([\u05D0-\u05EA][\u05D0-\u05EA\s\-]{1,25})',
        r'כבוד\s+(?:ה)?שופט(?:/ת|ת)?[:\s]+([\u05D0-\u05EA][\u05D0-\u05EA\s\-]{1,25})',
        r'בפני[:\s]+(?:כבוד\s+)?(?:ה)?שופט(?:/ת|ת)?[\s]+([\u05D0-\u05EA][\u05D0-\u05EA\s\-]{1,25})',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            name = m.group(1).strip().split('\n')[0].strip()
            # נקה מילים מיותרות שעלולות להידבק
            for junk in ['תיק', 'מספר', 'בית', 'משפט', 'תאריך', 'החלטה', 'פסק']:
                name = name.replace(junk, '').strip()
            if 2 <= len(name) <= 20:
                return name

    # ─── ניסיון 2: רשימת שופטי עליון (שמות ספציפיים) ───
    SUPREME_JUDGES = ["עמית", "סולברג", "ברק-ארז", "ברק ארז", "ברק ארץ", "מינץ", "וילנר", "גרוסקופף", "שטיין", "כבוב", "כשר", "רונן", "פוגלמן", "ברון", "אלרון", "חיות"]
    for j in SUPREME_JUDGES:
        if j in text:
            # נרמול "ברק ארז" לפורמט סטנדרטי אם תרצה
            if "ברק" in j and "ארז" in j: return "ברק-ארז"
            return j

    return ""


# ============================================================================
# Scan One Date Folder for Unstructured עליון Cases
# ============================================================================
def scan_elyon_date_folder(date_path, dt_obj):
    """
    סורק תיקיית תאריך של עליון.
    אם אין בה תיקיות-תיק (עם מספר תיק בשם), מחפש PDF וחולץ ממנו.
    מחזיר רשימה של {case, date, path, court, judge, source}
    """
    found = []
    items = os.listdir(date_path)

    # בדיקה: האם יש תיקיות תיק (עם מספר תיק בשם)?
    case_subdirs = []
    for item in items:
        item_path = os.path.join(date_path, item)
        if os.path.isdir(item_path):
            cn = extract_case_smart(item, dt_obj)
            if cn:
                case_subdirs.append((item, cn, item_path))

    if case_subdirs:
        # יש תיקיות תיק רגילות ← מחזיר אותן
        for (name, cn, path) in case_subdirs:
            # ניסיון חילוץ שופט משם התיקייה עצמו (למשל "הרכב ברק ארז")
            judge_from_name = extract_judge_from_text(name)
            
            # ניסיון חילוץ שופט מ-PDF בתוך התיקייה
            judge_from_pdf = ""
            try:
                for f in os.listdir(path):
                    if f.lower().endswith('.pdf') and not f.startswith('~'):
                        pdf_txt = extract_text_from_pdf(os.path.join(path, f))
                        if pdf_txt:
                            judge_from_pdf = extract_judge_from_text(pdf_txt)
                            if judge_from_pdf: break
            except: pass

            found.append({
                'case': cn,
                'date': dt_obj,
                'path': path,
                'court': 'עליון',
                'judge': judge_from_pdf or judge_from_name,
                'source': 'folder_name',
            })
    else:
        # [WARNING] אין תיקיות תיק ← מחפש PDF
        pdf_files = [f for f in items if f.lower().endswith('.pdf') and not f.startswith('~')]
        if not pdf_files:
             # התיקייה ריקה או שאין בה PDF - נדווח עליה כתיקייה לא מזוהה
             found.append({
                'case': "תיקייה לא מזוהה",
                'date': dt_obj,
                'path': date_path,
                'court': 'עליון',
                'judge': '',
                'source': 'folder_name',
                'warning': True,
                'error': f"תיקיית תאריך ללא תיקיות תיק וללא קבצי PDF: {os.path.basename(date_path)}"
            })
             
        for pdf_name in pdf_files:
            pdf_path = os.path.join(date_path, pdf_name)
            log(f"   📄 Scanning PDF for case number: {pdf_name}")
            txt = extract_text_from_pdf(pdf_path)
            cn = extract_case_smart(txt, dt_obj) if txt else None

            # גם ניסה לחלץ מתוך שם ה-PDF עצמו
            if not cn:
                cn = extract_case_smart(pdf_name, dt_obj)

            judge = extract_judge_from_text(txt) if txt else ""

            found.append({
                'case': cn or "תיק לא מזוהה ב-PDF",
                'date': dt_obj,
                'path': pdf_path,
                'court': 'עליון',
                'judge': judge,
                'source': 'pdf',
                'warning': True,  # ← מסמן שאין תיקייה מסודרת
            })
            log(f"   [SUCCESS] Extracted from PDF: case={cn}, judge={judge}")

    return found


# ============================================================================
# Main: Scan Server + Compare to Excel
# ============================================================================
def find_unlisted_cases():
    """
    סורק את השרת לתיקים שחסרים באקסל.
    מחזיר רשימה של תיקים לאישור.
    """
    log(f"[START] מחפש תיקים חסרים (חלון: {LOOKBACK_DAYS} ימים)...")

    # --- Load Excel ---
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except:
        excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    real_path = conf.JERUSALEM_EXCEL_PATH
    wb = None
    target_name = os.path.basename(real_path)
    try:
        for w in excel.Workbooks:
            if w.Name == target_name:
                wb = w
                break
    except: pass
    if not wb:
        try:
            wb = excel.Workbooks.Open(real_path, UpdateLinks=0)
        except Exception as e:
            log(f"[ERROR] לא ניתן לפתוח אקסל: {e}")
            return []

    # Copy to temp
    import tempfile, uuid
    temp_path = os.path.join(tempfile.gettempdir(), f"unlisted_temp_{uuid.uuid4().hex[:6]}.xlsx")
    try:
        wb.SaveCopyAs(temp_path)
    except Exception as e:
        log(f"[ERROR] שגיאה בשמירת עותק: {e}")
        return []

    try:
        wb_t = excel.Workbooks.Open(temp_path)
        ws = wb_t.Sheets("2023")
    except:
        log("[ERROR] לא ניתן לפתוח עותק זמני")
        return []

    # Clear filters
    try:
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
    except:
        pass

    # Read headers
    headers = {}
    r_head = 2
    for r in [1, 2]:
        vals = ws.Range(ws.Cells(r, 1), ws.Cells(r, 30)).Value[0]
        if any(vals):
            r_head = r
            for i, v in enumerate(vals):
                if v:
                    headers[str(v).strip()] = i + 1
            break

    col_case = headers.get("מספר תיק")
    col_date = headers.get("תאריך")
    col_judge_header = headers.get("שם שופט/ת") or headers.get("שופט")
    last_row = ws.Cells(ws.Rows.Count, col_case).End(-4162).Row

    raw_cases = ws.Range(ws.Cells(r_head + 1, col_case), ws.Cells(last_row, col_case)).Value
    raw_dates = ws.Range(ws.Cells(r_head + 1, col_date), ws.Cells(last_row, col_date)).Value
    raw_judges = ws.Range(ws.Cells(r_head + 1, col_judge_header), ws.Cells(last_row, col_judge_header)).Value if col_judge_header else None

    wb_t.Close(False)
    try:
        os.remove(temp_path)
    except:
        pass

    # Build Excel set (case, date)
    excel_set = set()
    excel_judges_map = {}
    for i in range(len(raw_cases)):
        try:
            c = str(raw_cases[i][0]).strip()
            d = normalize_date(raw_dates[i][0]) if raw_dates else None
            j = str(raw_judges[i][0]).strip() if raw_judges and raw_judges[i][0] else ""
            if j == "None": j = ""
            
            if c and c not in ("", "None"):
                excel_set.add((c, d))
                if j: excel_judges_map[c] = j
                
                # נרמולים נוספים
                nc = normalize_case(c)
                excel_set.add((nc, d))
                if j: excel_judges_map[nc] = j
                
                # גרסאות שנה קצרה/ארוכה
                parts = c.split('-')
                if len(parts) == 3:
                    p1, p2, p3 = parts
                    if len(p3) == 4 and p3.startswith("20"):
                        short = f"{p1}-{p2}-{p3[2:]}"
                        excel_set.add((short, d))
                        if j: excel_judges_map[short] = j
                    elif len(p3) == 2:
                        long_y = f"{p1}-{p2}-20{p3}"
                        excel_set.add((long_y, d))
                        if j: excel_judges_map[long_y] = j
        except:
            pass

    log(f"   [STATS] נטענו ~{len(excel_set)} וריאציות (רשומות + נרמולים) מהאקסל.")

    # --- Scan Server ---
    cutoff = datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)

    # [SUCCESS] תיקון חוצה-חודשים: אם חלון ה-lookback חוצה חודש, סרוק גם את החודש הקודם
    today = datetime.date.today()
    roots_to_scan = [conf.get_region_path("Jerusalem", datetime.datetime(today.year, today.month, 1))]
    if cutoff.month != today.month or cutoff.year != today.year:
        prev_month_date = datetime.datetime(cutoff.year, cutoff.month, 1)
        prev_root = conf.get_region_path("Jerusalem", prev_month_date)
        if os.path.exists(prev_root):
            roots_to_scan.append(prev_root)
            log(f"[WARNING] חלון lookback חוצה חודשים — סורק גם: {prev_root}")

    unlisted = []
    KNOWN_COURT_TYPES = ['עליון', 'מחוזי', 'שלום']  # [SUCCESS] רק סוגי בתי משפט ידועים — מונע סריקת 'דיונים'

    for root_j in roots_to_scan:
        if not os.path.exists(root_j):
            log(f"   [WARNING] נתיב לא קיים, מדלג: {root_j}")
            continue
        log(f"📂 סורק שרת: {root_j}")
        all_subdirs = [d for d in os.listdir(root_j) if os.path.isdir(os.path.join(root_j, d))]
        court_types = [d for d in all_subdirs if d in KNOWN_COURT_TYPES]
        skipped = [d for d in all_subdirs if d not in KNOWN_COURT_TYPES]
        if skipped:
            log(f"   ⏭️ מדלג על תיקיות לא-רלוונטיות: {skipped}")

        for court in court_types:
            court_path = os.path.join(root_j, court)
            for date_folder in os.listdir(court_path):
                date_path = os.path.join(court_path, date_folder)
                if not os.path.isdir(date_path):
                    continue

                # Parse date
                dt_obj = None
                try:
                    parts = date_folder.split('.')
                    if len(parts) == 3:
                        y = parts[2]
                        if len(y) == 2:
                            y = "20" + y
                        dt_obj = datetime.date(int(y), int(parts[1]), int(parts[0]))
                except:
                    continue

                if not dt_obj or dt_obj < cutoff:
                    continue

                # עליון ← לוגיקה מיוחדת (אולי אין תיקיית תיק)
                if court == "עליון":
                    elyon_items = scan_elyon_date_folder(date_path, dt_obj)
                    for item in elyon_items:
                        cn = item['case']
                        if "לא מזוהה" in str(cn):
                            unlisted.append(item)
                            continue
                            
                        # גיבוי באקסל לשם השופט
                        if not item.get('judge'):
                            item['judge'] = excel_judges_map.get(cn) or excel_judges_map.get(normalize_case(cn)) or ""
                            
                        # בדיקה אם קיים באקסל
                        if not _case_in_excel(cn, dt_obj, excel_set):
                            unlisted.append(item)
                    continue

                # מחוזי / שלום ← לוגיקה רגילה: סריקת תיקיות תיק
                for case_dir in os.listdir(date_path):
                    case_path = os.path.join(date_path, case_dir)
                    if not os.path.isdir(case_path):
                        continue

                    cn = extract_case_smart(case_dir, dt_obj)
                    is_pdf_extracted = False
                    judge_from_pdf = ""

                    # ─── חיפוש PDF בתוך תיקיית התיק לחילוץ שם שופט ───
                    # (גם אם מספר התיק נמצא בשם התיקייה — עדיין נחלץ שם שופט מה-PDF)
                    try:
                        pdf_files_in_case = [f for f in os.listdir(case_path)
                                             if f.lower().endswith('.pdf') and not f.startswith('~')]
                        for pdf_name in pdf_files_in_case:
                            pdf_path_full = os.path.join(case_path, pdf_name)
                            txt = extract_text_from_pdf(pdf_path_full)
                            if not txt:
                                continue
                            # אם אין עדיין מספר תיק — ננסה לחלץ גם אותו
                            if not cn:
                                cn_from_pdf = extract_case_smart(txt, dt_obj) or extract_case_smart(pdf_name, dt_obj)
                                if cn_from_pdf:
                                    cn = cn_from_pdf
                                    is_pdf_extracted = True
                            # תמיד ננסה לחלץ שם שופט
                            j = extract_judge_from_text(txt)
                            if j:
                                judge_from_pdf = j
                                break  # מצאנו — מספיק PDF אחד
                    except Exception:
                        pass

                    # ניסיון לחלץ תאריך מדויק מתוך שם התיקייה עצמו (אם נכתב יום שונה מתיקיית האב)
                    actual_dt_obj = dt_obj
                    try:
                        clean_dir = case_dir.replace(cn or '', '')
                        m_date = re.search(r'\b(\d{1,2})[\./\-](\d{1,2})[\./\-](\d{2,4})\b', clean_dir)
                        if m_date:
                            d, mth, y = int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3))
                            if y < 100: y += 2000
                            actual_dt_obj = datetime.date(y, mth, d)
                    except:
                        pass

                    if not cn:
                        # [WARNING] תיקייה קיימת בשרת אבל לא הצלחנו לחלץ ממנה מספר תיק
                        unlisted.append({
                            'case': f"תיקייה לא מזוהה: {case_dir}",
                            'date': actual_dt_obj,
                            'path': case_path,
                            'court': court,
                            'judge': judge_from_pdf,
                            'source': 'folder_name',
                            'warning': True,
                            'error': f"לא ניתן לחלץ מספר תיק משם התיקייה: {case_dir}"
                        })
                        continue
                        
                    # ─── חיפוש שם השופט משם התיקייה ───
                    # למשל: "פלקס 14634-12-21 כבוב מיום 25.2.26"
                    if not judge_from_pdf:
                        m_judge = re.search(r'\d{2,4}\s+([א-ת\-\s\'\"]+?)\s+מיום', case_dir)
                        if m_judge:
                            judge_from_pdf = m_judge.group(1).strip()
                        
                    # ─── חיפוש גיבוי באקסל של שם השופט ───
                    # אם לא חולץ שם שופט מה-PDF ולא משם התיקייה, מחפשים את התיק במאגר ההיסטורי של האקסל
                    if not judge_from_pdf and excel_judges_map:
                        judge_from_pdf = excel_judges_map.get(cn) or excel_judges_map.get(normalize_case(cn)) or ""

                    # ניסיון לחלץ תאריך מדויק מתוך שם התיקייה עצמו (אם נכתב יום שונה מתיקיית האב)
                    actual_dt_obj = dt_obj
                    try:
                        clean_dir = case_dir.replace(cn, '')
                        m_date = re.search(r'\b(\d{1,2})[\./\-](\d{1,2})[\./\-](\d{2,4})\b', clean_dir)
                        if m_date:
                            d, mth, y = int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3))
                            if y < 100: y += 2000
                            actual_dt_obj = datetime.date(y, mth, d)
                    except:
                        pass

                    if not _case_in_excel(cn, actual_dt_obj, excel_set):
                        unlisted.append({
                            'case': cn,
                            'date': actual_dt_obj,
                            'path': case_path,
                            'court': court,
                            'judge': judge_from_pdf,
                            'source': 'pdf' if is_pdf_extracted else 'folder_name',
                            'warning': is_pdf_extracted,
                        })

    log(f"\n[SUCCESS] סריקה הושלמה. נמצאו {len(unlisted)} תיקים חסרים.")
    return unlisted


def _case_in_excel(cn, dt_obj, excel_set):
    """בודק אם תיק+תאריך קיים במאגר האקסל (כולל נרמולים)."""
    if (cn, dt_obj) in excel_set:
        return True
    if (cn, None) in excel_set:
        return True
    nc = normalize_case(cn)
    if (nc, dt_obj) in excel_set:
        return True
    # גרסאות שנה קצרה/ארוכה
    parts = cn.split('-')
    if len(parts) == 3:
        p1, p2, p3 = parts
        if len(p3) == 4 and p3.startswith("20"):
            short = f"{p1}-{p2}-{p3[2:]}"
            if (short, dt_obj) in excel_set or (short, None) in excel_set:
                return True
        elif len(p3) == 2:
            long_y = f"{p1}-{p2}-20{p3}"
            if (long_y, dt_obj) in excel_set or (long_y, None) in excel_set:
                return True
    return False


# ============================================================================
# Send Approval Email
# ============================================================================
def _start_server_background():
    """מפעיל את שרת האישור כדי שיישאר באוויר גם אחרי שמשימת הלילה תסתיים."""
    import subprocess
    import sys
    try:
        # Launch independently so it doesn't die when the parent script exits
        script_path = os.path.join(os.path.dirname(__file__), "approval_server.py")
        subprocess.Popen(
            [sys.executable, script_path],
            creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS,
            close_fds=True
        )
        import time; time.sleep(1)  # נותן לשרת שנייה להתחיל
        log(f"🌐 שרת אישור הופעל ברקע (עצמאי)")
        return True
    except Exception as e:
        log(f"[WARNING] לא ניתן להפעיל שרת אישור: {e}")
        return False


def send_approval_email(unlisted_items):
    """
    מפעיל שרת אישור, שולח מייל עם כפתורי אישור/דחייה.
    מחזיר True אם נשלח בהצלחה.
    """
    # השאר תיקים ללא מספר כדי שהמשתמש יוכל לאשר ולהכניס אקסל ידני
    items_to_send = unlisted_items
    skipped_no_case = 0

    if not unlisted_items:
        log("[SUCCESS] אין תיקים חסרים לדיווח (כולם 'לא נמצא').")
        return True

    log(f"📧 שולח מייל אישור עם {len(unlisted_items)} תיקים...")

    # הפעל שרת אישור *לפני* שליחת המייל
    _start_server_background()

    # בנה URL לאישור
    try:
        import approval_server as _srv
        approve_url = _srv.build_approval_url()
        reject_url = _srv.build_reject_url()
    except:
        approve_url = f"http://localhost:5050/approve"
        reject_url = f"http://localhost:5050/reject"
        
    # --- אופציונלי: שלח עותק של התיקים החסרים גם ל-N8N (לפי בקשת המשתמש) ---
    try:
        import n8n_utils
        payload = {
            "event": "pending_approval",
            "approve_url": approve_url,
            "reject_url": reject_url,
            "items_count": len(unlisted_items),
            "items": [
                {
                    "case": i.get('case'),
                    "date": str(i.get('date')),
                    "court": i.get('court'),
                    "source": i.get('source')
                } for i in unlisted_items
            ]
        }
        n8n_utils.send_report(payload)
    except Exception as e:
        log(f"[WARNING] הערה: לא ניתן לשלוח מידע ל-N8N - {e}")

    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_TO
        mail.Subject = f"[WARNING] תיקים חסרים באקסל - {datetime.date.today().strftime('%d/%m/%Y')} - ממתינים לאישור"

        # Build table rows
        import urllib.parse
        rows_html = ""
        for i, item in enumerate(unlisted_items, 1):
            warn_icon = "[WARNING]" if item.get('warning') else "🔎"
            source_label = "PDF (חלץ מטקסט)" if item.get('source') == 'pdf' else "תיקיית שרת"
            judge_str = item.get('judge', '') or '-'
            c_enc = urllib.parse.quote(item['case'])
            app_one_url = f"{approve_url}&case={c_enc}"
            rej_one_url = f"{reject_url}&case={c_enc}"
            
            rows_html += f"""
            <tr style="background: {'#fff9e6' if item.get('warning') else 'white'}">
                <td style="padding:8px; border:1px solid #ddd; text-align:center;">{i}</td>
                <td style="padding:8px; border:1px solid #ddd; font-weight:bold; color:#c0392b;">{warn_icon} {item['case']}</td>
                <td style="padding:8px; border:1px solid #ddd;">{item['date'].strftime('%d/%m/%Y') if item['date'] else '?'}</td>
                <td style="padding:8px; border:1px solid #ddd;">{item['court']}</td>
                <td style="padding:8px; border:1px solid #ddd;">{judge_str}</td>
                <td style="padding:8px; border:1px solid #ddd; font-size:11px; color:#666;">
                    {source_label}<br>
                    <a href="{app_one_url}" style="color:green; text-decoration:none; font-weight:bold; display:inline-block; margin-top:5px;">[SUCCESS] אשר לבד</a> | 
                    <a href="{rej_one_url}" style="color:red; text-decoration:none; display:inline-block; margin-top:5px;">[ERROR] דחה</a>
                </td>
            </tr>"""

        # Save pending list to file (for execute step)
        os.makedirs(os.path.dirname(APPROVAL_LOG), exist_ok=True)
        with open(APPROVAL_LOG, "w", encoding="utf-8") as f:
            import json
            data_to_save = []
            for i, item in enumerate(unlisted_items):
                data_to_save.append({
                    'case': item['case'],
                    'date': str(item['date']),
                    'court': item.get('court', ''),
                    'judge': item.get('judge', ''),
                    'confidential': item.get('confidential', ''),
                    'status': item.get('status', ''),
                    'path': item.get('path', ''),
                    'source': item.get('source', ''),
                })
            json.dump(data_to_save, f, ensure_ascii=False, indent=2)
        log(f"📝 שמור לקובץ אישור: {APPROVAL_LOG}")

        html = f"""
        <html dir="rtl">
        <body style="font-family: 'Segoe UI', sans-serif; background:#f4f6f9; padding:20px;">
        <div style="max-width:800px; margin:0 auto; background:white; padding:25px;
                    border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,0.1);">

            <div style="text-align:center; border-bottom:3px solid #e67e22; padding-bottom:15px; margin-bottom:20px;">
                <h2 style="color:#c0392b; margin:0;">[WARNING] תיקים חסרים באקסל</h2>
                <p style="color:#7f8c8d; margin:5px 0;">{datetime.date.today().strftime('%d/%m/%Y')} - ממתין לאישורך</p>
            </div>

            <div style="background:#fff3cd; border:1px solid #ffc107; border-radius:8px;
                        padding:12px; margin-bottom:20px;">
                <strong>📋 נמצאו {len(unlisted_items)} תיקים בשרת שאינם מופיעים באקסל.</strong><br>
                לאחר עיון ברשימה, אנא כתוב בצ'אט: <strong>"אשר תיקים"</strong> כדי להוסיפם לאקסל.<br>
                אם יש תיקים שאין להוסיפם, פרט אחרי אישורך אילו להחריג.
            </div>

            <table style="width:100%; border-collapse:collapse; font-size:14px;">
                <thead>
                    <tr style="background:#2c3e50; color:white;">
                        <th style="padding:10px; border:1px solid #ddd;">#</th>
                        <th style="padding:10px; border:1px solid #ddd;">מספר תיק</th>
                        <th style="padding:10px; border:1px solid #ddd;">תאריך</th>
                        <th style="padding:10px; border:1px solid #ddd;">בית משפט</th>
                        <th style="padding:10px; border:1px solid #ddd;">שופט</th>
                        <th style="padding:10px; border:1px solid #ddd;">מקור</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>

            <!-- APPROVE / REJECT BUTTONS -->
            <div style="text-align:center; margin:25px 0;">
                <a href="{approve_url}"
                   style="display:inline-block; background:#27ae60; color:white; padding:14px 35px;
                          border-radius:8px; font-size:18px; font-weight:bold; text-decoration:none;
                          margin-left:15px; box-shadow:0 3px 8px rgba(39,174,96,.4);">[SUCCESS] אשר והוסף לאקסל</a>
                <a href="{reject_url}"
                   style="display:inline-block; background:#e74c3c; color:white; padding:14px 25px;
                          border-radius:8px; font-size:16px; text-decoration:none;
                          box-shadow:0 3px 8px rgba(231,76,60,.4);"> דחה</a>
            </div>

            <div style="background:#d5f5e3; border:1px solid #27ae60; border-radius:8px;
                        padding:12px; margin-top:10px; font-size:12px; color:#555;">
                💡 לחיצה על הכפתור תפתח דף אישור בדפדפן ותוסיף את התיקים לאקסל אוטומטית.<br>
                הכפתור פעיל למשך שעה אחת מהשליחה.<br>
                שורות מסומנות בצהוב ([WARNING]) = אין להן תיקייה מסודרת בשרת — כדאי לטפל בנפרד.
            </div>

            <p style="text-align:center; font-size:11px; color:#bdc3c7; margin-top:25px;">
                הודעה זו נוצרה אוטומטית ע״י AutoBot
            </p>
        </div>
        </body></html>
        """

        mail.HTMLBody = html
        mail.Send()
        log(f"[SUCCESS] מייל אישור נשלח אל {EMAIL_TO}")
        return True

    except Exception as e:
        log(f"[ERROR] שגיאה בשליחת מייל: {e}")
        return False


# ============================================================================
# Execute: Add Approved Cases to Excel
# ============================================================================
import gsheets_utils

def _get_confidentiality(pdf_path):
    """
    Attempts to detect if a PDF is confidential. 
    Uses logic from orders_processor or gsheets_utils if available.
    """
    try:
        import fitz
        with fitz.open(str(pdf_path)) as doc:
            if doc.page_count < 2: return "לא חסוי"
            # Basic textual check for "חסוי" or "בדלתיים סגורות"
            text = doc.load_page(0).get_text()
            if "חסוי" in text or "דלתיים סגורות" in text:
                return "חסוי"
    except: pass
    return "לא חסוי"

def execute_add_to_excel(approved_items=None):
    """
    מוסיף תיקים מאושרים לאקסל.
    אם approved_items=None ← קורא מקובץ ה-pending_approval.txt.
    """
    import json

    if approved_items is None:
        if not os.path.exists(APPROVAL_LOG):
            log("[ERROR] לא נמצא קובץ אישור. הרץ קודם find_and_notify().")
            return
        with open(APPROVAL_LOG, "r", encoding="utf-8") as f:
            approved_items = json.load(f)

    if not approved_items:
        log("[SUCCESS] אין תיקים להוסיף.")
        return

    log(f"📝 מוסיף {len(approved_items)} תיקים לאקסל...")

    try:
        excel = win32.GetActiveObject("Excel.Application")
    except:
        excel = win32.Dispatch("Excel.Application")
    # מסיר את excel.Visible = True כי זה זורק שגיאה
    excel.DisplayAlerts = False

    real_path = conf.JERUSALEM_EXCEL_PATH
    wb = None
    target_name = os.path.basename(real_path)
    try:
        for w in excel.Workbooks:
            if w.Name == target_name:
                wb = w
                break
    except: pass
    
    try:
        if not wb:
            wb = excel.Workbooks.Open(real_path, UpdateLinks=0)
        ws = wb.Sheets("2023")
    except Exception as e:
        log(f"[ERROR] לא ניתן לפתוח אקסל: {e}")
        return

    # Read headers
    headers = {}
    r_head = 2
    for r in [1, 2]:
        vals = ws.Range(ws.Cells(r, 1), ws.Cells(r, 40)).Value[0]
        if any(vals):
            r_head = r
            for i, v in enumerate(vals):
                if v:
                    headers[str(v).strip()] = i + 1
            break

    col_case = headers.get("מספר תיק")
    col_date = headers.get("תאריך")
    col_status = headers.get("האם בוטל?") or headers.get("סטטוס")
    col_court = headers.get("בית משפט") or headers.get("ביהמ\"ש")
    col_city = headers.get("עיר") or headers.get("מחוז")
    col_judge = headers.get("שם שופט/ת") or headers.get("שופט")
    col_type = headers.get("סוג") or headers.get("רגיל/דחוף") or headers.get("דחיפות")
    col_secret = headers.get("חסוי") or headers.get("חסוי?")
    col_words = headers.get("מספר מילים") or headers.get("מילים למתמלל")

    # [SUCCESS] תיקון קריטי: מציאת הטבלה לפי מיקום (לא לפי שם "Table1" שאולי שגוי)
    # זה מבטיח שהשורות נוספות *בתוך* הטבלה ויופיעו בסינון
    has_table = False
    list_obj = None
    try:
        if ws.ListObjects.Count > 0:
            list_obj = ws.ListObjects(1)  # הטבלה הראשונה בגיליון, לא משנה שמה
            has_table = True
            log(f"   📋 נמצאה טבלה: '{list_obj.Name}' — שורות יתווספו בתוכה")
        else:
            last_row = ws.Cells(ws.Rows.Count, col_case).End(-4162).Row
            row_idx = last_row + 1
            log(f"   [WARNING] אין טבלה מוגדרת — מוסיף אחרי שורה {last_row}")
    except Exception as e:
        log(f"   [WARNING] שגיאה בזיהוי טבלה: {e} — מוסיף לסוף")
        last_row = ws.Cells(ws.Rows.Count, col_case).End(-4162).Row
        row_idx = last_row + 1

    added = 0
    for item in approved_items:
        try:
            cn = item['case']
            # מאפשר להמשיך גם אם 'לא נמצא' כדי שהמשתמש יוכל להשלים באקסל

            if has_table:
                new_row = list_obj.ListRows.Add()
                row_idx = new_row.Range.Row
            
            d_str = item.get('date', '')
            d_obj = None
            try:
                d_obj = datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
            except:
                pass

            court = item.get('court', '')

            # ALWAYS copy format from row above to ensure perfect styling
            try:
                # Copy entire row (formats + values) to new row
                ws.Rows(row_idx - 1).Copy(ws.Rows(row_idx))
            except:
                pass

            # Clear content first
            try:
                ws.Rows(row_idx).ClearContents()
            except:
                pass

            # Write values
            if col_case:
                ws.Cells(row_idx, col_case).Value = cn
            if col_date and d_obj:
                # [SUCCESS] OLE Automation date: מספר ימים מאז 30/12/1899
                # זו הדרך הנכונה — מספר שלם, אין בעיית UTC/timezone, אקסל יציג תאריך אמיתי
                ole_date = (d_obj - datetime.date(1899, 12, 30)).days
                ws.Cells(row_idx, col_date).Value = ole_date
                ws.Cells(row_idx, col_date).NumberFormat = "DD/MM/YYYY"

            # Auto-fill based on court type
            if col_court:
                if court == "עליון":
                    ws.Cells(row_idx, col_court).Value = 'ביהמ"ש העליון'
                elif court == "מחוזי":
                    ws.Cells(row_idx, col_court).Value = 'ירושלים מחוזי'
                elif court == "שלום":
                    ws.Cells(row_idx, col_court).Value = 'ירושלים שלום'

            if col_city:
                # If we have hints that it's South, use appropriate city or stay empty
                if "באר שבע" in court or "דרום" in court:
                     ws.Cells(row_idx, col_city).Value = "" # Leave for manual or detect
                else:
                     ws.Cells(row_idx, col_city).Value = "ירושלים"
            
            if col_secret:
                # Try to detect confidentiality from PDF (only write once — no second override below!)
                pdf_path = item.get('path')
                status_secret = "לא חסוי"
                if pdf_path and os.path.isdir(str(pdf_path)):
                    try:
                        pdfs = [f for f in os.listdir(pdf_path) if f.lower().endswith('.pdf')]
                        if pdfs:
                            status_secret = _get_confidentiality(os.path.join(pdf_path, pdfs[0]))
                    except Exception:
                        pass
                elif pdf_path and str(pdf_path).lower().endswith('.pdf'):
                    status_secret = _get_confidentiality(pdf_path)
                ws.Cells(row_idx, col_secret).Value = status_secret

            if col_judge and item.get('judge'):
                ws.Cells(row_idx, col_judge).Value = item['judge']
            if col_status:
                ws.Cells(row_idx, col_status).Value = "לא בוטל"
            if col_type:
                ws.Cells(row_idx, col_type).Value = "רגיל"
            # [SUCCESS] תוקן: הוסר כתיבה כפולה ל-col_secret שהחליפה את תוצאת בדיקת ה-PDF

            # Highlight in light green
            ws.Rows(row_idx).Interior.Color = 13434828

            log(f"   [SUCCESS] נוסף שורה {row_idx}: {cn} ({d_str})")
            added += 1
            if not has_table:
                row_idx += 1

        except Exception as e:
            log(f"   [ERROR] שגיאה בהוספת {item.get('case', '?')}: {e}")

    try:
        wb.Save()
        log(f"\n[SUCCESS] נשמר! נוספו {added} שורות לאקסל.")
    except Exception as e:
        log(f"[ERROR] שגיאה בשמירה: {e}")

    # Remove approval log after execution
    try:
        os.remove(APPROVAL_LOG)
    except:
        pass


# ============================================================================
# Main Entry Points
# ============================================================================
def find_and_notify():
    """
    Entry point ראשי:
    1. מוצא תיקים חסרים
    2. שולח מייל לאישור
    """
    items = find_unlisted_cases()
    if items:
        send_approval_email(items)
    else:
        log("[SUCCESS] לא נמצאו תיקים חסרים. הכל מסונכרן!")
    return items


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "execute":
        log("🔥 מצב ביצוע: מוסיף תיקים מאושרים לאקסל...")
        execute_add_to_excel()
    else:
        log("[SEARCH] מצב סריקה: מחפש תיקים חסרים ושולח מייל...")
        find_and_notify()
