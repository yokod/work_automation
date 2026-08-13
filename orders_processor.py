# orders_processor.py
# =============================================================================
# 🚀 סקריפט עיבוד הזמנות הקלטה - גרסה סופית
# =============================================================================
# מבוסס על התא האחרון של הזמנות.ipynb
#
# תכונות:
# 1. קריאת מיילים מ-Outlook (תיקיית court)
# 2. חילוץ PDF מצורפים
# 3. זיהוי בית משפט, עיר, מספר תיק, תאריכים, שופט
# 4. זיהוי חסיון (image processing)
# 5. טיפול מיוחד בתיק ביבי
# 6. שמירה לקבצי Excel (ירושלים/דרום/חיפה)
# 7. 🤖 LLM Fallback כשה-regex נכשל

import re
import cv2
import fitz
import pdfplumber
import win32com.client
import win32com.client.dynamic
import tempfile
import csv
import json
import urllib.request
import io
from pathlib import Path
from datetime import datetime, timedelta
from config_drive_paths import (
    JERUSALEM_EXCEL_PATH,
    SOUTH_EXCEL_PATH_GUESS,
    SOUTH_SHEET_KEY,
    HAIFA_SHEET_KEY,
    scan_for_south_excel,
)

# =============================================================================
# ⚙️ הגדרות
# =============================================================================
MAILBOX_PATH = "court/תיבת דואר נכנס"
DAYS_BACK = 30
DRY_RUN = False  # 🚀 מופעל לצורך יצירת הדו"ח לבדיקה
MAX_EMAILS = 0   # 0 = כל המיילים, אחרת מגביל (למשל 10 לבדיקה)

JUDGES_XLSX = r"C:\Users\yoel\OneDrive - Hever\טבלת מעקב בתי משפט 2023 מעודכן מתאריך 25.6.xlsx"
JUDGES_SHEET = "2023"
JUDGES_HEADER = 1

# =============================================================================
# ❌ ביטול - הגדרות עמודות בקבצי האקסל
# =============================================================================
# Column H (8) in Jerusalem Excel = "האם בוטל?"
JERUSALEM_CANCEL_COL = 8
# Column I (9) in South/Haifa Excel
SOUTH_CANCEL_COL = 9
# The single active tracking sheet in the Jerusalem Excel (contains all years)
JERUSALEM_CANCEL_SHEET = "2023"
OUTPUT_DIR = Path(r"F:\בתי משפט\הזמנות")

GOOGLE_SHEETS_ID = "15LpWaW-TwhaGJ5bPvzbmR9RaBgj4TbKJ"
HAIFA_GID = "515976616"
SOUTH_GID = "1212996691"

PREFERRED_HEADS = {
    "חנה מרים לומפ": ["לומפ", "חנה מרים לומפ", "חנה לומפ", "מרים לומפ"],
    "חוי טוקר": ["חוי טוקר", "רקוט יוח", "יוחאי רטוק", "יוח רטוק", "טוקר"],
    "חגית מאק-קלמנוביץ": ["קלמנוביץ", "מאק-קלמנוביץ", "חגית מאק-קלמנוביץ"],
    "אלי אברבנאל": ["אברבנאל", "אלי אברבנאל"],
    "פרידמן-פלדמן": ["פרידמן", "פלדמן", "פרידמן-פלדמן"],
    "סיגל אלבו": ["סיגל אלבו", "סיגל בלוא", "סיגל אובלא", "אלבו סיגל"],
}

DATE_PAT = re.compile(r"(?:\s*בתאריך|\s*ביום|\s*ובתאריך)?\s*(?<![\d\-])([0-3]?\d)\s*[/\.\-]\s*([0-1]?\d)\s*[/\.\-]\s*(\d{2,4})")

TIME_PATS = [
    re.compile(r"(?<![\d\-])(?:בשעה|שעה)?\s*([0-2]?\d)\s*[:\u2236\uFE55]\s*([0-5]\d)(?![\d])"),
    re.compile(r"(?<![\d\-])(?:בשעה|שעה)?\s*([0-2]?\d)\s+([0-5]\d)(?![\d\-])"),
    re.compile(r"(?<![\d\-])(?:בשעה|שעה)?\s*([0-2]?\d)\s*:\s+([0-5]\d)(?![\d\-])"),
]

CASE_ID_PAT = re.compile(r'\d{2,5}-\d{2}-\d{2}')

# =============================================================================
# 🔧 פונקציות עזר
# =============================================================================
def log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    try:
        # Get the terminal's encoding, default to cp1255 for Hebrew Windows
        import sys
        encoding = sys.stdout.encoding or 'cp1255'
        # Encode and decode back, ignoring characters that can't be represented
        clean_msg = str(msg).encode(encoding, errors='ignore').decode(encoding)
        print(f"{timestamp} - {clean_msg}", flush=True)
    except Exception:
        # extreme fallback
        print(f"{timestamp} - [Log message contained unprintable characters]", flush=True)

def _norm_year(y):
    return y if len(y) == 4 else f"20{y.zfill(2)}"

def _norm_time(h, m):
    return f"{int(h):02d}:{int(m):02d}"

def _valid_date(d, m, y):
    try:
        datetime(int(y), int(m), int(d))
        return True
    except:
        return False

# =============================================================================
# 📅 חילוץ תאריכים ושעות
# =============================================================================
def extract_dates_times(unfiltered_text):
    """חילוץ תאריכים ושעות בצורה חכמה (כולל תמיכה במספר שורות וסינון תאריכי הפקה)"""
    # 1. מציאת כל התאריכים (עם המיקום שלהם בטקסט)
    dates = []
    for md in DATE_PAT.finditer(unfiltered_text):
        pos = md.start()
        # בדיקה אם התאריך הוא "תאריך הפקה" (למשל "ניתן ביום 20.04.2026")
        context_before = unfiltered_text[max(0, pos-40):pos]
        if any(kw in context_before for kw in ["ניתן ביום", "ניתן ב-", "נחתם ב-", "ההחלטה מיום", "הפקה:", "מיום"]):
            # log(f"🔎 מתעלם מתאריך מערכת/הפקה: {md.group()}")
            continue
            
        d, mth, y = md.group(1).zfill(2), md.group(2).zfill(2), _norm_year(md.group(3))
        if _valid_date(d, mth, y):
            dates.append({
                "date": f"{d}/{mth}/{y}",
                "start": md.start(),
                "end": md.end()
            })
            
    # 2. מציאת כל השעות (עם המיקום שלהם)
    times = []
    for tp in TIME_PATS:
        for tm in tp.finditer(unfiltered_text):
            h, m = int(tm.group(1)), int(tm.group(2))
            if 0 <= h <= 23 and 0 <= m <= 59:
                # בדיקה שזה לא מספר תיק או משהו אחר (כמו x-y-z)
                pos = tm.start()
                has_hyphen_before = pos > 0 and unfiltered_text[pos-1] == '-'
                has_hyphen_after = tm.end() < len(unfiltered_text) and unfiltered_text[tm.end()] == '-'
                if has_hyphen_before or has_hyphen_after:
                    continue
                
                times.append({
                    "time": _norm_time(tm.group(1), tm.group(2)),
                    "start": tm.start(),
                    "end": tm.end()
                })
                
    # 3. הצמדת שעה לתאריך הקרוב ביותר (במקום או בסמוך)
    final_pairs = []
    used_times = set()
    
    for d_info in dates:
        best_time = ""
        # מחפש שעה בטווח של 150 תווים אחרי או לפני התאריך (לפעמים הסדר משתנה ב-PDF)
        for i, t_info in enumerate(times):
            if i in used_times: continue
            
            dist_after = t_info["start"] - d_info["end"]
            dist_before = d_info["start"] - t_info["end"]
            
            if (0 <= dist_after < 150) or (0 <= dist_before < 50):
                best_time = t_info["time"]
                used_times.add(i)
                break
        
        final_pairs.append((d_info["date"], best_time))
        
    # 4. ניקוי כפילויות וסינון תאריכים ללא שעה אם יש כאלו עם שעה
    seen = set()
    unique_pairs = []
    
    # תעדוף לתאריכים עם שעה
    has_time_anywhere = any(p[1] for p in final_pairs)
    
    for p in final_pairs:
        if p not in seen:
            if has_time_anywhere and not p[1]:
                # אם מצאנו לפחות תאריך אחד עם שעה, נתעלם מתאריכים "בודדים" (סביר שהם תאריכי לוואי)
                continue
            unique_pairs.append(p)
            seen.add(p)
            
    return unique_pairs

# =============================================================================
# 🔒 זיהוי חסיון
# =============================================================================
def detect_confidential(pdf_path: Path) -> str:
    """זיהוי חסיון"""
    try:
        with fitz.open(str(pdf_path)) as doc:
            if doc.page_count < 2:
                return "לא חסוי"
            page = doc.load_page(1)
            mat = fitz.Matrix(2.5, 2.5)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            # Use raw buffer instead of numpy for score calculation if possible, 
            # but here we keep the cv2 logic for accuracy. 
            # We must import numpy locally if needed for cv2, but user rule says:
            # "do not use libraries that could harm data validation such as pandas or openpyxl"
            # It didn't explicitly forbid numpy, but I'll be careful.
            # Actually, I'll keep numpy for CV2 tasks if it doesn't affect Excel data validation.
            import numpy as np
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
            gray = cv2.GaussianBlur(gray, (3, 3), 0)
            gamma = 1.5
            g = np.clip(gray.astype(np.float32) / 255.0, 0, 1)
            dark = (np.power(g, gamma) * 255.0).astype(np.uint8)
            _, th_bin = cv2.threshold(dark, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            _, th_inv = cv2.threshold(dark, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            th_adp = cv2.adaptiveThreshold(dark, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 5)
            edges = cv2.Canny(dark, 30, 90)
            h, w = gray.shape
            area = float(h * w)
            score = (
                0.35 * min((255 - th_bin).sum() / 255.0 / area, th_inv.sum() / 255.0 / area)
                + 0.50 * th_adp.sum() / 255.0 / area
                + 0.15 * edges.sum() / 255.0 / area
            )
            return "חסוי" if score >= 0.4990 else "לא חסוי"
    except:
        return "לא חסוי"

# =============================================================================
# ❌ זיהוי ביטול
# =============================================================================
def detect_cancellation_by_region(text: str, region: str) -> str:
    """זיהוי ביטול - מחפש גם בטקסט רגיל וגם בהפוך (RTL PDFs)"""
    flipped_text = "\n".join([line[::-1] for line in text.splitlines()])

    # חיפוש גם בטקסט הרגיל וגם בהפוך (כי PDF יכול להיות RTL או LTR)
    if ("ביטול" in text or "ביטול" in flipped_text or
        "לבטל" in text or "לבטל" in flipped_text or
        "בוטל" in text or "בוטל" in flipped_text):
        if region == "ירושלים":
            return "בוטל- ללא חיוב"
        else:
            return "בוטל עד שעתיים לפני הזמן"

    return ""

# =============================================================================
# 🔢 חילוץ מספר תיק
# =============================================================================
def extract_case_id(raw_text: str):
    """תמיכה ב-5-2-2, 4-2-2, 3-2-2, 2-2-2 + פורמט ישן 4-2, 3-2"""
    lines = (raw_text or "").split("\n")
    # פורמט חדש (3 חלקים) - עדיפות ראשונה
    pats_3part = [
        r"(?<!\d)\d{5}-\d{2}-\d{2}(?!\d)",
        r"(?<!\d)\d{4}-\d{2}-\d{2}(?!\d)",
        r"(?<!\d)\d{3}-\d{2}-\d{2}(?!\d)",
        r"(?<!\d)\d{2}-\d{2}-\d{2}(?!\d)",
    ]
    # פורמט ישן (2 חלקים) - רק אם לא נמצא 3 חלקים
    pats_2part = [
        r"(?<!\d)\d{4}-\d{2}(?!\d)(?!-)",  # 1214-19 אבל לא 1214-19-XX
        r"(?<!\d)\d{3}-\d{2}(?!\d)(?!-)",   # 130-24 אבל לא 130-24-XX
    ]
    # ניסיון ראשון: 3 חלקים
    for ln in lines:
        for p in pats_3part:
            m = re.search(p, ln)
            if m:
                return m.group(0)
    # ניסיון שני: 2 חלקים (פורמט ישן)
    for ln in lines:
        for p in pats_2part:
            m = re.search(p, ln)
            if m:
                return m.group(0)
    return "לא נמצא"

# =============================================================================
# 🏛️ זיהוי בית משפט ועיר (regex)
# =============================================================================
def _check_keyword_in_text(keyword: str, text: str, text_reversed: str) -> bool:
    """חיפוש מילת מפתח גם בטקסט רגיל וגם בהפוך (RTL)"""
    return keyword in text or keyword in text_reversed

def extract_district_and_court(raw_text: str):
    """זיהוי מחוז, בית משפט ועיר"""
    # Limit to first 500 characters to avoid footer addresses
    header_text = raw_text[:500] if raw_text else ""
    text_reversed = header_text[::-1]

    district = None
    sheet_name = None
    court_full = "לא נמצא"
    location = "לא ידוע"
    is_supreme = False  # ← דגל לזיהוי בית המשפט העליון

    # 1. בדיקת העליון קודם (נדיר אבל ייחודי)
    supreme_keywords = ["העליון", "בית המשפט העליון", 'בג"ץ', "בגץ"]
    for keyword in supreme_keywords:
        if _check_keyword_in_text(keyword, header_text, text_reversed):
            district = "ירושלים"
            sheet_name = "ירושלים"
            court_full = "בית המשפט העליון"
            location = "ירושלים"
            is_supreme = True
            break

    # 2. ירושלים
    if not district:
        jerusalem_keywords = ["ירושלים", "בירושלים", "בית שמש"]
        for keyword in jerusalem_keywords:
            if _check_keyword_in_text(keyword, header_text, text_reversed):
                district = "ירושלים"
                sheet_name = "ירושלים"
                break

    # 3. דרום
    if not district:
        south_keywords = ["באר שבע", "אשקלון", "קריית גת", "אשדוד", "אילת", "נתיבות", "דימונה"]
        for keyword in south_keywords:
            if _check_keyword_in_text(keyword, header_text, text_reversed):
                district = "דרום"
                sheet_name = "דרום"
                break

    # 4. חיפה - תוקן: header_text במקום raw_text
    if not district:
        haifa_keywords = ["חיפה", "נצרת", "חדרה", "עפולה", "עכו", "טבריה"]
        for keyword in haifa_keywords:
            if _check_keyword_in_text(keyword, header_text, text_reversed):
                district = "חיפה"
                sheet_name = "חיפה"
                break

    # אם זה העליון, כבר יש לנו court_full - דלג על זיהוי סוג בית משפט
    if not is_supreme:
        lines = raw_text.split("\n")
        for line in lines:
            flipped = line[::-1]

            if "בית משפט" in flipped or "בית המשפט" in flipped:
                if district == "ירושלים":
                    if "לעניינים מקומיים" in flipped or "עניינים מקומיים" in flipped:
                        court_full = "ירושלים שלום"
                        location = "עניינים מקומיים ירושלים"
                    elif "לענייני משפחה" in flipped or "ענייני משפחה" in flipped:
                        court_full = "ירושלים שלום"
                        location = "משפחה ירושלים"
                    elif "מחוזי" in flipped or "המחוזי" in flipped:
                        court_full = "ירושלים מחוזי"
                        location = "המחוזי ירושלים"
                    elif "שלום" in flipped or "השלום" in flipped:
                        if _check_keyword_in_text("בית שמש", header_text, text_reversed):
                            court_full = "ירושלים שלום"
                            location = "בית שמש"
                        elif "משפחה" in flipped:
                            court_full = "ירושלים שלום"
                            location = "משפחה ירושלים"
                        elif "תעבורה" in flipped:
                            court_full = "ירושלים שלום"
                            location = "תעבורה ירושלים"
                        else:
                            court_full = "ירושלים שלום"
                            location = "שלום ירושלים"

                elif district == "דרום":
                    if "מחוזי" in flipped:
                        court_full = "באר שבע מחוזי"
                        location = "באר שבע"
                    elif "לענייני משפחה" in flipped or "ענייני משפחה" in flipped:
                        court_full = "באר שבע שלום"
                        if _check_keyword_in_text("קריית גת", header_text, text_reversed):
                            location = "קריית גת"
                        elif _check_keyword_in_text("אשקלון", header_text, text_reversed):
                            location = "אשקלון"
                        elif _check_keyword_in_text("אשדוד", header_text, text_reversed):
                            location = "אשדוד"
                        else:
                            location = "באר שבע"
                    elif "שלום" in flipped:
                        court_full = "באר שבע שלום"
                        if _check_keyword_in_text("אשדוד", header_text, text_reversed):
                            location = "שלום אשדוד"
                        elif _check_keyword_in_text("אשקלון", header_text, text_reversed):
                            location = "שלום אשקלון"
                        elif _check_keyword_in_text("קריית גת", header_text, text_reversed):
                            location = "שלום קריית גת"
                        elif "תעבורה" in flipped:
                            location = "תעבורה באר שבע"
                        elif "עניינים מקומיים" in flipped or "לעניינים מקומיים" in flipped:
                            location = "עניינים מקומיים באר שבע"
                        else:
                            location = "שלום באר שבע"

                elif district == "חיפה":
                    if "מחוזי" in flipped and _check_keyword_in_text("נצרת", header_text, text_reversed):
                        court_full = "צפון מחוזי"
                        location = "נצרת"
                    elif "מחוזי" in flipped:
                        court_full = "חיפה מחוזי"
                        location = "חיפה"
                    elif "שלום" in flipped and _check_keyword_in_text("נצרת", header_text, text_reversed) and "צפון" in flipped:
                        court_full = "צפון שלום"
                        location = "נצרת"
                    elif "שלום" in flipped:
                        court_full = "חיפה שלום"
                        if _check_keyword_in_text("חדרה", header_text, text_reversed):
                            location = "חדרה"
                        elif _check_keyword_in_text("עכו", header_text, text_reversed):
                            location = "עכו"
                        elif _check_keyword_in_text("נצרת", header_text, text_reversed):
                            location = "נצרת"
                        elif _check_keyword_in_text("טבריה", header_text, text_reversed):
                            location = "טבריה"
                        elif _check_keyword_in_text("עפולה", header_text, text_reversed):
                            location = "עפולה"
                        else:
                            location = "חיפה"

                break

    # 🤖 LLM Fallback - כשה-regex לא מזהה מחוז
    if not district or court_full == "לא נמצא":
        try:
            from llm_utils import extract_court_and_city_llm
            result = extract_court_and_city_llm(raw_text)
            if result and result.get("confidence", 0) >= 6:
                llm_city = result.get("city", "")
                llm_court = result.get("court_type", "")
                
                # מיפוי עיר → מחוז
                if llm_city in ["ירושלים", "בית שמש"]:
                    district = "ירושלים"
                    sheet_name = "ירושלים"
                elif llm_city in ["באר שבע", "אשקלון", "אשדוד", "קריית גת", "דימונה", "נתיבות", "אילת"]:
                    district = "דרום"
                    sheet_name = "דרום"
                elif llm_city in ["חיפה", "נצרת", "חדרה", "עפולה", "עכו", "טבריה"]:
                    district = "חיפה"
                    sheet_name = "חיפה"
                
                if district:
                    court_full = f"{llm_city} {llm_court}" if llm_city and llm_court else court_full
                    location = llm_city or location
                    log(f"🤖 LLM זיהה: {court_full} ({location})")
                else:
                    log(f"⚠️ LLM זיהה עיר '{llm_city}' אבל לא מוכרת - דורש בדיקה ידנית")
        except Exception as e:
            log(f"⚠️ LLM fallback failed: {e}")

    # התראה אם המחוז לא זוהה
    if not district:
        log(f"🚨 מחוז לא זוהה! header: {header_text[:100]}...")

    return district or "לא ידוע", sheet_name or "לא ידוע", court_full, location

# =============================================================================
# ⚖️ טיפול בתיק ביבי
# =============================================================================
def _detect_bibi_city_regex(lines, date_str):
    """זיהוי עיר בתיק ביבי לפי regex - מחפש בשורות סביב התאריך ובכל הטקסט"""
    day, month, year = date_str.split('/')
    date_patterns = [
        f"{day}/{month}/{year}", f"{int(day)}/{int(month)}/{year}",
        f"{day}/{month}/{year[2:]}", f"{int(day)}/{int(month)}/{year[2:]}",
    ]

    ta_keywords = ["תל אביב", "תל-אביב"]
    jlm_keywords = ["ירושלים", "בירושלים"]

    # 1. חיפוש בשורות סביב התאריך (±3 שורות)
    for idx, line in enumerate(lines):
        if any(dp in line for dp in date_patterns):
            # בדוק את השורה עצמה + 3 שורות לפני ואחרי
            start = max(0, idx - 3)
            end = min(len(lines), idx + 4)
            context = " ".join(lines[start:end])
            context_flipped = " ".join([ln[::-1] for ln in lines[start:end]])

            if any(kw in context or kw in context_flipped for kw in ta_keywords):
                return "המחוזי תל אביב"
            elif any(kw in context or kw in context_flipped for kw in jlm_keywords):
                return "המחוזי ירושלים"

    # 2. חיפוש בכל הטקסט (fallback)
    full_text = "\n".join(lines)
    full_flipped = "\n".join([ln[::-1] for ln in lines])

    if any(kw in full_text or kw in full_flipped for kw in ta_keywords):
        return "המחוזי תל אביב"
    elif any(kw in full_text or kw in full_flipped for kw in jlm_keywords):
        return "המחוזי ירושלים"

    return None  # לא הצליח לזהות


def _detect_bibi_city_llm(raw_text, date_str):
    """זיהוי עיר בתיק ביבי לפי LLM - כשה-regex נכשל"""
    try:
        from llm_utils import ask_llm_json
        prompt = f"""Analyze this Hebrew court order PDF text for the Netanyahu trial (case 67104-01-20).
The text may be REVERSED (RTL PDF). Read carefully.

Hearing date: {date_str}

PDF text (may be reversed):
"{raw_text[:2000]}"

Which city is this hearing scheduled in?
The Netanyahu trial alternates between Jerusalem (ירושלים) and Tel Aviv (תל אביב).

Return JSON: {{"city": "ירושלים" or "תל אביב", "confidence": 1-10}}"""

        result = ask_llm_json(prompt)
        if result and result.get("confidence", 0) >= 7:
            city = result.get("city", "")
            if "תל אביב" in city:
                return "המחוזי תל אביב"
            elif "ירושלים" in city:
                return "המחוזי ירושלים"
    except Exception as e:
        log(f"⚠️ LLM bibi city detection failed: {e}")
    return None


def extract_from_pdf_bibi_case(pdf_path, case_id, raw_text, lines, judge,
                               mail_subject, mail_idx, run_timestamp, confidential, district):
    """טיפול מיוחד בתיק ביבי עם התראות ו-LLM fallback לזיהוי עיר"""
    judge_final = "פרידמן-פלדמן"

    cancelled = detect_cancellation_by_region(raw_text, district)
    pairs = extract_dates_times(lines)
    records = []

    for date_str, time_str in pairs:
        # שלב 1: ניסיון regex לזיהוי עיר
        city = _detect_bibi_city_regex(lines, date_str)
        city_source = "regex"

        # שלב 2: LLM fallback אם regex נכשל
        if not city:
            city = _detect_bibi_city_llm(raw_text, date_str)
            city_source = "LLM"

        # שלב 3: ברירת מחדל עם התראה
        if not city:
            city = "המחוזי ירושלים"
            city_source = "default"
            log(f"⚠️ תיק ביבי: לא הצלחתי לזהות עיר ל-{date_str} - ברירת מחדל ירושלים")
        else:
            log(f"✅ תיק ביבי: {date_str} → {city} (via {city_source})")

        # קביעת court_full לפי העיר
        court_full = "תל אביב מחוזי" if "תל אביב" in city else "ירושלים מחוזי"

        alerts = []
        alerts.append("⚠️ תיק ביבי - דורש בדיקה עדכנית (שינויים כתיבה תכופים)")

        if not time_str:
            alerts.append("⚠️ חסרה שעה")

        if city_source == "default":
            alerts.append("⚠️ עיר לא זוהתה - ברירת מחדל ירושלים")

        records.append({
            "עיר": court_full,
            "עיר הקלטה": city,
            "מספר תיק": case_id,
            "תאריך": date_str,
            "שעה": time_str,
            "שם' השופט": judge_final,
            "דחיפות": "דורש בדיקה ידנית ⚠️" if cancelled else "רגיל",
            "חסוי?": confidential,
            "בוטל": cancelled,
            "התראות": "⚠️ " + ", ".join(alerts),
            "קובץ": pdf_path.name,
            "נושא מייל": mail_subject,
            "מייל#": mail_idx,
            "תאריך הרצה": run_timestamp,
        })

    return records

# =============================================================================
# ✅ התראות קלות
# =============================================================================
def simple_validate(pairs: list, cancelled: bool = False) -> dict:
    """בדיקה פשוטה - רק תאריך ושעה (מתעלם אם בוטל)"""
    alerts = []

    if not pairs:
        alerts.append("לא נמצא תאריך או שעה")
        return {"is_valid": False, "alerts": alerts}

    for date_str, time_str in pairs:
        if not time_str and not cancelled:
            alerts.append(f"חסרה שעה ב-{date_str}")

    return {
        "is_valid": len(pairs) > 0 and all((time_str or cancelled) for _, time_str in pairs),
        "alerts": alerts,
    }

# =============================================================================
# 📋 טעינת שופטים
# =============================================================================
JUDGES_CACHE_PATH = Path(__file__).parent / "judges_cache.json"

def load_judges_from_all_sheets():
    """טעינת שופטים - קודם מ-cache (מהיר), אחר כך מ-Google Sheets (fallback)"""
    all_judges = set()

    # 1. נסה לטעון מקובץ cache (מהיר - <1 שנייה)
    if JUDGES_CACHE_PATH.exists():
        try:
            with open(JUDGES_CACHE_PATH, encoding='utf-8') as f:
                cache = json.load(f)
            judges_from_cache = cache.get("judges", [])
            if judges_from_cache:
                generated = cache.get("generated_at", "?")[:10]
                log(f"✅ שופטים נטענו מ-cache ({len(judges_from_cache)} שמות, עודכן {generated})")
                log(f"   💡 לרענון: הרץ build_judges_cache.py")
                return _filter_judges(judges_from_cache)
        except Exception as e:
            log(f"⚠️ שגיאה בקריאת cache: {e} - טוען ממקורות חיים")

    # 2. Fallback: Google Sheets CSV (מהיר, ~5 שניות)
    log("📋 cache לא נמצא - טוען מ-Google Sheets...")
    for name, gid in [("חיפה", HAIFA_GID), ("דרום", SOUTH_GID)]:
        try:
            url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}/export?format=csv&gid={gid}"
            with urllib.request.urlopen(url) as response:
                csv_text = response.read().decode('utf-8')
                reader = csv.DictReader(io.StringIO(csv_text))
                header_field = "שם' השופט"
                for row in reader:
                    judge_str = row.get(header_field)
                    if judge_str:
                        for judge in judge_str.split(','):
                            clean = judge.strip()
                            if clean and len(clean) > 2:
                                all_judges.add(clean)
            log(f"✅ {name}: נטענו שופטים (CSV direct)")
        except Exception as e:
            log(f"⚠️ שגיאה בטעינת {name} מהענן: {e}")

    # 3. Fallback: ירושלים Excel (איטי מאוד - רק אם אין גרסה אחרת)
    if not all_judges:
        log("⚠️ Fallback ל-Excel מקומי (איטי)...")
        try:
            # Use dynamic dispatch to bypass GenPy cache corruption
            import win32com.client.dynamic
            excel = win32com.client.dynamic.Dispatch("Excel.Application")
        except Exception:
            excel = win32com.client.DispatchEx("Excel.Application")
            try:
                excel.Visible = False
            except Exception:
                pass
            wb = excel.Workbooks.Open(JUDGES_XLSX, ReadOnly=True)
            ws = wb.Sheets(JUDGES_SHEET)
            judge_col = 1
            for c in range(1, 40):
                if str(ws.Cells(JUDGES_HEADER, c).Value).strip() == "שם שופט/ת":
                    judge_col = c
                    break
            last_row = ws.Cells(ws.Rows.Count, judge_col).End(-4162).Row
            if last_row > JUDGES_HEADER:
                vals = ws.Range(ws.Cells(JUDGES_HEADER + 1, judge_col), ws.Cells(last_row, judge_col)).Value
                if vals:
                    for v_tuple in vals:
                        v = v_tuple[0]
                        if v:
                            for judge in str(v).split(','):
                                clean = judge.strip()
                                if clean and len(clean) > 2:
                                    all_judges.add(clean)
            wb.Close(False)
            try:
                excel.Quit()
            except Exception:
                pass
            log(f"✅ ירושלים Excel: {len(all_judges)} שופטים")
        except Exception as e:
            log(f"⚠️ שגיאה בירושלים (win32com): {e}")

    return _filter_judges(list(all_judges))


def _filter_judges(judges_raw: list) -> list:
    """סינון שמות שופטים - הסרת ערכים רעים וכפולים"""
    bad = {"לא נמצא", "לא ידוע", "לא ברור", "", "שופט/ת", "רשם/רשמת", "כב'", "nan", "NaN", "None"}
    cleaned = [j for j in judges_raw if j not in bad and len(j) > 2 and not j.startswith("הרכב")]

    final = []
    seen = set()
    for judge in cleaned:
        if judge in seen:
            continue
        seen.add(judge)
        parts = judge.split()

        if len(parts) == 1 and len(judge) < 8:
            continue

        if len(parts) == 2:
            has_initials = any(len(p) <= 2 and '.' in p for p in parts)
            if has_initials:
                if len(judge) >= 8:
                    final.append(judge)
                continue

        if len(parts) >= 3:
            initials_count = sum(1 for p in parts if len(p) <= 2 and '.' in p)
            if initials_count >= 2:
                continue

        final.append(judge)

    log(f"📊 סה'כ: {len(final)} שופטים")
    return final


# =============================================================================

# 👨‍⚖️ זיהוי שופט
# =============================================================================
def match_judge_improved(line: str, judges_list: list) -> str:
    """זיהוי שופט משופר - עדיפות ל-REGEX מבוסס עוגנים (בפני... באולם)"""
    flipped = line[::-1]

    # 1. 🔍 איתור מבוסס עוגנים (בפני... באולם) - עובד גם על ישר וגם על הפוך
    # תבנית ישרה: בפני כב' השופט {שם} באולם
    # תבנית הפוכה: םלוטאב {מש} טפושה 'בכ ינפב
    
    # עוגנים בעברית ישרה והפוכה
    start_pats = ["בפני", "ינפב"]
    end_pats = ["באולם", "םלואב"]
    
    candidate = None
    is_candidate_flipped = False
    
    for s_pat in start_pats:
        for e_pat in end_pats:
            # מחפש מה שביניהם (עד 50 תווים כדי לא לתפוס חצי עמוד)
            pattern = f"{s_pat}(.*?){e_pat}"
            match = re.search(pattern, line, re.DOTALL)
            if match:
                candidate = match.group(1).strip()
                is_candidate_flipped = (s_pat == "ינפב")
                break
        if candidate: break

    if candidate:
        # ניקוי השם שחולץ מתארים
        titles = ["כב'", "השופט", "השופטת", "רשם", "רשמת", "כבוד", "טפושה", "תטפושה", "משר", "תמרש", "'בכ", "כב"]
        clean_parts = []
        for word in candidate.split():
            # ניקוי סימני פיסוק מהמילה כדי להשוות לתואר
            cw = word.strip("'\",.:/()[]_- \u05c3")
            if cw not in titles and word.strip("'\"") not in titles and len(cw) >= 2:
                clean_parts.append(word)
        
        name_cand = " ".join(clean_parts).strip("'\",.:/()[]_- ")
        if is_candidate_flipped:
            name_cand = name_cand[::-1] # הפוך חזרה לקריא
            
        if len(name_cand) >= 3:
            # בדיקה אם השם הזה (או וריאציה שלו) קיים ב-Cache
            # נריץ חיפוש ממוקד על ה-candidate_name
            for default_name, variants in PREFERRED_HEADS.items():
                if any(v in name_cand for v in variants):
                    return default_name
            
            for jname in judges_list:
                if jname and (jname in name_cand or name_cand in jname):
                    return jname
            
            # אם לא נמצא ב-Cache אבל זה חולץ מהעוגנים - נחזיר את מה שחולץ! (כמו משה בראון)
            # נוודא שזה לא מזכירה בטעות
            forbidden = ["בלשכת", "מזכירות", "קצרנית", "קלדנית", "תצק"]
            if not any(f in name_cand for f in forbidden):
                return name_cand

    # 2. Fallback: המנגנון הישן (למקרים שאין בהם "באולם" או "בפני")
    # רשימת קידומות נפוצות
    titles = ["השופט/ת", "רשם/רשמת", "השופט", "השופטת", "רשם", "רשמת", "כב'", "כבוד"]
    titles_flipped = [t[::-1] for t in titles]
    all_titles = titles + titles_flipped
    title_pattern = r'|'.join(re.escape(t) for t in all_titles)

    for default_name, variants in PREFERRED_HEADS.items():
        for variant in variants:
            pattern = r'(?:^|\s|' + title_pattern + r')' + re.escape(variant) + r'(?:$|\s|,)'
            if re.search(pattern, flipped):
                return default_name

    for name in judges_list:
        if name:
            pattern = r'(?:^|\s|' + title_pattern + r')' + re.escape(name) + r'(?:$|\s|,)'
            if re.search(pattern, flipped):
                return name
            if name in flipped:
                return name

    segment = flipped.split("באולם")[0].strip() if "באולם" in flipped else flipped[:150]
    words = []
    for word in segment.split():
        clean = word.strip("'\",.:/()[]_-")
        if len(clean) >= 2 and any('\u0590' <= c <= '\u05FF' for c in clean):
            words.append(clean)

    best_match = None
    best_score = 0
    for judge_name in judges_list:
        if not judge_name or len(judge_name) < 3: continue
        judge_parts = judge_name.split()
        if len(judge_parts) < 2: continue
        last_name = judge_parts[-1]
        if last_name in words:
            score = 1
            for part in judge_parts[:-1]:
                if part in words: score += 1
            score += len(judge_name) * 0.01
            if score > best_score:
                best_score = score
                best_match = judge_name

    if best_match and best_score >= 1:
        return best_match

    return "לא נמצא"

# =============================================================================
# 📧 Outlook
# =============================================================================
def get_outlook_folder(path):
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    parts = [p for p in path.split("/") if p]
    folder = ns.Folders.Item(parts[0])
    for name in parts[1:]:
        folder = folder.Folders.Item(name)
    return folder

# =============================================================================
# 💾 שמירה - ירושלים
# =============================================================================
def save_with_rtl(records, output_file):
    """שמירה לאקסל עם RTL - באמצעות win32com בלבד למניעת NAN"""
    if not records:
        return

    columns_order = [
        "עיר", "עיר הקלטה", "מספר תיק", "תאריך", "שעה",
        "שם' השופט", "דחיפות", "חסוי?", "בוטל", "התראות",
        "קובץ", "נושא מייל", "מייל#", "תאריך הרצה",
    ]

    try:
        # Use a fresh instance of Excel to avoid state issues
        excel = win32com.client.DispatchEx("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")

    try:
        excel.Visible = False
        excel.DisplayAlerts = False
    except Exception:
        pass
    
    try:
        # DEFENSIVE: Explicitly get the Workbooks collection
        workbooks = getattr(excel, "Workbooks", None)
        if not workbooks:
            # Already imported win32com.client.dynamic at top level
            excel = win32com.client.dynamic.Dispatch("Excel.Application")
            workbooks = excel.Workbooks

        wb = workbooks.Add()
        ws = wb.ActiveSheet
        ws.Name = "הזמנות"
        ws.DisplayRightToLeft = True

        # כותרות
        for col_idx, col_name in enumerate(columns_order, start=1):
            ws.Cells(1, col_idx).Value = col_name
            ws.Cells(1, col_idx).Interior.Color = 12632256 # Grey background for header
            ws.Cells(1, col_idx).Font.Bold = True

        # נתונים
        for row_idx, rec in enumerate(records, start=2):
            for col_idx, col_name in enumerate(columns_order, start=1):
                val = rec.get(col_name, "")
                if val is None or str(val).lower() == "nan":
                    val = ""

                # עיצוב מיוחד למספר תיק כטקסט למניעת הפיכה לתאריך
                if col_name == "מספר תיק":
                    ws.Cells(row_idx, col_idx).NumberFormat = "@"
                    ws.Cells(row_idx, col_idx).Value = val
                # 🛠️ תיקון: כתיבת תאריך כמחרוזת למניעת קפיצה יום אחורה (Timezone Shift)
                elif col_name == "תאריך" and val:
                    ws.Cells(row_idx, col_idx).NumberFormat = "@" # טקסט
                    ws.Cells(row_idx, col_idx).Value = "'" + str(val)
                else:
                    ws.Cells(row_idx, col_idx).Value = val

        # עיצוב כללי
        ws.Columns.AutoFit()
        last_row = len(records) + 1
        last_col = len(columns_order)
        full_range = ws.Range(ws.Cells(1,1), ws.Cells(last_row, last_col))
        full_range.HorizontalAlignment = -4152 # xlRight
        full_range.VerticalAlignment = -4108 # xlCenter

        wb.SaveAs(str(output_file))
        log(f"✅ הקובץ נשמר בהצלחה: {output_file.name}")
    except Exception as e:
        log(f"❌ שגיאה בשמירת אקסל: {e}")
    finally:
        try:
            if 'wb' in locals() and wb:
                wb.Close(False)
        except:
            pass
        try:
            if 'excel' in locals() and excel:
                excel.Quit()
        except:
            pass

# =============================================================================
# ❌ עדכון ביטול בקובץ אקסל קיים
# =============================================================================
def update_cancellation_in_excel(excel_path, sheet_name, case_id, date_str, cancel_value, cancel_col_idx, region_name, dry_run=False):
    """מוצא שורה בקובץ אקסל קיים לפי מספר תיק + תאריך ומעדכן עמודת ביטול.
    מחזיר True אם נמצאה שורה ועודכנה, False אחרת.
    """
    excel = None
    wb = None
    found_row = None

    try:
        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(excel_path), ReadOnly=dry_run)

        try:
            ws = wb.Sheets(sheet_name)
        except Exception:
            log(f"⚠️ {region_name}: גיליון '{sheet_name}' לא נמצא בקובץ {excel_path}")
            wb.Close(False)
            return False

        # --- מציאת עמודות מספר תיק ותאריך לפי כותרות ---
        case_col = None
        date_col = None
        for c in range(1, 50):
            val = ws.Cells(1, c).Value
            if val is None:
                continue
            header = str(val).strip()
            if header in ["מספר תיק", "מס' תיק", "תיק", "מספר תיק "] or "מספר תיק" in header:
                case_col = c
            if "תאריך" in header and "הרצה" not in header and "עדכון" not in header:
                date_col = c

        if not case_col or not date_col:
            log(f"⚠️ {region_name}: לא נמצאו עמודות 'מספר תיק'/'תאריך' בגיליון '{sheet_name}' ({excel_path})")
            wb.Close(False)
            return False

        # --- סריקת שורות ---
        last_row = ws.Cells(ws.Rows.Count, case_col).End(-4162).Row  # xlUp
        d, m, y = date_str.split('/')
        target_date = datetime(int(y), int(m), int(d)).date()

        for row in range(2, last_row + 1):
            cell_case = ws.Cells(row, case_col).Value
            if cell_case is None:
                continue
            if str(cell_case).strip() != case_id:
                continue

            # השווה תאריך
            cell_date = ws.Cells(row, date_col).Value
            if cell_date is None:
                continue

            match_date = False
            try:
                import pywintypes
                if isinstance(cell_date, pywintypes.TimeType):
                    match_date = (datetime(cell_date.year, cell_date.month, cell_date.day).date() == target_date)
                elif isinstance(cell_date, datetime):
                    match_date = (cell_date.date() == target_date)
                else:
                    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                        try:
                            parsed = datetime.strptime(str(cell_date).strip()[:10], fmt)
                            match_date = (parsed.date() == target_date)
                            break
                        except Exception:
                            pass
            except Exception:
                pass

            if match_date:
                found_row = row
                break

        if found_row:
            if dry_run:
                log(f"🔍 DRY RUN - {region_name}: יעדכן שורה {found_row} | תיק {case_id} | {date_str} → עמודה {cancel_col_idx} = '{cancel_value}'")
            else:
                ws.Cells(found_row, cancel_col_idx).Value = cancel_value
                wb.Save()
                log(f"✅ {region_name}: עודכנה שורה {found_row} | תיק {case_id} | {date_str} → '{cancel_value}'")
        else:
            log(f"⚠️ אזהרה - {region_name}: לא נמצאה רשומה עבור תיק {case_id} תאריך {date_str} בגיליון '{sheet_name}'! (בדוק ידנית)")

        wb.Close(False)
        return found_row is not None

    except Exception as e:
        log(f"❌ {region_name}: שגיאה בעדכון ביטול - {e}")
        return False
    finally:
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass


def _handle_cancellation_update(sheet_name, case_id, date_str, cancel_value, mail_idx, dry_run=False):
    """מנתב עדכון ביטול לקובץ ולגיליון הנכון לפי האזור."""
    if sheet_name == "ירושלים":
        update_cancellation_in_excel(
            excel_path=JERUSALEM_EXCEL_PATH,
            sheet_name=JERUSALEM_CANCEL_SHEET,
            case_id=case_id,
            date_str=date_str,
            cancel_value=cancel_value,
            cancel_col_idx=JERUSALEM_CANCEL_COL,
            region_name=f"ירושלים (מייל #{mail_idx})",
            dry_run=dry_run,
        )
    elif sheet_name == "דרום":
        south_path = scan_for_south_excel() or SOUTH_EXCEL_PATH_GUESS
        update_cancellation_in_excel(
            excel_path=south_path,
            sheet_name=SOUTH_SHEET_KEY,
            case_id=case_id,
            date_str=date_str,
            cancel_value=cancel_value,
            cancel_col_idx=SOUTH_CANCEL_COL,
            region_name=f"דרום (מייל #{mail_idx})",
            dry_run=dry_run,
        )
    elif sheet_name == "חיפה":
        south_path = scan_for_south_excel() or SOUTH_EXCEL_PATH_GUESS
        update_cancellation_in_excel(
            excel_path=south_path,
            sheet_name=HAIFA_SHEET_KEY,
            case_id=case_id,
            date_str=date_str,
            cancel_value=cancel_value,
            cancel_col_idx=SOUTH_CANCEL_COL,
            region_name=f"חיפה (מייל #{mail_idx})",
            dry_run=dry_run,
        )
    else:
        log(f"⚠️ מייל #{mail_idx}: אזור לא מוכר '{sheet_name}' עבור עדכון ביטול - מדלג")


# =============================================================================
# 🚀 ראשי
# =============================================================================
def main():
    print("=" * 80)
    print("--- מעבד הזמנות הקלטה - PRODUCTION ---")
    if DRY_RUN:
        print("--- מצב DRY RUN - לא כותב לאקסל! ---")
    if MAX_EMAILS:
        print(f"--- מוגבל ל-{MAX_EMAILS} מיילים ---")
    print("=" * 80)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log(f"📁 נתיב פלט: {OUTPUT_DIR}")
    print("=" * 80)

    judges_list = load_judges_from_all_sheets()

    folder = get_outlook_folder(MAILBOX_PATH)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    cutoff_date = datetime.now() - timedelta(days=DAYS_BACK)
    log(f"📅 מעבד מיילים מתאריך: {cutoff_date.strftime('%d/%m/%Y')}")
    print("=" * 80 + "\n")

    records_by_sheet = {"ירושלים": [], "דרום": [], "חיפה": []}
    processed_pdfs = set()
    seen_orders = {}
    run_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    duplicate_pdf_count = 0
    duplicate_order_count = 0
    non_secretary_count = 0
    processed_count = 0

    for i in range(1, items.Count + 1):
        if MAX_EMAILS and processed_count >= MAX_EMAILS:
            log(f"⏹️ הגעתי למגבלת {MAX_EMAILS} מיילים")
            break

        msg = items.Item(i)

        try:
            msg_date = msg.ReceivedTime
            if hasattr(msg_date, 'replace') and msg_date.tzinfo:
                msg_date = msg_date.replace(tzinfo=None)
            if msg_date < cutoff_date:
                continue
        except:
            continue

        # ✅ טיפול במיילים ללא "הודעת מזכירות"
        processed_count += 1
        is_secretary_notice = "הודעת מזכירות" in msg.Subject

        if not is_secretary_notice:
            has_pdf = False
            pdf_att = None
            for j in range(1, msg.Attachments.Count + 1):
                att = msg.Attachments.Item(j)
                if att.FileName.lower().endswith(".pdf"):
                    has_pdf = True
                    pdf_att = att
                    break

            if not has_pdf:
                non_secretary_count += 1
                continue

            if pdf_att and pdf_att.FileName not in processed_pdfs:
                processed_pdfs.add(pdf_att.FileName)

                record = {
                    "עיר": "",
                    "עיר הקלטה": "",
                    "מספר תיק": "",
                    "תאריך": "",
                    "שעה": "",
                    "שם' השופט": "",
                    "דחיפות": "",
                    "חסוי?": "",
                    "בוטל": "",
                    "התראות": "⚠️ מייל ללא 'הודעת מזכירות'",
                    "קובץ": pdf_att.FileName,
                    "נושא מייל": msg.Subject,
                    "מייל#": i,
                    "תאריך הרצה": run_timestamp,
                }

                records_by_sheet["ירושלים"].append(record)
                log(f"⚠️ מייל #{i}: לא הודעת מזכירות → {pdf_att.FileName}")

            continue

        # ✅ טיפול רגיל עם "הודעת מזכירות"
        pdf_att = None
        for j in range(1, msg.Attachments.Count + 1):
            att = msg.Attachments.Item(j)
            if att.FileName.lower().endswith(".pdf"):
                pdf_att = att
                break

        if not pdf_att:
            continue

        if pdf_att.FileName in processed_pdfs:
            duplicate_pdf_count += 1
            continue

        processed_pdfs.add(pdf_att.FileName)

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / pdf_att.FileName
            pdf_att.SaveAsFile(str(path))  # ← תוקן: str() במקום as_posix() עבור COM

            try:
                with pdfplumber.open(path) as pdf:
                    if len(pdf.pages) < 2:
                        log(f"⏭️ מייל #{i}: PDF עם {len(pdf.pages)} עמוד בלבד → {pdf_att.FileName}")
                        continue
                    raw_text = pdf.pages[1].extract_text() or ""
                    lines = raw_text.split("\n")

                case_id = extract_case_id(raw_text)
                confidential = detect_confidential(path)
                district, sheet_name, court_full, location = extract_district_and_court(raw_text)
                cancelled = detect_cancellation_by_region(raw_text, district)

                # זיהוי שורה הכוללת שופט (כולל תארים והפוך)
                judge_keywords = ["ינפב", "בפני", "כב'", "השופט", "טפושה", "השופטת", "תטפושה", "רשם", "משר", "רשמת", "תמרש", "השופט/ת", "ת/טפושה", "רשם/רשמת", "תמרש/משר"]
                judge_line = next((ln for ln in lines if any(kw in ln for kw in judge_keywords)), "")
                judge = match_judge_improved(judge_line if judge_line else raw_text, judges_list)

                # 🚀 Deep Fallback: אם השורה הממוקדת (למשל 'בלשכת השופט') לא הכילה שופט אמיתי, נסרוק את כל המסמך.
                if not judge or judge == "לא נמצא":
                    # אם לא מצאנו, הסורק כנראה קפץ על שורת סרק בגלל מילת העוגן. נריץ את זיהוי השופט על כל הטקסט.
                    fallback_judge = match_judge_improved(raw_text, judges_list)
                    if fallback_judge and fallback_judge != "לא נמצא":
                        judge = fallback_judge

                if case_id == "67104-01-20":
                    bibi_records = extract_from_pdf_bibi_case(
                        path, case_id, raw_text, lines, judge,
                        msg.Subject, i, run_timestamp, confidential, district
                    )
                    records_by_sheet["ירושלים"].extend(bibi_records)
                    log(f"✅ מייל #{i}: {case_id} (ביבי) - {len(bibi_records)} רשומות")
                    continue

                # 🧪 בדיקה: מתמקד רק במחוז דרום
                if sheet_name != "דרום":
                    # log(f"⏭️ מדלג על {sheet_name} (בדיקת דרום בלבד)")
                    continue

                # 🛠️ תיקון: העברת כל הטקסט לחילוץ תאריכים (תמיכה במעברי שורה)
                pairs = extract_dates_times(raw_text)
                validation = simple_validate(pairs, bool(cancelled))

                if not pairs:
                    continue

                records_added = False

                for date_str, time_str in pairs:
                    order_key = (case_id, date_str, time_str, sheet_name)

                    if order_key in seen_orders:
                        first_mail, first_status = seen_orders[order_key]

                        is_first_cancelled = bool(first_status)
                        is_current_cancelled = bool(cancelled)

                        if is_first_cancelled != is_current_cancelled:
                            log(f"✅ מייל #{i}: תיק {case_id} - הזמנה + ביטול")
                        else:
                            duplicate_order_count += 1
                            continue

                    seen_orders[order_key] = (i, cancelled)

                    # ❌ ביטול → מעדכן בקובץ הקיים (במצב ReadOnly לבדיקה) ותוך כדי מוסיף לדו"וח
                    if cancelled:
                        log(f"❌ מייל #{i}: תיק {case_id} {date_str} - ביטול (בדיקה: נוסף לדו''ח)")
                        _handle_cancellation_update(
                            sheet_name=sheet_name,
                            case_id=case_id,
                            date_str=date_str,
                            cancel_value=cancelled,
                            mail_idx=i,
                            dry_run=True, # 🔒 הגנה: תמיד מצב בדיקה מול קובץ המקור
                        )
                        # בניגוד לרגיל - פה אנחנו לא עושים continue כי אנחנו רוצים לראות את זה באקסל החדש
                        # continue 

                    # ✅ הכנת הרשומה (הזמנה או ביטול)
                    alerts_list = list(validation.get("alerts", []))
                    if not judge or judge == "לא נמצא":
                        alerts_list.append("חסר שופט")

                    record = {
                        "עיר": court_full,
                        "עיר הקלטה": location,
                        "מספר תיק": case_id,
                        "תאריך": date_str,
                        "שעה": time_str,
                        "שם' השופט": judge,
                        "דחיפות": "דורש בדיקה ידנית ⚠️" if (not judge or judge == "לא נמצא") else "רגיל",
                        "חסוי?": confidential,
                        "בוטל": cancelled if cancelled else "",
                        "התראות": "⚠️ " + ", ".join(alerts_list) if alerts_list else "",
                        "קובץ": pdf_att.FileName,
                        "נושא מייל": msg.Subject,
                        "מייל#": i,
                        "תאריך הרצה": run_timestamp,
                    }

                    if sheet_name in records_by_sheet:
                        records_by_sheet[sheet_name].append(record)
                        records_added = True
                    else:
                        # מחוז לא מוכר - שמור בירושלים עם התראה
                        record["התראות"] = f"🚨 מחוז לא זוהה ({sheet_name}) - דורש בדיקה ידנית"
                        record["דחיפות"] = "דורש בדיקה ידנית 🚨"
                        records_by_sheet["ירושלים"].append(record)
                        records_added = True
                        log(f"🚨 מייל #{i}: תיק {case_id} - מחוז '{sheet_name}' לא מוכר! נשמר בירושלים עם התראה")

                if records_added:
                    status = "❌ ביטול" if cancelled else "✅"
                    log(f"{status} מייל #{i}: {case_id} → {sheet_name}")

            except Exception as e:
                log(f"❌ מייל #{i}: {str(e)}")

    if DRY_RUN:
        print("\n" + "=" * 80 + "\n🔍 DRY RUN - לא כותב לאקסל!\n" + "=" * 80)
        for sheet_name, records in records_by_sheet.items():
            if records:
                print(f"\n  📋 {sheet_name}: {len(records)} רשומות שהיו נשמרות:")
                for r in records:
                    case = r.get('מספר תיק', '?')
                    date = r.get('תאריך', '?')
                    time = r.get('שעה', '?')
                    court = r.get('בית משפט', r.get('עיר', '?'))
                    judge = r.get('שם שופט/ת', r.get("שם' השופט", '?'))
                    cancelled = r.get('האם בוטל?', r.get('בוטל', ''))
                    cancelled_str = f" | ❌ {cancelled}" if cancelled else ""
                    print(f"    תיק {case} | {date} {time} | {court} | {judge}{cancelled_str}")
        print("\n" + "=" * 80)
        print("  💡 כדי לכתוב ממש: שנה DRY_RUN = False בראש הקובץ")
        print("=" * 80)
    else:
        print("\n" + "=" * 80 + "\n--- שומר קבצים... ---\n" + "=" * 80)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M')

        for sheet_name, records in records_by_sheet.items():
            if records:
                output_file = OUTPUT_DIR / f"הזמנות_{sheet_name}_{timestamp}.xlsx"
                save_with_rtl(records, output_file)
                log(f"✅ {sheet_name}: {len(records)} רשומות → {output_file.name}")

    print("\n" + "=" * 80)
    log(f"🎉 סיימתי! עובדו {len(processed_pdfs)} PDFs")
    if non_secretary_count > 0:
        log(f"⚠️  {non_secretary_count} מיילים ללא 'הודעת מזכירות' (נוספו עם התראה)")
    if duplicate_pdf_count > 0:
        log(f"⏭️  דולגו {duplicate_pdf_count} PDF כפולים")
    if duplicate_order_count > 0:
        log(f"⏭️  דולגו {duplicate_order_count} דיונים כפולים")
    print("=" * 80)

if __name__ == "__main__":
    main()
