# file_saver.py
# =============================================================================
# 📁 מערכת שמירת קבצים - מבוססת על הלוגיקה המנצחת של run_word_count_local.py
# =============================================================================

import os
import shutil
import datetime
from pathlib import Path

# =============================================================================
# 📝 לוגים
# =============================================================================
LOG_DIR = r"d:\yoel\projects\auto\logs"
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"file_saver_log_{datetime.datetime.now().strftime('%Y-%m-%d')}.txt")

def log(msg):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    entry = f"{timestamp} - {msg}"
    print(entry, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(entry + "\n")
    except:
        pass

# =============================================================================
# 🛠️ פונקציות עזר
# =============================================================================

def get_hebrew_month_name(month_num):
    months = {
        1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל", 5: "מאי", 6: "יוני",
        7: "יולי", 8: "אוגוסט", 9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר"
    }
    return months.get(month_num, "")

def determine_court_type(text):
    """
    זיהוי סוג בית משפט לפי טקסט (פשוט ויעיל).
    """
    if not text:
        return "אחר"
    
    # בדיקת 1000 תווים ראשונים
    start_text = text[:1000] 
    
    if "מחוזי" in start_text:
        return "מחוזי"
    elif "שלום" in start_text:
        return "שלום"
    elif "משפחה" in start_text:
        return "שלום"  # בית משפט למשפחה -> שלום
    elif "תעבורה" in start_text:
        return "שלום"  # בית משפט לתעבורה -> שלום
    elif "נוער" in start_text:
        if "מחוזי" in start_text:
            return "מחוזי"
        else:
            return "שלום" # ברירת מחדל לנוער -> שלום
    elif "עבודה" in start_text or "לעבודה" in start_text:
        return "עבודה"
    elif "עליון" in start_text:
        return "עליון"
    
    return "אחר"

# =============================================================================
# 💾 הפונקציה הראשית לשמירה
# =============================================================================

def save_file_to_hierarchy(file_path, details, region_name, court_type=None, text_content=None):
    """
    מעתיק את הקובץ למיקום הנכון בהיררכיה:
    F:\Region\Year\Month\Region\Court\Date_Folder
    דוגמה: F:\ירושלים\2026\01- ינואר\ירושלים\מחוזי\1.1.26
    """
    try:
        # 1. וידוא שיש תאריך
        if not details.get('date'):
            log(f"⚠️ לא ניתן לשמור קובץ ללא תאריך: {os.path.basename(file_path)}")
            return False

        file_date = details['date']
        
        # 2. אם לא סופק סוג בית משפט, ננסה לזהות מהטקסט
        if not court_type and text_content:
            court_type = determine_court_type(text_content)
        if not court_type:
             court_type = "אחר" # ברירת מחדל

        # 3. בניית הנתיב
        year = str(file_date.year)
        month_num = file_date.month
        month_str = f"{month_num:02d}- {get_hebrew_month_name(month_num)}"
        
        # נסה פורמטים שונים לתיקיית התאריך
        # עדיפות: 1. d.m.yy (1.1.26)
        #           2. dd.mm.yyyy (01.01.2026)
        #           3. d.m.yyyy (1.1.2026)
        #           4. dd.mm.yy (01.01.26)
        
        possible_folders = [
            f"{file_date.day}.{file_date.month}.{str(year)[-2:]}", # 1.1.26
            f"{file_date.day:02d}.{file_date.month:02d}.{year}",   # 01.01.2026
            f"{file_date.day}.{file_date.month}.{year}",           # 1.1.2026
            f"{file_date.day:02d}.{file_date.month:02d}.{str(year)[-2:]}" # 01.01.26
        ]
        
        # תרגום שם האזור לעברית לנתיב
        region_hebrew_sub = "ירושלים" if region_name == "Jerusalem" else "דרום"
        
        # נתיב בסיס קבוע (הכל תחת ירושלים בכונן F)
        base_path = os.path.join(r"F:\ירושלים", year, month_str, region_hebrew_sub, court_type)
        
        target_dir = None
        
        # חיפוש תיקיית תאריך קיימת
        if os.path.exists(base_path):
            try:
                # קוראים את התיקיות שקיימות בתוך מחוזי/שלום
                existing_folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]
                
                for folder_name in existing_folders:
                    # מנסים למצוא חפיפה מלאה של מספרים לתאריך הנועד של התיק. מתעלמים מאפסים מובילים
                    # דוגמה: file_date 01/02/2026 -> מחפשים בסבטרינגים 1.2.26 או 1.02.26 בתיקייה
                    parts = folder_name.replace('-', '.').split('.')
                    if len(parts) >= 3:
                        try:
                            f_day = int(parts[0])
                            f_month = int(parts[1])
                            f_year = int(parts[2])
                            if f_year < 100: f_year += 2000
                            
                            if f_day == file_date.day and f_month == file_date.month and f_year == file_date.year:
                                target_dir = os.path.join(base_path, folder_name)
                                log(f"✅ נמצאה תיקיית תאריך תואמת באחסון (דינאמית): {folder_name}")
                                break
                        except ValueError:
                            pass
            except Exception as e:
                log(f"⚠️ שגיאה בקריאת תיקיות מ-{base_path}: {e}")
                
        else:
            log(f"⚠️ נתיב בסיס לא קיים: {base_path}")
            return False
        
        # אם לא נמצאה תיקיית תאריך נימנע משמירה כדי לא לעשות בלאגן
        if not target_dir:
            target_dir = os.path.join(base_path, possible_folders[0])
            log(f"⚠️ תיקיית תאריך לא נמצאה. נתיב צפוי: {target_dir}")
            return False # לא שומרים אם אין תיקייה מוכנה (בטיחות)

        if not os.path.exists(r"F:"):
             log(f"❌ שגיאה: כונן F: לא נמצא!")
             return False

        # 4. חיפוש תת-תיקייה לפי מספר תיק
        # דוגמה: F:\ירושלים\...\מחוזי\27.1.26\14465-09-25
        case_num = details.get('case_num')
        if case_num and target_dir and os.path.exists(target_dir):
            try:
                subfolders = [f for f in os.listdir(target_dir) if os.path.isdir(os.path.join(target_dir, f))]
                
                # חיפוש מדויק או חלקי של מספר התיק
                for sub in subfolders:
                    if case_num in sub:
                        target_dir = os.path.join(target_dir, sub)
                        log(f"✅ נמצאה תת-תיקייה לתיק: {sub}")
                        break
            except Exception as e:
                log(f"⚠️ שגיאה בבדיקת תת-תיקיות ב-{target_dir}: {e}")

        # 5. ביצוע ההעתקה
        filename = os.path.basename(file_path)
        target_path = os.path.join(target_dir, filename)
        
        if os.path.exists(target_path):
            msg = f"File exists: {target_path}"
            log(f"⚠️ {msg}. Skipping.")
            return False, msg

        shutil.copy2(file_path, target_path)
        log(f"📂 הקובץ הועתק בהצלחה ל: {target_path}")
        return True, target_path

    except Exception as e:
        msg = f"Error saving file: {e}"
        log(f"❌ {msg}")
        return False, msg
