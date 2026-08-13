# -*- coding: utf-8 -*-
import os
import datetime

# --- DYNAMIC DRIVE DETECTION ---
def scan_for_drive_root(possible_letters=["F", "Y", "G", "H", "I", "O", "E", "D"]):
    """Finds the base drive letter where 'בתי משפט' hierarchy lives."""
    for letter in possible_letters:
        drive = f"{letter}:\\"
        if os.path.exists(os.path.join(drive, "בתי משפט")):
            return drive
    return "F:\\" # Fallback

DRIVE_ROOT = scan_for_drive_root()

# Excel Paths
JERUSALEM_EXCEL_PATH = r"C:\Users\yoel\OneDrive - Hever\טבלת מעקב בתי משפט 2023 מעודכן מתאריך 25.6.xlsx"
JERUSALEM_SHEET_NAME = "2023"  # שם הגיליון הפעיל באקסל ירושלים

# Note: South Excel might move, so we might need dynamic discovery, but this is the last known location.
SOUTH_EXCEL_PATH_GUESS = r"I:\האחסון שלי\שרות א חדש.xlsx"
SOUTH_SHEET_KEY = "שרות א באר שבע ודרום"
SOUTH_CANCELLATIONS_SHEET = "ביטולים באר שבע"
HAIFA_SHEET_KEY = "שרות א חיפה וצפון"

# שרת דיונים ליואל — תיקים של OTHER_REGION (יובל דדו + יוליה)
YOAL_SERVER_PATH = os.path.join(DRIVE_ROOT, "בתי משפט", "דיונים ליואל")
ORDERS_BASE_PATH = os.path.join(DRIVE_ROOT, "בתי משפט", "הזמנות")
YOAL_RECIPIENTS  = "recipient1@example.com; recipient2@example.com"
YOAL_LOOKBACK_DAYS = 14   # סורקים 14 ימים אחורה


# --- FUNCTIONS ---

def get_year_root(year=None):
    """Returns the main root folder for a given year (e.g. F:\ירושלים\2026)."""
    if not year: year = datetime.datetime.now().year
    return os.path.join(DRIVE_ROOT, "ירושלים", str(year))

def get_month_folder_name(date_obj=None):
    """Returns just the folder name '02- פברואר'."""
    if not date_obj: date_obj = datetime.datetime.now()
    
    month_num = date_obj.month
    month_names = ["", "01- ינואר", "02- פברואר", "03- מרץ", "04- אפריל", "05- מאי", "06- יוני", 
                   "07- יולי", "08- אוגוסט", "09- ספטמבר", "10- אוקטובר", "11- נובמבר", "12- דצמבר"]
    
    return month_names[month_num]

def get_region_path(region="Jerusalem", date_obj=None):
    """
    Returns the full path to a region's folder for a given date.
    Region: 'Jerusalem' (-> 'ירושלים') or 'South' (-> 'דרום').
    Structure: F:\ירושלים\2026\02- פברואר\דרום
    """
    if not date_obj: date_obj = datetime.datetime.now()
    
    year_root = get_year_root(date_obj.year)
    month_name = get_month_folder_name(date_obj)
    
    month_path = os.path.join(year_root, month_name)
    
    if region.lower() in ["jerusalem", "ירושלים"]:
        return os.path.join(month_path, "ירושלים")
    elif region.lower() in ["south", "דרום", "darom"]:
        return os.path.join(month_path, "דרום")
        
    return month_path # Fallback

def scan_for_south_excel():
    """Attempts to find the South Excel file dynamically."""
    candidates = [
        "שרות א חדש.xlsx",
        "עותק של שרות א חדש.xlsx"
    ]
    
    search_paths = [
        r"C:\Users\yoel\OneDrive - Hever",
        r"Y:\Besh",
        r"Y:\Besh\שונות",
        r"Y:",
        r"I:\האחסון שלי",
        r"G:\האחסון שלי",
        r"H:\האחסון שלי",
        r"F:\האחסון שלי",
        r"F:",
        r"G:",
        r"H:",
    ]
    
    for base in search_paths:
        if not os.path.exists(base): continue
        for name in candidates:
            path = os.path.join(base, name)
            if os.path.exists(path): return path
            
    return None

def get_active_excel_path(region_name):
    """Unified helper to get the current Excel tracking sheet for a region."""
    if region_name.lower() in ["jerusalem", "ירושלים"]:
        return JERUSALEM_EXCEL_PATH
    elif region_name.lower() in ["south", "דרום"]:
        return scan_for_south_excel()
    return None

def verify_paths():
    """Debug function to check what exists."""
    print("Verifying Paths (Corrected Structure)...")
    
    now = datetime.datetime.now()
    
    j_path = get_region_path("Jerusalem", now)
    print(f"   Jerusalem Path ({now.strftime('%m/%Y')}): {j_path}")
    print(f"      -> {'Exists' if os.path.exists(j_path) else 'Missing'}")
    
    s_path = get_region_path("South", now)
    print(f"   South Path     ({now.strftime('%m/%Y')}): {s_path}")
    print(f"      -> {'Exists' if os.path.exists(s_path) else 'Missing'}")
    
    if os.path.exists(s_path):
         print(f"      Subfolders in South: {os.listdir(s_path)[:5]}")
         
    s_excel = scan_for_south_excel()
    if s_excel: print(f"   South Excel found at: {s_excel}")
    else: print("   South Excel NOT found (Check OneDrive path?)")

if __name__ == "__main__":
    verify_paths()
