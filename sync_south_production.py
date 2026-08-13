"""
🔄 South Sync & Process - Production
=====================================
Special script for South District (Darom).

1. Transfers files from Shared Folder (Y:\Besh) to F: drive.
   - Logic: Y:\Besh\YYYY\Month\Date\Case -> F:\Jerusalem\YYYY\Month\South\[District/Magistrate]\Date\Case
   - Determines [District/Magistrate] by checking Excel column 'City' ('Beer Sheva District' vs 'Beer Sheva Magistrate').
   
2. Updates Excel for South:
   - Marks "Not Cancelled" (if row is white).
   - Fills "Normal" urgency (if empty).
   - Fills Recording Duration (Column X).
   
3. Reports missing files (White rows in Excel with no files in F).
"""
import os
import shutil
import datetime
import win32com.client as win32
import re
import config_drive_paths as conf
from sync_missing_rows_prod import calculate_total_audio_duration # Reuse logic

# Configuration
SHARED_FOLDER_ROOT = r"Y:\Besh" 
LOOKBACK_DAYS = 5

def get_shared_folder_path(date_obj):
    """Returns path like Y:\Besh\2026\פברואר 2026\17.02.2026"""
    year = str(date_obj.year)
    
    # Month name Hebrew mapping
    month_names = ["", "ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", 
                   "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]
    month_str = f"{month_names[date_obj.month]} {year}"
    
    day_str = f"{date_obj.day:02d}.{date_obj.month:02d}.{year}"
    
    return os.path.join(SHARED_FOLDER_ROOT, year, month_str, day_str)

def get_dest_folder_path(date_obj, court_type):
    """
    Returns path in F: like F:\ירושלים\2026\02- פברואר\דרום\מחוזי\17.2.26
    Scans for EXISTING date folder first to avoid duplicates like '15.02.2026' vs '15.2.26'.
    """
    region_path = conf.get_region_path("South", date_obj) # ...\02- פברואר\דרום
    parent = os.path.join(region_path, court_type)
    
    # Target date formats to check
    # 1. d.m.yy (Preferred new) -> 17.2.26
    # 2. dd.mm.yyyy (Old style) -> 17.02.2026
    # 3. dd.mm.yy -> 17.02.26
    # 4. d.m.yyyy -> 17.2.2026
    
    # Preferred new: d.m.yy -> 17.2.26
    preferred_name = f"{date_obj.day}.{date_obj.month}.{str(date_obj.year)[-2:]}"
    preferred_path = os.path.join(parent, preferred_name)

    # Check if parent exists
    if os.path.exists(parent):
        for d_name in os.listdir(parent):
            # Try to parse folder name as date (robustly)
            try:
                # Common formats: 17.2.26, 17.02.2026, 2026-02-17
                d_found = None
                
                # Format 1: Dot separated
                if '.' in d_name:
                    parts = d_name.split('.')
                    if len(parts) == 3:
                        y = int(parts[2])
                        if y < 100: y += 2000
                        d_found = datetime.date(y, int(parts[1]), int(parts[0]))
                        
                # Format 2: Dash separated (ISO or reversed)
                elif '-' in d_name:
                    parts = d_name.split('-')
                    if len(parts) == 3:
                        # Check if first part is Year
                        if len(parts[0]) == 4:
                            d_found = datetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
                        else:
                            y = int(parts[2])
                            if y < 100: y += 2000
                            d_found = datetime.date(y, int(parts[1]), int(parts[0]))
                            
                if d_found and d_found == date_obj:
                    return os.path.join(parent, d_name) # Found existing!
                    
            except: pass
            
    # Default: Create preferred (d.m.yy)
    return preferred_path

def normalize_case(c):
    return str(c).strip().replace(' ', '')

def normalize_date_to_obj(d_val):
    """Convert Excel date values or strings into a proper datetime.date object."""
    if not d_val: return None
    if hasattr(d_val, 'date'): return d_val.date()
    if isinstance(d_val, datetime.datetime): return d_val.date()
    
    # It's a string or other type
    s = str(d_val).strip().split(" ")[0]
    s = s.replace('\u200e', '').replace('\u200f', '')
    
    for fmt in ["%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%d.%m.%y"]:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except: continue
    
    return None

def determine_court_type(city_val):
    """Decides 'מחוזי' or 'שלום' based on Excel City column."""
    val = str(city_val).strip()
    if 'מחוזי' in val:
        return 'מחוזי'
    return 'שלום' # Default for Magistrate/Family/Traffic/Labor

def ensure_workbook_open(excel_app, wb_obj, path):
    """Checks if workbook is accessible. If not, re-opens it."""
    try:
        # Try a dummy access
        _ = wb_obj.Name
        return wb_obj
    except:
        print("⚠️ Excel workbook seems closed. Re-opening...")
        try:
            return excel_app.Workbooks.Open(path)
        except Exception as e:
            print(f"❌ Failed to re-open Excel: {e}")
            raise e

def check_south_with_return():
    """
    Main function for Nightly Job.
    Returns a list of report strings.
    """
    report_lines = []
    print("🚀 Starting South Sync (Production)...")
    
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except:
        excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    try: excel.Visible = True
    except: pass
    
    target_path = conf.SOUTH_EXCEL_PATH_GUESS
    wb = None
    target_name = os.path.basename(target_path)
    try:
        for w in excel.Workbooks:
            if w.Name == target_name:
                wb = w
                break
    except: pass
    if not wb:
        try:
            wb = excel.Workbooks.Open(target_path, UpdateLinks=False)
        except Exception as e:
            return [f"❌ Error opening South Excel: {e}"]
            
    # Map Headers
    ws = wb.Sheets(conf.SOUTH_SHEET_KEY)
    
    headers = {}
    r_head = 1
    
    # Scan first 3 rows
    for r in range(1, 4):
        vals = ws.Range(ws.Cells(r, 1), ws.Cells(r, 30)).Value[0]
        if any(vals):
            r_head = r
            for c_idx, v in enumerate(vals):
                if v is None: continue
                s = str(v).strip()
                if s and s != 'None': headers[s] = c_idx + 1
            break
            
    col_case = headers.get("מספר תיק")
    col_date = headers.get("תאריך")
    col_city = headers.get("עיר")
    col_status = headers.get("האם בוטל?") or headers.get("סטטוס") or headers.get("בוטל")
    col_urgency = headers.get("דחיפות")
    col_duration_decimal = 24 # X (Fallback to fixed)
    
    # Load Data
    try:
        last_row = ws.Cells(ws.Rows.Count, col_case).End(-4162).Row
        raw_cases = ws.Range(ws.Cells(r_head+1, col_case), ws.Cells(last_row, col_case)).Value
        raw_dates = ws.Range(ws.Cells(r_head+1, col_date), ws.Cells(last_row, col_date)).Value
        col_city_idx = col_city if col_city else 1
        raw_cities = ws.Range(ws.Cells(r_head+1, col_city_idx), ws.Cells(last_row, col_city_idx)).Value
    except Exception as e:
        return [f"❌ Error reading Excel data: {e}"]
    
    excel_map = {}
    if raw_cases and raw_dates:
        for i in range(len(raw_cases)):
            try:
                c = normalize_case(raw_cases[i][0])
                d = normalize_date_to_obj(raw_dates[i][0])
                
                city = str(raw_cities[i][0]) if raw_cities and raw_cities[i] and raw_cities[i][0] else ""
                
                if c:
                    excel_map[(c, d)] = {
                        'row': r_head + 1 + i,
                        'city': city
                    }
            except: pass

    # ====================================================================
    # STEP 1: Transfer Y -> F
    # ====================================================================
    transferred_count = 0
    today = datetime.date.today()
    
    for i in range(LOOKBACK_DAYS + 1):
        d_check = today - datetime.timedelta(days=i)
        y_path = get_shared_folder_path(d_check)
        
        if os.path.exists(y_path):
            for case_folder in os.listdir(y_path):
                src_full = os.path.join(y_path, case_folder)
                if not os.path.isdir(src_full): continue
                
                # Match logic
                case_clean = normalize_case(case_folder)
                found_entry = None
                
                # 1. Direct match
                if (case_clean, d_check) in excel_map:
                    found_entry = excel_map[(case_clean, d_check)]
                else:
                    # 2. Fuzzy match
                    for (ex_c, ex_d), data in excel_map.items():
                        if ex_d == d_check and ex_c in case_clean:
                            found_entry = data
                            break
                
                if not found_entry: continue
                
                court_type = determine_court_type(found_entry['city'])
                dest_parent = get_dest_folder_path(d_check, court_type)
                dest_full = os.path.join(dest_parent, case_folder)
                
                # Check exist
                if os.path.exists(dest_full): continue
                
                try:
                    if not os.path.exists(dest_parent): os.makedirs(dest_parent)
                    shutil.copytree(src_full, dest_full)
                    transferred_count += 1
                except Exception as e:
                    report_lines.append(f"❌ נכשל העברה מ-Y: {case_clean} ({e})")

    if transferred_count > 0:
        report_lines.append(f"📦 הועברו {transferred_count} תיקים מ-Y ל-F (דרום).")

    # ====================================================================
    # STEP 2: Update Excel
    # ====================================================================
    updates_stat = 0
    updates_dur = 0
    missing_list = []
    
    cutoff = today - datetime.timedelta(days=LOOKBACK_DAYS)
    
    try:
        # Re-open safely if closed
        wb = ensure_workbook_open(excel, wb, target_path)
        ws = wb.Sheets(conf.SOUTH_SHEET_KEY)
        
        for (case, d), data in excel_map.items():
            if not d or d < cutoff or d > today: continue
            
            row = data['row']
            court_type = determine_court_type(data['city'])
            dest_parent = get_dest_folder_path(d, court_type)
            
            # Find in F
            found_path = None
            if os.path.exists(dest_parent):
                for f_name in os.listdir(dest_parent):
                    if case in normalize_case(f_name):
                        found_path = os.path.join(dest_parent, f_name)
                        break
            
            is_colored = False
            # Check Color logic
            try:
                row_color = ws.Cells(row, col_case).Interior.Color
                is_colored = (row_color != 16777215 and row_color != 0)
            except: pass
            
            if found_path:
                # Found in F
                total_hours, files_count = calculate_total_audio_duration(found_path)
                
                # Update Duration
                if total_hours > 0:
                    current_dur = ws.Cells(row, col_duration_decimal).Value
                    if not current_dur:
                        ws.Cells(row, col_duration_decimal).Value = round(total_hours, 2)
                        updates_dur += 1
                
                # Update Status if white
                if not is_colored and col_status:
                    current_stat = ws.Cells(row, col_status).Value
                    if not current_stat:
                        ws.Cells(row, col_status).Value = "לא בוטל"
                        updates_stat += 1
                
                 # Update Urgency if white
                if not is_colored and col_urgency:
                    current_urg = ws.Cells(row, col_urgency).Value
                    if not current_urg:
                        ws.Cells(row, col_urgency).Value = "רגיל"
                        
            else:
                # Missing from F
                if not is_colored:
                    # Don't report if today (give it time to arrive)
                    if d == today:
                        continue
                    missing_list.append(f"{case} ({d}) - {data['city']}")

        wb.Save()
        if updates_dur > 0:
            report_lines.append(f"✅ דרום: שעות עודכנו ל-{updates_dur} תיקים.")
        
        if missing_list:
            if len(missing_list) > 20:
                 report_lines.append(f"⚠️ חסר בדרום (F): {len(missing_list)} תיקים (רשימה מקוצרת):")
                 for m in missing_list[:20]: report_lines.append(f"   - {m}")
            else:
                report_lines.append(f"⚠️ חסר בדרום (F):")
                for m in missing_list: report_lines.append(f"   - {m}")
            
    except Exception as e:
        report_lines.append(f"❌ שגיאה בעדכון אקסל דרום: {e}")

    finally:
        try:
            if wb:
                wb.Close(False)
        except: pass
        try:
            if excel and not excel.Workbooks.Count:
                excel.Quit()
        except: pass

    return report_lines

if __name__ == "__main__":
    msgs = check_south_with_return()
    for m in msgs: print(m)
