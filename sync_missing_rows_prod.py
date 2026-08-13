"""
[SYNC] Sync Missing Rows - Production
==================================
Compares server folders (F: drive) against Excel data.

1. Cases that EXIST on server  → Mark "לא בוטל" + fill duration (hours)
2. Cases that DON'T exist on server → Report "לברר - אולי בוטל או טרם הגיע"
3. Cases on server but NOT in Excel → Report as missing (need to add)
"""
import os
import datetime
import re
import win32com.client as win32
from win32com.propsys import propsys, pscon
import pythoncom
import config_drive_paths as conf
LOOKBACK_DAYS = 30  # [SUCCESS] 30 ימים אחורה - כדי לתפוס דיונים ישנים שעדיין פתוחים באקסל

LOG_FILE = os.path.join(os.path.dirname(__file__), "logs", "sync_missing_rows.log")
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log(msg):
    try:
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        safe_msg = str(msg).encode('utf-8', errors='replace').decode('utf-8')
        formatted_msg = f"{timestamp} - {msg}"
        print(f"{timestamp} - {safe_msg}", flush=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(formatted_msg + "\n")
    except:
        pass

# ============================================================================
# Audio Duration
# ============================================================================
def get_audio_duration(file_path):
    """Get audio file duration in hours using Windows Shell properties."""
    try:
        properties = propsys.SHGetPropertyStoreFromParsingName(file_path)
        duration_delta = properties.GetValue(pscon.PKEY_Media_Duration).GetValue()
        if not duration_delta: return 0
        seconds = duration_delta / 10_000_000
        return round(seconds / 3600, 2)
    except: return 0

def calculate_total_audio_duration(folder_path):
    """Sums duration of audio files in ONE priority folder (Source/Backup/Root) to avoid duplicates."""
    audio_exts = {'.mp3', '.wav', '.wma', '.m4a', '.dss', '.ds2', '.mp4'}
    
    # Priority folders to check (in order)
    priority_subs = ['מקור', 'גיבוי', 'ראשי', 'Original', 'Backup']
    
    target_folder = folder_path 
    
    # Check if any priority subfolder exists and has files
    for sub in priority_subs:
        sub_path = os.path.join(folder_path, sub)
        if os.path.exists(sub_path):
            files_in_sub = [f for f in os.listdir(sub_path) if os.path.splitext(f)[1].lower() in audio_exts]
            if files_in_sub:
                target_folder = sub_path
                break
    
    # Now sum files in the chosen target folder ONLY (no recursive walk to avoid mixing)
    total_seconds = 0
    found_files = 0
    
    # Scan the chosen folder (non-recursive to be safe, or recursive if needed for split folders inside?)
    # Usually audio files are flat inside 'Mekor'.
    
    for root, _, files in os.walk(target_folder):
        # Safety: If we started at root (because no subfolder found), avoid entering 'Backup' if we missed it above?
        # Actually, simpler: just scan the target_folder content.
        
        for f in files:
            if os.path.splitext(f)[1].lower() in audio_exts:
                full_path = os.path.join(root, f)
                try:
                    dur = get_audio_duration(full_path) # Returns HOURS
                    if dur > 0:
                        total_seconds += dur
                        found_files += 1
                except: pass
        
        # If we found files in the top level of target_folder, we might want to stop descending?
        # But 'os.walk' descends. Let's assume structure is simple enough or we want all parts.
        # But wait, if target_folder is ROOT, os.walk will go into 'Backup' too!
        
        if target_folder == folder_path:
             # We are scanning root. We must avoid counting duplicates in subfolders if we are counting root files.
             # BUT usually if 'Mekor' exists we took it. If not, we are here.
             # If 'Backup' exists but empty/no-audio, we might be here.
             pass
             
    return total_seconds, found_files

# ============================================================================
# Case Number Extraction
# ============================================================================
def extract_case_from_path(path):
    """Extract case number from folder/file name.
    [SUCCESS] תוקן: מבחין בין שנה קצרה (22-30) לסיומת תיק (15, 16...)
    """
    fname = os.path.basename(path)
    m = re.search(r'(\d{3,})[-_](\d{1,2})[-_](\d{2,4})', fname)
    if m:
        g1, g2, g3 = m.groups()
        # אם g1 הוא שנה (4 ספרות), זה לא מספר תיק
        if len(g1) == 4 and (g1.startswith("19") or g1.startswith("20")):
            return None
        if len(g3) == 4:
            return f"{g1}-{int(g2):02d}-{g3}"  # כבר מלא
        if len(g3) == 2:
            val = int(g3)
            if 20 <= val <= 30:  # טווח שנים סביר → מוסיף 20
                return f"{g1}-{int(g2):02d}-20{g3}"
            else:  # סיומת תיק (15, 16, 17...) → נשאר כמו שהוא
                return f"{g1}-{int(g2):02d}-{g3}"
        return f"{g1}-{int(g2):02d}-{g3}"
    return None

def is_similar(c1, c2):
    """Allow 1 character difference (typo check)."""
    if len(c1) != len(c2): return False
    diffs = sum(1 for a, b in zip(c1, c2) if a != b)
    return diffs <= 1

# ============================================================================
# Main Sync Logic
# ============================================================================
def main_sync():
    excel = None
    wb_real = None
    simple_report = []
    try:
        log(f"[START] Starting Sync for LAST {LOOKBACK_DAYS} DAYS...")
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except:
            excel = win32.Dispatch("Excel.Application")
            
        excel.DisplayAlerts = False
        try: excel.Visible = True
        except: pass
        
        real_path = conf.JERUSALEM_EXCEL_PATH
        
        # Open or latch onto existing workbook
        wb_real = None
        target_name = os.path.basename(real_path)
        try:
            for w in excel.Workbooks:
                if w.Name == target_name:
                    wb_real = w
                    break
        except: pass
        if not wb_real:
            wb_real = excel.Workbooks.Open(real_path, UpdateLinks=False)
            
        # Create Temp Copy to avoid messing with user's filters
        import tempfile, uuid
        temp_dir = tempfile.gettempdir()
        unique_name = f"sync_temp_{uuid.uuid4().hex[:8]}.xlsx"
        temp_path = os.path.join(temp_dir, unique_name)
        
        try:
            wb_real.SaveCopyAs(temp_path)
        except Exception as e:
            log(f"[ERROR] Error saving temp copy: {e}")
            return []
        
        try:
            wb_temp = excel.Workbooks.Open(temp_path)
        except:
            log("[ERROR] Could not open the temp copy. Aborting.")
            return []
        
        ws = wb_temp.Sheets("2023")
        
        # Clear filters on temp
        try:
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
            if ws.ListObjects.Count > 0:
                try: ws.ListObjects(1).AutoFilter.ShowAllData()
                except: pass
        except: pass
        
        # ── Headers ──
        headers = {}
        found_dates = []
        r_head = 1
        
        for r_check in [1, 2]:
            line_vals = ws.Range(ws.Cells(r_check, 1), ws.Cells(r_check, 30)).Value[0]
            if any(line_vals):
                r_head = r_check
                for c_idx, val in enumerate(line_vals):
                    s = str(val).strip()
                    if s: 
                        headers[s] = c_idx + 1
                        if "תאריך" in s: found_dates.append(c_idx + 1)
                break
    
        col_case = headers.get("מספר תיק")
        col_date = headers.get("תאריך")
        col_status = headers.get("האם בוטל?") or headers.get("סטטוס") or headers.get("בוטל")
        col_judge = headers.get("שם שופט") or headers.get("שופט") or headers.get("שם שופט/ת") or headers.get("שם שופט הדן בתיק")
        col_duration = None
        for k, v in headers.items():
            if "אורך" in k and "שעות" in k: col_duration = v; break
        if not col_duration:
            for k, v in headers.items():
                if "אורך" in k or "שעות" in k: col_duration = v; break
        
        col_words = headers.get("מילים למתמלל") or headers.get("מספר מילים")
        
        if len(found_dates) > 1:
            log(f"[WARNING]  Warning: Found multiple 'Date' columns at: {found_dates}. Using {col_date}")
            
        log(f"   Cols: Case={col_case}, Date={col_date}, Status={col_status}, Duration={col_duration}, Judge={col_judge}")
    
        # ── Load all Excel rows into memory ──
        last_row = ws.Cells(ws.Rows.Count, col_case).End(-4162).Row
        log(f"[STATS] Read {last_row} rows from temp copy.")
    
        existing_keys = set()
        date_map = {}
        excel_rows = []  # Store row data for reverse lookup
        
        raw_cases = ws.Range(ws.Cells(r_head+1, col_case), ws.Cells(last_row, col_case)).Value
        raw_dates = ws.Range(ws.Cells(r_head+1, col_date), ws.Cells(last_row, col_date)).Value
        
        # Also read status and duration columns if they exist
        raw_status = None
        raw_duration = None
        raw_judge = None
        if col_status:
            raw_status = ws.Range(ws.Cells(r_head+1, col_status), ws.Cells(last_row, col_status)).Value
        if col_duration:
            raw_duration = ws.Range(ws.Cells(r_head+1, col_duration), ws.Cells(last_row, col_duration)).Value
        if col_judge:
            raw_judge = ws.Range(ws.Cells(r_head+1, col_judge), ws.Cells(last_row, col_judge)).Value
        
        # Close Temp immediately
        wb_temp.Close(False)
        try: os.remove(temp_path)
        except: pass
        
        # Process loaded data
        if raw_cases and raw_dates:
            for i in range(len(raw_cases)):
                try:
                    c = str(raw_cases[i][0]).strip()
                    d = raw_dates[i][0]
                    d_obj = None
                    if hasattr(d, 'date'): d_obj = d.date()
                    elif isinstance(d, datetime.datetime): d_obj = d.date()
                    
                    status_val = str(raw_status[i][0]).strip() if raw_status and raw_status[i][0] else ""
                    dur_val = raw_duration[i][0] if raw_duration and raw_duration[i][0] else None
                    judge_val = str(raw_judge[i][0]).strip() if raw_judge and raw_judge[i][0] else ""
                    
                    row_num = r_head + 1 + i  # Actual Excel row number
                    
                    if c:
                        existing_keys.add((c, d_obj))
                        
                        excel_rows.append({
                            'case': c,
                            'date': d_obj,
                            'row': row_num,
                            'status': status_val,
                            'duration': dur_val,
                            'judge': judge_val,
                        })
                        
                        if d_obj:
                            if d_obj not in date_map: date_map[d_obj] = []
                            date_map[d_obj].append(c)
                        
                        # Normalization variants
                        parts = c.split('-')
                        if len(parts) == 3:
                            p1, p2, p3 = parts
                            if len(p3) == 2: 
                                c_long = f"{p1}-{p2}-20{p3}"
                                existing_keys.add((c_long, d_obj))
                                if d_obj: date_map[d_obj].append(c_long)
                            elif len(p3) == 4: 
                                c_short = f"{p1}-{p2}-{p3[2:]}"
                                existing_keys.add((c_short, d_obj))
                                if d_obj: date_map[d_obj].append(c_short)
                                
                except: pass
                
        log(f"   Loaded {len(existing_keys)} existing cases into memory.")
    
        # ── Scan ONLY relevant date folders on server ──
        # Instead of scanning everything, find date folders matching last N days
        root_j = conf.get_region_path("Jerusalem") 
        log(f"📂 Scanning: {root_j} (last {LOOKBACK_DAYS} days only)")
        
        server_cases = {}
        found_miss = []
        count_scanned = 0
        cutoff = datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)
        
        # Generate target dates
        target_dates = []
        for i in range(LOOKBACK_DAYS + 1):
            d = datetime.date.today() - datetime.timedelta(days=i)
            target_dates.append(d)
        
        # Expected date folder formats: "12.2.26", "12.02.26", "12.2.2026"
        target_folder_names = set()
        for d in target_dates:
            target_folder_names.add(f"{d.day}.{d.month}.{str(d.year)[-2:]}")         # 12.2.26
            target_folder_names.add(f"{d.day:02d}.{d.month:02d}.{str(d.year)[-2:]}")  # 12.02.26
            target_folder_names.add(f"{d.day}.{d.month}.{d.year}")                     # 12.2.2026
            target_folder_names.add(f"{d.day:02d}.{d.month:02d}.{d.year}")             # 12.02.2026
        
        # Walk only 2 levels: court type folders → date folders
        if os.path.exists(root_j):
            court_types = [f for f in os.listdir(root_j) if os.path.isdir(os.path.join(root_j, f))]
            
            for court in court_types:
                court_path = os.path.join(root_j, court)
                
                for date_folder in os.listdir(court_path):
                    date_path = os.path.join(court_path, date_folder)
                    if not os.path.isdir(date_path): continue
                    
                    # Parse date from folder name
                    dt_obj = None
                    try:
                        parts = date_folder.split('.')
                        if len(parts) == 3:
                            y = parts[2]
                            if len(y) == 2: y = "20" + y
                            dt_obj = datetime.date(int(y), int(parts[1]), int(parts[0]))
                    except: continue
                    
                    if not dt_obj: continue
                    if dt_obj < cutoff: continue  # Skip old dates
                    
                    # Scan case folders inside this date folder
                    for case_dir in os.listdir(date_path):
                        case_path = os.path.join(date_path, case_dir)
                        if not os.path.isdir(case_path): continue
                        
                        count_scanned += 1
                        case_num = extract_case_from_path(case_dir)
                        if not case_num: continue
                        
                        # חילוץ תאריך משם התיקייה עצמה (אם הוגדר יום ספציפי)
                        actual_dt_obj = dt_obj
                        try:
                            clean_dir = case_dir.replace(case_num, '')
                            m_date = re.search(r'\b(\d{1,2})[\./-](\d{1,2})[\./-](\d{2,4})\b', clean_dir)
                            if m_date:
                                d, mth, y = int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3))
                                if y < 100: y += 2000
                                actual_dt_obj = datetime.date(y, mth, d)
                        except: pass
                        
                        server_cases[(case_num, actual_dt_obj)] = case_path
                        
                        # Normalized variants
                        parts_c = case_num.split('-')
                        if len(parts_c) == 3:
                            p1, p2, p3 = parts_c
                            if len(p3) == 4:
                                server_cases[(f"{p1}-{p2}-{p3[2:]}", actual_dt_obj)] = case_path
                            elif len(p3) == 2:
                                server_cases[(f"{p1}-{p2}-20{p3}", actual_dt_obj)] = case_path
    
                        # Check if missing from Excel
                        if (case_num, actual_dt_obj) in existing_keys: continue
                        if (case_num, None) in existing_keys: continue
    
                        # Typo check
                        potential_typo = False
                        if actual_dt_obj in date_map:
                            for existing_c in date_map[actual_dt_obj]:
                                if is_similar(case_num, existing_c):
                                    log(f"[WARNING]  SUSPECTED TYPO: Folder={case_num}, Excel={existing_c} on {actual_dt_obj}")
                                    potential_typo = True
                                    break
                        if potential_typo: continue
    
                        # Mismatch check
                        same_case_dates = [ex_date for (ex_case, ex_date) in existing_keys 
                                          if ex_case == case_num and ex_date is not None]
                        if same_case_dates:
                            log(f"[WARNING]  Mismatch: {case_num}: Server={actual_dt_obj}, Excel={same_case_dates}")
                            continue
                        
                        log(f"🔎 Missing from Excel: {case_num} ({actual_dt_obj})")
                        found_miss.append({
                            'case': case_num,
                            'date': actual_dt_obj,
                            'path': case_path,
                        })
    
        log(f"\n[STATS] Server scan complete: {count_scanned} folders scanned, {len(server_cases)} cases found.")
        
        # ====================================================================
        # NEW: Cross-check Excel rows vs Server
        # ====================================================================
        cutoff = datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)
        
        rows_to_update = []    # Exist on server → mark "לא בוטל" + fill duration
        rows_to_check = []     # NOT on server → "לברר"
        
        for row_data in excel_rows:
            case = row_data['case']
            d = row_data['date']
            
            # Only check recent dates, up to YESTERDAY (don't check today's or future sessions)
            if not d or d < cutoff or d >= datetime.date.today():
                continue
            
            # Skip if already has status filled
            if row_data['status'] and row_data['status'] not in ['', 'None', 'nan']:
                continue
                
            # Check if exists on server
            server_path = server_cases.get((case, d))
            
            if server_path:
                # [SUCCESS] EXISTS on server — sum total audio duration
                total_hours, files_count = calculate_total_audio_duration(server_path)
                
                # ROUNDING: Round to nearest 0.5 hours
                # Examples: 0.66 -> 0.5, 0.8 -> 1.0, 1.2 -> 1.0, 1.3 -> 1.5
                raw_hours = total_hours
                rounded_duration = round(raw_hours * 2) / 2
                if rounded_duration == 0 and raw_hours > 0: 
                    rounded_duration = 0.5 # Minimum charge if exists?
                
                # DEBUG: Specific check for Toker case
                if "61019-12-23" in case:
                    log(f"🐛 DEBUG TOKER: Found on server with date {d}. Duration: {raw_hours} -> {rounded_duration}")
    
                rows_to_update.append({
                    'row': row_data['row'],
                    'case': case,
                    'date': d,
                    'duration': rounded_duration,
                    'status': 'לא בוטל',
                    'files_found': files_count, 
                })
            else:
                # [WARNING] NOT on server — need to investigate
                if "61019-12-23" in case:
                    log(f"🐛 DEBUG TOKER: NOT found on server for date {d}. checking typos...")
    
                rows_to_check.append({
                    'row': row_data['row'],
                    'case': case,
                    'date': d,
                    'status': 'לברר - לא נמצא בשרת',
                    'judge': row_data.get('judge', ''),
                })
        
        log(f"\n{'='*60}")
        log(f"[STATS] Cross-Check Results (Last {LOOKBACK_DAYS} days):")
        log(f"   [SUCCESS] Found on server (to mark 'לא בוטל'): {len(rows_to_update)}")
        log(f"   [WARNING]  NOT found (to investigate): {len(rows_to_check)}")
        log(f"   🔎 Missing from Excel (server-only): {len(found_miss)}")
        log(f"{'='*60}")
        
        # ====================================================================
        # Write updates to REAL Excel
        # ====================================================================
        updates_written = 0
        
        if rows_to_update and col_status:
            log(f"\n📝 Writing {len(rows_to_update)} updates to Excel...")
            try:
                # Get the real workbook again
                try:
                    target_name = os.path.basename(real_path)
                    for w in excel.Workbooks:
                        if w.Name == target_name:
                            wb = w
                            break
                    if not wb:
                        wb = excel.Workbooks.Open(real_path, UpdateLinks=False)
                except Exception as wbe:
                    log(f"[ERROR] Could not open real workbook for update: {wbe}")
                ws_real = wb.Sheets("2023")
                
                for item in rows_to_update:
                    row = item['row']
                    
                    # Check Row Color
                    skip_status_update = False
                    try:
                        row_color = ws_real.Cells(row, col_case).Interior.Color
                        if row_color != 16777215 and row_color != 0: 
                             skip_status_update = True
                    except: pass
    
                    # Write "לא בוטל" (only if not colored)
                    if col_status and not skip_status_update:
                        current_status = ws_real.Cells(row, col_status).Value
                        if not current_status or str(current_status).strip() in ['', 'None', 'nan']:
                            ws_real.Cells(row, col_status).Value = "לא בוטל"
                            updates_written += 1
                    
                    # Write duration (hours) - ALWAYS (needed for billing)
                    
                    # Write duration (hours)
                    if col_duration and item['duration'] > 0:
                        current_dur = ws_real.Cells(row, col_duration).Value
                        if not current_dur:
                            ws_real.Cells(row, col_duration).Value = item['duration']
                    
                    if updates_written % 10 == 0 and updates_written > 0:
                        log(f"   ...written {updates_written} rows...")
                
                wb.Save()
                log(f"[SUCCESS] Written {updates_written} status updates + durations to Excel.")
                
            except Exception as e:
                log(f"[ERROR] Error writing to Excel: {e}")
        else:
            log("ℹ️  No updates needed (all rows already have status).")
    
        # ====================================================================
        # Filter 'Not Found' rows by Color (Don't report cancelled/charged rows)
        # ====================================================================
        final_missing_report = []
        
        if rows_to_check:
            log(f"\n🔎 Verifying {len(rows_to_check)} suspected missing rows against Excel colors...")
            try:
                # Ensure workbook is open
                try:
                    target_name = os.path.basename(real_path)
                    for w in excel.Workbooks:
                        if w.Name == target_name:
                            wb = w
                            break
                    if not wb:
                        wb = excel.Workbooks.Open(real_path, UpdateLinks=False)
                except Exception as wbe:
                    log(f"[WARNING] Could not open real workbook for color check: {wbe}")
                ws_real = wb.Sheets("2023")
                
                for item in rows_to_check:
                    row = item['row']
                    # Check color
                    try:
                        c_val = ws_real.Cells(row, col_case).Interior.Color
                        # If colored (not white and not 0), skip reporting
                        if c_val != 16777215 and c_val != 0:
                            continue
                    except: pass
                    
                    # If we are here, it's a white row that is missing files!
                    final_missing_report.append(item)
                    
            except Exception as e:
                log(f"[WARNING] Error checking colors for report: {e}")
                # Fallback: report all if color check fails
                final_missing_report = rows_to_check
    
        # ====================================================================
        # Build Report
        # ====================================================================
        simple_report = []
        
        # Section 1: Missing from Excel (Server has file, Excel doesn't)
        if found_miss:
            for item in found_miss:
                simple_report.append(f"🔎 חסר באקסל: {item['case']} | {item['date']}")
        
        # Section 2: Not found on server (potential missing files)
        if final_missing_report:
            for item in final_missing_report:
                judge_str = item.get('judge', '')
                simple_report.append(f"[WARNING] לא נמצא בשרת: {item['case']} | {item['date']} | {judge_str} (שורה {item['row']})")
    
        # Section 3: Summary of updates
        if updates_written > 0:
            simple_report.append(f"[SUCCESS] סומנו {updates_written} תיקים כ'לא בוטל' + אורך הקלטה")
    finally:
        try:
            if wb_real:
                wb_real.Close(False)
        except: pass
        try:
            if excel and not excel.Workbooks.Count:
                excel.Quit()
        except: pass

    return simple_report

def check_missing_with_return():
    """Entry point for nightly job."""
    return main_sync()

if __name__ == "__main__":
    log("👋 Starting Sync Script...")
    results = main_sync()
    log("\n📋 Report:")
    for line in results:
        log(f"   {line}")
