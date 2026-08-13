"""
Transcription Queue Manager v3.1
מערכת חלוקה ושיבוץ דיונים אוטומטית למתמללים - גרסה מיוצבת (תיקון כותרות)
"""

import os
import sys
import datetime
import time
import json
import re
import shutil
import win32com.client as win32
import sys_process_utils
from typing import List, Dict, Any, Optional
import llm_utils
import holidays
import gsheets_utils

# Setup Israel Holidays (Jewish holidays included)
IL_HOLIDAYS = holidays.IL()

# --- Configuration ---
import config_drive_paths
JERUSALEM_EXCEL = config_drive_paths.JERUSALEM_EXCEL_PATH
SOUTH_EXCEL = config_drive_paths.scan_for_south_excel() or r"I:\האחסון שלי\שרות א חדש.xlsx"
METADATA_EXCEL = r"C:\Users\yoel\OneDrive - Hever\Metadata_Transcribers.xlsx"
ASSIGNMENTS_EXCEL = r"C:\Users\yoel\OneDrive - Hever\איושים 2026.xlsx"

# 🚀 Dynamic Drive Scanning for Audio Folders
def get_drive_root(possible_letters=["F", "G", "H", "I", "O", "E"]):
    for letter in possible_letters:
        if os.path.exists(f"{letter}:\\בתי משפט"): return f"{letter}:\\"
    return "F:\\" # Fallback

DRIVE_ROOT = get_drive_root()
SOURCE_BASE_F = os.path.join(DRIVE_ROOT, "בתי משפט", "דיונים ליואל")
DEST_BASE_O = r"O:\My Files\בית משפט1" # Keep O: as likely static if mapped
if not os.path.exists(DEST_BASE_O):
    for letter in ["G", "H", "I", "K", "O"]:
        test_path = f"{letter}:\\My Files\\בית משפט1"
        if os.path.exists(test_path):
            DEST_BASE_O = test_path
            break
            
DEST_FALLBACK_F = os.path.join(DRIVE_ROOT, "FileCloud", "Transcribers")

LOG_DIR = r"d:\yoel\projects\auto\logs\transcription_manager"
if not os.path.exists(LOG_DIR): os.makedirs(LOG_DIR)

# Defaults
GLOBAL_DEFAULT_RATE = 1.5
SLA_DAYS = {"חריג": 1, "מיידי": 3, "דחוף": 5, "רגיל": 10}

# --- DISABLE SWITCH (User requested to stop sending material to transcribers) ---
DISABLE_SWITCH = False 
# -------------------------------------------------------------------------------

def log(msg, level="INFO"):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try: print(f"[{ts}] [{level}] {msg}")
    except: pass
    with open(os.path.join(LOG_DIR, "system.log"), "a", encoding="utf-8") as f:
        f.write(f"[{ts}] [{level}] {msg}\n")

def get_business_days_ahead(start_date: datetime.date, days_count: int) -> datetime.date:
    """Business days skipped Fri, Sat and IL Holidays."""
    curr = start_date
    added = 0
    while added < days_count:
        curr += datetime.timedelta(days=1)
        if curr.weekday() not in [4, 5] and curr not in IL_HOLIDAYS:
            added += 1
    return curr

def normalize_date(val):
    if not val: return None
    if isinstance(val, datetime.datetime): return val.date()
    if isinstance(val, datetime.date): return val
    if hasattr(val, 'date'): return val.date()
    try:
        if isinstance(val, str):
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d.%m.%Y"]:
                try: return datetime.datetime.strptime(val.strip().split(" ")[0], fmt).date()
                except: continue
    except: pass
    return None

def verify_folder_copy(src, dest_f):
    if not os.path.exists(dest_f):
        return False, []
    
    src_files = {}
    for root, _, files in os.walk(src):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, src)
            try:
                src_files[rel_path] = os.path.getsize(full_path)
            except Exception as e:
                log(f"Error reading source file size for {f}: {e}", "WARNING")
                src_files[rel_path] = -1

    dest_files = {}
    for root, _, files in os.walk(dest_f):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, dest_f)
            try:
                dest_files[rel_path] = os.path.getsize(full_path)
            except:
                dest_files[rel_path] = -2

    missing_or_mismatched = []
    for rel_path, src_size in src_files.items():
        if rel_path not in dest_files:
            missing_or_mismatched.append(rel_path)
        elif src_size != dest_files[rel_path]:
            missing_or_mismatched.append(rel_path)
            
    if missing_or_mismatched:
        return False, missing_or_mismatched
    return True, []

class ExcelManager:
    def __init__(self, visible=False):
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = visible
        self.excel.DisplayAlerts = False
        self.workbooks = []

    def open_workbook(self, path):
        import os
        import tempfile
        import shutil
        abs_path = os.path.abspath(path)
        for w in self.excel.Workbooks:
            try:
                if os.path.abspath(w.FullName).lower() == abs_path.lower():
                    return w
            except:
                pass
        abs_path = os.path.abspath(path)
        target_basename = os.path.basename(abs_path)
        try:
            wb = sys_process_utils.robust_com_call(lambda: self.excel.Workbooks.Open(abs_path, ReadOnly=True, UpdateLinks=False))
            if wb is None:
                raise Exception("Workbooks.Open returned None")
            self.workbooks.append(wb)
            return wb
        except Exception as e:
            try:
                import uuid
                td = tempfile.mkdtemp(prefix="transq_")
                tmp = os.path.join(td, f"tmp_{uuid.uuid4().hex[:8]}_{target_basename}")
                shutil.copy2(abs_path, tmp)
                wb = sys_process_utils.robust_com_call(lambda: self.excel.Workbooks.Open(tmp, ReadOnly=True, UpdateLinks=False))
                if wb is None:
                    raise Exception("Workbooks.Open returned None for tmp file")
                self.workbooks.append(wb)
                return wb
            except Exception as e2:
                log(f"Failed to open workbook {abs_path} (direct and temp): {e2}", "ERROR")
                raise e2

    def get_sheet_data(self, wb, sheet_name=None, last_n_rows=1000):
        try:
            ws = wb.Sheets(sheet_name) if sheet_name else wb.Sheets(1)
            
            # 🚀 Fix: Dynamic header row detection (Jerusalem has headers on Row 2 usually)
            r_header = 1
            val1 = str(ws.Cells(1, 1).Value).strip()
            val2 = str(ws.Cells(1, 2).Value).strip()
            if val1 in ["None", ""] and val2 in ["None", ""]:
                r_header = 2
            
            last_row = ws.Cells(ws.Rows.Count, 3).End(-4162).Row
            if last_row < r_header + 1: last_row = ws.UsedRange.Rows.Count
            last_col = ws.UsedRange.Columns.Count

            # Get header range and flatten it
            header_range = ws.Range(ws.Cells(r_header, 1), ws.Cells(r_header, last_col)).Value
            header = header_range[0] if isinstance(header_range, tuple) else header_range

            start_row = max(r_header + 1, last_row - last_n_rows)
            data = ws.Range(ws.Cells(start_row, 1), ws.Cells(last_row, last_col)).Value
            return header, data, start_row
        except Exception as e:
            log(f"Error reading sheet '{sheet_name}': {e}", "ERROR")
            return None, None, 0

    def close_all(self):
        for wb in self.workbooks: 
            try: wb.Close(False)
            except: pass
def verify_folder_copy(src, dest):
    """
    Verifies that all files in src were successfully copied to dest.
    Compares file names and sizes.
    Returns (True, "") or (False, error_message).
    """
    if not os.path.exists(dest):
        return False, "Destination folder does not exist."
        
    src_files = {}
    for root, _, files in os.walk(src):
        for f in files:
            rel_path = os.path.relpath(os.path.join(root, f), src)
            try:
                src_files[rel_path] = os.path.getsize(os.path.join(root, f))
            except:
                pass
            
    dest_files = {}
    for root, _, files in os.walk(dest):
        for f in files:
            rel_path = os.path.relpath(os.path.join(root, f), dest)
            try:
                dest_files[rel_path] = os.path.getsize(os.path.join(root, f))
            except:
                pass
            
    missing = []
    size_mismatch = []
    
    for rel_path, src_size in src_files.items():
        if rel_path not in dest_files:
            missing.append(rel_path)
        elif dest_files[rel_path] != src_size:
            size_mismatch.append(rel_path)
            
    if missing or size_mismatch:
        err = ""
        if missing:
            err += f"Missing files: {', '.join(missing[:5])} (total {len(missing)}). "
        if size_mismatch:
            err += f"Size mismatch (corrupted): {', '.join(size_mismatch[:5])} (total {len(size_mismatch)}). "
        return False, err
        
    return True, ""


class TranscriptionManager:
    def __init__(self, region="south"):
        self.region = region
        self.ex = ExcelManager()
        self.transcribers = {} 
        self.queue = []
        self.ai_cache = {}

    def load_transcribers(self):
        """Compiles transcribers from Jerusalem Excel Roster and calculates current load."""
        log("Loading transcriber metadata from Jerusalem Excel Roster...")
        wb = self.ex.open_workbook(JERUSALEM_EXCEL)
        header, data, _ = self.ex.get_sheet_data(wb, "רשימת מתמללים", 150)
        if not header: return
        
        cols = {str(h).strip(): i for i, h in enumerate(header) if h}
        idx_name = cols.get("שם מועמד/מתרגם") or cols.get("מועמד") or cols.get("מתמלל") or 7
        idx_email = cols.get("מייל") or cols.get("דוא\"ל") or 8
        idx_rate = cols.get("שעות הקלטה") or cols.get("שעות") or cols.get("קצב הקלטה (שעות ליום)") or 9
        idx_region = cols.get("אזור מועדף") or cols.get("אזור") or 10
        idx_max_cases = cols.get("מגבלת תיקים פתוחים", -1)
        idx_daily_h = cols.get("מגבלת שעות יומית", -1)
        idx_prefs = cols.get("העדפות מיוחדות (AI Check)") or cols.get("אזור מועדף") or -1
        idx_opath = cols.get("תיקיית O") or cols.get("נתיב תיקייה") or -1
        
        self.transcribers = {}
        for row in data:
            try:
                name = str(row[idx_name] or "").strip()
                email = str(row[idx_email] or "").strip()
                if not name or "@" not in email: continue
                
                # Try to get values from optional columns or use defaults
                rate_val = GLOBAL_DEFAULT_RATE
                if idx_rate != -1 and idx_rate < len(row) and row[idx_rate] not in (None, ""):
                    try: rate_val = float(row[idx_rate])
                    except: pass
                    
                region_val = "הכל"
                if idx_region != -1 and idx_region < len(row) and row[idx_region] not in (None, ""):
                    region_val = str(row[idx_region]).strip()
                    
                max_cases_val = 5
                if idx_max_cases != -1 and idx_max_cases < len(row) and row[idx_max_cases] not in (None, ""):
                    try: max_cases_val = int(row[idx_max_cases])
                    except: pass
                    
                max_daily_h_val = 4
                if idx_daily_h != -1 and idx_daily_h < len(row) and row[idx_daily_h] not in (None, ""):
                    try: max_daily_h_val = float(row[idx_daily_h])
                    except: pass
                    
                prefs_val = "הכל"
                if idx_prefs != -1 and idx_prefs < len(row) and row[idx_prefs] not in (None, ""):
                    prefs_val = str(row[idx_prefs]).strip()
                    
                opath_val = ""
                if idx_opath != -1 and idx_opath < len(row) and row[idx_opath] not in (None, ""):
                    opath_val = str(row[idx_opath]).strip()
                
                self.transcribers[name] = {
                    "name": name, "email": email,
                    "rate": rate_val,
                    "region": region_val,
                    "max_cases": max_cases_val,
                    "max_daily_h": max_daily_h_val,
                    "special_prefs": prefs_val,
                    "o_path": opath_val,
                    "current_load_hours": 0, "active_cases": 0
                }
            except: continue

        # Calculate live load from tracking sheet
        log(f"Reading tracking sheet for active load ({self.region})...")
        target_excel = SOUTH_EXCEL if self.region == "south" else JERUSALEM_EXCEL
        target_sheet = "שרות א באר שבע ודרום" if self.region == "south" else "2023"
        wb_track = self.ex.open_workbook(target_excel)
        h_t, d_t, _ = self.ex.get_sheet_data(wb_track, target_sheet, 1000)
        if not h_t: return
        
        c_t = {str(h).strip(): i for i, h in enumerate(h_t) if h}
        
        idx_t = c_t.get("מתמלל") or c_t.get("שם מתמלל") or c_t.get("מתמלל/ת")
        idx_r = c_t.get("תאריך מסירה") or c_t.get("מסירה") or c_t.get("תאריך החזרה בפועל")
        idx_h = c_t.get("אורך ( שעות) הקלטה") or c_t.get("שעות הקלטה") or c_t.get("שעות")
        if idx_t is None:
            log(f"⚠️ [CRITICAL] 'מתמלל' column NOT found in {target_sheet}. Possible Headers: {list(c_t.keys())[:10]}...", "ERROR")
            return


        for row in d_t:
            if not isinstance(row, (tuple, list)): continue # Skip extra rows
            try:
                t_name = str(row[idx_t] or "").strip()
                if t_name in self.transcribers:
                    # Case is active if NOT returned ("תאריך מסירה" is empty)
                    is_active = not (row[idx_r])
                    if is_active:
                        h_val = 0
                        try:
                            raw_h = row[idx_h]
                            if isinstance(raw_h, datetime.datetime): h_val = raw_h.hour + (raw_h.minute/60.0)
                            else: h_val = float(raw_h or 0)
                        except: pass
                        self.transcribers[t_name]["current_load_hours"] += h_val
                        self.transcribers[t_name]["active_cases"] += 1
            except: continue

    def load_pending_hearings(self):
        """Finds hearings needing assignment."""
        log(f"Finding pending hearings for {self.region}...")
        target_excel = SOUTH_EXCEL if self.region == "south" else JERUSALEM_EXCEL
        target_sheet = "שרות א באר שבע ודרום" if self.region == "south" else "2023"
        wb = self.ex.open_workbook(target_excel)
        header, data, start_row = self.ex.get_sheet_data(wb, target_sheet, 500)
        if not header: return
        
        cols = {str(h).strip(): i for i, h in enumerate(header) if h}
        
        idx_case = cols.get("מספר תיק")
        idx_stat = cols.get("בוטל") or cols.get("האם בוטל?")
        idx_trans = cols.get("מתמלל") or cols.get("שם מתמלל") or cols.get("מתמלל/ת")
        idx_sent = cols.get("תאריך שליחת פרוטוקול לתמלול") or cols.get("שליחה לתמלול")
        idx_hours = cols.get("אורך ( שעות) הקלטה") or cols.get("שעות הקלטה")
        idx_date = cols.get("תאריך")
        idx_urgency = cols.get("דחיפות")
        idx_judge = cols.get("שם שופט") or cols.get("שם שופט/ת") or cols.get("שם' השופט")

        self.queue = []
        if idx_case is None: return

        for i, row in enumerate(data):
            if not isinstance(row, (tuple, list)): continue
            # Pending = Active court (בוטל is None/empty/לא בוטל), No transcriber, No sent-date
            status = str(row[idx_stat] or "").strip()
            trans = str(row[idx_trans] or "").strip()
            sent_ts = str(row[idx_sent] or "").strip()
            
            if (status == "לא בוטל") and not trans and not sent_ts:
                try:
                    c_id = str(row[idx_case])
                    if not c_id or "None" in c_id: continue
                    
                    h_date = normalize_date(row[idx_date]) or datetime.date.today()
                    urgency = str(row[idx_urgency] or "רגיל").strip()
                    judge_name = str(row[idx_judge] or "לא צוין").strip() if idx_judge is not None else "לא צוין"
                    h_val = 0
                    if idx_hours is not None:
                        raw_h = row[idx_hours]
                        if isinstance(raw_h, datetime.datetime): h_val = raw_h.hour + (raw_h.minute/60.0)
                        else: h_val = float(raw_h or 0)
                    
                    if h_val > 0:
                        deadline = get_business_days_ahead(h_date, SLA_DAYS.get(urgency, 10))
                        if isinstance(deadline, datetime.datetime): deadline = deadline.date()
                        self.queue.append({
                            "case": c_id, "date": h_date, "urgency": urgency,
                            "hours": h_val, "deadline": deadline, "row_in_sheet": start_row + i,
                            "court": self.region, "judge": judge_name
                        })
                except: continue
        log(f"Queue ready: {len(self.queue)} hearings.")

    def check_ai_preference(self, t_name, pref, hearing):
        if not pref or pref == "הכל": return True
        cache_key = f"{t_name}_{hearing['case']}"
        if cache_key in self.ai_cache: return self.ai_cache[cache_key]
        
        prompt = f"Does hearing Case {hearing['case']} (Court: {hearing['court']}) VIOLATE transcriber restriction '{pref}'? Answer YES or NO."
        response = llm_utils.ask_llm(prompt)
        result = (str(response).strip().upper() == "NO")
        self.ai_cache[cache_key] = result
        return result

    def select_best_transcriber(self, hearing):
        candidates = []
        today = datetime.date.today()
        for name, t in self.transcribers.items():
            if t["active_cases"] >= t["max_cases"]: continue
            region_map = {"south": "דרום", "jerusalem": "ירושלים"}
            if t["region"] != "הכל" and t["region"] != region_map.get(self.region, self.region): continue
            if not self.check_ai_preference(name, t["special_prefs"], hearing): continue
            
            eff_rate = min(t["rate"], t["max_daily_h"])
            eta = get_business_days_ahead(today, int((t["current_load_hours"] + hearing["hours"]) / max(0.1, eff_rate)) + 1)
            if eta <= hearing["deadline"]:
                candidates.append({"name": name, "eta": eta, "buffer": (hearing["deadline"] - eta).days})
        
        if candidates: return sorted(candidates, key=lambda x: x["buffer"], reverse=True)[0]
        return None

    def find_hearing_folder(self, case, target_date):
        fid = str(case).strip()
        region_folder = "ירושלים" if self.region == "jerusalem" else "דרום"
        year = str(target_date.year) if target_date else str(datetime.date.today().year)
        month_str = f"{target_date.month:02d}" if target_date else None

        # Build possible date strings to match against folder names (e.g. "02.06.26", "2.6.26", "02_06_2026")
        date_strs = set()
        if target_date:
            d, m, y, Y = target_date.day, target_date.month, str(target_date.year)[2:], target_date.year
            for ds in [
                f"{d:02d}.{m:02d}.{y}", f"{d}.{m}.{y}", f"{d:02d}.{m:02d}.{Y}", f"{d}.{m}.{Y}",
                f"{d:02d}_{m:02d}_{y}", f"{d}_{m}_{y}", f"{d:02d}_{m:02d}_{Y}", f"{d}_{m}_{Y}",
                f"{d:02d}-{m:02d}-{y}", f"{d}-{m}-{y}", f"{d:02d}-{m:02d}-{Y}", f"{d}-{m}-{Y}"
            ]:
                date_strs.add(ds.lower())

        # Search the primary region first, then fall back to all other known region folders.
        regions_to_try = [region_folder]
        for r in ["ירושלים", "חיפה", "בתי משפט"]:
            if r not in regions_to_try:
                regions_to_try.append(r)

        # Helper to search targeted paths
        def search_targeted(base_path, max_depth=3):
            if not os.path.exists(base_path):
                return None
            
            queue = [(base_path, 0)]
            visited = set()
            
            while queue:
                curr_path, depth = queue.pop(0)
                if curr_path in visited:
                    continue
                visited.add(curr_path)
                
                try:
                    entries = os.listdir(curr_path)
                except Exception:
                    continue
                
                for entry in entries:
                    full_entry_path = os.path.join(curr_path, entry)
                    if not os.path.isdir(full_entry_path):
                        continue
                    
                    entry_lower = entry.lower()
                    if entry_lower in date_strs:
                        try:
                            children = os.listdir(full_entry_path)
                            for child in children:
                                if fid in child:
                                    child_path = os.path.join(full_entry_path, child)
                                    if os.path.isdir(child_path):
                                        return child_path
                        except Exception:
                            pass
                    
                    if depth < max_depth:
                        queue.append((full_entry_path, depth + 1))
            return None

        # Try targeted search in regions
        for reg_folder in regions_to_try:
            base_dir = os.path.join(DRIVE_ROOT, reg_folder, year)
            search_dirs = []

            if month_str and os.path.exists(base_dir):
                try:
                    for f in os.listdir(base_dir):
                        if (f.startswith(month_str) or f.startswith(str(target_date.month) + "-")) and os.path.isdir(os.path.join(base_dir, f)):
                            search_dirs.append(os.path.join(base_dir, f))
                except: pass

            search_dirs.append(base_dir)

            for s_dir in search_dirs:
                res = search_targeted(s_dir, max_depth=3)
                if res:
                    log(f"Audio found in region folder '{reg_folder}': {res}")
                    return res

        # Fallback to old static directory
        if os.path.exists(SOURCE_BASE_F):
            res = search_targeted(SOURCE_BASE_F, max_depth=2)
            if res:
                return res

        return None

    def execute_assignment(self, match, h):
        self.last_error = ""
        original_region = self.region
        t = self.transcribers[match["name"]]
        log(f"🚀 Assigning {h['case']} -> {t['name']}")
        src = self.find_hearing_folder(h["case"], h["date"])
        if not src: 
            self.last_error = f"Audio folder not found for {h['case']}"
            log(f"⚠️ Audio not found for {h['case']}", "ERROR")
            return False
            
        dest_base = t.get("o_path")
        if not dest_base:
            potential_names = [t["name"]] + t["name"].split()
            for n in potential_names:
                test_dir = os.path.join(DEST_BASE_O, n)
                if os.path.exists(test_dir):
                    dest_base = test_dir
                    break
            if not dest_base and os.path.exists(DEST_BASE_O):
                try:
                    existing_folders = [f for f in os.listdir(DEST_BASE_O) if os.path.isdir(os.path.join(DEST_BASE_O, f))]
                    for folder in existing_folders:
                        for n in potential_names:
                            if len(n) > 2 and n in folder:
                                dest_base = os.path.join(DEST_BASE_O, folder)
                                break
                        if dest_base: break
                except: pass
            if not dest_base:
                dest_base = os.path.join(DEST_BASE_O, t["name"])

        if not os.path.exists(os.path.dirname(dest_base)): 
            dest_base = os.path.join(DEST_FALLBACK_F, t["name"])
        
        dest_f = os.path.join(dest_base, os.path.basename(src))
        
        try:
            if os.path.exists(dest_f) and len(os.listdir(dest_f)) > 0:
                log(f"Destination folder already exists and is not empty. Running verification before skipping copy: {dest_f}")
                is_verified, failed_files = verify_folder_copy(src, dest_f)
                if not is_verified:
                    log(f"Existing folder copy was incomplete. Missing or mismatched files: {failed_files}. Attempting fallback copy...", "WARNING")
                    for rel_path in failed_files:
                        file_src = os.path.join(src, rel_path)
                        file_dest = os.path.join(dest_f, rel_path)
                        try:
                            os.makedirs(os.path.dirname(file_dest), exist_ok=True)
                            shutil.copy2(file_src, file_dest)
                            log(f"Fallback copied: {rel_path}")
                        except Exception as e_fallback:
                            log(f"Fallback copy failed for {rel_path}: {e_fallback}", "ERROR")
                    
                    is_verified, failed_files = verify_folder_copy(src, dest_f)
                    if not is_verified:
                        raise Exception(f"Folder verification failed for existing folder. Missing files: {failed_files}")
            else:
                import shutil
                if os.path.exists(dest_f):
                    shutil.rmtree(dest_f, ignore_errors=True)
                os.makedirs(dest_f, exist_ok=True)
                
                # Use robocopy for robust large-file copying to FileCloud virtual drives.
                # /E = copy subdirs including empty, /COPY:D = copy Data only (updates timestamps to NOW)
                import subprocess
                cmd = f'robocopy "{src}" "{dest_f}" /E /COPY:D /R:3 /W:2 /NP /MT:4'
                try:
                    # robocopy returns 1 if successful, 0 if nothing copied. Anything >= 8 is a real error.
                    result = subprocess.run(cmd, shell=True, capture_output=True)
                    if result.returncode >= 8:
                        log(f"Robocopy error: {result.returncode} - {result.stderr.decode('utf-8', 'ignore')}", "ERROR")
                        raise Exception("Audio copy failed")
                except Exception as e:
                    log(f"Error copying audio with robocopy: {e}", "ERROR")
                    raise e

                # Verify copy
                is_verified, failed_files = verify_folder_copy(src, dest_f)
                if not is_verified:
                    log(f"Verification failed after robocopy. Missing or mismatched files: {failed_files}. Attempting fallback copy...", "WARNING")
                    for rel_path in failed_files:
                        file_src = os.path.join(src, rel_path)
                        file_dest = os.path.join(dest_f, rel_path)
                        try:
                            os.makedirs(os.path.dirname(file_dest), exist_ok=True)
                            shutil.copy2(file_src, file_dest)
                            log(f"Fallback copied: {rel_path}")
                        except Exception as e_fallback:
                            log(f"Fallback copy failed for {rel_path}: {e_fallback}", "ERROR")
                    
                    is_verified, failed_files = verify_folder_copy(src, dest_f)
                    if not is_verified:
                        raise Exception(f"Folder verification failed after fallback copy. Missing files: {failed_files}")
            
            # Force FileCloud sync trigger by touching all copied files
            try:
                for root, dirs, files in os.walk(dest_f):
                    for f in files:
                        p = os.path.join(root, f)
                        try:
                            os.utime(p, None)
                        except Exception as e:
                            log(f"Error touching file {f} for sync trigger: {e}", "WARNING")
            except Exception as e:
                log(f"Error in sync trigger loop: {e}", "WARNING")
            # Detect true region from the audio folder path and override self.region
            # so write_excel always writes to the correct Excel sheet.
            if src:
                norm_src = src.replace('\\', '/').lower()
                if '/\u05d3\u05e8\u05d5\u05dd/' in norm_src or '\u05d3\u05e8\u05d5\u05dd' in src:
                    self.region = 'south'
                elif '/\u05d9\u05e8\u05d5\u05e9\u05dc\u05d9\u05dd/' in norm_src or '\u05d9\u05e8\u05d5\u05e9\u05dc\u05d9\u05dd' in src:
                    self.region = 'jerusalem'
                if self.region != original_region:
                    log(f"Region corrected from '{original_region}' to '{self.region}' based on audio path.")

            # Action
            self.write_excel(h["case"], h["date"], t["name"], court=h.get("court"), judge=h.get("judge"), hours=h.get("hours"))
            self.send_email(t, h)
            return True
        except Exception as e:
            err_str = str(e).lower()
            if "-2147023174" in err_str or "rpc" in err_str:
                self.last_error = "שגיאה בתקשורת עם אקסל (RPC server unavailable). נא לנסות שוב או לבדוק תהליכי אקסל תקועים."
            else:
                self.last_error = f"Copy or Excel update failed: {e}"
            log(f"Copy failed: {e}", "ERROR")
            if os.path.exists(dest_f): shutil.rmtree(dest_f, ignore_errors=True)
            return False
        finally:
            self.region = original_region  # Always restore original region

    def write_excel(self, case, date_obj, name, court=None, judge=None, hours=None):
        """
        Updates the Excel sheet when an assignment is made.
        """
        import win32com.client
        import datetime
        import os
        import tempfile
        import shutil
        import pythoncom
        import time

        target_excel = SOUTH_EXCEL if self.region == "south" else JERUSALEM_EXCEL
        target_sheet = config_drive_paths.SOUTH_SHEET_KEY if self.region == "south" else config_drive_paths.JERUSALEM_SHEET_NAME
        abs_path = os.path.abspath(target_excel)
        target_basename = os.path.basename(target_excel)

        if not os.path.exists(abs_path):
            log(f"Excel file not found at {abs_path}", "ERROR")
            return

        try:
            pythoncom.CoInitialize()
        except: pass

        excel = None
        wb = None
        opened_by_us = False
        td = None
        tmp_path = None
        row_found = None

        try:
            def _get_excel():
                try: return win32com.client.GetActiveObject("Excel.Application")
                except Exception:
                    e_app = win32com.client.Dispatch("Excel.Application")
                    e_app.Visible = False
                    return e_app
            excel = sys_process_utils.robust_com_call(_get_excel)
            sys_process_utils.robust_com_call(lambda: setattr(excel, 'DisplayAlerts', False))

            # Check if already open
            target_basename_lower = target_basename.lower()
            for w in excel.Workbooks:
                try:
                    if w.Name.lower() == target_basename_lower or os.path.basename(w.FullName).lower() == target_basename_lower:
                        wb = w
                        break
                except:
                    pass

            # If already open in read-only mode (e.g., opened by ExcelManager for loading),
            # close it and re-open in read-write mode so we can save.
            if wb:
                try:
                    if wb.ReadOnly:
                        log(f"Workbook {target_basename} is read-only — closing and re-opening read-write.", "WARNING")
                        wb.Close(False)
                        wb = None
                except Exception:
                    pass

            if wb and not excel.Visible:
                opened_by_us = True

            if not wb:
                def _open_wb():
                    opened = excel.Workbooks.Open(abs_path, ReadOnly=False, UpdateLinks=0)
                    if opened.ReadOnly:
                        opened.Close(False)
                        raise Exception("קובץ האקסל נעול לקריאה בלבד (כנראה פתוח אצלך או אצל משתמש אחר, או שהוא מסתנכרן הרגע). נא לסגור אותו ולנסות שוב.")
                    return opened
                wb = sys_process_utils.robust_com_call(_open_wb)
                opened_by_us = True

            ws = wb.Sheets(target_sheet)

            # Dynamic header row detection
            r_header = 1
            if str(ws.Cells(1, 1).Value).strip() in ["None", ""] and str(ws.Cells(1, 2).Value).strip() in ["None", ""]:
                r_header = 2
            last_col = ws.UsedRange.Columns.Count
            header_range = ws.Range(ws.Cells(r_header, 1), ws.Cells(r_header, last_col)).Value
            header = header_range[0] if isinstance(header_range, tuple) else header_range
            cols = {str(h).strip(): i + 1 for i, h in enumerate(header) if h}

            c_case  = cols.get("מספר תיק")
            c_date  = cols.get("תאריך")
            c_trans = cols.get("מתמלל/ת") or cols.get("מתמלל") or cols.get("שם מתמלל")
            c_sent  = cols.get("שליחה לתמלול") or cols.get("תאריך שליחת פרוטוקול לתמלול")

            if not c_case or not c_date or not c_trans or not c_sent:
                msg = f"Required columns not found in '{target_sheet}': Case={c_case} Date={c_date} Trans={c_trans} Sent={c_sent}. Available: {list(cols.keys())[:15]}"
                log(msg, "ERROR")
                if opened_by_us and wb:
                    try: wb.Close(False)
                    except: pass
                raise Exception(msg)

            target_date = normalize_date(date_obj)

            # Find row: Find() with cycle detection, then linear fallback
            rng = ws.Columns(c_case)
            found = rng.Find(case, LookAt=2)
            first_addr = None

            while found:
                addr = found.Address
                if normalize_date(ws.Cells(found.Row, c_date).Value) == target_date:
                    row_found = found.Row
                    break
                if first_addr is None:
                    first_addr = addr
                elif addr == first_addr:
                    break  # Full cycle - no date match
                found = rng.FindNext(found)

            if not row_found:
                log(f"Find() missed {case} ({target_date}), running linear scan...", "WARNING")
                last_row_fb = ws.Cells(ws.Rows.Count, c_case).End(-4162).Row
                for r in range(r_header + 1, last_row_fb + 1):
                    cell_case = ws.Cells(r, c_case).Value
                    if cell_case and str(cell_case).strip() == str(case).strip():
                        if normalize_date(ws.Cells(r, c_date).Value) == target_date:
                            row_found = r
                            log(f"Linear scan found row {r} for {case} ({target_date})")
                            break

            if row_found:
                def _update_existing():
                    ws.Cells(row_found, c_trans).Value = name
                    ws.Cells(row_found, c_sent).Value = datetime.datetime.now()
                    wb.Save()
                sys_process_utils.robust_com_call(_update_existing, max_retries=30, delay=2.0)
                log(f"Excel values saved: {case} ({target_date}) row {row_found} -> '{name}'")
            else:
                log(f"Case {case} ({target_date}) not found in sheet '{target_sheet}'. Appending a new row...", "WARNING")
                def _append_new():
                    # Find list object inside the retry to ensure it evaluates at the time of modification
                    local_row_found = None
                    try:
                        if ws.ListObjects.Count > 0:
                            list_obj = ws.ListObjects(1)
                            new_row = list_obj.ListRows.Add()
                            local_row_found = new_row.Range.Row
                    except:
                        pass
                    
                    if not local_row_found:
                        last_row = ws.Cells(ws.Rows.Count, c_case).End(-4162).Row
                        local_row_found = last_row + 1

                    ws.Cells(local_row_found, c_case).Value = case
                    ws.Cells(local_row_found, c_date).Value = target_date
                    ws.Cells(local_row_found, c_trans).Value = name
                    ws.Cells(local_row_found, c_sent).Value = datetime.datetime.now()
                    
                    # Try optional fields
                    c_court = cols.get("בית משפט") or cols.get("ביהמ\"ש")
                    c_judge = cols.get("שם שופט/ת") or cols.get("שופט")
                    c_status = cols.get("האם בוטל?") or cols.get("סטטוס")
                    c_hours = None
                    for k, v in cols.items():
                        if "אורך" in k or "שעות" in k:
                            c_hours = v
                            break

                    if c_court and court:
                        ws.Cells(local_row_found, c_court).Value = 'ביהמ"ש העליון' if 'עליון' in str(court) else court
                    if c_judge and judge:
                        ws.Cells(local_row_found, c_judge).Value = judge
                    if c_status:
                        ws.Cells(local_row_found, c_status).Value = "לא בוטל"
                    if c_hours and hours:
                        try:
                            ws.Cells(local_row_found, c_hours).Value = float(hours)
                        except:
                            pass
                    
                    # Set row color to indicate it was dynamically added/assigned
                    try:
                        ws.Rows(local_row_found).Interior.Color = 13434828
                    except:
                        pass

                    wb.Save()
                    return local_row_found
                row_found = sys_process_utils.robust_com_call(_append_new, max_retries=30, delay=2.0)
                log(f"Excel new row created and saved: {case} ({target_date}) row {row_found} -> '{name}'")

            if opened_by_us:
                wb.Close(True)
                wb = None
                log(f"Excel updated successfully: {case} ({target_date}) in sheet '{target_sheet}'")

        except Exception as e:
            err_str = str(e).lower()
            if "rejected by callee" in err_str or "-2147418111" in err_str or "rpc_e_call_rejected" in err_str or "-2147467259" in err_str:
                msg = "שגיאה קריטית: אקסל במצב עריכה (תא פעיל) או מציג תיבת דו-שיח חוסמת.\nאנא לחץ על ESC באקסל כדי לצאת ממצב עריכה ולסגור הודעות פתוחות, ואז נסה שוב."
                log(f"Excel write error (Edit Mode): {msg}", "ERROR")
                raise Exception(msg) from e
            elif "-2147023174" in err_str or "rpc server is unavailable" in err_str:
                msg = "השרת של אקסל לא זמין (RPC server is unavailable). ייתכן שאקסל נסגר באופן בלתי צפוי. נא לנסות שוב."
                log(f"Excel write error: {msg}", "ERROR")
                raise Exception(msg) from e
            log(f"Excel write error: {e}", "ERROR")
            raise e
        finally:
            if opened_by_us and wb:
                try: wb.Close(False)
                except: pass

            if excel:
                try:
                    if excel.Workbooks.Count == 0:
                        excel.Quit()
                except: pass
            if td:
                shutil.rmtree(td, ignore_errors=True)

    def send_email(self, t, h):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            
            def _send():
                import win32com.client as win32
                ol = win32.Dispatch("Outlook.Application")
                m = ol.CreateItem(0)
                m.To = t["email"]
                
                h_date_str = h['date'].strftime('%d/%m/%Y') if hasattr(h['date'], 'strftime') else str(h['date'])
                judge_str = h.get('judge', 'לא צוין')
                duration_str = f"{round(h.get('hours', 0), 2)} שעות"
                
                m.Subject = f"תיק תמלול חדש: {h['case']} (דיון מיום {h_date_str})"
                m.HTMLBody = f"<div dir='rtl' style='font-family: Arial, sans-serif; font-size: 14px;'>שלום {t['name']},<br><br>הקבצים הועלו לשרת עבור התיק הבא לביצוע:<br><br><ul><li><b>מספר תיק:</b> {h['case']}</li><li><b>תאריך הדיון:</b> {h_date_str}</li><li><b>שם שופט/ת:</b> {judge_str}</li><li><b>אורך הדיון:</b> {duration_str}</li></ul><br>נא לאשר קבלה.<br><br>בברכה,<br>מערכת שיבוצים</div>"
                m.Send()
            
            sys_process_utils.robust_com_call(_send)
            log(f"📧 Email sent to {t['email']}")

        except: log("Email failed (ensure Outlook is open)")

    def run(self, dry_run=True, requeue=False):
        if DISABLE_SWITCH:
            log("🛑 [DISABLE SWITCH] Script is currently DISABLED by user request. Skipping all assignments.", "WARNING")
            return
            
        log(f"--- STARTING RUN (Region: {self.region}, Dry: {dry_run}) ---")
        self.load_transcribers()
        self.load_pending_hearings()
        
        matches = []
        for h in self.queue:
            best = self.select_best_transcriber(h)
            if best:
                matches.append({"hearing": h, "transcriber": best})
                self.transcribers[best["name"]]["current_load_hours"] += h["hours"]
                log(f"Match Found: {h['case']} -> {best['name']}")
        
        if not dry_run:
            for m in matches: self.execute_assignment(m["transcriber"], m["hearing"])
            log(f"Run completed. {len(matches)} assignments processed.")
        else:
            log(f"Dry Run end. Found {len(matches)} potential matches.")
        self.ex.close_all()

if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--region", default="south")
    p.add_argument("--no-dry-run", action="store_true")
    p.add_argument("--requeue", action="store_true")
    args = p.parse_args()
    
    TranscriptionManager(region=args.region).run(dry_run=not args.no_dry_run, requeue=args.requeue)
