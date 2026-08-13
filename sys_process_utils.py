import os
import time
import datetime
import win32com.client as win32
import psutil
import win32gui
import win32process

LOCK_FILE = r"d:\yoel\projects\auto\logs\excel_access.lock"

def log(msg):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"{timestamp} - {msg}", flush=True)

def acquire_lock():
    """Tries to create a lock file. Returns True if successful."""
    if os.path.exists(LOCK_FILE):
        # Check if the process is actually still running (PID could be in lock file)
        try:
            with open(LOCK_FILE, "r") as f:
                pid = int(f.read().strip())
            if psutil.pid_exists(pid):
                return False
        except:
            pass # Old or empty lock file
    
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))
    return True

def release_lock():
    """Removes the lock file."""
    if os.path.exists(LOCK_FILE):
        try:
            os.remove(LOCK_FILE)
        except:
            pass

def is_work_hours():
    """Returns True if current time is between 06:00 and 20:00."""
    now = datetime.datetime.now().time()
    start = datetime.time(6, 0)
    end = datetime.time(20, 0)
    return start <= now <= end

def kill_excel_force(force=False):
    """Forcefully kills Excel processes."""
    if not force and is_work_hours():
        log("Skipping Excel taskkill because it's currently work hours (06:00-20:00).")
        return

    log("Attempting to close Excel workbooks gracefully first...")
    close_all_excel_vba()
    time.sleep(2)

    log("Forcefully killing Excel processes (Nightly Cleanup)...")
    os.system("taskkill /F /IM excel.exe /T >nul 2>&1")
    time.sleep(2)

def kill_word_force(force=False):
    """Forcefully kills Word processes."""
    if not force and is_work_hours():
        log("Skipping Word taskkill because it's currently work hours (06:00-20:00).")
        return

    log("Forcefully killing Word processes (Nightly Cleanup)...")
    os.system("taskkill /F /IM winword.exe /T >nul 2>&1")
    time.sleep(2)


def close_all_excel_vba():
    """Closes all workbooks and quits Excel via COM."""
    try:
        excel = win32.GetActiveObject("Excel.Application")
        log(f"[INFO] Found active Excel instance. Closing {excel.Workbooks.Count} workbooks...")
        for wb in list(excel.Workbooks):
            try:
                wb.Save()
                wb.Close(False)
            except: pass
        excel.Quit()
        log("[OK] Excel Quit successfully.")
    except:
        pass # No excel or already closed


def release_excel_workbook(workbook_paths: list):
    """
    Safely releases a locked Excel workbook via COM without killing Excel.
    Closes only the specific workbooks in `workbook_paths` (by filename match).
    Safe to call during work hours — does NOT kill the Excel process.
    Discards any unsaved changes (SaveChanges=False) to break the lock.
    """
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        return  # Excel not running at all — nothing to release

    closed_any = False
    for wb in list(excel.Workbooks):
        try:
            wb_path = wb.FullName.lower()
            for target in workbook_paths:
                if os.path.basename(target).lower() in wb_path or wb_path == target.lower():
                    log(f"[RELEASE] Closing locked workbook: {wb.FullName}")
                    wb.Close(SaveChanges=False)
                    closed_any = True
                    break
        except Exception as e:
            log(f"[WARN] Could not close workbook: {e}")

    if closed_any:
        time.sleep(1)  # Allow OneDrive to release the cloud lock
        log("[RELEASE] Workbook lock released.")

def kill_invisible_excel():
    """
    Finds any running EXCEL.EXE processes that do NOT have a visible window,
    and terminates them. Safe to call during work hours.
    """
    try:
        visible_pids = set()
        def enum_windows_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                visible_pids.add(pid)
        win32gui.EnumWindows(enum_windows_callback, None)
        
        killed_count = 0
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() == 'excel.exe':
                    pid = proc.info['pid']
                    if pid not in visible_pids:
                        log(f"[CLEANUP] Terminating background/invisible Excel process: PID {pid}")
                        try:
                            proc.kill()
                            killed_count += 1
                        except psutil.AccessDenied:
                            log(f"[WARN] Access Denied killing PID {pid}. Trying force taskkill...")
                            os.system(f"taskkill /F /PID {pid} >nul 2>&1")
                            killed_count += 1
            except Exception:
                pass
        if killed_count > 0:
            time.sleep(1)
    except Exception as e:
        log(f"[WARN] Error running kill_invisible_excel: {e}")

def kill_excel_if_stuck():
    """Wrapper that calls kill_invisible_excel to clean up stuck background Excel instances."""
    kill_invisible_excel()

import functools
import pywintypes

def com_retry(max_retries=10, delay=1.0):
    """
    Decorator that retries a function if it fails with 'Call was rejected by callee'.
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_err = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if "rejected by callee" in str(e).lower() or "-2147418111" in str(e) or "RPC_E_CALL_REJECTED" in str(e):
                        last_err = e
                        time.sleep(delay)
                    else:
                        raise
            
            # If we exhausted retries due to a locked/edit mode Excel
            raise Exception("שגיאה קריטית: אקסל במצב עריכה (תא פעיל) או מציג תיבת דו-שיח חוסמת.\nאנא לחץ על ESC באקסל כדי לצאת ממצב עריכה ולסגור הודעות פתוחות, ואז נסה שוב.") from last_err
        return wrapper
    return decorator

def robust_com_call(func, *args, max_retries=10, delay=1.0, **kwargs):
    """Helper to run a single COM operation with retry logic without defining a new function."""
    last_err = None
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if "rejected by callee" in str(e).lower() or "-2147418111" in str(e) or "RPC_E_CALL_REJECTED" in str(e) or "0x800a01a8" in str(e):
                last_err = e
                time.sleep(delay)
            else:
                raise
                
    # If we exhausted retries due to a locked/edit mode Excel
    log(f"[CRITICAL] COM Call failed after {max_retries} retries: {last_err}")
    raise Exception(f"שגיאה קריטית: אקסל במצב עריכה (תא פעיל) או מנותק. פרטים: {last_err}")
