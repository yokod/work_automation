"""
# Nightly Automation Job
# =========================
# Runs every night at 00:00 via Task Scheduler.
# 1. Processes new Word files (word count + Excel update)
# 2. Checks for missing rows (files on disk not in Excel)
# 3. Runs Supreme Court (Elyon) extraction
# 4. Sends a RICH HTML summary report via Outlook email.
"""
import os
import re
import datetime
import json
import shutil
import time
import traceback
import subprocess
import tempfile
import win32com.client as win32
import sys

# Force UTF-8 for console output to prevent UnicodeEncodeErrors
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    pass

import crash_notifier
import config_drive_paths as conf
import sys_process_utils

# --- CONFIG ---
EMAIL_TO = "your_email@example.com"
LOG_DIR = r"d:\yoel\projects\auto\logs"
BACKUP_DIR = r"d:\yoel\projects\auto\backups"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

JERUSALEM_EXCEL = conf.JERUSALEM_EXCEL_PATH
SOUTH_EXCEL = conf.scan_for_south_excel() or conf.SOUTH_EXCEL_PATH_GUESS

# ---------------------------------------------------------------------------
# Isolated step runner — each step gets its own process + timeout
# ---------------------------------------------------------------------------
def run_step_isolated(step_name: str, module: str, func_expr: str,
                      timeout: int, default=None):
    """
    Spawns a child process that imports `module` and evaluates `func_expr`.
    The child writes its return value to a temp JSON file; the parent reads it.
    If the child exceeds `timeout` seconds it is killed and `default` is returned.
    Stdout/stderr from the child are appended to the nightly log file.
    """
    result_file = os.path.join(LOG_DIR, f"_step_{step_name}.json")

    code = (
        "import sys, json, traceback\n"
        f"sys.path.insert(0, r'{SCRIPT_DIR}')\n"
        f"import {module}\n"
        "try:\n"
        f"    _result = {func_expr}\n"
        f"    open(r'{result_file}', 'w', encoding='utf-8').write(\n"
        "        json.dumps({'ok': True, 'data': _result}, ensure_ascii=False, default=str))\n"
        "except Exception as _e:\n"
        f"    open(r'{result_file}', 'w', encoding='utf-8').write(\n"
        "        json.dumps({'ok': False, 'error': str(_e),\n"
        "                    'tb': traceback.format_exc()}, ensure_ascii=False))\n"
    )

    with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False,
                                     encoding='utf-8') as tf:
        tf.write(code)
        wrapper = tf.name

    try:
        log(f"[STEP] '{step_name}' starting (timeout={timeout}s)...")
        # On Windows, cmd.exe's >> redirect opens the log file without FILE_SHARE_WRITE,
        # so re-opening it from this process raises PermissionError. The child
        # inherits stdout (= the log file) and its log() calls already write there
        # via print(), so no explicit stdout capture is needed.
        proc = subprocess.Popen([sys.executable, wrapper])
        try:
            proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()
            log(f"[TIMEOUT] '{step_name}' exceeded {timeout}s — killed, continuing.")
            return False, default

        if os.path.exists(result_file):
            try:
                with open(result_file, 'r', encoding='utf-8') as rf:
                    res = json.load(rf)
                if res.get('ok'):
                    log(f"[STEP] '{step_name}' finished OK.")
                    return True, res['data']
                log(f"[ERROR] '{step_name}' failed: {res.get('error', '?')}")
                tb = res.get('tb', '')
                if tb:
                    log(tb[:600])
                return False, default
            except Exception as parse_err:
                log(f"[ERROR] Could not read result for '{step_name}': {parse_err}")
        else:
            log(f"[ERROR] '{step_name}' produced no result file.")
        return False, default
    finally:
        try: os.remove(wrapper)
        except: pass
        try: os.remove(result_file)
        except: pass

ONEDRIVE_EXE = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe")

def onedrive_stop():
    os.system("taskkill /F /IM OneDrive.exe >nul 2>&1")
    time.sleep(4)

def onedrive_start():
    if os.path.exists(ONEDRIVE_EXE):
        subprocess.Popen([ONEDRIVE_EXE])

def prepare_excel_for_write():
    """Restart OneDrive so it pulls the latest SharePoint version, wait for
    sync to complete, then stop it so the COM write is uncontested."""
    if not os.path.exists(ONEDRIVE_EXE):
        log("[ONEDRIVE] OneDrive.exe not found — skipping sync preparation.")
        return
    log("[ONEDRIVE] Restarting OneDrive to pull latest SharePoint version...")
    onedrive_stop()
    onedrive_start()
    log("[ONEDRIVE] Waiting 60s for OneDrive to finish syncing...")
    time.sleep(60)
    log("[ONEDRIVE] Stopping OneDrive so COM write is uncontested...")
    onedrive_stop()
    log("[ONEDRIVE] Ready — Excel write will proceed with no sync conflict.")

def resume_onedrive():
    """Restart OneDrive after the COM write so it uploads the changes."""
    if not os.path.exists(ONEDRIVE_EXE):
        return
    onedrive_start()
    log("[ONEDRIVE] OneDrive restarted — changes uploading to SharePoint.")

def clean_slate():
    """Kills any stuck Word/Excel/Git processes to ensure a fresh start.
    Safe to call at midnight — no user is active during the nightly run."""
    log("[CLEANUP] Cleaning up stuck processes (Word/Excel/Git)...")
    try:
        # Use guarded kill functions from sys_process_utils
        sys_process_utils.kill_excel_force()
        sys_process_utils.kill_word_force()
        os.system("taskkill /F /IM git.exe /T >nul 2>&1")
        # Give processes time to fully close
        time.sleep(3)
    except Exception as e:
        log(f"[WARNING] Process cleanup failed: {e}")

def log(msg):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    formatted_msg = f"{timestamp} - {msg}"
    try:
        print(formatted_msg, flush=True)
    except Exception:
        pass
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        log_file = os.path.join(LOG_DIR, f"nightly_{datetime.datetime.now().strftime('%Y-%m-%d')}.log")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(formatted_msg + "\n")
    except:
        pass

def backup_excel_files():
    """Create daily backups of Excel files before updating."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    
    warnings = []
    
    # Backup Jerusalem Excel
    if os.path.exists(JERUSALEM_EXCEL):
        backup_name = f"jerusalem_backup_{today}.xlsx"
        backup_path = os.path.join(BACKUP_DIR, backup_name)
        if not os.path.exists(backup_path):  # Don't overwrite today's backup
            try:
                shutil.copy2(JERUSALEM_EXCEL, backup_path)
                log(f"[SAVE] Backup created: {backup_name}")
            except Exception as e:
                log(f"[WARNING] Failed to backup Jerusalem Excel: {e}")
                warnings.append(f"Backup Failed (Jerusalem): {e}")
    
    # Backup South Excel
    if SOUTH_EXCEL and os.path.exists(SOUTH_EXCEL):
        backup_name = f"south_backup_{today}.xlsx"
        backup_path = os.path.join(BACKUP_DIR, backup_name)
        if not os.path.exists(backup_path):
            try:
                shutil.copy2(SOUTH_EXCEL, backup_path)
                log(f"[SAVE] Backup created: {backup_name}")
            except Exception as e:
                log(f"[WARNING] Failed to backup South Excel: {e}")
                warnings.append(f"Backup Failed (South): {e}")
    else:
        log("[WARNING] South Excel drive not found! South cases will not be updated tonight.")
        warnings.append("[WARNING] כונן דרום לא נמצא! תיקי דרום לא יעודכנו הלילה.")
    
    # Cleanup old backups (keep last 7 days)
    try:
        for f in os.listdir(BACKUP_DIR):
            fpath = os.path.join(BACKUP_DIR, f)
            if os.path.isfile(fpath):
                age_days = (datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getmtime(fpath))).days
                if age_days > 7:
                    os.remove(fpath)
                    log(f"[CLEANUP] Removed old backup: {f}")
    except Exception as e:
        log(f"[WARNING] Error cleaning up old backups: {e}")
        
    return warnings

def get_daily_sent_emails():
    """Retrieve list of emails actually sent today from the logs."""
    sent_list = []
    json_path = os.path.join(LOG_DIR, "sent_emails.json")
    if not os.path.exists(json_path):
        return []
    
    # The nightly script runs at 00:00, so we want to collect emails sent "yesterday".
    target_date_str = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # Filter for items sent strictly yesterday (which makes up the whole work day)
            for item in data:
                sent_at = item.get("sent_at", "")
                if sent_at.startswith(target_date_str):
                    sent_list.append(item)
    except Exception as e:
        log(f"[WARNING] Error reading sent_emails.json: {e}")
        
    return sent_list

def send_email_report(report_data, html_attachment=None):
    """Send a RICH HTML email report via Outlook."""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_TO
        mail.Subject = f"דוח אוטומציה לילי - {datetime.datetime.now().strftime('%d/%m/%Y')}"
        
        # Unpack data
        success = report_data.get('success', [])
        failed = report_data.get('failed', [])
        missing = report_data.get('missing', [])
        warnings = report_data.get('warnings', [])
        audit_results = report_data.get('audit_results', [])
        elyon_stats = report_data.get('elyon', {'added_count': 0, 'details': []})
        emails_sent = report_data.get('emails_sent', [])
        unlisted_cases = report_data.get('unlisted', [])
        delayed = report_data.get('delayed', [])
        approve_url = report_data.get('approve_url', 'http://localhost:5050/approve')
        
        # CSS Styles
        style_box = "background: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef; margin-bottom: 20px;"
        style_table = "width: 100%; border-collapse: collapse; font-size: 14px;"
        style_th = "background: #34495e; color: white; padding: 10px; text-align: right;"
        style_td = "padding: 8px; border-bottom: 1px solid #ddd;"
        
        missing_html = ""
        if missing:
            attachment_banner = ""
            if html_attachment:
                attachment_banner = """
                <p>יש לערוך אישור על ידי מנהל לדיונים שלא נמצאו בשרת:</p>
                <div style="background: #e67e22; color: white; padding: 12px 25px; border-radius: 6px; font-weight: bold; display: inline-block; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                    [ATTACHMENT] פתח את הקובץ המצורף (Manual_Review_Form.html)
                </div>"""
                
            # Safely escape text lines
            missing_lines_html = "".join([f"<li>{str(m).replace('<', '&lt;').replace('>', '&gt;')}</li>" for m in missing])
            
            missing_html = f"""
            <div style="text-align: center; background: #fff9e6; padding: 20px; border: 1px dashed #e67e22; border-radius: 8px;">
                <h4 style="margin-top: 0; color: #d35400;">[EXCLAMATION] נמצאו הודעות מסנכרון יומי ({len(missing)} שורות)</h4>
                <div style="text-align: right; background: white; padding: 10px; margin-top: 10px; margin-bottom: 15px; border: 1px solid #ccc; max-height: 200px; overflow-y: auto;">
                    <ul style="margin: 0; padding-right: 20px; font-size: 13px; color: #333;">
                        {missing_lines_html}
                    </ul>
                </div>
                {attachment_banner}
            </div>
            """
        else:
            missing_html = '<p style="color: #27ae60;">הכל מסונכרן! [OK]</p>'

        # Build HTML Body
        html = f"""
        <html dir="rtl">
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background-color: #f4f6f9;">
            <div style="max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                
                <!-- HEADER -->
                <div style="text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 15px; margin-bottom: 20px;">
                    <h1 style="color: #2c3e50; margin: 0;">דוח סיכום יומי</h1>
                    <p style="color: #7f8c8d; font-size: 14px;">{datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                </div>

                <!-- SUMMARY STATS -->
                <div style="display: flex; justify-content: space-around; margin-bottom: 25px;">
                    <div style="text-align: center;">
                        <span style="font-size: 24px; color: #3498db;">[MAIL] {len(emails_sent)}</span><br>
                        <span style="font-size: 11px; color: #555;">מיילים שנשלחו</span>
                    </div>
                    <div style="text-align: center;">
                        <span style="font-size: 24px; color: #27ae60;">[SUCCESS] {len(success)}</span><br>
                        <span style="font-size: 11px; color: #555;">ספירת מילים</span>
                    </div>
                     <div style="text-align: center;">
                        <span style="font-size: 24px; color: #e67e22;">[WARNING] {len(missing)}</span><br>
                        <span style="font-size: 11px; color: #555;">חריגים</span>
                    </div>
                    <div style="text-align: center;">
                        <span style="font-size: 24px; color: #c0392b;">[ERROR] {len(failed)}</span><br>
                        <span style="font-size: 11px; color: #555;">כישלונות</span>
                    </div>
                     <div style="text-align: center;">
                        <span style="font-size: 24px; color: #8e44ad;">[SUPREME] {elyon_stats['added_count']}</span><br>
                        <span style="font-size: 11px; color: #555;">עליון (חדש)</span>
                    </div>
                     <div style="text-align: center;">
                        <span style="font-size: 24px; color: #d35400;">[DELAY] {len(delayed)}</span><br>
                        <span style="font-size: 11px; color: #555;">איחורי תמלול</span>
                    </div>
                </div>

                <!-- WARNINGS -->
                {f'<div style="background: #fff3cd; color: #856404; padding: 10px; border-radius: 5px; margin-bottom: 20px; border-right: 5px solid #ffc107;"><strong>[WARNING] התראות מערכת:</strong><ul>{"".join([f"<li>{w}</li>" for w in warnings])}</ul></div>' if warnings else ''}

                </div>

                <!-- EMAILS SENT TODAY -->
                <div style="{style_box}">
                    <h3 style="color: #3498db; border-bottom: 1px solid #ddd; padding-bottom: 5px;">
                        [MAIL] מיילים שנשלחו היום (ללקוחות / אפיקי שידור)
                    </h3>
                    {f'<table style="{style_table}"><thead><tr><th style="{style_th}">תיק</th><th style="{style_th}">תאריך דיון</th><th style="{style_th}">נמען</th><th style="{style_th}">מחוז</th></tr></thead><tbody>' + 
                     ''.join([f"<tr><td style='{style_td}'>{x.get('case_num','')}</td><td style='{style_td}'>{x.get('date','')}</td><td style='{style_td}'>{x.get('recipient','')}</td><td style='{style_td}'>{x.get('district','')}</td></tr>" for x in emails_sent]) + 
                     '</tbody></table>' if emails_sent else '<p style="color: #7f8c8d;">לא נשלחו מיילים חדשים היום.</p>'}
                </div>

                <!-- SUCCESS -->
                <div style="{style_box}">
                    <h3 style="color: #27ae60; border-bottom: 1px solid #ddd; padding-bottom: 5px;">
                        [SUCCESS] קבצים שעובדו (Word ספירת מילים)
                    </h3>
                    {f'<table style="{style_table}"><thead><tr><th style="{style_th}">קובץ</th><th style="{style_th}">תיק</th><th style="{style_th}">תאריך</th><th style="{style_th}">מילים</th></tr></thead><tbody>' + 
                     ''.join([f"<tr><td style='{style_td}'>{x.get('file','')}</td><td style='{style_td}'>{x.get('case','')}</td><td style='{style_td}'>{x.get('date','')}</td><td style='{style_td}'><b>{x.get('words','')}</b></td></tr>" for x in success]) + 
                     '</tbody></table>' if success else '<p style="color: #7f8c8d;">אין קבצים חדשים לעיבוד היום.</p>'}
                </div>
                
                 <!-- ELYON -->
                <div style="{style_box}">
                    <h3 style="color: #8e44ad; border-bottom: 1px solid #ddd; padding-bottom: 5px;">
                        [SUPREME] סריקת עליון (Supreme Court)
                    </h3>
                    <p style="margin:5px 0;">נוספו <b>{elyon_stats['added_count']}</b> דיונים חדשים לאקסל.</p>
                       {f'<ul style="font-size:13px; color:#555;">' + ''.join([f"<li>{d}</li>" for d in elyon_stats['details']]) + '</ul>' if elyon_stats['details'] else ''}
                </div>

                <!-- MISSING / SYNC ISSUES -->
                <div style="{style_box}">
                    <h3 style="color: #e67e22; border-bottom: 1px solid #ddd; padding-bottom: 5px;">
                         [SEARCH] פערי סנכרון (שרת/אקסל)
                    </h3>
                     {missing_html}
                </div>

                <!-- UNLISTED CASES (Requires Approval) -->
                {f'''
                <div style="background: #eaf2f8; padding: 15px; border-radius: 8px; border: 1px solid #3498db; margin-bottom: 20px;">
                    <h3 style="color: #2980b9; margin-top:0;">[INFO] תיקים חדשים שממתינים לאישור</h3>
                    <p style="font-size: 13px;">נמצאו <b>{len(unlisted_cases)}</b> תיקים חדשים בשרת (למשל מהעליון) שאינם מופיעים באקסל.</p>
                    <table style="{style_table}">
                        <thead><tr><th style="{style_th}">תיק</th><th style="{style_th}">תאריך</th><th style="{style_th}">שופט</th><th style="{style_th}">פעולה</th></tr></thead>
                        <tbody>{''.join([f"<tr><td style='{style_td}'>{c.get('case')}</td><td style='{style_td}'>{c.get('date')}</td><td style='{style_td}'>{c.get('judge')}</td><td style='{style_td}'><a href='{approve_url}&case={c.get('case')}' style='color:green; font-weight:bold; text-decoration:none;'>✔️ אשר</a> &nbsp;|&nbsp; <a href='{approve_url.replace('/approve', '/reject')}&case={c.get('case')}' style='color:red; font-weight:bold; text-decoration:none;'>❌ דחה</a></td></tr>" for c in unlisted_cases])}</tbody>
                    </table>
                    <div style="text-align:center; margin:15px 0;">
                        <a href="{approve_url}" style="display:inline-block; background:#27ae60; color:white; padding:12px 25px; border-radius:6px; text-decoration:none; font-weight:bold; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">אשר את כולם</a>
                        <a href="{approve_url.replace('/approve', '/reject')}" style="display:inline-block; background:#e74c3c; color:white; padding:12px 25px; border-radius:6px; text-decoration:none; font-weight:bold; box-shadow: 0 2px 4px rgba(0,0,0,0.2); margin-right:10px;">דחה את כולם</a>
                    </div>
                </div>
                ''' if unlisted_cases else ''}

                <!-- AUDIT RESULTS (Unsent files) -->
                {f'''
                <div style="background: #fff4f4; padding: 15px; border-radius: 8px; border: 1px solid #f5b7b1; margin-bottom: 20px;">
                    <h3 style="color: #c0392b; margin-top:0;">[AUDIT] קבצים בשרת שלא נשלחו (7 ימים אחרונים)</h3>
                    <p style="font-size: 12px; color: #666;">קבצים אלו נמצאים בתיקיות המחוזות אך לא נמצא תיעוד למשלוח שלהם.</p>
                    <table style="{style_table}">
                        <thead><tr><th style="background: #c0392b; color: white; padding: 8px;">קובץ</th><th style="background: #c0392b; color: white; padding: 8px;">סיבה / סטטוס</th></tr></thead>
                        <tbody>{''.join([f"<tr><td style='{style_td}'>{x.get('file','')}</td><td style='{style_td}'>{x.get('reason','')}</td></tr>" for x in audit_results[:50]])}</tbody>
                    </table>
                    {f'<p style="text-align:center; font-size:12px;">...ועוד {len(audit_results)-50} קבצים נוספים</p>' if len(audit_results) > 50 else ''}
                </div>
                ''' if audit_results else ''}

                <!-- DELAYED TRANSCRIPTIONS -->
                {f'''
                <div style="background: #fdf2e9; padding: 15px; border-radius: 8px; border: 1px solid #e67e22; margin-bottom: 20px;">
                    <h3 style="color: #d35400; margin-top:0;">[DELAY] התרעת דיונים שעבר זמן תמלולם</h3>
                    <p style="font-size: 13px; color: #555;">תיקים אלו חצו את מועד הגשת התמלול (10 ימי עסקים) או נמצאים בטווח של יומיים לפניו, וטרם הוזנה להם ספירת מילים.</p>
                    <table style="{style_table}">
                        <thead><tr><th style="background: #e67e22; color: white; padding: 8px;">תיק</th><th style="background: #e67e22; color: white; padding: 8px;">תאריך דיון</th><th style="background: #e67e22; color: white; padding: 8px;">תאריך יעד</th><th style="background: #e67e22; color: white; padding: 8px;">שופט</th><th style="background: #e67e22; color: white; padding: 8px;">מחוז</th><th style="background: #e67e22; color: white; padding: 8px;">סטטוס</th></tr></thead>
                        <tbody>{''.join([f"<tr><td style='{style_td}'>{c.get('case')}</td><td style='{style_td}'>{c.get('date')}</td><td style='{style_td}'>{c.get('due_date')}</td><td style='{style_td}'>{c.get('judge')}</td><td style='{style_td}'>{c.get('region')}</td><td style='{style_td}; font-weight:bold; color:#d35400;'>{c.get('status_desc')}</td></tr>" for c in delayed])}</tbody>
                    </table>
                </div>
                ''' if delayed else ''}

                <!-- FAILURES -->
                {f'''
                <div style="background: #fadbd8; padding: 15px; border-radius: 8px; border: 1px solid #f5b7b1; margin-bottom: 20px;">
                    <h3 style="color: #c0392b; margin-top:0;">[ERROR] כישלונות בעיבוד קבצים (מערכת)</h3>
                    <table style="{style_table}">
                        <thead><tr><th style="background: #c0392b; color: white; padding: 8px;">קובץ</th><th style="background: #c0392b; color: white; padding: 8px;">סיבה</th></tr></thead>
                        <tbody>{''.join([f"<tr><td style='{style_td}'>{x.get('file','')}</td><td style='{style_td}'>{x.get('reason','')}</td></tr>" for x in failed])}</tbody>
                    </table>
                </div>
                ''' if failed else ''}

                <div style="text-align: center; margin-top: 30px; font-size: 11px; color: #bdc3c7;">
                    <hr>
                    דוח זה נוצר אוטומטית ע"י 'AutoBot'. נא לא להשיב.
                </div>
            </div>
        </body>
        </html>
        """
        
        mail.HTMLBody = html
        if html_attachment and os.path.exists(html_attachment):
            mail.Attachments.Add(html_attachment)
        mail.Send()
        log(f"[MAIL] Report sent to {EMAIL_TO}")
        
    except Exception as e:
        log(f"[ERROR] Failed to send email: {e}")

def main_nightly():
    log("LOG ========================================")
    log("Starting Nightly Automation Job")
    log("LOG ========================================")
    
    clean_slate()  # Ensure no stuck apps
    
    report = {
        'success': [],
        'failed': [],
        'missing': [],
        'warnings': [],
        'audit_results': [],
        'elyon': {'added_count': 0, 'details': []},
        'emails_sent': [],
        'delayed': []
    }
    
    # Step 0: Backup Excel files
    log("\n--- Step 0: Backing up Excel files ---")
    warnings = backup_excel_files()
    report['warnings'] = warnings

    # Step 0.5: Collect today's emails
    log("\n--- Step 0.5: Collecting sent emails log ---")
    report['emails_sent'] = get_daily_sent_emails()
    
    # Step 1: Process Word files (count words + update Excel)
    log("\n--- Step 1: Processing Word files ---")
    prepare_excel_for_write()
    ok1, results = run_step_isolated(
        "step1_wordcount", "mail_word_count_prod",
        "mail_word_count_prod.run_batch_with_stats()",
        timeout=30*60,
        default={'success': [], 'failed': [], 'skipped_yoel': []}
    )
    if ok1 and results:
        report['success'] = results.get('success', [])
        report['failed'] = results.get('failed', [])
        if any(isinstance(f, dict) and 'READ_ONLY' in str(f.get('reason', '')) for f in report['failed']):
            report['warnings'].append("[LOCK] Jerusalem Excel was READ-ONLY during nightly run — word counts NOT saved. Close Excel before midnight!")
        if any('SAVE_ERROR' in str(f) for f in results.get('success', []) + results.get('failed', [])):
            report['warnings'].append("[ONEDRIVE_CONFLICT] ⚠️ אקסל ירושלים לא נשמר — בעיית סנכרון OneDrive! ספירות המילים נשמרו זמנית ב-pending_word_counts.json. יש לפתוח את קובץ האקסל, לפתור את הקונפליקט (כפתור 'שמור עותק' או 'בטל שינויים'), ולהריץ מחדש.")
    else:
        report['failed'].append({'file': 'STEP_1_SYSTEM_ERROR', 'reason': 'Subprocess failed or timed out'})

    # Step 1.5: Process Excel Queue (Queue Manager)
    log("\n--- Step 1.5: Process Excel Queue Manager ---")
    run_step_isolated(
        "step1_5_queue", "process_excel_queue",
        "process_excel_queue.process_queue()",
        timeout=10*60, default=None
    )


    # Step 1.6: בדיקת מיילים ישנים שממתינים בתור
    log("\n--- Step 1.6: Checking unsent email queue ---")
    ok1_6, unsent_warnings = run_step_isolated(
        "step1_6_queue", "send_district_emails",
        "send_district_emails.check_unsent_old_hearings()",
        timeout=2*60, default=[]
    )
    if ok1_6 and unsent_warnings:
        for w in unsent_warnings:
            report['warnings'].append(w)

    # Step 1.65: Pending drafts report — auto-detect sent drafts and show what's still waiting
    log("\n--- Step 1.65: Pending Drafts Report ---")
    run_step_isolated(
        "step1_65_drafts", "pending_drafts_report",
        "pending_drafts_report.run()",
        timeout=3*60, default=None
    )

    # Step 1.7: Fix mismatched dates in Word protocol files
    log("\n--- Step 1.7: Auto-fixing mismatched dates in Word files ---")
    run_step_isolated(
        "step1_7_dates", "nightly_fix_dates_in_word",
        "nightly_fix_dates_in_word.check_and_fix_pending_files()",
        timeout=10*60, default=None
    )

    # Kill any residual Excel left by step 1 (a COM crash in mail_word_count_prod
    # leaves Excel.exe alive but corrupt; steps 2/2.7/3 would inherit that zombie
    # instance and fail immediately on first property access with OLE 0x800a01a8).
    sys_process_utils.kill_excel_force()
    time.sleep(2)

    # Step 2: CHECK Missing Rows (Jerusalem Sync)
    log("\n--- Step 2: Checking for missing rows (Jerusalem) ---")
    ok2, jerusalem_missing = run_step_isolated(
        "step2_jerusalem", "sync_missing_rows_prod",
        "sync_missing_rows_prod.check_missing_with_return()",
        timeout=15*60, default=[]
    )
    if ok2 and jerusalem_missing:
        report['missing'].extend(["--- ירושלים ---"])
        report['missing'].extend(jerusalem_missing)
    elif not ok2:
        report['warnings'].append("שגיאה בסנכרון ירושלים (timeout או קריסה)")
        
    # Step 2.5: הוסר — extract_court_data מוחלף על-ידי find_unlisted_and_notify (Step 2.7)
    # הסיבות: שימש pandas/openpyxl (אסור), כתב מחוץ לטבלה, לא דרש אישור, כיסה עליון בלבד

    # Step 2.7: Find Unlisted Cases + Send Approval Email
    log("\n--- Step 2.7: Searching for cases on server but not in Excel ---")
    ok2_7, unlisted = run_step_isolated(
        "step2_7_unlisted", "find_unlisted_and_notify",
        "find_unlisted_and_notify.find_and_notify(send_email=False)",
        timeout=10*60, default=[]
    )
    if ok2_7 and unlisted:
        report['unlisted'] = unlisted
        log(f"[INFO] Found {len(unlisted)} unlisted cases for unified report.")
        # Start approval server and build URL with token so the email link works
        try:
            import find_unlisted_and_notify as _fun
            _fun._start_server_background()
            import approval_server as _srv
            report['approve_url'] = _srv.build_approval_url()
            log(f"[WEB] Approval server started, URL built with token.")
        except Exception as _e:
            log(f"[WARNING] Could not start approval server: {_e}")
    elif not ok2_7:
        report['warnings'].append("שגיאה בחיפוש תיקים חסרים (timeout או קריסה)")
        
    # Step 2.9: Verification & Refill Word Counts (Jerusalem & Supreme)
    log("\n--- Step 2.9: Verifying and refilling missing word counts (Jerusalem & Supreme) ---")
    ok2_9, _ = run_step_isolated(
        "step2_9_refill", "verify_and_refill_word_counts",
        "verify_and_refill_word_counts.run_verification()",
        timeout=10*60, default=None
    )

    # All Excel writes are done — restart OneDrive so it uploads all changes
    resume_onedrive()

    sys_process_utils.kill_excel_force()
    time.sleep(2)

    # Step 3: South Sync (Transfer + Excel Update)
    log("\n--- Step 3: South Sync (Transfer + Excel) ---")
    ok3, south_report = run_step_isolated(
        "step3_south", "sync_south_production",
        "sync_south_production.check_south_with_return()",
        timeout=15*60, default=[]
    )
    if ok3 and south_report:
        report['missing'].extend(["--- דרום ---"])
        report['missing'].extend(south_report)
    elif not ok3:
        report['warnings'].append("שגיאה בסנכרון דרום (timeout או קריסה)")

    # Step 3.2: South Word Counts
    log("\n--- Step 3.2: South Word Counts ---")
    ok3_2, _ = run_step_isolated(
        "step3_2_south_words", "final_fill_missing_south",
        "final_fill_missing_south.main_fill()",
        timeout=20*60, default=None
    )
    if not ok3_2:
        report['warnings'].append("שגיאה בספירת מילים דרום (timeout או קריסה)")

    # Step 3.4: Comprehensive Audit (Daily check for unsent files)
    log("\n--- Step 3.4: Comprehensive Audit (Checking for unsent files) ---")
    ok3_4, unsent_files = run_step_isolated(
        "step3_4_audit", "comprehensive_audit",
        "comprehensive_audit.run_comprehensive_audit_final()",
        timeout=10*60, default=[]
    )
    if ok3_4 and unsent_files:
        filtered_unsent = [
            uf for uf in unsent_files
            if not any(x in uf['file'].lower() for x in ['verbatim', 'protocol', 'formatted', 'shot', 'temp'])
        ]
        report['audit_results'] = filtered_unsent
        if filtered_unsent:
            report['warnings'].append(f"[AUDIT] נמצאו {len(filtered_unsent)} קבצים בשרת שלא נשלחו (ראו פירוט בטבלה למטה)")

    # Step 3.5: Generate Manual Review Report (Excel Form)
    html_path = None
    try:
        # Extract structured data from the text report
        structured_missing = []
        for line in report['missing']:
            # Format: [WARNING] לא נמצא בשרת: 12345-01-26 | 2026-02-15 | שופט (שורה 10)
            if "לא נמצא בשרת" in line and "|" in line:
                try:
                    parts = line.split('|')
                    case_part = parts[0].split(':')[-1].strip()
                    if len(parts) >= 3:
                        date_part = parts[1].strip()
                        judge_part = parts[2].split('(')[0].strip()
                    else:
                        date_part = parts[1].split('(')[0].strip()
                        judge_part = ""
                    structured_missing.append({'case': case_part, 'date': date_part, 'judge': judge_part})
                except:
                    pass
        
        if structured_missing:
            # Generate Excel Report (Office Use) - legacy support
            import generate_missing_report
            report_path = generate_missing_report.generate_report(structured_missing)
            if report_path:
                    report['warnings'].append(f"[FILE] נוצר דוח אקסל לטיפול: {report_path}")

            # Prepare enriched data for the hosted HTML form
            enriched_missing = []
            for item in structured_missing:
                # Find row number from the original report line
                row_val = 0
                for line in report['missing']:
                    if item['case'] in line and "שורה" in line:
                        try:
                            row_val = int(line.split('שורה')[-1].split(')')[0].strip())
                            break
                        except: pass
                
                enriched_missing.append({
                    'case': item['case'],
                    'date': item['date'],
                    'row': row_val,
                    'judge': item.get('judge', '')
                })

            if enriched_missing and report_path:
                # Generate HTML Form (Office Use as an Attachment)
                import generate_html_report_form
                html_path = os.path.join(os.path.dirname(report_path), "Manual_Review_Form.html")
                if generate_html_report_form.generate_html_form(enriched_missing, html_path):
                    report['warnings'].append(f"[WEB] קובץ עריכת סטטוס נוצר ויצורף למייל: {html_path}")

    except Exception as e:
        log(f"[WARNING] Failed to generate manual report: {e}")

    # Step 3.6: Check for delayed transcriptions
    log("\n--- Step 3.6: Check for delayed transcriptions ---")
    ok3_6, delayed_cases = run_step_isolated(
        "step3_6_delayed", "check_delayed_transcriptions",
        "check_delayed_transcriptions.find_delayed_cases()",
        timeout=10*60, default=[]
    )
    if ok3_6 and delayed_cases:
        report['delayed'] = delayed_cases
        report['warnings'].append(f"[DELAY] נמצאו {len(delayed_cases)} דיונים שעברו או מתקרבים למועד הגשתם וטרם תומללו!")

    # Step 4: Send Email Report
    log("\n--- Step 4: Sending report ---")
    
    html_attachment = html_path if html_path and os.path.exists(html_path) else None
    
    send_email_report(report, html_attachment=html_attachment)
    
    # Step 5: Close Excel (cleanup)
    log("\n--- Step 5: Closing Excel (Hard Cleanup) ---")
    try:
        from sys_process_utils import kill_excel_force
        kill_excel_force()
        log("[SUCCESS] Excel processes killed.")
    except Exception as e:
        log(f"[WARNING] Could not close Excel: {e}")
    
    # Step 6: Transcribe new cases — DISABLED
    log("\n--- Step 6: Transcribing new cases (SKIPPED — disabled) ---")

    log("LOG ========================================")
    log("Nightly Job Complete!")
    log("LOG ========================================")

if __name__ == "__main__":
    try:
        main_nightly()
    except Exception as e:
        import traceback
        err_msg = str(e)
        st = traceback.format_exc()
        try:
            log(f"[FATAL] run_nightly_job.py crashed: {err_msg}")
        except:
            pass
        import crash_notifier
        crash_notifier.send_crash_notification("run_nightly_job.py", err_msg, st)
        raise e
