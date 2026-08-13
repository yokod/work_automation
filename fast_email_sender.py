import os
import re
import datetime
import shutil
import sys

# Force UTF-8 for console output
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    pass
import pythoncom
import win32com.client as win32
import file_validator
import gsheets_utils
import send_district_emails
import config_drive_paths
from mail_word_count_prod import bump_fail_count, clear_fail_count, MAX_FILE_ATTEMPTS

# --- Configuration ---
INPUT_FOLDER = r"C:\Users\yoel\OneDrive - Hever\Jerusalem"
PENDING_WC_FOLDER = os.path.join(INPUT_FOLDER, "Pending_Word_Count")
REVIEW_FOLDER = os.path.join(INPUT_FOLDER, "Requires_Manual_Review")
OTHER_REGION_FOLDER = os.path.join(INPUT_FOLDER, "Pending_Other_Region")
ALREADY_SENT_DIR = os.path.join(INPUT_FOLDER, "Already_Sent_To_Skip")
DUPLICATES_FOLDER = os.path.join(INPUT_FOLDER, "Duplicates_Skipped")

DRY_RUN = False  # ACTIVE
MAX_FILES = 20

def log(msg):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    formatted_msg = f"{timestamp} - {msg}"
    try:
        print(formatted_msg, flush=True)
    except:
        pass
    
    try:
        log_file = r"d:\yoel\projects\auto\logs\fast_sender_log.txt"
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(formatted_msg + "\n")
    except:
        pass

def safe_move(src_path, dest_folder):
    import time
    if not os.path.exists(src_path): return
    filename = os.path.basename(src_path)
    if not os.path.exists(dest_folder): os.makedirs(dest_folder)
    dest_path = os.path.join(dest_folder, filename)
    
    # Retry logic for Windows file locking
    for i in range(3):
        try:
            if os.path.exists(dest_path) and os.path.abspath(src_path) != os.path.abspath(dest_path):
                # File already exists at destination  move src to Duplicates instead of silently deleting it
                dupes_folder = os.path.join(os.path.dirname(dest_folder), "Duplicates_Skipped")
                os.makedirs(dupes_folder, exist_ok=True)
                shutil.move(src_path, os.path.join(dupes_folder, filename))
                log(f"   [DUPE] Destination already exists, moved to Duplicates_Skipped: {filename}")
            else:
                shutil.move(src_path, dest_path)
            return # Success
        except Exception as e:
            if i < 2:
                time.sleep(1) # Wait a second for COM to release handle
                continue
            log(f"   [ERROR] safe_move failed after retries: {e}")

def extract_case_check_regex(text):
    if not text: return None, None
    # 1. Supreme Court Prefix check (growskopf files often are Supreme)
    supreme_prefixes = r'ע"א|בג"ץ|ער"מ|בש"פ|דנ"א|רע"א|רע"פ|עע"מ|ערעור אזרחי|ערעור פלילי|ערעור מינהלי|רשות ערעור|ע"פ|עע"מ|עב"ל'
    supreme_m = re.search(fr'(?:{supreme_prefixes})\s*(?:מס(?:\'|פר)?\s+)?(\d{{2,6}})[/-](\d{{2,4}})', text)
    if supreme_m:
        g1, g2 = supreme_m.groups()
        return f"{g1}-{g2}", supreme_m.group(0)

    # 2. General patterns including underscore (robust logic from mail_word_count_prod)
    patterns = [
        r'(\d{2,6})[-/_](\d{1,2})[-/_](\d{2,6})',  # 12345-01-23 or 23_01_12345
        r'(\d{2,6})[/-](\d{2,5})(?![/-]\d)',        # 12345-23
        r'(\d{2,4})[_/-](\d{3,8})'                   # 24_854
    ]
    
    candidates = []
    for pat in patterns:
        for match in re.finditer(pat, text):
            candidates.append(match)
    
    if not candidates:
        return None, None
        
    candidates.sort(key=lambda m: m.start())
    match = candidates[0]
    val = match.group(0)
    
    parts = re.split(r'[-/_]', val)
    if len(parts) == 3:
        # Check if year is first or last
        if len(parts[0]) <= 2 and len(parts[2]) >= 3:
            normalized = f"{parts[2]}-{int(parts[1]):02d}-{parts[0]}"
        else:
            normalized = f"{parts[0]}-{int(parts[1]):02d}-{parts[2]}"
    else:
        p1, p2 = parts[0], parts[1]
        # YY_NUMBER -> NUMBER-YY
        if len(p1) <= 2 and len(p2) >= 3:
            normalized = f"{p2}-{p1}"
        else:
            normalized = f"{p1}-{p2}"
            
    return normalized, val


def quick_extract_text(file_path):
    """Extract ONLY headers + first page text from a Word file.
    Deliberately avoids reading the full document to prevent party address
    mentions (in the body) from causing false district detection.
    """
    import pythoncom
    import win32com.client as win32
    import gc
    
    pythoncom.CoInitialize()
    word = None
    doc = None
    text = ""
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0 
        
        doc = word.Documents.Open(file_path, ReadOnly=True, AddToRecentFiles=False)
        # 1. Read all section headers (court name, judge name almost always here)
        try:
            for section in doc.Sections:
                for header in section.Headers:
                    h = (header.Range.Text or "")[:500]
                    if h.strip(): text += " " + h
        except: pass
        # 2. Read only the first ~2000 chars of the document body (approx. first page)
        # This is where the court caption always appears. Party details come much later.
        try:
            end = min(2000, doc.Range().End)
            body_start = doc.Range(0, end).Text
            text += " " + (body_start or "")
        except:
            try:
                text += " " + (doc.Range().Text or "")[:2000]
            except: pass
            
        # 3. Read the LAST ~1500 chars (for completion validation)
        try:
            doc_end = doc.Range().End
            if doc_end > 2500:
                text += "\n\n... [GAP] ...\n\n"
                tail_start = max(2000, doc_end - 1500)
                text_tail = doc.Range(tail_start, doc_end).Text
                text += (text_tail or "")
        except: pass
            
        if not text.strip() or len(text.strip()) < 10:
            log(f"   [WARN] extracted text is too short or empty for {os.path.basename(file_path)}")
    except Exception as e:
        log(f"   [WARN] quick_extract_text failed for {os.path.basename(file_path)}: {e}")
    finally:
        try:
            if doc:
                doc.Close(False)
                del doc
        except: pass
        try:
            if word:
                word.Quit()
                del word
        except: pass
        gc.collect()
    return text.strip() or ""

def get_yoel_folder_cases():
    cases = set()
    path = r"F:\ירושלים\דיונים ליואל"
    if os.path.exists(path):
        try:
            for f in os.listdir(path):
                case, _ = extract_case_check_regex(f)
                if case: cases.add(case)
        except: pass
    return cases

def get_open_dates_from_excel(case_numbers: list) -> dict:
    hearing_details = {c: {} for c in case_numbers}
    if not case_numbers: return hearing_details
    pythoncom.CoInitialize()
    excel = None
    created_excel = False
    try:
        excel = win32.DispatchEx("Excel.Application")
        created_excel = True
        try: excel.DisplayAlerts = False
        except: pass
        
        for region in ["jerusalem", "south"]:
            path = config_drive_paths.get_active_excel_path(region)
            if not path or not os.path.exists(path): continue
            wb = None
            try:
                wb = excel.Workbooks.Open(path, ReadOnly=True)
                sheet_name = "2023" if region == "jerusalem" else "שרות א באר שבע ודרום"
                try:
                    ws = wb.Sheets(sheet_name)
                except:
                    ws = wb.Sheets(1)
                data = ws.UsedRange.Value
                if not data: continue
                
                # Determine header row dynamically
                headers = None
                header_row_idx = 0
                for r_idx in range(min(5, len(data))):
                    row_data = data[r_idx]
                    if any(row_data[i] and ("תיק" in str(row_data[i]) or "Case" in str(row_data[i])) for i in range(len(row_data)) if row_data[i]):
                        headers = row_data
                        header_row_idx = r_idx
                        break
                
                if not headers:
                    headers = data[0]
                    header_row_idx = 0
                
                c_idx = next((i for i, h in enumerate(headers) if h and ("תיק" in str(h) or "Case" in str(h))), None)
                d_idx = next((i for i, h in enumerate(headers) if h and ("תאריך" in str(h) or "Date" in str(h))), None)
                s_idx = next((i for i, h in enumerate(headers) if h and ("מילים" in str(h) or "Words" in str(h))), None)
                st_idx = next((i for i, h in enumerate(headers) if h and ("סטטוס" in str(h) or "Status" in str(h) or "בוטל" in str(h))), None)
                del_idx = next((i for i, h in enumerate(headers) if h and ("תאריך מסירה" in str(h) or "נשלח ללקוח" in str(h))), None)
                
                if c_idx is not None and d_idx is not None:
                    start_row = ws.UsedRange.Row
                    start_col = ws.UsedRange.Column
                    for r_offset, row in enumerate(data[header_row_idx + 1:]):
                        if len(row) <= max(c_idx, d_idx): continue
                        raw_c = str(row[c_idx]).strip() if row[c_idx] else ""
                        case_clean = raw_c.replace("/", "-").lstrip("0").strip()
                        if case_clean in hearing_details:
                            val = row[d_idx]
                            date_obj = None
                            if isinstance(val, datetime.datetime): date_obj = val.date()
                            elif hasattr(val, "date"): date_obj = val.date()
                            elif val:
                                s = str(val).strip().split(" ")[0]
                                s = s.replace('\u200e', '').replace('\u200f', '')
                                for fmt in ["%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%d.%m.%y", "%Y.%m.%d"]:
                                    try:
                                        date_obj = datetime.datetime.strptime(s, fmt).date()
                                        break
                                    except: continue
                            
                            if date_obj:
                                wc = row[s_idx] if s_idx is not None and len(row) > s_idx and row[s_idx] else 0
                                status_val = str(row[st_idx]).strip() if (st_idx is not None and len(row) > st_idx and row[st_idx]) else "לא בוטל"
                                delivered_val = str(row[del_idx]).strip() if (del_idx is not None and len(row) > del_idx and row[del_idx]) else ""
                                is_delivered = bool(delivered_val and delivered_val not in ("0", "None", "False", "nan"))
                                
                                is_canceled = False
                                if region == "south":
                                    try:
                                        excel_row_num = start_row + header_row_idx + 1 + r_offset
                                        cell_color = float(ws.Cells(excel_row_num, start_col + c_idx).Interior.Color)
                                        cell_color_index = int(ws.Cells(excel_row_num, start_col + c_idx).Interior.ColorIndex)
                                        # If the cell has a color that is not white (16777215.0) and not no-fill (-4142 or 2), consider it canceled
                                        if cell_color != 16777215.0 and cell_color_index not in [-4142, 2]:
                                            is_canceled = True
                                    except Exception as color_err:
                                        pass
                                
                                is_open = (wc == 0 or wc == "" or str(wc).strip() == "0") and (status_val == "לא בוטל" or not status_val) and not is_delivered and not is_canceled
                                hearing_details[case_clean][date_obj] = {
                                    'region': region,
                                    'is_open': is_open,
                                    'is_canceled': is_canceled,
                                    'district': None
                                }
            except Exception as inner_e:
                log(f"   [WARN] Failed to process {region} workbook: {inner_e}")
            finally:
                if wb: wb.Close(False)
    except Exception as e:
        log(f"   [ERROR] Excel scan failed: {e}")
    finally:
        if created_excel and excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()
    return hearing_details

def run_fast_sender():
    import pythoncom
    pythoncom.CoInitialize()
    try:
        import sys_process_utils
        sys_process_utils.kill_invisible_excel()
    except:
        pass
    # Force draft mode for this sender (was previously at module load — moved
    # here so importing this module doesn't mutate send_district_emails state).
    send_district_emails.ENABLE_DRAFT_MODE = False
    send_district_emails.ENABLE_CLIENT_EMAILS = True
    try:
        log(">>> Fast Email Sender starting ...")
        if not os.path.exists(INPUT_FOLDER):
            log(f"Error: INPUT_FOLDER not found and is not a directory: {INPUT_FOLDER}")
            return

        files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
        if not files: return
        if MAX_FILES: files = files[:MAX_FILES]

        cases_dict = {}
        for f in files:
            case_num, r_match = extract_case_check_regex(f)
            if not case_num:
                log(f"   [WARN] Could not identify case number — likely another department, moving: {f}")
                safe_move(os.path.join(INPUT_FOLDER, f), OTHER_REGION_FOLDER)
                continue
            if case_num not in cases_dict: cases_dict[case_num] = []
            cases_dict[case_num].append({'file': f, 'path': os.path.join(INPUT_FOLDER, f), 'case_num': case_num})

        hearing_details = get_open_dates_from_excel(list(cases_dict.keys()))

        for case_num, items in cases_dict.items():
            for item in items:
                f = item['file']
                f_path = item['path']
                try:  # --- Per-file isolation ---
                    # --- Retry-limit gate ---
                    from mail_word_count_prod import _load_fail_counter
                    fail_count = _load_fail_counter().get(f, 0)
                    if fail_count >= MAX_FILE_ATTEMPTS:
                        log(f"   [RETRY-LIMIT] {f} has failed {fail_count} times. Moving to Manual Review.")
                        safe_move(f_path, REVIEW_FOLDER)
                        continue

                    log(f"Processing: {f}")
                    
                    txt_content = quick_extract_text(f_path)
                    
                    # Identification logic - search for all dates and pick the one most likely to be the hearing date
                    dates = list(re.finditer(r'(\d{1,2})[\./-](\d{1,2})[\./-](\d{2,4})', f))
                    date_obj = None
                    if dates:
                        # Filter out dates that are substrings of the case number (to avoid matching 15-04-21 inside 16115-04-21)
                        valid_date_matches = []
                        for m in dates:
                            m_str = m.group(0)
                            if case_num and m_str in case_num.replace("/", "-"):
                                continue
                            valid_date_matches.append(m)
                        
                        if valid_date_matches:
                            # Take the last valid date match (standard practice for these filenames)
                            last_match = valid_date_matches[-1]
                            try:
                                d, m, y = int(last_match.group(1)), int(last_match.group(2)), int(last_match.group(3))
                                if y < 100: y += 2000
                                # Basic validation to ensure it's not a garbage date
                                if 1 <= m <= 12 and 1 <= d <= 31:
                                    date_obj = datetime.date(y, m, d)
                            except: pass

                    if not date_obj:
                        filename_for_date = f.replace(r_match, "") if r_match else f
                        date_match_2part = re.search(r'(\d{1,2})[\./-](\d{1,2})', filename_for_date)
                        if date_match_2part:
                            try:
                                d, m = int(date_match_2part.group(1)), int(date_match_2part.group(2))
                                if 1 <= m <= 12 and 1 <= d <= 31:
                                    date_obj = datetime.date(datetime.date.today().year, m, d)
                            except: pass

                    if not date_obj and txt_content:
                        try:
                            snippet = txt_content[:1000]
                            content_date_match = re.search(r'(\d{1,2})[\./-](\d{1,2})[\./-](\d{2,4})', snippet)
                            if content_date_match:
                                d, m, y = int(content_date_match.group(1)), int(content_date_match.group(2)), int(content_date_match.group(3))
                                if y < 100: y += 2000
                                if 1 <= m <= 12 and 1 <= d <= 31:
                                    date_obj = datetime.date(y, m, d)
                        except: pass

                    # --- Chronological check ---
                    if date_obj and case_num in hearing_details:
                        open_earlier_dates = []
                        for edate, edetails in hearing_details[case_num].items():
                            # Ignore extremely old open hearings (older than 90 days) to avoid blocking on stale/forgotten Excel entries
                            if edate < date_obj and edetails.get('is_open') and (date_obj - edate).days <= 90:
                                open_earlier_dates.append(edate)
                        
                        if open_earlier_dates:
                            earliest = min(open_earlier_dates)
                            log(f"   [CHRONO-WAIT] {f} (date: {date_obj}) is waiting for earlier open hearing on {earliest}.")
                            wait_folder = os.path.join(INPUT_FOLDER, "Waiting_For_Previous_Hearing")
                            safe_move(f_path, wait_folder)
                            
                            try:
                                import win32com.client as win32
                                outlook = win32.Dispatch("Outlook.Application")
                                mail = outlook.CreateItem(0)
                                mail.To = "your_email@example.com"
                                mail.Subject = f"[המתנה לדיון קודם] תיק {case_num} - קובץ {f}"
                                mail.HTMLBody = f"""
                                <div dir="rtl" style="font-family: Arial; text-align: right;">
                                    <h2 style="color: #ff9800;">קובץ הועבר להמתנה (יש דיון קודם פתוח)</h2>
                                    <p><b>קובץ:</b> {f}</p>
                                    <p><b>תיק:</b> {case_num} | <b>תאריך הדיון בקובץ:</b> {date_obj}</p>
                                    <p><b>תאריך מוקדם שעדיין פתוח:</b> {earliest}</p>
                                    <hr>
                                    <p style="font-size: 13px; color: #666;">הקובץ הועבר לתיקיית: Waiting_For_Previous_Hearing</p>
                                </div>
                                """
                                mail.Save() # Creates a draft instead of sending
                            except Exception as e:
                                log(f"   [WARN] Failed to create draft email for chrono wait: {e}")
                            
                            continue

                    success_or_reason = False
                    try:
                        success_or_reason = send_district_emails.send_transcription_email(
                            file_path=f_path,
                            case_num=case_num,
                            date_obj=date_obj,
                            word_count=0,
                            text_content=txt_content,
                            filename=f
                        )
                    except Exception as se:
                        log(f"   [ERROR] send_transcription_email failed: {se}")
                        success_or_reason = "CRASH"

                    if DRY_RUN:
                        # DRY_RUN mode: log but do NOT move the file
                        log(f"   [DRY RUN] Finished processing {f}  file stays in place for real run.")
                    elif success_or_reason is True:
                        # Email was actually sent (or queued)  safe to archive
                        clear_fail_count(f)
                        safe_move(f_path, PENDING_WC_FOLDER)
                        log(f"   [OK] Sent and moved to Pending_Word_Count: {f}")

                    else:
                        reason = success_or_reason if isinstance(success_or_reason, str) else "UNKNOWN_FAILURE"
                        count, limit_hit = bump_fail_count(f)
                        log(f"   [FAIL] Email blocked for {f}. Reason: {reason}. Attempt {count}/{MAX_FILE_ATTEMPTS}.")
                        if not txt_content or len(txt_content.strip()) < 10:
                            reason = "TEXT_EXTRACTION_FAILED"
                            send_district_emails.send_manual_review_notification(f, case_num, str(date_obj) if date_obj else "", reason)
                        if reason == "VALIDATION_FAILED":
                            log(f"   [VALIDATION_FAILED] Moving {f} to Manual Review immediately.")
                            fail_folder = os.path.join(REVIEW_FOLDER, "Validation_Failed")
                            safe_move(f_path, fail_folder)
                        elif limit_hit:
                            log(f"   [RETRY-LIMIT] Moving {f} to Manual Review after {count} failed attempts.")
                            safe_move(f_path, REVIEW_FOLDER)
                        else:
                            log(f"   [RETRY] Will retry on next run ({MAX_FILE_ATTEMPTS - count} attempts remaining).")

                except Exception as file_err:  # --- Per-file exception handler ---
                    count, limit_hit = bump_fail_count(f)
                    log(f"   [ERROR] Crash processing '{f}' (attempt {count}/{MAX_FILE_ATTEMPTS}): {file_err}")
                    if limit_hit:
                        log(f"   [RETRY-LIMIT] Moving {f} to Manual Review after crash.")
                        safe_move(f_path, REVIEW_FOLDER)
        log("[QUEUE] Processing remaining items in email queue...")
        while send_district_emails.process_one_queued_email():
            pass

    except Exception as e:
        import traceback
        err_msg = str(e)
        st = traceback.format_exc()
        log(f"[FATAL] run_fast_sender crashed: {err_msg}")
        import crash_notifier
        crash_notifier.send_crash_notification("fast_email_sender.py", err_msg, st)
        raise e

if __name__ == "__main__":
    run_fast_sender()
