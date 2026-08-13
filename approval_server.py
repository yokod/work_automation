"""
approval_server.py
==================
שרת מקומי תמידי המאזין על http://localhost:5050.

תומך בשתי זרימות:
  GET  /approve?token=X  ← אישור תיקים חסרים מהמייל הלילי
  GET  /reject?token=X   ← דחיית תיקים חסרים
  POST /status-update    ← עדכון סטטוס מטופס HTML (דיונים לא בשרת)

שימוש:
  python approval_server.py   ← שרת תמידי (הפעל עם Windows)
"""

import os
import sys
import json
import datetime
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# ----
import sys
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    pass

PORT = 5050
TOKEN_FILE = r"d:\yoel\projects\auto\logs\approval_token.txt"
PENDING_FILE = r"d:\yoel\projects\auto\logs\pending_approval.txt"
DONE_FILE = r"d:\yoel\projects\auto\logs\approval_done.txt"
STATUS_DATA_FILE = r"d:\yoel\projects\auto\logs\status_update_data.json"

_server_instance = None


def generate_token():
    import secrets
    token = secrets.token_urlsafe(16)
    os.makedirs(os.path.dirname(TOKEN_FILE), exist_ok=True)
    with open(TOKEN_FILE, "w") as f:
        f.write(token)
    return token


def read_token():
    try:
        with open(TOKEN_FILE) as f:
            return f.read().strip()
    except:
        return None


def read_pending():
    try:
        with open(PENDING_FILE, encoding="utf-8") as f:
            return json.load(f)
    except:
        return []


def run_export_dashboard_data():
    import subprocess
    import sys
    try:
        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "export_dashboard_data.py")
        subprocess.run([sys.executable, script_path], check=True, capture_output=True)
        return True
    except Exception as e:
        print(f"Error running export_dashboard_data.py: {e}")
        return False


class ApprovalHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        pass  # Suppress default HTTP logs

    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)
        token_param = params.get("token", [""])[0]
        valid_token = read_token()

        # ── /approve ──────────────────────────────────────────
        if parsed.path == "/approve":
            if token_param != valid_token:
                self._send_html(403, self._page_error("❌ אסימון לא תקין. הכפתור אולי פג תוקף."))
                return

            items = read_pending()
            if not items:
                self._send_html(200, self._page_error("⚠️ לא נמצאו תיקים ממתינים. אולי כבר אושרו?"))
                return

            case_to_approve = params.get("case", [""])[0]
            if case_to_approve:
                items_to_approve = [it for it in items if it.get("case") == case_to_approve]
                items_remaining = [it for it in items if it.get("case") != case_to_approve]
            else:
                items_to_approve = items
                items_remaining = []

            if not items_to_approve:
                self._send_html(200, self._page_error(f"⚠️ התיק {case_to_approve} לא נמצא ברשימה. אולי כבר אושר או נדחה?"))
                return

            # Add to Excel in background
            result_holder = {"ok": False, "added": 0, "errors": []}
            self._execute_add(items_to_approve, result_holder)

            if result_holder["ok"]:
                with open(PENDING_FILE, "w", encoding="utf-8") as f:
                    json.dump(items_remaining, f, ensure_ascii=False, indent=2)
                html = self._page_success(items_to_approve, result_holder["added"])
            else:
                html = self._page_error(f"❌ שגיאה בהוספה לאקסל: {result_holder['errors']}")

            self._send_html(200, html)

            if not items_remaining:
                t = threading.Timer(3.0, self._shutdown)
                t.start()

        # ── /status-update (GET - should be POST, but some clients use GET) ──
        elif parsed.path == "/status-update-page":
            # Serve a simple page telling user to submit the form
            self._send_html(200, "<html><body>Use POST /status-update</body></html>")

        # ── /quick-status (Direct from Email) ─────────────────
        elif parsed.path == "/quick-status":
            row_str = params.get('row', [''])[0]
            status = params.get('status', [''])[0]
            duration = params.get('duration', [''])[0]
            note = params.get('note', [''])[0]
            
            if not row_str or not status:
                self._send_html(400, self._page_error("❌ חסרים נתוני שורה או סטטוס."))
                return
                
            try:
                row = int(row_str)
            except:
                self._send_html(400, self._page_error("❌ השורה חייבת להיות מספר."))
                return
            
            # If status is "בוטל עם חיוב" and duration is missing, show the tiny form!
            if status == "בוטל עם חיוב" and not duration:
                options = ''.join([f'<option value="{i/2}">{i/2}</option>' for i in range(1, 19)])
                html_form = f"""
                <html dir="rtl"><head><meta charset="utf-8"><title>הזנת שעות חיוב</title>
                <style>
                  body{{font-family:'Segoe UI',sans-serif;background:#f4f7f6;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;}}
                  .box{{background:white;padding:35px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.12);text-align:center;width:90%;max-width:400px;}}
                  h2{{color:#2c3e50;margin-top:0;margin-bottom:20px;}}
                  select, input{{width:100%;padding:12px;margin:10px 0;border:1px solid #ccc;border-radius:6px;font-size:16px;box-sizing:border-box;}}
                  select:focus, input:focus{{border-color:#3498db;outline:none;}}
                  button{{background:#27ae60;color:white;border:none;padding:14px 20px;font-size:16px;font-weight:bold;border-radius:6px;cursor:pointer;width:100%;margin-top:10px;transition:background 0.3s;}}
                  button:hover{{background:#219150;}}
                </style></head>
                <body><div class="box">
                  <h2>כמה שעות חיוב נפסקו?</h2>
                  <form action="/quick-status" method="GET">
                    <input type="hidden" name="row" value="{row}">
                    <input type="hidden" name="status" value="{status}">
                    <select name="duration" required>
                        <option value="" disabled selected>בחר מספר שעות...</option>
                        {options}
                    </select>
                    <input type="text" name="note" placeholder="הערות אישיות לחשבון (אופציונלי)">
                    <button type="submit">💾 שמור עדכון (שורה {row}) באקסל</button>
                  </form>
                </div></body></html>
                """
                self._send_html(200, html_form)
                return
                
            # Execute Update
            update_data = [{
                'row': row,
                'status': status,
                'duration': duration,
                'note': note
            }]
            
            res = self._process_status_update(update_data)
            if res.get("ok"):
                self._send_html(200, self._page_success([{'case': f'שורה {row}', 'date': '-', 'court': status}], res.get("updated", 1)))
            else:
                self._send_html(500, self._page_error(f"❌ שגיאה בעדכון שורה {row}: {res.get('error')}"))
            return

        # ── /reject ───────────────────────────────────────────
        elif parsed.path == "/reject":
            if token_param != valid_token:
                self._send_html(403, self._page_error("❌ אסימון לא תקין."))
                return
            
            items = read_pending()
            case_to_reject = params.get("case", [""])[0]
            
            # --- הוספת התיק לרשימת ההתעלמות ---
            def add_to_reject_list(case_num):
                if not case_num: return
                reject_log = r"d:\yoel\projects\auto\logs\rejected_cases.json"
                rejected = []
                try:
                    if os.path.exists(reject_log):
                        with open(reject_log, "r", encoding="utf-8") as f:
                            rejected = json.load(f)
                except: pass
                if case_num not in rejected:
                    rejected.append(case_num)
                try:
                    os.makedirs(os.path.dirname(reject_log), exist_ok=True)
                    with open(reject_log, "w", encoding="utf-8") as f:
                        json.dump(rejected, f, ensure_ascii=False, indent=2)
                except: pass
            
            if case_to_reject:
                items_remaining = [it for it in items if it.get("case") != case_to_reject]
                with open(PENDING_FILE, "w", encoding="utf-8") as f:
                    json.dump(items_remaining, f, ensure_ascii=False, indent=2)
                add_to_reject_list(case_to_reject)
                self._send_html(200, self._page_error(f"🚫 התיק {case_to_reject} נדחה והוסר מהרשימה. הוא לא יקפוץ יותר."))
                if not items_remaining:
                    t = threading.Timer(3.0, self._shutdown)
                    t.start()
            else:
                for it in items:
                    add_to_reject_list(it.get("case"))
                try:
                    os.remove(PENDING_FILE)
                except:
                    pass
                self._send_html(200, self._page_rejected())
                t = threading.Timer(3.0, self._shutdown)
                t.start()

        # ── /status-form (Hosted HTML Form) ───────────────────
        elif parsed.path == "/status-form":
            try:
                import generate_html_report_form
                if not os.path.exists(STATUS_DATA_FILE):
                    self._send_html(200, self._page_error("⚠️ לא נמצאו דיונים הממתינים לעדכון סטטוס."))
                    return
                
                with open(STATUS_DATA_FILE, "r", encoding="utf-8") as f:
                    cases_to_show = json.load(f)
                
                html = generate_html_report_form.get_form_html(cases_to_show)
                self._send_html(200, html)
            except Exception as e:
                self._send_html(500, self._page_error(f"❌ שגיאה בטעינת הטופס: {e}"))

        # ── /sync ─────────────────────────────────────────────
        elif parsed.path == "/sync":
            ok = run_export_dashboard_data()
            if not ok:
                self._send_json(500, {"ok": False, "error": "Failed to run export_dashboard_data.py"})
                return
            try:
                json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_data.json")
                with open(json_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                self._send_json(200, {"ok": True, "data": payload})
            except Exception as e:
                self._send_json(500, {"ok": False, "error": f"Failed to read JSON: {e}"})
            return

        else:
            self._send_html(404, self._page_error("404 - לא נמצא"))

    def do_OPTIONS(self):
        """מטפל בבקשות CORS preflight."""
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        """מטפל בשליחת טופס עדכון סטטוס (/status-update) או שיבוץ תיק (/assign)."""
        parsed = urlparse(self.path)

        if parsed.path == "/status-update":
            try:
                length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(length)
                import json
                data = json.loads(body.decode('utf-8'))
            except Exception as e:
                self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
                return

            # Process status updates
            result = self._process_status_update(data)
            if result["ok"]:
                self._send_json(200, result)
            else:
                self._send_json(500, result)
        elif parsed.path == "/assign":
            try:
                length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(length)
                import json
                data = json.loads(body.decode('utf-8'))
            except Exception as e:
                self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
                return

            # Process manual assignment
            result = self._process_manual_assignment(data)
            if result["ok"]:
                self._send_json(200, result)
            else:
                self._send_json(500, result)
        else:
            self._send_json(404, {"ok": False, "error": "Not found"})

    def _process_manual_assignment(self, data):
        """מבצע שיבוץ ידני של תיק למתמלל."""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        case = data.get("case")
        transcriber = data.get("transcriber")
        region = data.get("region")
        date_str = data.get("date")
        hours = data.get("hours", 0.0)
        judge = data.get("judge", "")

        if not case or not transcriber or not region:
            return {"ok": False, "error": "Missing case, transcriber, or region"}

        # Normalize region string
        region_clean = "south" if "דרום" in region or region.lower() == "south" else "jerusalem"

        try:
            from transcription_queue_manager import TranscriptionManager, normalize_date
            mgr = TranscriptionManager(region=region_clean)
            mgr.load_transcribers()

            matched_t = None
            import re
            def clean_name(s):
                s = re.sub(r'[\u200e\u200f\u202a-\u202e]', '', str(s))
                return re.sub(r'\s+', '', s).strip()
                
            t_search = clean_name(transcriber)
            for t_name, t_data in mgr.transcribers.items():
                if t_search == clean_name(t_name):
                    matched_t = t_data
                    break
            
            if not matched_t:
                return {"ok": False, "error": f"Transcriber '{transcriber}' not found in roster."}

            mgr.load_pending_hearings()
            
            cases_to_assign = [c.strip() for c in case.split(",") if c.strip()]
            success_count = 0
            failed_cases = []
            
            for sub_case in cases_to_assign:
                hearing = None
                for h in mgr.queue:
                    if h["case"].strip() == sub_case:
                        hearing = h
                        break
                        
                if not hearing:
                    if not date_str:
                        failed_cases.append(f"{sub_case} (Not found in pending, and no date)")
                        continue
                    hdate = normalize_date(date_str)
                    if not hdate:
                        failed_cases.append(f"{sub_case} (Invalid date)")
                        continue
                    try:
                        hours_val = float(hours) if len(cases_to_assign) == 1 else 0.0
                    except:
                        hours_val = 0.0
                    hearing = {
                        "case": sub_case,
                        "date": hdate,
                        "court": region_clean,
                        "hours": hours_val,
                        "urgency": "רגיל",
                        "deadline": hdate + datetime.timedelta(days=10),
                        "judge": judge
                    }
                else:
                    if len(cases_to_assign) == 1 and hours:
                        try:
                            hearing["hours"] = float(hours)
                        except:
                            pass
                    if judge and not hearing.get("judge"):
                        hearing["judge"] = judge

                success = mgr.execute_assignment(matched_t, hearing)
                if success:
                    success_count += 1
                else:
                    err_msg = getattr(mgr, "last_error", "שגיאה לא ידועה")
                    failed_cases.append(f"{sub_case} ({err_msg})")

            if success_count > 0:
                run_export_dashboard_data()
                if failed_cases:
                    return {"ok": True, "message": f"Successfully assigned {success_count} cases to {matched_t['name']}. Failed cases: {', '.join(failed_cases)}"}
                return {"ok": True, "message": f"Successfully assigned {success_count} cases to {matched_t['name']}"}
            else:
                return {"ok": False, "error": f"Assignment failed for all cases. Errors: {', '.join(failed_cases)}"}
        except Exception as e:
            import traceback
            return {"ok": False, "error": f"Exception: {str(e)}", "trace": traceback.format_exc()}

    def _process_status_update(self, items):
        """
        מעדכן אקסל לפי רשימת סטטוסים שהתקבלה מהטופס.
        items: [{case, date, row, status, duration, note}, ...]
        """
        excel = None
        wb = None
        created_excel = False
        opened_wb = False
        try:
            import win32com.client as win32
            import config_drive_paths as conf
            import os
            import sys_process_utils

            real_path = conf.JERUSALEM_EXCEL_PATH
            sys_process_utils.kill_excel_if_stuck()
            
            def _get_excel():
                try: return win32.GetActiveObject("Excel.Application")
                except Exception: return win32.Dispatch("Excel.Application")
                
            excel = sys_process_utils.robust_com_call(_get_excel)
            created_excel = False # Hard to track, assume we might need to quit later

            try:
                sys_process_utils.robust_com_call(lambda: setattr(excel, 'Visible', True))
            except Exception as e:
                print(f"   [WARN] Could not set excel.Visible: {e}")
            try:
                sys_process_utils.robust_com_call(lambda: setattr(excel, 'DisplayAlerts', False))
            except Exception as e:
                print(f"   [WARN] Could not set excel.DisplayAlerts: {e}")

            target_name = os.path.basename(real_path)
            try:
                for w in excel.Workbooks:
                    if w.Name == target_name:
                        wb = w
                        break
            except:
                pass

            if not wb:
                def _open_wb():
                    return excel.Workbooks.Open(real_path, UpdateLinks=0)
                wb = sys_process_utils.robust_com_call(_open_wb)
                opened_wb = True

            ws = wb.Sheets("2023")

            # Read headers
            headers = {}
            r_head = 2
            for r in [1, 2]:
                vals = ws.Range(ws.Cells(r, 1), ws.Cells(r, 40)).Value[0]
                if any(vals):
                    r_head = r
                    for i, v in enumerate(vals):
                        if v: headers[str(v).strip()] = i + 1
                    break

            col_status = headers.get("האם בוטל?") or headers.get("סטטוס") or 8
            col_duration = None
            for k, v in headers.items():
                if "אורך" in k and "שעות" in k:
                    col_duration = v
                    break
            if not col_duration:
                for k, v in headers.items():
                    if "אורך" in k or "שעות" in k:
                        col_duration = v
                        break

            # Map status text to Excel value
            STATUS_MAP = {
                "בוטל עם חיוב": "בוטל- עם חיוב",
                "בוטל ללא חיוב": "בוטל- ללא חיוב",
                "התקיים (חסר הקלטה)": "לא בוטל",
                "התקיים (יש הקלטה - בטיפול)": "לא בוטל",
                "נדחה": "נדחה",
                "אחר": "לברר",
            }

            updated = 0
            for item in items:
                row = item.get("row", 0)
                status_raw = item.get("status", "")
                duration = item.get("duration", "")
                note = item.get("note", "")

                if not row or not status_raw:
                    continue

                excel_status = STATUS_MAP.get(status_raw, status_raw)

                try:
                    ws.Cells(row, col_status).Value = excel_status
                    if duration and col_duration:
                        try:
                            ws.Cells(row, col_duration).Value = float(duration)
                        except:
                            pass
                    updated += 1
                    print(f"   [OK] Row {row}: {item.get('case')} → {excel_status}")
                except Exception as e:
                    print(f"   [ERROR] Error updating row {row}: {e}")

            wb.Save()
            print(f"[SUCCESS] Status update saved: {updated}/{len(items)} rows")
            if updated > 0:
                run_export_dashboard_data()
            return {"ok": True, "updated": updated, "total": len(items)}

        except Exception as e:
            print(f"[ERROR] Status update error: {e}")
            err_str = str(e).lower()
            if any(x in err_str for x in ["displayalerts", "visible", "rejected", "busy", "0x8001010a", "0x800ac472"]):
                return {"ok": False, "error": "אקסל במצב עריכה (תא פעיל) או מציג תיבת דו-שיח. נא ללחוץ על ESC באקסל כדי לצאת ממצב עריכה ולסגור הודעות פתוחות, ואז לנסות שוב."}
            return {"ok": False, "error": str(e)}
        finally:
            if opened_wb and wb:
                try:
                    wb.Close(SaveChanges=True)
                except:
                    pass
            if created_excel and excel:
                try:
                    excel.Quit()
                except:
                    pass

    def _send_json(self, code, data):
        import json
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _execute_add(self, items, result_holder):
        """מוסיף תיקים לאקסל ושומר תוצאות ב-result_holder."""
        try:
            import find_unlisted_and_notify
            find_unlisted_and_notify.execute_add_to_excel(items)
            result_holder["ok"] = True
            result_holder["added"] = len(items)
            # Log
            os.makedirs(os.path.dirname(DONE_FILE), exist_ok=True)
            with open(DONE_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "approved_at": str(datetime.datetime.now()),
                    "items": items
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            result_holder["ok"] = False
            err_str = str(e).lower()
            if any(x in err_str for x in ["displayalerts", "visible", "rejected", "busy", "0x8001010a", "0x800ac472"]):
                msg = "אקסל במצב עריכה (תא פעיל) או מציג תיבת דו-שיח. נא ללחוץ על ESC באקסל כדי לצאת ממצב עריכה ולסגור הודעות פתוחות, ואז לנסות שוב."
            else:
                msg = str(e)
            result_holder["errors"].append(msg)

    def _shutdown(self):
        global _server_instance
        if _server_instance:
            t = threading.Thread(target=_server_instance.shutdown)
            t.daemon = True
            t.start()

    def _send_html(self, code, html):
        body = html.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _page_success(self, items, added):
        rows = ""
        for item in items:
            rows += f"<tr><td>{item.get('case','')}</td><td>{item.get('date','')}</td><td>{item.get('court','')}</td></tr>"
        return f"""
        <html dir="rtl"><head><meta charset="utf-8"><title>אושר ✅</title>
        <style>
          body{{font-family:'Segoe UI',sans-serif;background:#f0fff4;display:flex;
                justify-content:center;align-items:center;min-height:100vh;margin:0}}
          .box{{background:white;padding:40px;border-radius:12px;
                box-shadow:0 4px 20px rgba(0,0,0,.1);max-width:600px;width:90%;text-align:center}}
          h1{{color:#27ae60;font-size:2em}} table{{width:100%;border-collapse:collapse;margin-top:20px;text-align:right}}
          th{{background:#2c3e50;color:white;padding:10px}} td{{padding:8px;border-bottom:1px solid #eee}}
          .badge{{background:#27ae60;color:white;padding:5px 15px;border-radius:20px;font-size:14px}}
        </style></head>
        <body><div class="box">
          <h1>✅ אושר והוסף לאקסל!</h1>
          <p>נוספו <span class="badge">{added} תיקים</span> לאקסל בהצלחה.</p>
          <table><thead><tr><th>מספר תיק</th><th>תאריך</th><th>בית משפט</th></tr></thead>
          <tbody>{rows}</tbody></table>
          <p style="color:#999;font-size:12px;margin-top:20px">חלון זה ייסגר אוטומטית.</p>
        </div></body></html>
        """

    def _page_rejected(self):
        return """
        <html dir="rtl"><head><meta charset="utf-8"><title>נדחה</title>
        <style>body{font-family:'Segoe UI',sans-serif;background:#fff5f5;
               display:flex;justify-content:center;align-items:center;min-height:100vh}
               .box{background:white;padding:40px;border-radius:12px;
                    box-shadow:0 4px 20px rgba(0,0,0,.1);text-align:center}
               h1{color:#c0392b}</style></head>
        <body><div class="box">
          <h1>🚫 הבקשה נדחתה</h1>
          <p>התיקים לא יוכנסו לאקסל.</p>
        </div></body></html>"""

    def _page_error(self, msg):
        return f"""
        <html dir="rtl"><head><meta charset="utf-8"><title>שגיאה</title>
        <style>body{{font-family:'Segoe UI',sans-serif;background:#fff5f5;
               display:flex;justify-content:center;align-items:center;min-height:100vh}}
               .box{{background:white;padding:40px;border-radius:12px;text-align:center}}
               h1{{color:#c0392b}}</style></head>
        <body><div class="box"><h1>{msg}</h1></div></body></html>"""


def start_server(timeout_minutes=4320):  # 72 שעות (3 ימים) כדי שהשרת לא ייסגר לפני שהמשתמש מגיע בבוקר
    """
    מפעיל את שרת האישור.
    timeout_minutes: כמה דקות לחכות לאישור לפני כיבוי אוטומטי.
    """
    global _server_instance
    
    # השתמש בטוקן הקיים אם יש, כדי שקישורים ממיילים ישנים ימשיכו לעבוד!
    token = read_token()
    if not token:
        token = generate_token()
        
    try:
        print(f"[WEB] שרת אישור פועל: http://localhost:{PORT}")
        print(f"[KEY] אסימון: {token}")
    except Exception:
        pass

    _server_instance = HTTPServer(("", PORT), ApprovalHandler)
    _server_instance.timeout = 1

    # Run with timeout
    deadline = datetime.datetime.now() + datetime.timedelta(minutes=timeout_minutes)
    try:
        while datetime.datetime.now() < deadline:
            _server_instance.handle_request()
    except Exception:
        pass
    finally:
        _server_instance.server_close()
        print("[STOP] שרת אישור כובה.")


def build_approval_url():
    """בונה URL לאישור (לשימוש במייל)."""
    token = read_token() or generate_token()
    return f"http://localhost:{PORT}/approve?token={token}"


def build_reject_url():
    """בונה URL לדחייה."""
    token = read_token() or generate_token()
    return f"http://localhost:{PORT}/reject?token={token}"


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "stop":
        print("כיבוי ידני לא זמין מכאן. הסקריפט כבה אוטומטית לאחר אישור.")
    else:
        start_server()
