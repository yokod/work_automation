import os
import sys
import win32com.client as win32

# הוספת הנתיב של הפרויקט ל-sys.path
sys.path.append(r"d:\yoel\projects\auto")

from send_district_emails import detect_district, DISTRICT_CONTACTS

DONE_FOLDER = r"C:\Users\yoel\OneDrive - Hever\Jerusalem\Done"

def run_real_dry_run():
    print("\n" + "="*70)
    print("🧪 [DRY RUN] בדיקת זיהוי AI על קבצים אמיתיים מה-Done")
    print("="*70)
    
    # שליפת 10 קבצים אחרונים
    files = [f for f in os.listdir(DONE_FOLDER) if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
    files = sorted(files, key=lambda f: os.path.getmtime(os.path.join(DONE_FOLDER, f)), reverse=True)[:10]

    if not files:
        print("❌ לא נמצאו קבצים בתיקיית ה-Done.")
        return

    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        
        results = []
        for fname in files:
            fpath = os.path.join(DONE_FOLDER, fname)
            print(f"\n📄 מעבד: {fname}...")
            
            try:
                doc = word.Documents.Open(fpath, ReadOnly=True, Visible=False)
                # קריאת 4000 תווים ראשונים (הכותרת)
                text_content = doc.Content.Text[:4000]
                doc.Close(False)
                
                # הפעלת הלוגיקה החדשה (AI + Regex fallback)
                district_key = detect_district(text_content, fname)
                recipient = DISTRICT_CONTACTS.get(district_key, "❌ לא זוהה מחוז")
                
                results.append({
                    "file": fname,
                    "district": district_key,
                    "recipient": recipient
                })
                
            except Exception as e:
                print(f"   ⚠️ שגיאה בקריאת הקובץ: {e}")
                continue
        
        word.Quit()
        
        print("\n" + "="*70)
        print("📊 סיכום הבדיקה (מה היה קורה במציאות):")
        print("="*70)
        for r in results:
            status = "✅" if r['district'] else "❓"
            print(f"{status} {r['file'][:50]}...")
            print(f"   -> מחוז שזוהה: {r['district']}")
            print(f"   -> יעד שליחה: {r['recipient']}")
            print("-" * 40)

    except Exception as e:
        print(f"❌ שגיאה כללית: {e}")

if __name__ == "__main__":
    run_real_dry_run()
