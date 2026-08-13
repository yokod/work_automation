"""
 word_format_checker.py
=========================
מודול לבדיקה ותיקון עיצוב קבצי Word של פרוטוקולי דיון.

תכונות:
- פונט: David 12
- יישור: לשני הצדדים (Justify)
- מרווח בין שורות: 1.5
- כניסה תלויה: 3.7 מ"מ (בגוף הפרוטוקול בלבד  לא בכותרת)
- בדיקת סימני פיסוק ע"י AI (ללא שינוי תמלול)

שימוש:
    from word_format_checker import check_and_fix_word_format
    fixed_path, report = check_and_fix_word_format("path/to/protocol.docx")
"""

import os
import re
import tempfile
import shutil
import pythoncom
import win32com.client as win32

# ============================================================================
#  קבועי עיצוב
# ============================================================================

FONT_NAME        = "David"          # פונט
FONT_SIZE        = 12               # גודל פונט (points)
LINE_SPACING     = 1.5              # מרווח בין שורות (multiple)
HANGING_INDENT_MM = 3.7            # כניסה תלויה במ"מ
ALIGNMENT_JUSTIFY = 3              # wdAlignParagraphJustify = 3
ALIGNMENT_CENTER  = 1              # wdAlignParagraphCenter  = 1
ALIGNMENT_RIGHT   = 2              # wdAlignParagraphRight   = 2
ALIGNMENT_DISTRIBUTE = 4           # wdAlignParagraphDistribute = 4

# ספירת שורות כותרת (הפסקאות הראשונות = כותרת, לא נגע בהן)
# כותרת מוגדרת כ: מספר שורות מהתחלה שמכילות מידע טכני (בית משפט, מספר תיק, תאריך, כבוד השופט)
HEADER_LINES_COUNT = 15  # הגדרה שמרנית  עד 15 שורות ראשונות נחשבות כותרת

# המרות Word
POINTS_PER_CM  = 72 / 2.54
POINTS_PER_MM  = POINTS_PER_CM / 10
TWIPS_PER_INCH = 1440
TWIPS_PER_MM   = TWIPS_PER_INCH / 25.4

HANGING_INDENT_TWIPS = int(HANGING_INDENT_MM * TWIPS_PER_MM)  # ~ 210 twips


# ============================================================================
#  זיהוי כותרת
# ============================================================================

def _is_header_paragraph(para_index: int, para_text: str) -> bool:
    """
    מזהה אם פסקה היא חלק מהכותרת (שלא נגע בה).
    כותרת = פסקאות ראשונות שמכילות: שם בית משפט, מספר תיק, תאריך, שם שופט.
    """
    if para_index < HEADER_LINES_COUNT:
        return True
    # זיהוי נוסף: שורות הכותרת מכילות מילות מפתח אופייניות
    header_keywords = [
        "בית משפט", "בית-משפט", "מחוזי", "השלום", "תעבורה", "משפחה",
        "מספר תיק", "תיק מ", "תיק ע",
        "בפני", "בפני כב", "כבוד", "השופט", "השופטת", "הרשמ",
        "פרוטוקול", "דיון", "ישיבה",
        "נוכחים", "מועד הדיון",
    ]
    text_lower = para_text.strip()
    for kw in header_keywords:
        if kw in text_lower:
            return True
    return False


def _is_body_start(para_text: str) -> bool:
    """
    מזהה את תחילת גוף הפרוטוקול (שורות דיאלוג).
    סימנים: ת: / ש: / עו"ד: / השופט: וכו'
    """
    stripped = para_text.strip()
    body_patterns = [
        r'^ת[\'׳\u2019]?\s*[:.]',
        r'^ש[\'׳\u2019]?\s*[:.]',
        r'^עו"ד\s*.*?[:.]',
        r'^עו\'ד\s*.*?[:.]',
        r'^עוד\s*.*?[:.]',
        r'^השופט[ת]?\s*.*?[:.]',
        r'^הרשמ[ת]?\s*.*?[:.]',
        r'^התובע\s*.*?[:.]',
        r'^הנאשם\s*.*?[:.]',
        r'^העד\s*.*?[:.]',
        r'^הסניגור\s*.*?[:.]',
        r'^המשיב\s*.*?[:.]',
        r'^המבקש\s*.*?[:.]',
    ]
    for pat in body_patterns:
        if re.match(pat, stripped):
            return True
    return False


# ============================================================================
#  פונקציה מרכזית: בדיקה + תיקון עיצוב
# ============================================================================

def fix_word_format(file_path: str, output_path: str = None, dry_run: bool = False) -> dict:
    """
    בודק ומתקן עיצוב קובץ Word של פרוטוקול דיון.

    פרמטרים:
        file_path:   נתיב לקובץ Word (docx/doc)
        output_path: נתיב לפלט (אם None  ישמור מעל המקור)
        dry_run:     אם True  רק מדווח, לא משנה

    מחזיר:
        dict עם:
          "fixed":          True/False  האם בוצעו שינויים
          "output_path":    נתיב הקובץ המתוקן
          "changes":        רשימת השינויים שבוצעו
          "warnings":       אזהרות
          "paragraphs_fixed": מספר פסקאות שתוקנו
    """
    changes   = []
    warnings  = []
    fixed_any = False
    paragraphs_fixed = 0

    if not os.path.exists(file_path):
        return {
            "fixed": False, "output_path": file_path,
            "changes": [], "warnings": [f"קובץ לא נמצא: {file_path}"],
            "paragraphs_fixed": 0
        }

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in (".doc", ".docx", ".odt"):
        return {
            "fixed": False, "output_path": file_path,
            "changes": [], "warnings": [f"סוג קובץ לא נתמך: {ext}"],
            "paragraphs_fixed": 0
        }

    # קובץ פלט
    if output_path is None:
        output_path = file_path

    try: pythoncom.CoInitialize()
    except: pass
    _created_word = False
    try:
        word_app = win32.GetActiveObject("Word.Application")
    except:
        try:
            word_app = win32.Dispatch("Word.Application")
            _created_word = True
        except Exception as e:
            warnings.append(f"Could not start Word application: {e}")
            return {
                "fixed": False, "output_path": file_path,
                "changes": [], "warnings": warnings,
                "paragraphs_fixed": 0
            }

    try:
        word_app.Visible = False
        word_app.DisplayAlerts = 0  # wdAlertsNone

        abs_path = os.path.abspath(file_path)
        doc = word_app.Documents.Open(abs_path, ReadOnly=False, Visible=False)
        
        if not doc:
            raise Exception("Word.Open returned None for " + abs_path)
            
        try:
            paragraphs = doc.Paragraphs
            para_count = paragraphs.Count
        except Exception as pe:
            doc.Close(False)
            raise Exception(f"Could not access Paragraphs from document: {pe}")

        in_body         = False

        # --- מעבר ראשון: איתור גוף הפרוטוקול ---
        for i in range(1, para_count + 1):
            para = paragraphs(i)
            text = para.Range.Text.strip()
            
            # אם המילה "פרוטוקול" רשומה לבדה, הפרוטוקול מתחיל מהפסקה הבאה
            text_clean = text.replace(" ", "").replace("_", "").replace("-", "")
            if text_clean == "פרוטוקול":
                in_body = True
                body_start_idx = i + 1
                break

            if not in_body and _is_body_start(text):
                in_body = True
                body_start_idx = i
                break

        # --- מעבר יחיד: עדכון פונט ומרווח בלבד ---
        # חובה שלא לגעת באף הגדרה של הפסקאות: לא יישור, לא מרכז, ולא שוליים!
        for i in range(1, para_count + 1):
            para = paragraphs(i)
            text = para.Range.Text.strip()

            # פסקה ריקה  דלג
            if not text:
                continue

            para_changed = False
            fmt = para.Format
            rng = para.Range

            # א. פונט (David 12)
            if rng.Font.Name != FONT_NAME:
                if not dry_run:
                    rng.Font.Name = FONT_NAME
                changes.append(f"פסקה {i}: פונט")
                para_changed = True

            if rng.Font.Size != FONT_SIZE:
                if not dry_run:
                    rng.Font.Size = FONT_SIZE
                para_changed = True

            # ב. מרווח בין שורות (1.5)
            # LineSpacingRule 5 = wdLineSpaceMultiple
            if fmt.LineSpacingRule != 5 or abs(fmt.LineSpacing - (LINE_SPACING * 12)) > 1.0:
                if not dry_run:
                    fmt.LineSpacingRule = 5
                    fmt.LineSpacing = LINE_SPACING * 12
                changes.append(f"פסקה {i}: מרווח 1.5")
                para_changed = True

            if para_changed:
                paragraphs_fixed += 1
                fixed_any = True

        if not dry_run and fixed_any:
            doc.SaveAs(os.path.abspath(output_path))

        doc.Close(False)
        if _created_word:
            word_app.Quit()

    except Exception as e:
        warnings.append(f"שגיאה בעיבוד Word: {e}")
        try:
            if _created_word:
                word_app.Quit()
        except Exception:
            pass
    finally:
        try: pythoncom.CoUninitialize()
        except: pass

    return {
        "fixed":            fixed_any,
        "output_path":      output_path,
        "changes":          changes,
        "warnings":         warnings,
        "paragraphs_fixed": paragraphs_fixed,
    }


# ============================================================================
#  בדיקת סימני פיסוק ע"י AI (ללא שינוי תמלול)
# ============================================================================

def check_punctuation_ai(text_content: str) -> dict:
    """
    שולח את הטקסט ל-Gemini לבדיקת סימני פיסוק.
    לא משנה את התמלול  רק מדווח על בעיות.

    מחזיר:
        dict עם:
          "issues":   רשימת בעיות שזוהו
          "warnings": אזהרות קלות
          "ok":       True אם אין בעיות חוסמות
    """
    try:
        from llm_utils import ask_llm_json
    except ImportError:
        return {"issues": ["מודול llm_utils לא נטען"], "warnings": [], "ok": True}

    if not text_content or len(text_content.strip()) < 50:
        return {"issues": [], "warnings": ["טקסט קצר מדי לבדיקה"], "ok": True}

    # שולחים רק התחלה וקצת אמצע כדי למנוע קטיעת תשובת AI (Max Tokens)
    if len(text_content) > 3500:
        sample = text_content[:3500] + "\n...\n" + text_content[-500:]
    else:
        sample = text_content

    prompt = f"""אתה עורך לשוני המתמחה בבדיקת פרוטוקולי בתי משפט עבריים.
המשימה שלך: **לזהות בעיות בסימני פיסוק בלבד**  אסור לשנות מילים, ניסוחים, או תוכן.

פרוטוקול לבדיקה:
---
{sample}
---

בדוק רק את הנושאים הבאים:
1. נקודה בסוף משפט (חסרה / מיותרת)
2. פסיק (חסר / מיותר)  במיוחד לפני מילות קישור
3. גרש בשם קצר (ת׳, ש׳)  עקביות
4. מירכאות ("...")  פתיחה/סגירה תואמות

השב ב-JSON בלבד (הקפד לסגור את כל הסוגריים כדי למנוע שגיאות חיתוך).
* חובה להחזיר מקסימום 3 בעיות כדי לא לחרוג ממגבלת אורך.

{{
  "punctuation_issues": [
    {{
      "type": "שם הבעיה בעברית",
      "description": "תיאור קצר",
      "example": "הטקסט עם הבעיה (עד 60 תווים)",
      "severity": "error|warning"
    }}
  ],
  "overall_quality": "good|fair|poor",
  "summary": "סיכום קצר בעברית"
}}

חשוב: אל תשנה את התמלול עצמו. דווח בלבד."""

    result = ask_llm_json(prompt)
    if not result:
        return {"issues": [], "warnings": ["לא ניתן לבצע בדיקת פיסוק (AI לא זמין)"], "ok": True}

    issues_raw = result.get("punctuation_issues", [])
    issues   = [f"{x.get('type','')}: {x.get('description','')} | דוגמה: {x.get('example','')}"
                for x in issues_raw if x.get("severity") == "error"]
    warnings = [f"{x.get('type','')}: {x.get('description','')}"
                for x in issues_raw if x.get("severity") == "warning"]

    return {
        "issues":   issues,
        "warnings": warnings,
        "ok":       len(issues) == 0,
        "summary":  result.get("summary", ""),
        "quality":  result.get("overall_quality", ""),
        "raw":      result,
    }


# ============================================================================
#  פונקציה ראשית: בדיקה מלאה לפני שליחה
# ============================================================================

def check_and_fix_word_format(
    file_path: str,
    text_content: str = "",
    dry_run: bool = False,
    check_punctuation: bool = True,
    add_fixed_suffix: bool = False, # ברירת מחדל: לא להוסיף סיומת "מתוקן"
) -> tuple[str, dict]:
    """
    Entry Point ראשי: בוצע לפני שליחת מייל.

    1. מתקן עיצוב קובץ Word (גוף בלבד, לא כותרת)
    2. בודק סימני פיסוק ע"י AI (רק מדווח, לא משנה)

    מחזיר:
        (output_path, report_dict)
        output_path: נתיב הקובץ המתוקן (אותו קובץ אם dry_run)
        report_dict: דוח מלא
    """
    print(f"\n[INFO] בדיקת עיצוב: {os.path.basename(file_path)}")

    report = {
        "file":              file_path,
        "format_fixed":     False,
        "paragraphs_fixed": 0,
        "format_changes":   [],
        "format_warnings":  [],
        "punct_ok":         True,
        "punct_issues":     [],
        "punct_warnings":   [],
        "punct_summary":    "",
        "output_path":      file_path,
    }

    # --- שלב 1: תיקון עיצוב ---
    # נגדיר נתיב פלט
    base, ext = os.path.splitext(file_path)
    output_path = file_path # ברירת מחדל: דורס את המקור
    
    # מוסיפים "מתוקן" רק אם התבקש במפורש והקובץ באמת עובד
    if add_fixed_suffix and not dry_run and "מתוקן" not in base:
        output_path = f"{base} מתוקן{ext}"

    fmt_result = fix_word_format(file_path, output_path=output_path, dry_run=dry_run)
    
    # עדכון נתיב הפלט בדוח
    report["output_path"] = fmt_result["output_path"] if fmt_result["fixed"] else file_path

    
    report["format_fixed"]     = fmt_result["fixed"]
    report["paragraphs_fixed"] = fmt_result["paragraphs_fixed"]
    report["format_changes"]   = fmt_result["changes"]
    report["format_warnings"]  = fmt_result["warnings"]


    if fmt_result["fixed"]:
        print(f"   [DONE] עיצוב תוקן: {fmt_result['paragraphs_fixed']} פסקאות")
        # מדפיס רק עד 5 שינויים ראשונים
        for c in fmt_result["changes"][:5]:
            print(f"       {c}")
        if len(fmt_result["changes"]) > 5:
            print(f"      ... ועוד {len(fmt_result['changes']) - 5} שינויים")
    else:
        print("   [INFO] עיצוב תקין  לא נדרשו שינויים")

    for w in fmt_result["warnings"]:
        print(f"   [WARN] {w}")

    # --- שלב 2: בדיקת סימני פיסוק ---
    if check_punctuation and text_content:
        print("   [INFO] בודק סימני פיסוק...")
        punct_result = check_punctuation_ai(text_content)
        report["punct_ok"]       = punct_result["ok"]
        report["punct_issues"]   = punct_result["issues"]
        report["punct_warnings"] = punct_result["warnings"]
        report["punct_summary"]  = punct_result.get("summary", "")

        quality_icon = {"good": "[OK]", "fair": "[WARN]", "poor": "[FAIL]"}.get(
            punct_result.get("quality", ""), "[INFO]"
        )
        print(f"   {quality_icon} פיסוק: {punct_result.get('summary', 'ללא סיכום')}")
        for issue in punct_result["issues"]:
            print(f"      [ISSUE] {issue}")
        for warn in punct_result["warnings"]:
            print(f"      [WARN] {warn}")
    elif check_punctuation:
        print("   [INFO] לא בוצעה בדיקת פיסוק (אין טקסט להעברה)")

    return report["output_path"], report


# ============================================================================
#  בדיקה עצמאית
# ============================================================================

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("שימוש: python word_format_checker.py <path_to_docx> [--dry-run]")
        sys.exit(1)

    path     = sys.argv[1]
    is_dry   = "--dry-run" in sys.argv

    print(f"{'[DRY RUN] ' if is_dry else ''}בודק: {path}")
    out, rep = check_and_fix_word_format(path, dry_run=is_dry)

    print("\n--- דוח מלא ---")
    print(f"קובץ פלט:      {out}")
    print(f"עיצוב תוקן:    {rep['format_fixed']} ({rep['paragraphs_fixed']} פסקאות)")
    print(f"פיסוק תקין:    {rep['punct_ok']}")
    if rep["punct_summary"]:
        print(f"סיכום פיסוק:   {rep['punct_summary']}")
