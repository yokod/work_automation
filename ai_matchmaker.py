import os
import sys
import argparse
import json
import datetime
import re
import time

try:
    import PyPDF2
except ImportError:
    print("Please install PyPDF2: pip install PyPDF2")
    sys.exit(1)

try:
    from google import genai
    from google.genai import types
except ImportError:
    print("Please install google-genai: pip install google-genai")
    sys.exit(1)

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Please install Playwright: pip install playwright && playwright install")
    sys.exit(1)

# --- CONFIGURATION ---
API_KEY = os.environ.get("GEMINI_API_KEY", "")
USER_DATA_DIR = os.path.join(os.getcwd(), "playwright_profile")

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        return text
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None

def scrape_job_url(url):
    print(f"\nOpening browser to scrape job details from: {url}")
    with sync_playwright() as p:
        # Use persistent context to inherit saved logins (LinkedIn session etc)
        browser = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            args=['--start-maximized'],
            no_viewport=True
        )
        page = browser.new_page()
        page.goto(url)
        
        # Check if user needs to solve a captcha or login
        if "login" in page.url or "checkpoint" in page.url or "security" in page.url:
            print("Please log in or solve captcha in the browser window if needed.")
            print("Waiting 15 seconds...")
            time.sleep(15)
            
        print("Extracting job page text...")
        time.sleep(5) # Let content render
        
        title = page.title()
        
        # Attempt to find main job description container text
        selectors = [
            '.jobs-description', 
            '#job-details', 
            '.description__text', 
            '.job-description',
            'main'
        ]
        
        desc_text = ""
        for sel in selectors:
            el = page.query_selector(sel)
            if el:
                desc_text = el.text_content().strip()
                print(f"Found description container using selector: {sel}")
                break
                
        if not desc_text:
            desc_text = page.locator('body').text_content().strip()
            
        browser.close()
        
        # Clean title (e.g. remove " | LinkedIn")
        title = re.sub(r'\s*\|\s*LinkedIn.*', '', title, flags=re.IGNORECASE)
        title = re.sub(r'\s*\|\s*AllJobs.*', '', title, flags=re.IGNORECASE)
        
        return title, desc_text

def analyze_job_fit(resume_text, job_title, job_description):
    if API_KEY == "YOUR_GEMINI_API_KEY" or not API_KEY:
        print("\n[ERROR] Missing API Key! Please edit ai_matchmaker.py and add your GEMINI_API_KEY.")
        sys.exit(1)

    print("\nAnalyzing and Categorizing with Gemini AI...")
    try:
        client = genai.Client(api_key=API_KEY)
        
        prompt = f"""
You are an expert tech recruiter and AI career coach.
I am providing you with my resume, a job title, and a job description. 
Please evaluate the job fit, categorize it, and provide a structured JSON response.

Your response must be in valid JSON format only, matching this structure (do not output markdown tags or ```json wrappers, just raw JSON):
{{
  "score": <a number from 1 to 10 evaluating the fit>,
  "category": "<Must be exactly one of: 'AI & Automation', 'Project Management', 'Operations & Admin', or 'Other'>",
  "company": "<Extracted company name, or 'Unknown Company'>",
  "title": "<Cleaned job title>",
  "summary": "<2-3 sentence Hebrew summary of why this is a fit or not>",
  "missing_skills": ["skill 1", "skill 2"],
  "cv_tips": ["tip 1", "tip 2"]
}}

My Resume:
{resume_text}

Job Details:
Job Title: {job_title}
Job Description:
{job_description}
"""

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        
        # Clean up any potential markdown wrap in Gemini response
        res_text = response.text.strip()
        res_text = re.sub(r'^```json\s*', '', res_text)
        res_text = re.sub(r'\s*```$', '', res_text)
        
        return json.loads(res_text)
    except Exception as e:
        print(f"API or Parsing Error: {e}")
        print(f"Raw Gemini response was:\n{response.text if 'response' in locals() else 'None'}")
        return None

def main():
    parser = argparse.ArgumentParser(description="AI Job Matchmaker & Categorizer")
    parser.add_argument("input_source", help="Path to a job description text file OR a Job URL (e.g. LinkedIn link)")
    parser.add_argument("--resume", default="../resume/yoel_guy_english_resume.pdf", help="Path to your resume PDF")
    
    args = parser.parse_args()

    # Get absolute paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    resume_path = args.resume
    if not os.path.isabs(resume_path):
        resume_path = os.path.join(script_dir, resume_path)

    if not os.path.exists(resume_path):
        print(f"Could not find resume at: {resume_path}")
        return

    print("Extracting text from your resume...")
    resume_text = extract_text_from_pdf(resume_path)
    if not resume_text:
        return
    
    job_title = "Unknown Job"
    job_description = ""
    job_url = ""
    
    # Check if input is a URL
    if args.input_source.startswith("http://") or args.input_source.startswith("https://"):
        job_url = args.input_source
        job_title, job_description = scrape_job_url(job_url)
    else:
        # It's a text file
        job_desc_path = args.input_source
        if not os.path.isabs(job_desc_path):
            job_desc_path = os.path.join(script_dir, job_desc_path)
            
        if not os.path.exists(job_desc_path):
            print(f"Could not find job description file at: {job_desc_path}")
            return
            
        print(f"Reading job description from {os.path.basename(job_desc_path)}...")
        with open(job_desc_path, 'r', encoding='utf-8') as f:
            job_description = f.read()
            
        # Try to infer title from first line of text file
        lines = [l.strip() for l in job_description.split('\n') if l.strip()]
        if lines:
            job_title = lines[0]

    if not job_description:
        print("Error: Job description content is empty.")
        return

    analysis = analyze_job_fit(resume_text, job_title, job_description)
    if not analysis:
        return
        
    # Format notes field for the dashboard
    notes_hebrew = f"""=== ניתוח התאמה AI ===
ציון: {analysis.get('score', '-')}/10
קטגוריה: {analysis.get('category', 'Other')}
חברה: {analysis.get('company', '')}

סיכום:
{analysis.get('summary', '')}

מה חסר לי:
{chr(10).join(['- ' + s for s in analysis.get('missing_skills', [])])}

טיפים לקורות חיים:
{chr(10).join(['- ' + t for t in analysis.get('cv_tips', [])])}
"""

    # Print results to console
    print("\n================ AI MATCHMAKER ANALYSIS ================\n")
    print(notes_hebrew)
    print("========================================================\n")
    
    # Create a job record that can be directly imported
    new_job = {
        "id": f"ai_match_{int(time.time())}",
        "title": analysis.get("title", job_title),
        "company": analysis.get("company", "Unknown Company"),
        "url": job_url,
        "source": "AI Scraped" if job_url else "AI Text",
        "category": analysis.get("category", "Other"),
        "score": str(analysis.get("score", "")),
        "status": "new",
        "priority": "mid",
        "salary": "",
        "notes": notes_hebrew,
        "date": datetime.date.today().isoformat()
    }
    
    # Save directly to workspace scraped folder
    out_file = "scraped_ai_match.json"
    with open(out_file, 'w', encoding='utf-8') as out:
        json.dump([new_job], out, indent=2, ensure_ascii=False)
        
    print(f"Job saved to {out_file}! You can import it using the 'ייבא משרות (JSON)' button on your dashboard.")

if __name__ == "__main__":
    main()
