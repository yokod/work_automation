# -*- coding: utf-8 -*-
import json
import glob
import os
import sys

def main():
    # Make sure stdout supports UTF-8 on Windows
    if sys.platform.startswith('win'):
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

    # Get the parent directory of the script (which is the workspace directory)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    workspace_dir = os.path.dirname(script_dir)
    
    print(f"Workspace directory: {workspace_dir}")
    
    all_jobs = []
    seen = set()
    
    # Pattern to search for scraped JSON files in the workspace dir
    pattern = os.path.join(workspace_dir, "scraped_*.json")
    json_files = glob.glob(pattern)
    
    # Also include rag_project live_jobs.json if exists
    live_jobs_path = os.path.join(workspace_dir, "rag_project", "data", "live_jobs.json")
    if os.path.exists(live_jobs_path):
        json_files.append(live_jobs_path)
    
    # Strict Blacklist filter: Any title containing 'product' or 'מוצר' is purged
    STRICT_BLACKLIST = [
        "product", "מוצר",
        "qa", "tester", "בדיקות", "test engineer", "quality assurance",
        "project manager", "מנהל פרויקטים", "tpm", "program manager",
        "full stack", "fullstack", "backend", "devops", "frontend", "software engineer", "senior developer", "software developer",
        "ppc", "soc", "security analyst", "sales", "מכירות", "financial analyst", "helpdesk", "marketing", "שיווק",
        "embedded", "hardware", "rtos", "vlsi", "c++", "fpga", "board design", "linux"
    ]

    def is_eligible_title(t):
        t_lower = t.lower().strip()
        for b_word in STRICT_BLACKLIST:
            if b_word in t_lower:
                return False
        return True

    def calculate_ai_score(title, desc, reqs):
        full_text = f"{title} {desc} {reqs}".lower()
        score = 7.0
        
        # Penalties for mismatched platforms/technologies
        penalties = ["salesforce", "apex", "salesforce flows", "sap", "workday", "hubspot", "oracle erp", "servicenow"]
        for p in penalties:
            if p in full_text:
                score -= 2.0
                
        high_impact = ["n8n", "make", "zapier", "ai builder", "solutions engineer", "solution engineer", "rag", "agents", " סוכנים "]
        mid_impact = ["python", "automation", "אוטומציה", "integration", "אינטגרציה", "bizops", "technical operations", "gpt", "llm", "low-code", "no-code"]
        
        for kw in high_impact:
            if kw in full_text:
                score += 0.6
        for kw in mid_impact:
            if kw in full_text:
                score += 0.3
                
        # Clamp score between 4.5 and 9.6
        score = max(4.5, min(9.6, round(score, 1)))
        return str(score)


    print(f"Found {len(json_files)} scraped JSON files to sync:")
    for file_path in json_files:
        print(f" - {os.path.basename(file_path)}")
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    for job in data:
                        title = job.get("title", "").strip()
                        company = job.get("company", "").strip()
                        if not title or not company:
                            continue
                            
                        # Filter out blacklisted titles (any title with 'product' or 'מוצר')
                        if not is_eligible_title(title):
                            continue

                        # Avoid duplicates
                        key = (title.lower(), company.lower())
                        if key not in seen:
                            seen.add(key)

                            # Ensure defaults & AI Score
                            job_date = job.get("date") or job.get("date_posted") or ""
                            job_desc = job.get("notes") or job.get("description") or ""
                            job_reqs = job.get("requirements") or ""
                            
                            if job_reqs:
                                full_notes = f"{job_desc}\n\nדרישות:\n{job_reqs}".strip()
                            else:
                                full_notes = job_desc

                            raw_score = job.get("score") or job.get("ai_score")
                            if not raw_score or str(raw_score).strip() in ["", "-", "None"]:
                                final_score = calculate_ai_score(title, job_desc, job_reqs)
                            else:
                                final_score = str(raw_score).strip()

                            if not final_score or final_score in ["", "-", "None"]:
                                final_score = "7.5"


                            job_data = {
                                "id": job.get("id") or job.get("job_id"),
                                "title": title,
                                "company": company,
                                "url": job.get("url", ""),
                                "source": job.get("source", job.get("domain", "Local Scraper")),
                                "score": final_score,
                                "status": job.get("status", "saved"),
                                "priority": job.get("priority", "high"),
                                "salary": job.get("salary", ""),
                                "notes": full_notes,
                                "date": job_date
                            }


                            # Optional: auto-categorize in python too
                            job_data["category"] = job.get("category", "")
                            all_jobs.append(job_data)

                else:
                    print(f"   [Warning] {os.path.basename(file_path)} does not contain a list of jobs.")
        except Exception as e:
            print(f"   [Error] Failed to read {os.path.basename(file_path)}: {e}")
            
    # Write to scraped_jobs.js
    output_js_path = os.path.join(workspace_dir, "scraped_jobs.js")
    
    # Serialize to JSON then wrap in global javascript variable
    js_content = f"window.scrapedJobs = {json.dumps(all_jobs, indent=2, ensure_ascii=False)};\n"
    
    try:
        with open(output_js_path, "w", encoding="utf-8") as f:
            f.write(js_content)
        print(f"\nSuccess! Merged {len(all_jobs)} unique jobs into {os.path.basename(output_js_path)}")
    except Exception as e:
        print(f"\n[Error] Failed to write {output_js_path}: {e}")

if __name__ == "__main__":
    main()
