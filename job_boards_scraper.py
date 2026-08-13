import json
import datetime
import os
import time
from playwright.sync_api import sync_playwright

# Path to store the browser profile (so you don't have to login every time)
USER_DATA_DIR = os.path.join(os.getcwd(), "playwright_profile")

def save_jobs(jobs):
    if not jobs:
        print("No new jobs to save.")
        return
    output_file = "scraped_web_jobs.json"
    
    existing_jobs = []
    if os.path.exists(output_file):
        with open(output_file, 'r', encoding='utf-8') as f:
            try:
                existing_jobs = json.load(f)
            except:
                pass
                
    existing_jobs.extend(jobs)
    
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(existing_jobs, f, indent=2, ensure_ascii=False)
    print(f"\nSuccess! Saved {len(jobs)} jobs to {output_file}.")

def scrape_linkedin(page):
    print("\n--- LinkedIn Scraping ---")
    page.goto('https://www.linkedin.com/my-items/saved-jobs/?cardType=APPLIED')
    
    # Wait for the user to log in if they aren't
    if "login" in page.url or "checkpoint" in page.url:
        print("Please log in to LinkedIn in the browser window.")
        print("Waiting for you to log in...")
        page.wait_for_url('**/my-items/saved-jobs/?cardType=APPLIED', timeout=0) # wait indefinitely
    
    jobs = []
    page_num = 1
    
    while True:
        print(f"Scraping page {page_num}...")
        time.sleep(5) # Give page some time to fully render React components
        
        # Try multiple selectors that LinkedIn uses for job list cards
        selectors = [
            '.reusable-search__result-container',
            '.entity-result',
            '.entity-result__item',
            '.job-card-container',
            '[data-chameleon-result-urn]'
        ]
        
        job_cards = []
        for sel in selectors:
            cards = page.query_selector_all(sel)
            if cards:
                job_cards = cards
                break
                
        if not job_cards:
            print(f"Could not find job elements on page {page_num}.")
            break
            
        print(f"Found {len(job_cards)} job cards on page {page_num}.")
        
        for card in job_cards:
            try:
                raw_text = card.inner_text().strip()
                lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
                
                title = "Unknown Title"
                company = "Unknown Company"
                
                title_el = card.query_selector('.entity-result__title-text, .job-card-list__title, [class*="title-text"]')
                company_el = card.query_selector('.entity-result__primary-subtitle, .job-card-container__company-name, [class*="primary-subtitle"]')
                
                if title_el:
                    title = title_el.inner_text().split('\n')[0].strip()
                elif len(lines) > 0:
                    title = lines[0]
                    
                if company_el:
                    company = company_el.inner_text().split('\n')[0].strip()
                elif len(lines) > 1:
                    company = lines[1]
                    
                if "View" in title or "Click" in title or len(title) > 100:
                    if len(lines) > 0:
                        title = lines[0]
                
                # Check for duplicates on this run
                if not any(j['title'] == title and j['company'] == company for j in jobs):
                    jobs.append({
                        "id": f"li_{datetime.datetime.now().timestamp()}_{len(jobs)}",
                        "title": title,
                        "company": company,
                        "url": "",
                        "source": "LinkedIn",
                        "score": "",
                        "status": "applied",
                        "priority": "mid",
                        "salary": "",
                        "notes": "נשאב אוטומטית מהאתר - לינקדין",
                        "date": datetime.date.today().isoformat()
                    })
            except Exception as e:
                pass
                
        # Look for the Next pagination button
        next_btn = page.query_selector('button[aria-label="Next"], button:has-text("Next"), button:has-text("הבא")')
        if next_btn:
            try:
                # Scroll to button and click
                next_btn.scroll_into_view_if_needed()
                if next_btn.is_enabled():
                    print("Clicking 'Next' button...")
                    next_btn.click()
                    page_num += 1
                    time.sleep(3)
                else:
                    print("Next button is disabled. Reached the last page.")
                    break
            except Exception as e:
                print(f"Could not click next button: {e}")
                break
        else:
            print("No next button found. Reached the last page.")
            break
            
    print(f"Scraped a total of {len(jobs)} jobs from LinkedIn across all pages.")
    return jobs

def scrape_alljobs(page):
    print("\n--- AllJobs Scraping ---")
    page.goto('https://www.alljobs.co.il/')
    
    print("Please log in to AllJobs in the browser window.")
    print("Then navigate to your 'משרות ששלחתי' (Sent Jobs) page.")
    print("IMPORTANT: Click INSIDE the black Terminal window in VS Code before pressing Enter.")
    input("Once you are on the sent jobs page, click inside this Terminal and press Enter...")
    
    print("Waiting for page to stabilize...")
    time.sleep(5)
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except:
        pass
        
    jobs = []
    job_cards = []
    
    # Try up to 3 times in case of active navigation
    for attempt in range(3):
        try:
            selectors = [
                '.N_UserList_Row', 
                '.N_UserList_Row_Alternate', 
                'tr[id*="dgUserList"]', 
                '#dgUserList tr',
                '.job-list-item', 
                '.job-box',
                '.job-block',
                'table tr' # Generic fallback for any table row
            ]
            
            for sel in selectors:
                cards = page.query_selector_all(sel)
                if cards:
                    job_cards = cards
                    print(f"Found {len(cards)} elements matching AllJobs selector: {sel}")
                    break
            break # Success, exit retry loop
        except Exception as e:
            print(f"Attempt {attempt+1} failed due to navigation. Retrying in 2 seconds...")
            time.sleep(2)
                
    if not job_cards:
        try:
            print("\n[DEBUG] Selector failed. Dumping page table structures:")
            tables = page.query_selector_all('table')
            print(f"Found {len(tables)} tables on the page.")
            for t_idx, t in enumerate(tables):
                t_id = t.get_attribute('id') or 'No ID'
                t_class = t.get_attribute('class') or 'No Class'
                print(f"Table #{t_idx+1}: ID='{t_id}', Class='{t_class}'")
                
            rows = page.query_selector_all('tr')
            print(f"Found {len(rows)} total tr elements on the page.")
            if rows:
                print("First 3 tr text previews:")
                for r_idx, r in enumerate(rows[:5]):
                    print(f"  tr #{r_idx+1}: Class='{r.get_attribute('class') or 'No Class'}', Text='{r.text_content().strip()[:100]}'")
        except Exception as debug_error:
            print(f"Error gathering debug info: {debug_error}")
            
        return []
            
    for idx, card in enumerate(job_cards):
        try:
            tag_name = card.evaluate("el => el.tagName")
            class_name = card.evaluate("el => el.className")
            
            title = "Unknown Title"
            company = "Unknown Company"
            date_str = datetime.date.today().isoformat()
            
            # Check if it's a table row with cells
            cells = card.query_selector_all('td')
            if len(cells) >= 3:
                title = cells[0].text_content().strip()
                company = cells[1].text_content().strip()
                date_text = cells[2].text_content().strip()
                
                # Try to parse date (e.g. 25/06/2026 -> 2026-06-25)
                m = re.search(r'(\d{2})/(\d{2})/(\d{4})', date_text)
                if m:
                    date_str = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
            else:
                # Fallback to text lines
                raw_text = card.text_content().strip()
                lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
                if len(lines) >= 2:
                    title = lines[0]
                    company = lines[1]
            
            title = " ".join(title.split())
            company = " ".join(company.split())
            
            # Filter out header rows
            if 'תאריך' in title or 'משרה' in title or 'חברה' in title or 'תפקיד' in title:
                continue
                
            print(f"DEBUG - Card #{idx+1}: Title='{title}', Company='{company}', Date='{date_str}'")
            
            if title != "Unknown Title" and title != "" and company != "Unknown Company":
                jobs.append({
                    "id": f"aj_{datetime.datetime.now().timestamp()}_{len(jobs)}",
                    "title": title,
                    "company": company,
                    "url": "",
                    "source": "AllJobs",
                    "score": "",
                    "status": "applied",
                    "priority": "mid",
                    "salary": "",
                    "notes": "נשאב אוטומטית מהאתר - AllJobs",
                    "date": date_str
                })
        except Exception as parse_error:
            print(f"DEBUG - Error parsing AllJobs card: {parse_error}")
            pass
            
    print(f"Scraped {len(jobs)} jobs from AllJobs.")
    return jobs

def main():
    print("====================================")
    print("  Job Boards Scraper Selector")
    print("====================================")
    print("1. LinkedIn")
    print("2. AllJobs")
    print("====================================")
    choice = input("Enter your choice (1 or 2): ").strip()
    
    print("\nStarting Playwright Browser...")
    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            args=['--start-maximized'],
            no_viewport=True
        )
        page = browser.new_page()
        
        all_jobs = []
        
        if choice == '1':
            try:
                li_jobs = scrape_linkedin(page)
                all_jobs.extend(li_jobs)
            except Exception as e:
                print(f"Error scraping LinkedIn: {e}")
        elif choice == '2':
            try:
                aj_jobs = scrape_alljobs(page)
                all_jobs.extend(aj_jobs)
            except Exception as e:
                print(f"Error scraping AllJobs: {e}")
        else:
            print("Invalid choice. Exiting.")
            browser.close()
            return
            
        save_jobs(all_jobs)
        print("\nScraping complete.")
        input("Press Enter to close the browser and exit...")
        browser.close()

if __name__ == '__main__':
    main()
