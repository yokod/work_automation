import imaplib
import email
import json
import re
from email.header import decode_header
import datetime

# --- CONFIGURATION ---
EMAIL_ACCOUNT = "your_email@example.com"
# Generate an App Password in your Google Account:
# Security -> 2-Step Verification -> App passwords
APP_PASSWORD = "YOUR_APP_PASSWORD"
IMAP_SERVER = "imap.gmail.com"

# Search criteria: emails containing these words
SEARCH_KEYWORDS = ['Jobify', 'AllJobs', 'Drushim', 'LinkedIn', 'מועמדות', 'Application']

def connect_imap():
    print(f"Connecting to {IMAP_SERVER}...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, APP_PASSWORD)
    mail.select('inbox')
    return mail

def decode_subject(encoded_subject):
    if not encoded_subject:
        return ""
    decoded_parts = decode_header(encoded_subject)
    subject = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            subject += part.decode(encoding or 'utf-8', errors='ignore')
        else:
            subject += part
    return subject

def extract_job_info(subject, from_address):
    # A simple heuristics-based extraction
    company = ""
    title = ""
    source = ""
    
    subject = subject.replace('\n', ' ').replace('\r', '')
    
    if "LinkedIn" in from_address or "LinkedIn" in subject:
        source = "LinkedIn"
        m = re.search(r'application to (.*?) at (.*?) was', subject, re.IGNORECASE)
        if m:
            title = m.group(1).strip()
            company = m.group(2).strip()
    elif "jobify360" in from_address.lower() or "jobify" in subject.lower():
        source = "Jobify"
        company = "Jobify System" 
        title = subject
    elif "drushim" in from_address.lower() or "דרושים" in subject:
        source = "Drushim"
        company = "Drushim"
        title = subject
    elif "alljobs" in from_address.lower():
        source = "AllJobs"
        company = "AllJobs"
        title = subject
    else:
        source = "Email"
        title = subject
        
    return {"title": title, "company": company, "source": source}

def scrape_emails():
    if EMAIL_ACCOUNT == "YOUR_EMAIL@gmail.com":
        print("Please edit the script to add your EMAIL_ACCOUNT and APP_PASSWORD.")
        return

    try:
        mail = connect_imap()
    except Exception as e:
        print(f"Failed to connect: {e}")
        return

    jobs = []
    
    # We will search the last 30 days
    date_since = (datetime.date.today() - datetime.timedelta(30)).strftime("%d-%b-%Y")
    
    print(f"Searching for applied jobs since {date_since}...")
    
    status, messages = mail.search(None, f'(SINCE "{date_since}")')
    
    if status != 'OK':
        print("No messages found!")
        return
        
    mail_ids = messages[0].split()
    print(f"Found {len(mail_ids)} recent emails in inbox. Filtering for job applications...")

    # We iterate backwards to get the most recent first, limiting to last 200 emails for speed
    for i in reversed(mail_ids[-200:]):
        status, msg_data = mail.fetch(i, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject = decode_subject(msg["Subject"])
                from_addr = msg.get("From", "")
                
                # Check if it's an application email
                if any(k.lower() in subject.lower() or k.lower() in from_addr.lower() for k in SEARCH_KEYWORDS):
                    print(f"Found match: {subject}")
                    info = extract_job_info(subject, from_addr)
                    
                    job_entry = {
                        "id": f"email_{int(i)}",
                        "title": info["title"] or "משרה מהמייל",
                        "company": info["company"] or "חברה לא ידועה",
                        "url": "",
                        "source": info["source"],
                        "score": "",
                        "status": "applied",
                        "priority": "mid",
                        "salary": "",
                        "notes": "יובא אוטומטית מהמייל",
                        "date": datetime.date.today().isoformat()
                    }
                    jobs.append(job_entry)
                    
    mail.logout()
    
    if jobs:
        output_file = "scraped_email_jobs.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(jobs, f, indent=2, ensure_ascii=False)
        print(f"\nSuccess! Extracted {len(jobs)} jobs to {output_file}.")
    else:
        print("\nNo job applications found in recent emails.")

if __name__ == "__main__":
    scrape_emails()
