import os
import requests
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET

# Configuration settings
TARGET_COUNT = 25
EXCEL_FILE_PATH = "dubai_job_tracker.xlsx"

def init_excel_or_get_history():
    """Ensures the Excel tracking sheet exists and reads historical job descriptions to prevent duplication."""
    if not os.path.exists(EXCEL_FILE_PATH):
        df = pd.DataFrame(columns=["Title", "Workplace", "Email", "Experience", "Industry", "Qualification", "Nationality", "Gender", "Expiry Date", "Posted Date", "Job Type"])
        df.to_excel(EXCEL_FILE_PATH, index=False)
        return set()
    
    try:
        excel_file = pd.ExcelFile(EXCEL_FILE_PATH)
        history = set()
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet_name)
            if not df.empty and "Title" in df.columns:
                history.update(df["Title"].astype(str).str.lower().strip().tolist())
        return history
    except Exception as e:
        print(f"Warning initializing historical logs: {str(e)}")
        return set()

def fetch_dubai_jobs_pipeline(history):
    """Scrapes live Dubai jobs directly from LinkedIn's public RSS/Guest Feed API."""
    print("Connecting to live LinkedIn job index...")
    
    # URL targeting public job postings in the United Arab Emirates / Dubai
    url = "https://www.linkedin.com/jobs/rss/search?keywords=Dubai&location=United%20Arab%20Emirates&geoId=104305776"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    fresh_jobs = []
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code != 200:
            print(f"LinkedIn Feed error: Status code {response.status_code}")
            return fresh_jobs
            
        # Parse the XML structure from the feed
        root = ET.fromstring(response.content)
        
        for item in root.findall(".//item"):
            title = item.find("title").text if item.find("title") is not None else "N/A"
            link = item.find("link").text if item.find("link") is not None else "N/A"
            pub_date = item.find("pubDate").text if item.find("pubDate") is not None else datetime.now().strftime("%d-%m-%Y")
            
            # Clean up formatting strings
            clean_title = title.split(" at ")[0].strip()
            company = title.split(" at ")[1].strip() if " at " in title else "Unknown Company"
            
            # Skip duplicates against this week's history log
            if clean_title.lower().strip() in history:
                continue
                
            # Formulate the data entry dictionary matching your fields
            job_data = {
                "title": f"{clean_title} ({company})",
                "workplace": "See Description / Link",
                "email": f"Apply via Link: {link}",
                "experience": "As per job post details",
                "ind": "Sourced via LinkedIn",
                "qualification": "Relevant Degree or Certification",
                "nationality": "Any",
                "gender": "Any",
                "expiry_date": "ASAP",
                "posted_date": pub_date[:16], # Extract clean date string segment
                "job_type": "Full Time"
            }
            
            fresh_jobs.append(job_data)
            if len(fresh_jobs) >= TARGET_COUNT:
                break
                
    except Exception as e:
        print(f"Scraper execution glitch: {str(e)}")
        
    print(f"Scrape completed: Isolated {len(fresh_jobs)} brand-new openings from LinkedIn.")
    return fresh_jobs

def build_email_broadcast_payload(jobs):
    """Formats the extracted jobs into your clean, customized WhatsApp Community text block."""
    payload = "🚀 *RIGHTVOWS LIVE DUBAI JOB BROADCAST* 🚀\n\n"
    
    if not jobs:
        payload += "No new unique vacancies discovered today! Check back tomorrow for fresh updates.\n\n"
    else:
        for idx, job in enumerate(jobs, 1):
            payload += f"📌 *VACANCY NO {idx}: {job.get('title', 'N/A').upper()}*\n\n"
            payload += f"⛳ *Job Location:* UAE\n"
            payload += f"✅ *Gulf Experience:* Required\n"
            payload += f"⛳ *Work Place:* {job.get('workplace', 'N/A')}\n"
            payload += f"🕹️ *Visa Status:* Any\n"
            payload += f"💰 *Salary:* To be discussed\n\n"
            
            payload += f"📥 *Application/Email:* {job.get('email', 'N/A')}\n\n"
            
            payload += f"💼 *Experience:* {job.get('experience', 'As per job post')}\n"
            payload += f"🔎 *Industry:* {job.get('ind', 'N/A')}\n"
            payload += f"📚 *Qualification:* {job.get('qualification', 'Relevant Degree or Certification')}\n"
            payload += f"🌎 *Preferred Nationality:* {job.get('nationality', 'Any')}\n"
            payload += f"🚻 *Gender:* {job.get('gender', 'Any')}\n"
            payload += f"🕚 *Job Expiry Date:* {job.get('expiry_date', 'N/A')}\n"
            payload += f"📆 *Job Posted Date:* {job.get('posted_date', 'N/A')}\n"
            payload += f"🛡️ *Job Type:* {job.get('job_type', 'Full Time')}\n"
            payload += f"🚀 *Source of Vacancy:* RightVows Job Store (via LinkedIn)\n\n"
            payload += "----------------------------------------\n\n"
        
    payload += "Best Wishes,\n*HR Team*\n*RightVows*\n_Connecting Your Talent_"
    return payload

def save_jobs_to_excel(jobs):
    """Appends today's parsed data as a brand new dated sheet tab inside the master workbook."""
    if not jobs:
        return
        
    today_str = datetime.now().strftime("%d-%m-%Y")
    df_new = pd.DataFrame(jobs)
    df_new.columns = [c.replace('_', ' ').title() for c in df_new.columns]
    
    try:
        with pd.ExcelWriter(EXCEL_FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_new.to_excel(writer, sheet_name=today_str, index=False)
        print(f"Excel updated successfully. Added tab: {today_str}")
    except Exception:
        with pd.ExcelWriter(EXCEL_FILE_PATH, engine="openpyxl", mode="w") as writer:
            df_new.to_excel(writer, sheet_name=today_str, index=False)

if __name__ == "__main__":
    history_logs = init_excel_or_get_history()
    fresh_vacancies = fetch_dubai_jobs_pipeline(history_logs)
    
    target_subset = fresh_vacancies[:TARGET_COUNT]
    broadcast_text = build_email_broadcast_payload(target_subset)
    
    with open("latest_job_broadcast.txt", "w", encoding="utf-8") as f:
        f.write(broadcast_text)
    
    if target_subset:
        save_jobs_to_excel(target_subset)
