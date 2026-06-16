import os
import requests
import pandas as pd
from datetime import datetime

# Configuration settings
TARGET_COUNT = 25
EXCEL_FILE_PATH = "dubai_job_tracker.xlsx"

def init_excel_or_get_history():
    """Ensures the Excel tracking sheet exists and reads historical job descriptions to prevent duplication."""
    if not os.path.exists(EXCEL_FILE_PATH):
        # Create an empty template if it doesn't exist yet
        df = pd.DataFrame(columns=["Title", "Workplace", "Email", "Experience", "Industry", "Qualification", "Nationality", "Gender", "Expiry Date", "Posted Date", "Job Type"])
        df.to_excel(EXCEL_FILE_PATH, index=False)
        return set()
    
    try:
        # Read all sheet tabs to collect historical job titles/emails to avoid duplicates
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
    """
    Placeholder simulation for your web scraping infrastructure.
    Replace this internal logic with your real BeautifulSoup/Requests loop.
    """
    # This structure mirrors your live feed data requirements
    scraped_feed = [
        {
            "title": "Mechanical Engineer",
            "workplace": "On-site, Dubai Marina",
            "email": "careers@rightvowsmedia.com",
            "experience": "5-7 Years",
            "ind": "Engineering & Construction",
            "qualification": "B.E. / B.Tech in Mechanical Engineering",
            "nationality": "Any Nationality",
            "gender": "Any",
            "expiry_date": "30-07-2026",
            "posted_date": "17-06-2026",
            "job_type": "Full Time"
        }
        # Additional scraped items drop in here dynamically...
    ]
    
    # Filter out items that are already in your Excel history tracking vault
    fresh_jobs = []
    for job in scraped_feed:
        if job["title"].lower().strip() not in history:
            fresh_jobs.append(job)
            
    print(f"Scrape completed: Isolated {len(fresh_jobs)} brand-new openings.")
    return fresh_jobs

def build_email_broadcast_payload(jobs):
    """Formats the extracted jobs into your clean, customized WhatsApp Community text block."""
    payload = "🚀 *RIGHTVOWS LIVE DUBAI JOB BROADCAST* 🚀\n\n"
    
    for idx, job in enumerate(jobs, 1):
        payload += f"📌 *VACANCY NO {idx}: {job.get('title', 'N/A').upper()}*\n\n"
        payload += f"⛳ *Job Location:* UAE\n"
        payload += f"✅ *Gulf Experience:* Required\n"
        payload += f"⛳ *Work Place:* {job.get('workplace', 'N/A')}\n"
        payload += f"🕹️ *Visa Status:* Any\n"
        payload += f"💰 *Salary:* To be discussed\n\n"
        
        payload += f"📥 *Email:* Send CVs to {job.get('email', 'N/A')}\n\n"
        
        # Dynamic fields pulled directly from your job post structure
        payload += f"💼 *Experience:* {job.get('experience', 'As per job post')}\n"
        payload += f"🔎 *Industry:* {job.get('ind', 'N/A')}\n"
        payload += f"📚 *Qualification:* {job.get('qualification', 'Relevant Degree or Certification')}\n"
        payload += f"🌎 *Preferred Nationality:* {job.get('nationality', 'Any')}\n"
        payload += f"🚻 *Gender:* {job.get('gender', 'Any')}\n"
        payload += f"🕚 *Job Expiry Date:* {job.get('expiry_date', 'N/A')}\n"
        payload += f"📆 *Job Posted Date:* {job.get('posted_date', 'N/A')}\n"
        payload += f"🛡️ *Job Type:* {job.get('job_type', 'Full Time')}\n"
        payload += f"🚀 *Source of Vacancy:* RightVows Job Store\n\n"
        payload += "----------------------------------------\n\n"
        
    payload += "Best Wishes,\n*HR Team*\n*RightVows*\n_Connecting Your Talent_"
    return payload

def save_jobs_to_excel(jobs):
    """Appends today's parsed data as a brand new dated sheet tab inside the master workbook."""
    if not jobs:
        return
        
    today_str = datetime.now().strftime("%d-%m-%Y")
    df_new = pd.DataFrame(jobs)
    
    # Capitalize columns to match historical structures
    df_new.columns = [c.replace('_', ' ').title() for c in df_new.columns]
    
    try:
        # Read existing tabs so we append without overwriting historical sheets
        with pd.ExcelWriter(EXCEL_FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_new.to_excel(writer, sheet_name=today_str, index=False)
        print(f"Excel updated successfully. Added tab: {today_str}")
    except Exception:
        # Fallback mechanism if writing/appending throws an operational system error
        with pd.ExcelWriter(EXCEL_FILE_PATH, engine="openpyxl", mode="w") as writer:
            df_new.to_excel(writer, sheet_name=today_str, index=False)

if __name__ == "__main__":
    # Step 1: Initialize existing database logs
    history_logs = init_excel_or_get_history()
    
    # Step 2: Fetch and filter incoming vacancies
    fresh_vacancies = fetch_dubai_jobs_pipeline(history_logs)
    
    if fresh_vacancies:
        # Enforce target limits to capture up to 25 records max
        target_subset = fresh_vacancies[:TARGET_COUNT]
        
        # Step 3: Build the copy-paste WhatsApp broadcast text layout
        broadcast_text = build_email_broadcast_payload(target_subset)
        
        # Step 4: Write text payload to a permanent file for your GitHub Email Runner to extract
        with open("latest_job_broadcast.txt", "w", encoding="utf-8") as f:
            f.write(broadcast_text)
        print("Success: Broadcast block exported safely for email dispatch.")
        
        # Step 5: Log data cleanly inside your tracking workbook
        save_jobs_to_excel(target_subset)
    else:
        # Fallback dump so your email runner doesn't throw blank errors if no new jobs drop
        with open("latest_job_broadcast.txt", "w", encoding="utf-8") as f:
            f.write("No new unique vacancies discovered today! Check back tomorrow for fresh updates.")
        print("Pipeline Idle: No new vacancies found to log today.")
