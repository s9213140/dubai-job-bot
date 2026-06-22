import os
import requests
import pandas as pd
from datetime import datetime

TARGET_COUNT = 25
EXCEL_FILE_PATH = "dubai_job_tracker.xlsx"

def init_excel_or_get_history():
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
    except Exception:
        return set()

def fetch_dubai_jobs_pipeline(history):
    print("Connecting to public LinkedIn skilled job index (Past 24 Hours)...")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    fresh_jobs = []
    keywords = "Manager%20OR%20Engineer%20OR%20Developer%20OR%20Executive%20OR%20Specialist%20OR%20Analyst"
    
    # Expanded loop to check pages 1, 2, 3, and 4 to confidently reach 25 unique items
    for start_index in [0, 25, 50, 75]:
        if len(fresh_jobs) >= TARGET_COUNT:
            break
        url = f"https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search?keywords={keywords}&location=Dubai%2C%20United%20Arab%20Emirates&f_TPR=r86400&start={start_index}"
        try:
            response = requests.get(url, headers=headers, timeout=15)
            if response.status_code != 200:
                continue
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(response.text, 'html.parser')
            job_cards = soup.find_all("li")
            
            for card in job_cards:
                title_tag = card.find("h3", class_="base-search-card__title")
                company_tag = card.find("h4", class_="base-search-card__subtitle")
                link_tag = card.find("a", class_="base-card__full-link")
                date_tag = card.find("time", class_="job-search-card__listdate")
                
                if not title_tag:
                    continue
                    
                clean_title = title_tag.text.strip()
                company = company_tag.text.strip() if company_tag else "Professional Employer"
                link = link_tag['href'].split('?')[0] if link_tag else "https://linkedin.com"
                pub_date = date_tag.text.strip() if date_tag else "Posted Today"
                
                if clean_title.lower().strip() in history:
                    continue
                    
                job_data = {
                    "title": f"{clean_title} ({company})",
                    "workplace": "Dubai, UAE",
                    "email": f"Apply via LinkedIn: {link}",
                    "experience": "See job details on LinkedIn",
                    "ind": "Sourced via LinkedIn Professional Index",
                    "qualification": "Relevant Professional Degree",
                    "nationality": "Any",
                    "gender": "Any",
                    "expiry_date": "ASAP",
                    "posted_date": pub_date,
                    "job_type": "Full Time"
                }
                if job_data not in fresh_jobs:
                    fresh_jobs.append(job_data)
        except Exception:
            pass
            
    print(f"Scrape completed: Isolated {len(fresh_jobs)} fresh skilled jobs.")
    return fresh_jobs

def build_broadcast_payload(jobs):
    payload = "🚀 *RIGHTVOWS LIVE DUBAI SKILLED JOB BROADCAST* 🚀\n\n"
    if not jobs:
        payload += "No new unique skilled vacancies discovered in the last 24 hours! Check back shortly.\n\n"
    else:
        for idx, job in enumerate(jobs, 1):
            payload += f"📌 *VACANCY NO {idx}: {job['title'].upper()}*\n\n"
            payload += f"⛳ *Job Location:* Dubai, UAE\n"
            payload += f"✅ *Gulf Experience:* Preferred / Required\n"
            payload += f"⛳ *Work Place:* {job['workplace']}\n"
            payload += f"🕹️ *Visa Status:* Open / Any\n"
            payload += f"💰 *Salary:* Industry Standard (To be discussed)\n\n"
            payload += f"📥 *Application Link:* {job['email']}\n\n"
            payload += f"💼 *Experience:* {job['experience']}\n"
            payload += f"🔎 *Industry:* {job['ind']}\n"
            payload += f"🛡️ *Job Type:* {job['job_type']}\n\n"
            payload += "----------------------------------------\n\n"
    payload += "Best Wishes,\n*HR Team*\n*RightVows*"
    return payload

def save_jobs_to_excel(jobs):
    if not jobs:
        return
    today_str = datetime.now().strftime("%d-%m-%Y")
    df_new = pd.DataFrame(jobs)
    df_new.columns = [c.replace('_', ' ').title() for c in df_new.columns]
    try:
        with pd.ExcelWriter(EXCEL_FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_new.to_excel(writer, sheet_name=today_str, index=False)
        print(f"Excel tracker updated successfully for date: {today_str}")
    except Exception:
        with pd.ExcelWriter(EXCEL_FILE_PATH, engine="openpyxl", mode="w") as writer:
            df_new.to_excel(writer, sheet_name=today_str, index=False)
        print(f"Excel tracker created successfully for date: {today_str}")

if __name__ == "__main__":
    history_logs = init_excel_or_get_history()
    fresh_vacancies = fetch_dubai_jobs_pipeline(history_logs)
    target_subset = fresh_vacancies[:TARGET_COUNT]
    
    broadcast_text = build_broadcast_payload(target_subset)
    with open("whatsapp_broadcast.txt", "w", encoding="utf-8") as f:
        f.write(broadcast_text)
    print("WhatsApp broadcast text block generated successfully.")
        
    if target_subset:
        save_jobs_to_excel(target_subset)
