import os
import requests
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText

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
    except Exception as e:
        print(f"Warning initializing historical logs: {str(e)}")
        return set()

def fetch_dubai_jobs_pipeline(history):
    print("Connecting to live LinkedIn job index...")
    url = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search?keywords=Dubai&location=United%20Arab%20Emirates&start=0"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    fresh_jobs = []
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code != 200:
            print(f"LinkedIn Feed error: Status code {response.status_code}")
            return fresh_jobs
            
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
            company = company_tag.text.strip() if company_tag else "Unknown Company"
            link = link_tag['href'].split('?')[0] if link_tag else "https://linkedin.com"
            pub_date = date_tag.text.strip() if date_tag else datetime.now().strftime("%d-%m-%Y")
            
            if clean_title.lower().strip() in history:
                continue
                
            job_data = {
                "title": f"{clean_title} ({company})",
                "workplace": "Dubai, UAE",
                "email": f"Apply directly on LinkedIn: {link}",
                "experience": "See LinkedIn job description",
                "ind": "Sourced via LinkedIn Jobs",
                "qualification": "Relevant Degree / Experience",
                "nationality": "Any",
                "gender": "Any",
                "expiry_date": "ASAP",
                "posted_date": pub_date,
                "job_type": "Full Time"
            }
            fresh_jobs.append(job_data)
            if len(fresh_jobs) >= TARGET_COUNT:
                break
    except Exception as e:
        print(f"Scraper structural checkpoint: {str(e)}")
        
    print(f"Scrape completed: Isolated {len(fresh_jobs)} brand-new openings from LinkedIn.")
    return fresh_jobs

def build_email_broadcast_payload(jobs):
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

def send_direct_email(content):
    """Sends the broadcast text directly using secure built-in Python SMTP protocol."""
    print("Initiating direct secure email transmission...")
    sender = "sngithacv@gmail.com"
    receiver = "sngithacv@gmail.com"
    
    # Using your active application password credentials
    app_password = "qyfx xbnm xgzo ytyu" 
    
    msg = MIMEText(content, "plain", "utf-8")
    msg["Subject"] = "🚀 RIGHTVOWS DAILY WHATSAPP BROADCAST MATRICES"
    msg["From"] = f"RightVows Bot <{sender}>"
    msg["To"] = receiver
    
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, app_password)
            server.sendmail(sender, [receiver], msg.as_string())
        print("Email transmitted successfully straight to your inbox!")
    except Exception as e:
        print(f"Direct mail transport failure: {str(e)}")

def save_jobs_to_excel(jobs):
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
    
    # Send the email immediately from the script execution context
    send_direct_email(broadcast_text)
    
    if target_subset:
        save_jobs_to_excel(target_subset)
