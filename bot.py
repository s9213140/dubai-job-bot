import os
import re
import time
import hashlib
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup
import requests

# --- CONFIGURATION (Pulled securely from GitHub Secrets) ---
LI_AT_COOKIE = os.getenv("LI_AT_COOKIE")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

EXCEL_FILE = "dubai_job_tracker.xlsx"
TARGET_COUNT = 25

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Cookie": f"li_at={LI_AT_COOKIE};",
}

def generate_job_hash(title, company):
    """Generates a unique fingerprint to ensure no duplicate posts are processed."""
    combined_string = f"{str(title).strip().lower()}_{str(company).strip().lower()}"
    return hashlib.md5(combined_string.encode('utf-8')).hexdigest()[:10]

def init_excel_or_get_history():
    """Initializes a styled tracking ledger if missing, and loads existing job hashes to prevent duplicates."""
    existing_hashes = set()
    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb["Job Database"]
            for r in range(2, ws.max_row + 1):
                cell_val = ws.cell(row=r, column=1).value
                if cell_val:
                    existing_hashes.add(str(cell_val).strip())
            return existing_hashes
        except Exception:
            pass

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job Database"
    ws.views.sheetView[0].showGridLines = True
    
    headers_list = [
        "Job ID (Hash)", "Job Title", "Company", "Location", "Workplace", 
        "Corporate Email", "Experience", "Industry", "Qualification", 
        "Posted Date", "Scraped Date"
    ]
    ws.append(headers_list)
    ws.row_dimensions[1].height = 26
    
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for col_num, h_text in enumerate(headers_list, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        
    wb.save(EXCEL_FILE)
    return existing_hashes

def save_jobs_to_excel(jobs_to_save):
    """Appends newly sent jobs to the spreadsheet tracker with formatting."""
    if not jobs_to_save:
        return
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Job Database"]
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    current_row = ws.max_row + 1
    for job in jobs_to_save:
        row_data = [
            job["hash"], job["title"], job["company"], "Dubai, UAE", job["workplace"],
            job["email"], "Minimum 5 Years", job["ind"], "Relevant Degree or Certification",
            job["posted_date"], datetime.date.today().strftime("%d-%m-%Y")
        ]
        ws.append(row_data)
        ws.row_dimensions[current_row].height = 20
        
        for col_idx in range(1, len(row_data) + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if col_idx in [1, 10, 11]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
        current_row += 1

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
        
    wb.save(EXCEL_FILE)

def extract_email_from_text(text):
    """Scans description for corporate emails, excluding Gmail accounts."""
    emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    for email in emails:
        if "@gmail.com" not in email.lower():
            return email
    return None

def fetch_linkedin_dubai_jobs(history_hashes):
    """Queries LinkedIn specifically for Dubai jobs, applying filtration boundaries."""
    valid_jobs = []
    start_index = 0
    while len(valid_jobs) < TARGET_COUNT and start_index < 300:
        search_url = f"https://www.linkedin.com/voyager/api/voyagerCatalogDashJobPostings?geoId=105114972&keywords=Hiring&start={start_index}&count=25"
        try:
            response = requests.get(search_url, headers=headers)
            if response.status_code != 200:
                break
            elements = response.json().get("elements", [])
        except Exception:
            break
            
        if not elements:
            break

        for item in elements:
            if len(valid_jobs) >= TARGET_COUNT:
                break
            title = item.get("title", "Job Vacancy")
            company = item.get("companyName", "A Reputed Company")
            
            job_hash = generate_job_hash(title, company)
            if job_hash in history_hashes:
                continue

            description_text = BeautifulSoup(item.get("description", {}).get("text", ""), "html.parser").get_text()
            corporate_email = extract_email_from_text(description_text)

            if corporate_email:
                short_desc = description_text[:320].strip().replace("\n", " ") + "..."
                industry = item.get("industry", "Corporate / Specified") if item.get("industry") else "Construction / Infrastructure"
                job_object = {
                    "hash": job_hash, "title": title, "company": company, "desc": short_desc,
                    "workplace": "Dubai, UAE", "email": corporate_email, "ind": industry,
                    "posted_date": datetime.date.today().strftime("%d-%m-%Y"),
                    "expiry_date": (datetime.date.today() + datetime.timedelta(days=30)).strftime("%d-%m-%Y")
                }
                valid_jobs.append(job_object)
                history_hashes.add(job_hash)

        start_index += 25
        time.sleep(2)
    return valid_jobs

def build_whatsapp_payload(jobs_list):
    """Assembles the final text layout tailored perfectly to your RightVows branding configuration."""
    current_date = datetime.date.today().strftime("%d/%m/%Y")
    full_message = ""
    for job in jobs_list:
        full_message += f"""⭐ Hiring ⭐

⭐ Job Title: {job['title']} ⭐
 
📝 Job Description:
{job['desc']}

⛳ Job Location: UAE
✅ Gulf Experience: Required
⛳ Work Place: {job['workplace']}
🕹️ Visa Status: Any
💰 Salary: To be discussed

📥 Email: Send CVs to {job['email']}

💼 Experience: Minimum 5 Years
🔎 Industry: {job['ind']}
📚 Qualification: Relevant Degree or Certification
🌎 Preferred Nationality: Any
🚻 Gender: Any
🕚 Job Expiry Date: {job['expiry_date']}
📆 Job Posted Date: {job['posted_date']}
🛡️ Job Type: Full Time
🚀 Source of Vacancy: RightVows Job Store

Best Wishes,
HR Team
RightVows
Connecting Your Talent

📄 ATS CV Writing & Video CV Services: WhatsApp us at +971 543078783
📲 Download the RightVows Mobile App for faster job updates:
iOS 🍎 https://apple.co/2Z9oLQ6
Android 🤖 https://bit.ly/2BRBrSK
==================================\n"""
    
    full_message += f"\nPrepare Job Title Listings in the below Format \n\nRightVows Epilogue \n\nSummary of vacancies posted on {current_date} through RightVows JobStore\n\nJob Positions\n\n"
    for index, job in enumerate(jobs_list, 1):
        full_message += f"{index}. {job['title']}\n"
    
    full_message += """\nYes, we still update vacancies for you on a daily basis through RightVows Job Store\n\nTo get vacancies pls Join the biggest Job Search Whatsapp Group in GCC with 80 Nationalities \nJoining Link 🔃 https://rightvows.com/join/  (LifeTime Membership)\n\nand\n\nDownload our mobile application from 📲 App Store or Play Store \n\nApp Link 📲https://rightvows.app.link or search RightVows\n\nSubscribe to our YouTube channel for all updates  Link https://youtube.com/c/RightVows\n\nBest Wishes..\nHR Team\nRightVows\nConnecting Your Talent"""
    return full_message

def send_to_telegram(message):
    """Sends the complete structured block directly to your personal messaging screen via your bot access."""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    
    if len(message) > 4000:
        chunks = [message[i:i+4000] for i in range(0, len(message), 4000)]
        for chunk in chunks:
            payload = {"chat_id": TELEGRAM_CHAT_ID, "text": chunk}
            requests.post(url, json=payload)
            time.sleep(0.5)
    else:
        payload = {"chat_id": TELEGRAM_CHAT_ID, "text": message}
        requests.post(url, json=payload)

if __name__ == "__main__":
    history = init_excel_or_get_history()
    fresh_jobs = fetch_linkedin_dubai_jobs(history)
    if fresh_jobs:
        payload = build_whatsapp_payload(fresh_jobs[:TARGET_COUNT])
        send_to_telegram(payload)
        save_jobs_to_excel(fresh_jobs[:TARGET_COUNT])
    else:
        send_to_telegram("RightVows Check Complete: No new unique vacancies matching corporate criteria detected today.")
