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

    # Build stylized spreadsheet structure if it doesn't exist
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
        ws.column_dimensions[col_letter].width = max(max_len
