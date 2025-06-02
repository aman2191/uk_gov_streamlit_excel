import os
import re
import time
import sys
import csv
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (NoSuchElementException,
                                      TimeoutException,
                                      ElementClickInterceptedException)
from difflib import SequenceMatcher
import PyPDF2

# Initialize session state
if 'processing' not in st.session_state:
    st.session_state.processing = False
if 'stats' not in st.session_state:
    st.session_state.stats = {
        'total': 0,
        'processed': 0,
        'downloaded': 0,
        'failed': 0,
        'succeeded': [],
        'failed_companies': [],
        'unprocessed': []
    }
if 'log' not in st.session_state:
    st.session_state.log = []
if 'report_generated' not in st.session_state:
    st.session_state.report_generated = False

# Constants
current_date = datetime.now().strftime("%Y-%m-%d")
current_time = datetime.now().strftime("%H-%M-%S")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "PDF_FILES", current_date)
BASE_DIR = os.path.join(os.getcwd(), "REPORTS", current_date)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(BASE_DIR, exist_ok=True)

# Helper functions
def log_result(company_name, input_date, status, reason="", downloaded_pdf_name=""):
    log_entry = {
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'company': company_name,
        'date': input_date,
        'status': status,
        'reason': reason,
        'file': downloaded_pdf_name
    }
    st.session_state.log.append(log_entry)

def add_log_message(message):
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.log.append(f"[{timestamp}] {message}")

def similarity_ratio(a, b):
    return SequenceMatcher(None, a, b).ratio()

def parse_date(date_str):
    formats = ["%d %B %Y", "%d-%m-%Y", "%d/%m/%Y"]
    for fmt in formats:
        try:
            date_obj = datetime.strptime(date_str, fmt)
            return {
                "month_in_word": date_obj.strftime("%d %B %Y"),
                "month_in_num": date_obj.strftime("%d/%m/%Y"),
                "filename_date": date_obj.strftime("%Y%m%d")
            }
        except ValueError:
            continue
    raise ValueError(f"Invalid date format: {date_str}")

def get_pdf_content(url):
    add_log_message(f"📥 Downloading PDF from {url}")
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    return BytesIO(response.content)

def parse_pdf_content(pdf_content):
    reader = PyPDF2.PdfReader(pdf_content)
    text = ""
    charge_code = ""
    for page in reader.pages:
        page_text = page.extract_text() or ""
        text += page_text + "\n"
        if not charge_code:
            match = re.search(r"Charge code:\s*(\d{3,4}\s*\d{3,4}\s*\d{3,4})", page_text)
            if match:
                charge_code = match.group(1).replace(" ", "")
    return text, charge_code

def extract_pdf_info(pdf_text):
    normalized_pdf = ' '.join(pdf_text.replace('\n', ' ').split()).upper()
    company_match = re.search(r'COMPANY NAME:\s*(.*?)\s*COMPANY NUMBER:', normalized_pdf)
    company_name = company_match.group(1).strip() if company_match else None
    desc_match = re.search(r'BRIEF DESCRIPTION:\s*(.*?)(?=(CONTAINS|AUTHENTICATION OF FORM|CERTIFIED BY:|CERTIFICATION STATEMENT:))', normalized_pdf)
    brief_description = desc_match.group(1).strip() if desc_match else None
    if brief_description:
        stop_phrases = ['CONTAINS FIXED CHARGE', 'CONTAINS NEGATIVE PLEDGE', 
                       'CONTAINS FLOATING CHARGE', 'CONTAINS']
        for phrase in stop_phrases:
            if phrase in brief_description:
                brief_description = brief_description.split(phrase)[0].strip()
    date_match = re.search(r'DATE OF CREATION:\s*(\d{2}/\d{2}/\d{4})', normalized_pdf)
    month_in_num = date_match.group(1) if date_match else None
    entitled_match = re.search(r'PERSONS ENTITLED:\s*(.*?)(?=(CHARGE|DATE OF CREATION|BRIEF DESCRIPTION|AUTHENTICATION|CERTIFIED BY:|CERTIFICATION STATEMENT:))', normalized_pdf)
    persons_entitled = entitled_match.group(1).strip() if entitled_match else None
    return {
        'company_name': company_name,
        'brief_description': brief_description,
        'month_in_num': month_in_num,
        'persons_entitled': persons_entitled
    }

def check_pdf_conditions(pdf_text, date_info, company_name, persons_entitled, brief_description):
    result = extract_pdf_info(pdf_text)
    result_text = (result['brief_description'] or '').upper()
    input_text = (brief_description or '').upper()
    cmp_result_txt = (result['company_name'] or '').upper()
    in_cmp_result_txt = (company_name or '').upper()

    cmp_name_result = int(similarity_ratio(cmp_result_txt, in_cmp_result_txt) * 100)
    brief_description_score = int(similarity_ratio(result_text, input_text) * 100)
    persons_entitled_score = int(similarity_ratio((result['persons_entitled'] or '').upper(), persons_entitled)) * 100

    conditions_met = True
    if cmp_name_result < 80:
        conditions_met = False
    if persons_entitled_score < 80:
        conditions_met = False
    if brief_description not in result_text:
        conditions_met = False
    if date_info["month_in_num"] != result['month_in_num']:
        conditions_met = False
    return conditions_met

def sanitize_filename(name):
    return re.sub(r'[\\/:"*?<>|]', '_', name)

def save_pdf_file(pdf_content, filename, date_info):
    base_name = f"{filename}"
    safe_filename = sanitize_filename(base_name)
    full_path = os.path.join(DOWNLOAD_DIR, safe_filename)
    with open(full_path, "wb") as f:
        f.write(pdf_content.getvalue())
    add_log_message(f"💾 Saved PDF: {safe_filename}")
    time.sleep(1)
    return full_path

def get_company_info(company_name, persons_entitled, brief_description, input_date):
    company_name = company_name.upper()
    persons_entitled = persons_entitled.upper()
    brief_description = brief_description.upper()
    date_info = parse_date(input_date)
    success = False
    downloaded_filename = ""

    # service = Service(ChromeDriverManager().install())
    # options = webdriver.ChromeOptions()
    # options.add_argument("--headless=new")
    # options.add_argument("--start-maximized")
    # driver = webdriver.Chrome(service=service, options=options)
    
    chrome_driver_path = os.path.join(os.getcwd(), "chromedriver")
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    service = Service(executable_path=chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    
    try:
        driver.get("https://find-and-update.company-information.service.gov.uk/")
        time.sleep(2)
        wait = WebDriverWait(driver, 30)

        search_box = wait.until(EC.presence_of_element_located((By.ID, "site-search-text")))
        search_box.send_keys(company_name + Keys.RETURN)
        time.sleep(2)

        company_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//a[contains(., '{company_name}') and contains(@href, '/company/')]")))
        company_link.click()
        time.sleep(2)

        filing_history_tab = wait.until(EC.element_to_be_clickable((By.ID, "filing-history-tab")))
        filing_history_tab.click()
        time.sleep(2)

        charges_filter_label = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//label[@for='filter-category-mortgage']")))
        driver.execute_script("arguments[0].scrollIntoView(true);", charges_filter_label)
        time.sleep(1)

        try:
            charges_filter_label.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", charges_filter_label)

        time.sleep(2)

        try:
            wait.until(EC.presence_of_element_located((By.ID, "fhTable")))
        except TimeoutException:
            add_log_message(f"❌ No filings available for {company_name}")
            log_result(company_name, input_date, "Failed", "No filings available", "")
            return False

        rows = driver.find_elements(By.CSS_SELECTOR, "#fhTable tbody tr:not(:first-child)")
        for row in rows:
            try:
                description = row.find_element(By.CSS_SELECTOR, "td:nth-child(3)").text
                if date_info["month_in_word"].split()[1] in description:
                    pdf_link = row.find_element(By.CSS_SELECTOR, "a[href*='/document']")
                    pdf_url = pdf_link.get_attribute("href")
                    pdf_content = get_pdf_content(pdf_url)
                    pdf_text, charge_code = parse_pdf_content(pdf_content)

                    if check_pdf_conditions(pdf_text, date_info, company_name, 
                                          persons_entitled, brief_description):
                        add_log_message("✅ PDF Matched All Requirements!")
                        filename = f"{company_name}_{date_info['month_in_num']}_downloaded_file.pdf"
                        saved_path = save_pdf_file(pdf_content, filename, date_info)
                        downloaded_filename = os.path.basename(saved_path)
                        success = True
                        break
            except Exception as e:
                add_log_message(f"⚠️ Error processing filing: {str(e)}")
                log_result(company_name, input_date, "Failed", "Processing error", "")

        if not success:
            add_log_message(f"❌ No valid filings for {company_name}")
            log_result(company_name, input_date, "Failed", "No valid filings", "")
    except Exception as e:
        add_log_message(f"🚨 Critical error: {str(e)}")
        log_result(company_name, input_date, "Failed", "Critical error", "")
    finally:
        driver.quit()
    
    if success:
        log_result(company_name, input_date, "Success", "", downloaded_filename)
    return success

def generate_summary_file():
    try:
        unprocessed = [
            company for company in st.session_state.stats['unprocessed']
            if company not in st.session_state.stats['succeeded']
            and company not in st.session_state.stats['failed_companies']
        ]

        summary_data = {
            'Metric': [
                'Total companies in Excel',
                'Total processed',
                'Successfully downloaded',
                'Failed',
                'Unprocessed'
            ],
            'Value': [
                st.session_state.stats['total'],
                st.session_state.stats['processed'],
                st.session_state.stats['downloaded'],
                st.session_state.stats['failed'],
                len(unprocessed)
            ]
        }
        df_summary = pd.DataFrame(summary_data)
        df_success = pd.DataFrame({'Succeeded Companies': st.session_state.stats['succeeded']})
        df_failed = pd.DataFrame({'Failed Companies': st.session_state.stats['failed_companies']})
        df_unprocessed = pd.DataFrame({'Unprocessed Companies': unprocessed})

        summary_path = os.path.join(BASE_DIR, 'final_processed_company_count_results.xlsx')
        with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            df_success.to_excel(writer, sheet_name='Succeeded', index=False)
            df_failed.to_excel(writer, sheet_name='Failed', index=False)
            df_unprocessed.to_excel(writer, sheet_name='Unprocessed', index=False)

        add_log_message(f"✅ Excel summary report generated")
        return summary_path
    except Exception as e:
        add_log_message(f"⚠️ Failed to generate summary file: {str(e)}")
        return None

# Streamlit UI
st.title("UK Charge Report Scraper")
st.write("Upload an Excel file to process company information")

# File upload
uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx", "xls"])

# Processing controls
col1, col2 = st.columns(2)
with col1:
    start_btn = st.button("Start Processing", disabled=st.session_state.processing)
with col2:
    stop_btn = st.button("Stop Processing", disabled=not st.session_state.processing)

# Statistics
st.subheader("Processing Statistics")
stats_cols = st.columns(4)
with stats_cols[0]:
    st.metric("Total Companies", st.session_state.stats['total'])
with stats_cols[1]:
    st.metric("Processed", st.session_state.stats['processed'])
with stats_cols[2]:
    st.metric("Downloaded", st.session_state.stats['downloaded'])
with stats_cols[3]:
    st.metric("Failed", st.session_state.stats['failed'])

# Log display
st.subheader("Processing Log")
log_container = st.container(height=300)
for entry in st.session_state.log[-20:]:
    log_container.write(entry)

# Processing logic
if start_btn and uploaded_file is not None:
    st.session_state.processing = True
    st.session_state.report_generated = False
    df = pd.read_excel(uploaded_file)
    st.session_state.stats['total'] = len(df)
    st.session_state.stats['unprocessed'] = df['company_name'].tolist()

    for idx, row in df.iterrows():
        if not st.session_state.processing:
            break

        try:
            input_date = str(row['input_date']).split()[0]
            formatted_date = datetime.strptime(input_date, '%Y-%m-%d').strftime('%d/%m/%Y')
            company_name = row['company_name']

            success = get_company_info(
                company_name=company_name,
                persons_entitled=row['persons_entitled'],
                brief_description=row['brief_description'],
                input_date=formatted_date
            )

            st.session_state.stats['processed'] += 1
            if success:
                st.session_state.stats['downloaded'] += 1
                st.session_state.stats['succeeded'].append(company_name)
            else:
                st.session_state.stats['failed'] += 1
                st.session_state.stats['failed_companies'].append(company_name)

            st.session_state.stats['unprocessed'].remove(company_name)
            time.sleep(2)

        except Exception as e:
            add_log_message(f"⚠️ Error processing row {idx+1}: {str(e)}")
            st.session_state.stats['processed'] += 1
            st.session_state.stats['failed'] += 1

    st.session_state.processing = False
    st.session_state.report_generated = True
    st.rerun()

if stop_btn:
    st.session_state.processing = False
    add_log_message("🛑 Processing stopped by user")
    st.rerun()

# Report download
if st.session_state.report_generated:
    st.success("Processing completed!")
    report_path = generate_summary_file()
    
    if report_path:
        with open(report_path, "rb") as f:
            st.download_button(
                label="Download Summary Report",
                data=f,
                file_name="processing_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )