import os
import re
import time
import csv
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (NoSuchElementException, TimeoutException, ElementClickInterceptedException)
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
    add_log_message(f"\U0001F4E5 Downloading PDF from {url}")
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