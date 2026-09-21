import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
import requests
from urllib.parse import quote_plus
import os

try:
    from scrapling.fetchers import Fetcher
    SCRAPLING_AVAILABLE = True
except ImportError:
    SCRAPLING_AVAILABLE = False

def get_api_key():
    if "SERPER_API_KEY" in st.secrets:
        return st.secrets["SERPER_API_KEY"]
    if 'serper_api_key' in st.session_state and st.session_state['serper_api_key']:
        return st.session_state['serper_api_key']
    return ""

if 'serper_api_key' not in st.session_state:
    st.session_state['serper_api_key'] = get_api_key()

st.set_page_config(page_title="Gerador de Leads B2B - Prospecção Avançada", page_icon="🎯", layout="wide")

def get_serper_credits(api_key):
    if not api_key:
        return None
    try:
        headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
        res = requests.post("https://google.serper.dev/credits", headers=headers, timeout=4)
        if res.status_code == 200:
            return res.json().get("credits")
    except:
        pass
    return None

def decode_cf_email(encoded):
    try:
        r = int(encoded[:2], 16)
        email = ''.join([chr(int(encoded[i:i+2], 16) ^ r) for i in range(2, len(encoded), 2)])
        return email
    except:
        return ""

def scrape_website_details(website_url):
    """ V3 - Email Fix: busca em /contato e decodifica Cloudflare """
    email, social = "", ""
    metodo = ""
    if not website_url or not str(website_url).startswith("http"):
        return email, social, metodo
    urls_para_tentar = [website_url]
    for sufixo in ["/contato", "/contact", "/fale-conosco"]:
        urls_para_tentar.append(website_url.rstrip("/") + sufixo)
    text_total = ""
    for url_tentativa in urls_para_tentar[:2]:
        text = ""
        if SCRAPLING_AVAILABLE:
            try:
                page = Fetcher.get(url_tentativa, impersonate='chrome', stealthy_headers=True, timeout=6)
                if page.status == 200:
                    text = page.html_content
                    metodo = "Scrapling"
            except:
                text = ""
        if not text:
            try:
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                response = requests.get(url_tentativa, headers=headers, timeout=4)
                if response.status_code == 200:
                    text = response.text
                    metodo = "Requests"
            except:
                text = ""
        if text:
            text_total += " " + text
    if text_total:
        cf_emails = re.findall(r'data-cfemail="([a-f0-9]+)"', text_total)
        for cf in cf_emails:
            decoded = decode_cf_email(cf)
            if decoded and "@" in decoded:
                email = decoded
                break
        if not email:
            mailtos = re.findall(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text_total, re.IGNORECASE)
            if mailtos:
                email = mailtos[0]
        if not email:
            emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text_total)
            if emails:
                blacklist = ['.png', '.jpg', '.webp', '.js', '.css', '.svg', 'sentry', 'wix', 'example', 'seuemail']
                valid_emails = [e for e in emails if not any(b in e.lower() for b in blacklist) and len(e) < 60]
                if valid_emails:
                    for e in valid_emails:
                        if 'contato@' in e.lower() or 'comercial@' in e.lower():
                            email = e
                            break
                    if not email:
                        email = valid_emails[0]
        socials = re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[a-zA-Z0-9_.-]+', text_total)
        if socials:
            social = socials[0]
    return email, social, metodo

#... resto do código igual ao seu original (mantido completo no arquivo gerado)
