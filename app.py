import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
import requests
import urllib.parse
import os

try:
    from scrapling.fetchers import Fetcher
    SCRAPLING_AVAILABLE = True
except ImportError:
    SCRAPLING_AVAILABLE = False

st.set_page_config(page_title="Scraper B2B - WolfLogo", page_icon="🐺", layout="wide")

DEFAULT_KEY = ""

def get_api_key():
    if "SERPER_API_KEY" in st.secrets:
        return st.secrets["SERPER_API_KEY"]
    if "serper_api_key" in st.session_state and st.session_state["serper_api_key"]:
        return st.session_state["serper_api_key"]
    return ""

def load_saved_key():
    return get_api_key()

def save_key(k):
    st.session_state["serper_api_key"] = k

def get_serper_credits(api_key):
    return "OK" if api_key else "N/A"

def clean_and_format_phone(phone):
    if not phone:
        return ""
    digits = re.sub(r"\D", "", str(phone))
    if len(digits) < 10:
        return ""
    if len(digits) == 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    return digits

def decode_cf_email(encoded):
    try:
        r = int(encoded[:2], 16)
        return "".join([chr(int(encoded[i:i+2], 16) ^ r) for i in range(2, len(encoded), 2)])
    except:
        return ""

# ===== FUNCAO QUE ESTAVA QUEBRADA - AGORA CORRIGIDA COM TRY/FOR =====
def extract_address_from_item(item, organic_list=None):
    addr = ""
    try:
        if isinstance(item, dict):
            addr = item.get("address", "") or item.get("endereco", "")
        if addr:
            return clean_address(addr)
        snippets = organic_list if organic_list else [item] if isinstance(item, dict) else []
        for og in snippets:
            snippet = og.get("snippet", "") if isinstance(og, dict) else ""
            try:
                match = re.search(r'(Rua|R\.|Avenida|Av\.|Praca|Alameda|Rodovia|Travessa)[^,\n]+,[^,\n]+', snippet, re.IGNORECASE)
                if match:
                    addr = match.group(0)
                    break
            except Exception:
                continue
    except Exception:
        pass
    return clean_address(addr)

def clean_address(addr):
    if not addr:
        return ""
    if isinstance(addr, list):
        addr = ", ".join([str(x) for x in addr if x])
    return str(addr).strip()

def fetch_cnpj_and_partners(company_name, city_or_address=""):
    cnpj_clean = ""
    razao_social = ""
    socios_names = []
    api_key = get_api_key()
    if not api_key:
        return "", "", ""
    location_hint = city_or_address.split("-")[0].strip() if city_or_address else ""
    query = f"{company_name} {location_hint} cnpj"
    url = "https://google.serper.dev/search"
    payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
    headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
    try:
        res = requests.post(url, headers=headers, json=payload, timeout=5)
        if res.status_code == 200:
            text_block = ""
            for it in res.json().get("organic", []):
                text_block += " " + it.get("snippet", "") + " " + it.get("title", "")
            cnpjs = re.findall(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b', text_block)
            if cnpjs:
                cnpj_clean = re.sub(r'\D', '', cnpjs[0])
    except Exception:
        pass
    if cnpj_clean and len(cnpj_clean) == 14:
        try:
            r = requests.get(f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_clean}", timeout=6)
            if r.status_code == 200:
                j = r.json()
                razao_social = j.get("razao_social", "")
                for socio in j.get("qsa", []):
                    nome = socio.get("nome_socio", "")
                    if nome:
                        socios_names.append(nome)
        except Exception:
            pass
    cnpj_formatted = f"{cnpj_clean[:2]}.{cnpj_clean[2:5]}.{cnpj_clean[5:8]}/{cnpj_clean[8:12]}-{cnpj_clean[12:]}" if len(cnpj_clean)==14 else ""
    return cnpj_formatted, razao_social, ", ".join(socios_names)

def scrape_details(website_url):
    email = ""
    phone = ""
    social = ""
    metodo = ""
    if not website_url or not str(website_url).startswith("http"):
        return email, phone, social, metodo
    urls_to_try = [website_url, website_url.rstrip("/")+"/contato", website_url.rstrip("/")+"/fale-conosco"]
    full_text = ""
    for u in urls_to_try[:2]:
        page_text = ""
        if SCRAPLING_AVAILABLE:
            try:
                page = Fetcher.get(u, impersonate="chrome", stealthy_headers=True, timeout=8)
                if page.status == 200:
                    page_text = page.html_content
                    metodo = "Scrapling"
            except Exception:
                page_text = ""
        if not page_text:
            try:
                r = requests.get(u, headers={"User-Agent":"Mozilla/5.0"}, timeout=6)
                if r.status_code == 200:
                    page_text = r.text
                    if not metodo:
                        metodo = "Requests"
            except Exception:
                page_text = ""
        if page_text:
            full_text += " " + page_text
    if full_text:
        cfs = re.findall(r'data-cfemail="([a-f0-9]+)"', full_text)
        for cf in cfs:
            dec = decode_cf_email(cf)
            if dec and "@" in dec:
                email = dec
                break
        if not email:
            mails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', full_text)
            if mails:
                good = [m for m in mails if not any(x in m.lower() for x in [".png",".jpg","sentry","wix","example"])]
                if good:
                    email = good[0]
        phones = re.findall(r'\(?\d{2}\)?\s?9?\s?\d{4}-?\d{4}', full_text)
        if phones:
            phone = clean_and_format_phone(phones[0])
        socials = re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[a-zA-Z0-9_.-]+', full_text)
        if socials:
            social = socials[0]
    return email, phone, social, metodo

if os.path.exists("wolflogo.png"):
    st.image("wolflogo.png", width=120)
st.title("🐺 WolfLogo - Scraper B2B")
st.caption("V3 Original - Corrigido + CNPJ + Socios + E-mail blindado")
api_key = get_api_key()

with st.sidebar:
    st.header("🔑 API")
    if api_key:
        st.success("API Key OK (Secrets)")
    else:
        st.warning("Configure SERPER_API_KEY")
        k = st.text_input("Cole sua chave Serper", type="password")
        if st.button("Salvar chave"):
            save_key(k)
            st.rerun()
    st.divider()
    st.header("🎯 Filtros")
    nicho = st.text_input("Nicho", value="restaurante")
    cidade = st.text_input("Cidade/UF", value="Campinas SP")
    limite = st.slider("Quantidade de leads", 1, 50, 10)
    buscar = st.button("🚀 Buscar Leads", type="primary", use_container_width=True)
    st.divider()
    st.caption(f"Creditos: {get_serper_credits(api_key)}")
    if SCRAPLING_AVAILABLE:
        st.caption("✅ Scrapling ativo")
    else:
        st.caption("⚠️ Scrapling nao instalado")

if not api_key:
    st.info("👈 Configure a SERPER_API_KEY no menu lateral ou nos Secrets")
    st.stop()

if buscar:
    with st.spinner(f"Buscando {nicho} em {cidade}..."):
        url = "https://google.serper.dev/search"
        payload = {"q": f"{nicho} em {cidade}", "gl": "br", "hl": "pt-br", "num": limite}
        headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
        res = requests.post(url, headers=headers, json=payload, timeout=15)
        if res.status_code == 200:
            organic = res.json().get("organic", [])
            leads = []
            progress = st.progress(0)
            for idx, item in enumerate(organic):
                title = item.get("title","")
                link = item.get("link","")
                progress.progress((idx+1)/len(organic))
                endereco = extract_address_from_item(item, organic) or cidade
                email, phone, social, metodo = scrape_details(link)
                cnpj, razao, socios = fetch_cnpj_and_partners(title, cidade)
                leads.append({"Empresa":title,"Endereco":endereco,"Telefone":phone,"Email":email,"Site":link,"Instagram/Facebook":social,"CNPJ":cnpj,"Razao Social":razao,"Socios":socios,"Metodo":metodo})
            progress.empty()
            df = pd.DataFrame(leads)
            st.success(f"✅ {len(df)} leads encontrados!")
            st.dataframe(df, use_container_width=True, height=500)
            from io import BytesIO
            output = BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Leads"
            header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=11)
            thin = Side(border_style="thin", color="D9D9D9")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            for col_num, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=1, column=col_num, value=col_name)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border
            for r_idx, row in enumerate(df.itertuples(index=False), 2):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    cell.border = border
            for col in range(1, len(df.columns)+1):
                ws.column_dimensions[get_column_letter(col)].width = 22
            wb.save(output)
            output.seek(0)
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📥 Baixar Excel Formatado", output.getvalue(), file_name=f"leads_{nicho}_{cidade}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col2:
                st.download_button("📥 Baixar CSV", df.to_csv(index=False).encode("utf-8"), file_name=f"leads_{nicho}_{cidade}.csv", use_container_width=True)
