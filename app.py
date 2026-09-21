import streamlit as st
import requests
import re
import pandas as pd

DEFAULT_KEY = ""

def get_api_key():
    if "SERPER_API_KEY" in st.secrets:
        return st.secrets["SERPER_API_KEY"]
    if "serper_api_key" in st.session_state and st.session_state["serper_api_key"]:
        return st.session_state["serper_api_key"]
    return ""

try:
    from scrapling import Fetcher
    SCRAPLING_AVAILABLE = True
except:
    SCRAPLING_AVAILABLE = False

def decode_cf_email(encoded):
    try:
        r = int(encoded[:2], 16)
        return ''.join([chr(int(encoded[i:i+2], 16) ^ r) for i in range(2, len(encoded), 2)])
    except:
        return ""

def clean_address(addr):
    if not addr:
        return ""
    if isinstance(addr, list):
        addr = ", ".join([str(x) for x in addr if x])
    return str(addr).strip()

def get_full_address(company_name, city=""):
    addr = ""
    api_key = get_api_key()
    if not api_key:
        return addr
    try:
        url = "https://google.serper.dev/search"
        query = f"{company_name} {city} endereco"
        payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
        headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
        res = requests.post(url, headers=headers, json=payload, timeout=5)
        if res.status_code == 200:
            organic = res.json().get("organic", [])
            for og in organic:
                snippet = og.get("snippet", "") if isinstance(og, dict) else ""
                m = re.search(r'(Rua|Av\.|Avenida|Praca|Alameda|Rodovia)[^,]+,[^,]+', snippet, re.I)
                if m:
                    addr = m.group(0)
                    break
    except:
        pass
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
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    try:
        res = requests.post(url, headers=headers, json=payload, timeout=5)
        if res.status_code == 200:
            text_block = ""
            for item in res.json().get("organic", []):
                text_block += " " + item.get("snippet", "") + " " + item.get("title", "")
            cnpjs = re.findall(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b', text_block)
            if cnpjs:
                cnpj_clean = re.sub(r'\D', '', cnpjs[0])
    except:
        pass
    if cnpj_clean and len(cnpj_clean) == 14:
        try:
            r = requests.get(f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_clean}", timeout=5)
            if r.status_code == 200:
                j = r.json()
                razao_social = j.get("razao_social", "")
                for socio in j.get("qsa", []):
                    nome = socio.get("nome_socio", "")
                    if nome:
                        socios_names.append(nome)
        except:
            pass
    cnpj_formatted = f"{cnpj_clean[:2]}.{cnpj_clean[2:5]}.{cnpj_clean[5:8]}/{cnpj_clean[8:12]}-{cnpj_clean[12:]}" if len(cnpj_clean)==14 else ""
    return cnpj_formatted, razao_social, ", ".join(socios_names)

def scrape_website_details(website_url):
    email = ""
    social = ""
    metodo = ""
    if not website_url or not str(website_url).startswith("http"):
        return email, social, metodo
    urls = [website_url, website_url.rstrip("/")+"/contato"]
    full_text = ""
    for u in urls:
        page_text = ""
        if SCRAPLING_AVAILABLE:
            try:
                page = Fetcher.get(u, impersonate='chrome', timeout=6)
                if page.status == 200:
                    page_text = page.html_content
                    metodo = "Scrapling"
            except:
                page_text = ""
        if not page_text:
            try:
                r = requests.get(u, headers={'User-Agent':'Mozilla/5.0'}, timeout=5)
                if r.status_code == 200:
                    page_text = r.text
                    if not metodo:
                        metodo = "Requests"
            except:
                page_text = ""
        full_text += " " + page_text
    if full_text:
        cfs = re.findall(r'data-cfemail="([a-f0-9]+)"', full_text)
        for cf in cfs:
            d = decode_cf_email(cf)
            if d and "@" in d:
                email = d
                break
        if not email:
            ms = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', full_text)
            if ms:
                email = ms[0]
        ss = re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[a-zA-Z0-9_.-]+', full_text)
        if ss:
            social = ss[0]
    return email, social, metodo

st.set_page_config(page_title="Scraper B2B - WolfLogo", page_icon="🐺", layout="wide")
st.title("🐺 WolfLogo - Scraper B2B v4 Corrigido")
api_key = get_api_key()
if not api_key:
    st.warning("Configure SERPER_API_KEY nos Secrets")
    with st.sidebar:
        k = st.text_input("SERPER_API_KEY", type="password")
        if st.button("Salvar"):
            st.session_state['serper_api_key']=k
            st.rerun()
else:
    st.sidebar.success("API Key ok")
    st.session_state['serper_api_key']=api_key

if api_key:
    nicho = st.text_input("Nicho", "restaurante")
    cidade = st.text_input("Cidade", "Campinas SP")
    limite = st.slider("Leads", 1, 20, 5)
    if st.button("Buscar Leads"):
        with st.spinner("Buscando..."):
            url = "https://google.serper.dev/search"
            payload = {"q": f"{nicho} em {cidade}", "gl": "br", "hl": "pt-br", "num": limite}
            headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
            res = requests.post(url, headers=headers, json=payload, timeout=10)
            if res.status_code == 200:
                leads = []
                for item in res.json().get("organic", []):
                    title = item.get("title","")
                    link = item.get("link","")
                    email, social, metodo = scrape_website_details(link)
                    endereco = get_full_address(title, cidade)
                    cnpj, razao, socios = fetch_cnpj_and_partners(title, cidade)
                    leads.append({"Empresa":title,"Site":link,"Endereco":endereco,"Email":email,"Social":social,"CNPJ":cnpj,"Razao":razao,"Socios":socios,"Metodo":metodo})
                if leads:
                    df = pd.DataFrame(leads)
                    st.success(f"{len(leads)} leads")
                    st.dataframe(df, use_container_width=True)
                    st.download_button("Baixar CSV", df.to_csv(index=False).encode('utf-8'), "leads.csv", "text/csv")
