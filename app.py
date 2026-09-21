import streamlit as st
import pandas as pd
import re
import requests
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gerador de Leads B2B", page_icon="🎯", layout="wide")

DEFAULT_KEY = ""

def get_api_key():
    if "SERPER_API_KEY" in st.secrets:
        return st.secrets["SERPER_API_KEY"]
    if "serper_api_key" in st.session_state and st.session_state["serper_api_key"]:
        return st.session_state["serper_api_key"]
    return ""

def save_key(k):
    st.session_state["serper_api_key"] = k

def clean_address(addr):
    if not addr:
        return ""
    return str(addr).strip()

def extract_address_from_item(item, organic_list=None):
    addr = ""
    try:
        if isinstance(item, dict):
            addr = item.get("address", "") or item.get("endereco", "") or item.get("formattedAddress", "")
        if addr:
            return clean_address(addr)
        snippets = organic_list if organic_list else [item] if isinstance(item, dict) else []
        for og in snippets:
            snippet = og.get("snippet", "") if isinstance(og, dict) else ""
            try:
                match = re.search(r'(Rua|R\.|Avenida|Av\.|Praca|Alameda|Rodovia|Travessa)[^,\n]{5,},[^,\n]{3,}', snippet, re.IGNORECASE)
                if match:
                    addr = match.group(0)
                    break
            except Exception:
                continue
    except Exception:
        pass
    return clean_address(addr)

def clean_and_format_phone(phone):
    if not phone:
        return ""
    digits = re.sub(r"\D", "", str(phone))
    if len(digits) < 10:
        return ""
    if len(digits) >= 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:11]}"
    return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"

def decode_cf_email(e):
    try:
        r = int(e[:2], 16)
        return "".join([chr(int(e[i:i+2], 16) ^ r) for i in range(2, len(e), 2)])
    except:
        return ""

def fetch_cnpj_and_partners(company_name, city=""):
    cnpj = ""
    razao = ""
    socios = ""
    api_key = get_api_key()
    if not api_key:
        return cnpj, razao, socios
    try:
        q = f"{company_name} {city} cnpj"
        url = "https://google.serper.dev/search"
        payload = {"q": q, "gl": "br", "hl": "pt-br", "num": 3}
        headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
        res = requests.post(url, headers=headers, json=payload, timeout=6)
        if res.status_code == 200:
            txt = ""
            for it in res.json().get("organic", []):
                txt += " " + it.get("snippet","") + " " + it.get("title","")
            cnpjs = re.findall(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b", txt)
            if cnpjs:
                clean = re.sub(r"\D", "", cnpjs[0])
                if len(clean)==14:
                    cnpj = f"{clean[:2]}.{clean[2:5]}.{clean[5:8]}/{clean[8:12]}-{clean[12:]}"
                    try:
                        r2 = requests.get(f"https://brasilapi.com.br/api/cnpj/v1/{clean}", timeout=6)
                        if r2.status_code==200:
                            j=r2.json()
                            razao=j.get("razao_social","")
                            socios = ", ".join([s.get("nome_socio","") for s in j.get("qsa",[]) if s.get("nome_socio")])
                    except:
                        pass
    except:
        pass
    return cnpj, razao, socios

def scrape_site(url, do_email=True, do_social=True):
    email=""
    social=""
    if not url or not url.startswith("http"):
        return email, social
    full=""
    try:
        headers={"User-Agent":"Mozilla/5.0"}
        r=requests.get(url, headers=headers, timeout=6)
        if r.status_code==200:
            full=r.text
        try:
            r2=requests.get(url.rstrip("/")+"/contato", headers=headers, timeout=5)
            if r2.status_code==200:
                full+=" "+r2.text
        except:
            pass
    except:
        pass
    if full:
        if do_email:
            cfs=re.findall(r'data-cfemail="([a-f0-9]+)"', full)
            for cf in cfs:
                dec=decode_cf_email(cf)
                if dec and "@" in dec:
                    email=dec
                    break
            if not email:
                ms=re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', full)
                if ms:
                    good=[m for m in ms if not any(x in m.lower() for x in [".png",".jpg","wix","sentry","example"])]
                    if good:
                        email=good[0]
        if do_social:
            ss=re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[a-zA-Z0-9_.-]+', full)
            if ss:
                social=ss[0]
    return email, social

with st.sidebar:
    st.markdown("### 🔧 Configuração da API")
    st.caption("Chave Serper API")
    api_key_input = st.text_input("Chave Serper API", type="password", label_visibility="collapsed", value=get_api_key())
    if api_key_input!= get_api_key() and api_key_input:
        save_key(api_key_input)
    st.divider()
    st.markdown("### ⚙️ Configurações da Busca")
    st.caption("Termo de Busca e Bairro")
    termo = st.text_input("Termo", value="Loja de roupas Campinas SP Nuvemsh", label_visibility="collapsed")
    st.caption("Quantidade de Resultados")
    quantidade = st.number_input("Qtd", min_value=1, max_value=50, value=10, label_visibility="collapsed")
    st.divider()
    st.markdown("### 🔍 Opções de Enriquecimento")
    buscar_email = st.checkbox("Buscar E-mails nos Sites", value=True)
    buscar_social = st.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)
    buscar_cnpj = st.checkbox("Buscar CNPJ & Sócios (BrasilAPI)", value=True)
    st.markdown("<br>", unsafe_allow_html=True)
    iniciar = st.button("🚀 Iniciar Extração de Leads", type="primary", use_container_width=True)

st.markdown("## 🎯 Gerador de Leads B2B - Prospecção Avançada")
st.caption("Extração de dados de estabelecimentos, empresas e profissionais.")
api_key = get_api_key()
if not api_key:
    st.warning("Configure a chave Serper na lateral")
    st.stop()

if iniciar:
    with st.spinner(f"Extraindo {termo}..."):
        url = "https://google.serper.dev/search"
        payload = {"q": termo, "gl": "br", "hl": "pt-br", "num": quantidade}
        headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
        res = requests.post(url, headers=headers, json=payload, timeout=15)
        if res.status_code == 200:
            organic = res.json().get("organic", [])
            leads = []
            for item in organic:
                title = item.get("title","")
                link = item.get("link","")
                categoria = termo.split(" em ")[0] if " em " in termo else "Loja de moda feminina"
                endereco = extract_address_from_item(item, organic) or termo
                snippet = item.get("snippet","")
                m_phone = re.search(r'\(?\d{2}\)?\s?9?\d{4}-?\d{4}', snippet)
                telefone = m_phone.group(0) if m_phone else ""
                whatsapp = ""
                if telefone:
                    digits = re.sub(r"\D","", telefone)
                    whatsapp = f"https://wa.me/55{digits}" if len(digits)>=10 else ""
                email, social = "", ""
                if buscar_email or buscar_social:
                    email, social = scrape_site(link, buscar_email, buscar_social)
                cnpj, razao, socios = ("","","")
                if buscar_cnpj:
                    cnpj, razao, socios = fetch_cnpj_and_partners(title, termo)
                leads.append({"Categoria":categoria,"Endereco":endereco,"Telefone":telefone,"Whatsapp":whatsapp,"Email":email,"Redes Sociais":social,"CNPJ":cnpj,"Razao":razao,"Socios":socios,"Empresa":title,"Site":link})
            if leads:
                df = pd.DataFrame(leads)
                st.markdown(f'<div style="background-color:#1c4d2a; padding:12px; border-radius:8px; color:#a3e6b5;">✅ Sucesso! {len(df)} empresas extraídas com sucesso.</div>', unsafe_allow_html=True)
                tab1, tab2 = st.tabs(["📋 Tabela de Leads", "💬 Gerador de Script de Vendas"])
                with tab1:
                    st.markdown("### 📋 Prévia dos Resultados Reais")
                    st.dataframe(df, use_container_width=True, height=500)
                    from io import BytesIO
                    output = BytesIO()
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Leads"
                    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)
                    for col_num, col_name in enumerate(df.columns, 1):
                        c = ws.cell(row=1, column=col_num, value=col_name)
                        c.fill = header_fill
                        c.font = header_font
                    for r_idx, row in enumerate(df.itertuples(index=False), 2):
                        for c_idx, val in enumerate(row, 1):
                            ws.cell(row=r_idx, column=c_idx, value=val)
                    wb.save(output)
                    output.seek(0)
                    st.download_button("📥 Baixar Excel", output.getvalue(), file_name=f"leads_{termo}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    st.download_button("📥 Baixar CSV", df.to_csv(index=False).encode("utf-8"), file_name=f"leads_{termo}.csv", use_container_width=True)
