import streamlit as st
import pandas as pd
import re
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill
from io import BytesIO

st.set_page_config(page_title="Gerador de Leads B2B", page_icon="🎯", layout="wide")

def get_api_key():
    if "SERPER_API_KEY" in st.secrets:
        return st.secrets["SERPER_API_KEY"]
    if "serper_api_key" in st.session_state and st.session_state["serper_api_key"]:
        return st.session_state["serper_api_key"]
    return ""

def save_key(k):
    st.session_state["serper_api_key"] = k

def clean_address(addr):
    return str(addr).strip() if addr else ""

# ===== CORRIGIDO: NAO REPETE ENDERECO =====
def extract_address_from_item(item):
    try:
        # 1. Tenta campo direto
        if isinstance(item, dict):
            for k in ["address", "endereco", "formattedAddress"]:
                if item.get(k):
                    return clean_address(item.get(k))
        # 2. Extrai SO DO SNIPPET DESSE ITEM (nao da lista toda)
        snippet = item.get("snippet", "") if isinstance(item, dict) else ""
        # Procura endereco completo: Rua X, Numero - Bairro - Cidade
        patterns = [
            r'(Rua|Av\.|Avenida|Alameda|Rodovia|Travessa|Praça)\s[^,]{3,}\s*,?\s*\d+[^,]*,\s*[^,]+,\s*[^,]+',
            r'(Rua|Av\.|Avenida)\s[^,]{3,},\s*\d+.*?(?:Campinas|São Paulo|SP)'
        ]
        for pat in patterns:
            m = re.search(pat, snippet, re.IGNORECASE)
            if m:
                return clean_address(m.group(0)[:120])
    except:
        pass
    return ""

def clean_phone(p):
    if not p:
        return ""
    d = re.sub(r"\D", "", str(p))
    return d[-11:] if len(d) >= 11 else d

def format_phone(d):
    d = clean_phone(d)
    if len(d) == 11:
        return f"({d[:2]}) {d[2:7]}-{d[7:]}"
    if len(d) == 10:
        return f"({d[:2]}) {d[2:6]}-{d[6:]}"
    return d

def decode_cf_email(e):
    try:
        r = int(e[:2], 16)
        return "".join([chr(int(e[i:i+2], 16) ^ r) for i in range(2, len(e), 2)])
    except:
        return ""

def fetch_cnpj(title, city):
    cnpj = razao = socios = ""
    api_key = get_api_key()
    if not api_key:
        return cnpj, razao, socios
    try:
        res = requests.post("https://google.serper.dev/search",
            headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
            json={"q": f"{title} {city} cnpj", "gl": "br", "hl": "pt-br", "num": 2},
            timeout=6)
        if res.status_code == 200:
            txt = " ".join([i.get("snippet","") for i in res.json().get("organic",[])])
            cnpjs = re.findall(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b", txt)
            if cnpjs:
                clean = re.sub(r"\D","", cnpjs[0])
                if len(clean)==14:
                    cnpj = f"{clean[:2]}.{clean[2:5]}.{clean[5:8]}/{clean[8:12]}-{clean[12:]}"
                    try:
                        r2 = requests.get(f"https://brasilapi.com.br/api/cnpj/v1/{clean}", timeout=5)
                        if r2.status_code==200:
                            j=r2.json()
                            razao=j.get("razao_social","")
                            socios=", ".join([s.get("nome_socio","") for s in j.get("qsa",[])[:2] if s.get("nome_socio")])
                    except:
                        pass
    except:
        pass
    return cnpj, razao, socios

def scrape_site(url, do_email=True, do_social=True):
    email = social = ""
    if not url or not url.startswith("http"):
        return email, social
    html = ""
    try:
        h={"User-Agent":"Mozilla/5.0"}
        r=requests.get(url, headers=h, timeout=6)
        if r.status_code==200:
            html=r.text[:20000]
        try:
            r2=requests.get(url.rstrip("/")+"/contato", headers=h, timeout=4)
            if r2.status_code==200:
                html+=" "+r2.text[:10000]
        except:
            pass
    except:
        pass
    if html:
        if do_email:
            cfs=re.findall(r'data-cfemail="([a-f0-9]+)"', html)
            for cf in cfs:
                d=decode_cf_email(cf)
                if d and "@" in d and "example" not in d:
                    email=d
                    break
            if not email:
                ms=re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', html)
                # FILTRO FORTE CONTRA LIXO
                blacklist = [".png",".jpg",".js","wix","sentry","example","seuemail","@sentry","@2x","@1x"]
                for m in ms:
                    if len(m) < 60 and "@" in m and not any(b in m.lower() for b in blacklist) and not m.endswith(".png"):
                        if not m[0].isdigit():
                            email=m
                            break
        if do_social:
            # FILTRO FORTE: ignora /tr, /sharer, /2008, so numero
            raw = re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[A-Za-z0-9_.-]+', html)
            for s in raw:
                low=s.lower()
                if any(x in low for x in ["/tr", "/sharer", "/plugins", "/login", "/2008", "/reel", "/p/", "/watch"]):
                    continue
                # extrai username
                username = s.split("/")[-1].split("?")[0]
                if len(username) < 4 or username.isdigit():
                    continue
                social=s
                break
    return email, social

# SIDEBAR
with st.sidebar:
    st.markdown("### 🔧 Configuração da API")
    st.caption("Chave Serper API")
    ak = st.text_input("key", type="password", label_visibility="collapsed", value=get_api_key())
    if ak and ak!= get_api_key():
        save_key(ak)
    st.divider()
    st.markdown("### ⚙️ Configurações da Busca")
    st.caption("Termo de Busca e Bairro")
    termo = st.text_input("termo", value="Loja de roupas Campinas SP", label_visibility="collapsed")
    st.caption("Quantidade de Resultados")
    qtd = st.slider("qtd", 1, 50, 10, label_visibility="collapsed")
    st.divider()
    st.markdown("### 🔍 Opções de Enriquecimento")
    b_email = st.checkbox("Buscar E-mails nos Sites", value=True)
    b_social = st.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)
    b_cnpj = st.checkbox("Buscar CNPJ & Sócios (BrasilAPI)", value=True)
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
        res = requests.post("https://google.serper.dev/search",
            headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
            json={"q": termo, "gl": "br", "hl": "pt-br", "num": qtd},
            timeout=15)
        if res.status_code!= 200:
            st.error(f"Erro {res.status_code}: {res.text}")
            st.stop()
        organic = res.json().get("organic", [])
        leads=[]
        prog=st.progress(0)
        for idx, item in enumerate(organic):
            prog.progress((idx+1)/len(organic))
            title=item.get("title","")
            link=item.get("link","")
            # CORRIGE CATEGORIA: pega primeira parte do termo
            categoria = termo.replace(" em "," ").split(" ")[0].strip()
            # ENDERECO UNICO POR EMPRESA
            endereco = extract_address_from_item(item)
            if not endereco:
                endereco = "" # deixa vazio ao inves de repetir
            # TELEFONE DO SNIPPET
            snippet=item.get("snippet","")
            tel_match=re.search(r'\(?\d{2}\)?\s*9?\d{4}-?\d{4}', snippet)
            telefone=format_phone(tel_match.group(0)) if tel_match else ""
            whatsapp=f"https://wa.me/55{clean_phone(telefone)}" if telefone else ""
            email, social = scrape_site(link, b_email, b_social)
            cnpj, razao, socios = ("","","")
            if b_cnpj:
                cnpj, razao, socios = fetch_cnpj(title, termo)
            leads.append({
                "Categoria": categoria,
                "Empresa": title,
                "Endereço": endereco,
                "Telefone": telefone,
                "Whatsapp": whatsapp,
                "Email": email,
                "Redes Sociais": social,
                "CNPJ": cnpj,
                "Razão Social": razao,
                "Sócios": socios,
                "Site": link
            })
        prog.empty()
        if leads:
            df=pd.DataFrame(leads)
            # ORDEM CORRETA DAS COLUNAS igual print original
            ordem = ["Categoria","Empresa","Endereço","Telefone","Whatsapp","Email","Redes Sociais","CNPJ","Razão Social","Sócios","Site"]
            df=df[ordem]
            st.markdown(f'<div style="background:#1c4d2a;padding:12px;border-radius:8px;color:#a3e6b5">✅ Sucesso! {len(df)} empresas extraídas com sucesso.</div>', unsafe_allow_html=True)
            st.markdown("### 📋 Prévia dos Resultados Reais")
            st.dataframe(df, use_container_width=True, height=600)
            output=BytesIO()
            wb=openpyxl.Workbook()
            ws=wb.active
            ws.title="Leads"
            fill=PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
            font=Font(color="FFFFFF", bold=True)
            for c, name in enumerate(ordem,1):
                cell=ws.cell(row=1,column=c,value=name)
                cell.fill=fill
                cell.font=font
            for r, row in enumerate(df.itertuples(index=False),2):
                for c, v in enumerate(row,1):
                    ws.cell(row=r,column=c,value=v)
            wb.save(output)
            output.seek(0)
            st.download_button("📥 Baixar Excel Formatado", output.getvalue(), f"leads_{termo}.xlsx", use_container_width=True)
            st.download_button("📥 Baixar CSV", df.to_csv(index=False).encode("utf-8"), f"leads_{termo}.csv", use_container_width=True)
