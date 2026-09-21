import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
import requests
from urllib.parse import quote_plus, urlparse
import os

# ==============================================================================
# SCRAPER B2B + GESTÃO DINÂMICA DE CHAVE SERPER + CRÉDITOS EM TEMPO REAL
# ==============================================================================

# --- Integração com Scrapling (impersonation de TLS/browser real) ---
# Requisito: pip install "scrapling[fetchers]"
# Não precisa rodar "scrapling install" (que baixa browsers) para o uso abaixo,
# pois o Fetcher.get() é puro HTTP com impersonation, sem abrir navegador.
try:
    from scrapling.fetchers import Fetcher
    SCRAPLING_AVAILABLE = True
except ImportError:
    SCRAPLING_AVAILABLE = False

KEY_FILE = ".serper_key"
HUNTER_KEY_FILE = ".hunter_key"

def get_default_key(secret_name):
    """
    Busca uma chave padrão a partir do st.secrets (configurado no painel do
    Streamlit Cloud, em Settings > Secrets) — NUNCA hardcoded no código-fonte,
    já que este repositório é público no GitHub. Se não houver secret
    configurado (ex: rodando local sem .streamlit/secrets.toml), retorna
    string vazia e o usuário precisa colar a chave manualmente na sidebar.
    """
    try:
        return st.secrets[secret_name]
    except Exception:
        return ""

# Funções genéricas para carregar/salvar qualquer chave localmente
def load_saved_key(key_file, secret_name):
    if os.path.exists(key_file):
        try:
            with open(key_file, "r") as f:
                key = f.read().strip()
                if key:
                    return key
        except Exception:
            pass
    return get_default_key(secret_name)

def save_key(key_file, key_str):
    try:
        with open(key_file, "w") as f:
            f.write(key_str.strip())
    except Exception as e:
        st.error(f"Erro ao salvar a chave: {e}")

# Inicializa o estado das chaves API
if 'serper_api_key' not in st.session_state:
    st.session_state['serper_api_key'] = load_saved_key(KEY_FILE, "SERPER_API_KEY")
if 'hunter_api_key' not in st.session_state:
    st.session_state['hunter_api_key'] = load_saved_key(HUNTER_KEY_FILE, "HUNTER_API_KEY")

st.set_page_config(
    page_title="Gerador de Leads B2B - Prospecção Avançada",
    page_icon="🎯",
    layout="wide"
)

# Estilização para Dark Mode
st.markdown("""
<style>
    .main-header { font-size: 2.2rem; color: #60A5FA; font-weight: 700; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1.1rem; color: #D1D5DB; margin-bottom: 1.5rem; }
    .stButton>button { background-color: #2563EB; color: white; border-radius: 6px; padding: 0.5rem 1.5rem; font-weight: 600; border: none; }
    .stButton>button:hover { background-color: #1D4ED8; color: white; }
    .footer-box {
        background-color: #1E293B;
        border-top: 2px solid #3B82F6;
        padding: 15px;
        border-radius: 8px;
        margin-top: 40px;
        color: #F3F4F6;
    }
</style>
""", unsafe_allow_html=True)

# Função para checar créditos restantes na API Serper (Usando POST)
# Nota: removida a função get_serper_credits() que existia aqui — o endpoint
# "https://google.serper.dev/credits" nunca foi um endpoint real e documentado
# da API (confirmado: a lista oficial de endpoints do Serper cobre search,
# images, videos, places, maps, reviews, news, shopping, lens, patents,
# autocomplete e scrape — sem nada de conta/créditos). Por isso a consulta
# sempre falhava, com qualquer chave. Ver rodapé do app para o substituto:
# contagem real de créditos gastos por extração, calculada pelo próprio app.

# Função para checar buscas restantes na API Hunter.io
def get_hunter_credits(api_key):
    if not api_key:
        return None
    try:
        res = requests.get(f"https://api.hunter.io/v2/account?api_key={api_key}", timeout=4)
        if res.status_code == 200:
            data = res.json()
            searches = data.get("data", {}).get("requests", {}).get("searches", {})
            available = searches.get("available")
            used = searches.get("used")
            if available is not None:
                return available, used
    except Exception:
        pass
    return None, None

# --- BARRA LATERAL ---
with st.sidebar:
    if os.path.exists("wolflogo.png"):
        st.image("wolflogo.png", use_container_width=True)
    
    st.header("🔑 Configuração da API")
    
    # Campo para alterar a chave Serper
    current_key = st.session_state['serper_api_key']
    new_key = st.text_input("Chave Serper API", value=current_key, type="password", help="Cole sua chave ativa aqui")
    
    if new_key != current_key:
        st.session_state['serper_api_key'] = new_key
        save_key(KEY_FILE, new_key)
        st.success("Nova Chave Salva!")
        st.rerun()

    # Campo para alterar a chave Hunter.io
    current_hunter_key = st.session_state['hunter_api_key']
    new_hunter_key = st.text_input("Chave Hunter.io API", value=current_hunter_key, type="password", help="Opcional — usada só quando o site do lead não expõe e-mail direto (plano free do Hunter tem cota mensal bem limitada)")

    if new_hunter_key != current_hunter_key:
        st.session_state['hunter_api_key'] = new_hunter_key
        save_key(HUNTER_KEY_FILE, new_hunter_key)
        st.success("Nova Chave Hunter Salva!")
        st.rerun()

    st.markdown("---")
    st.header("⚙️ Configurações da Busca")
    termo_busca = st.text_input("Termo de Busca e Bairro", value="Pizzarias Campinas SP Bairro Castelo")
    qtd_resultados = st.number_input("Quantidade de Resultados", min_value=1, max_value=20, value=10, step=1)

    st.markdown("---")
    st.header("🔍 Opções de Enriquecimento")
    enriquecer_emails = st.checkbox("Buscar E-mails nos Sites", value=True)
    enriquecer_redes = st.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)
    buscar_socios = st.checkbox("Buscar CNPJ & Sócios (BrasilAPI)", value=True)
    usar_hunter = st.checkbox(
        "Completar E-mail via Hunter.io (só quando faltar)",
        value=False,
        help="Consome cota do Hunter (25 buscas/mês no plano free) — só é chamado pros leads que ainda estão sem e-mail depois do site e da busca Serper. Desligado por padrão pra você não gastar cota sem querer."
    )

    # --- Estimativa de créditos Serper antes de rodar ---
    # Mínimo: só a busca Places (1 crédito, cobre o lote inteiro).
    # Máximo: Places + 1 busca fallback por lead (se faltar contato) +
    # 1 busca de CNPJ por lead (se a opção estiver marcada).
    est_min = 1
    est_max = 1 + qtd_resultados + (qtd_resultados if buscar_socios else 0)
    st.caption(f"💰 Estimativa de custo Serper: **{est_min} a {est_max} créditos** nesta extração (varia conforme quantos leads já vêm com contato direto do Google Places).")
    
    st.markdown("<br>", unsafe_allow_html=True)
    btn_extrair = st.button("🚀 Iniciar Extração de Leads", use_container_width=True)

# --- CABEÇALHO PRINCIPAL ---
st.markdown('<div class="main-header">🎯 Gerador de Leads B2B - Prospecção Avançada</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Extração de dados de estabelecimentos, empresas e profissionais.</div>', unsafe_allow_html=True)

def clean_and_format_phone(phone_str):
    if not phone_str or pd.isna(phone_str):
        return "", ""
    digits = re.sub(r'\D', '', str(phone_str))
    
    if len(digits) > 11 and digits.startswith("55"):
        digits = digits[2:]
        
    if len(digits) < 8:
        return str(phone_str), ""
    
    formatted_phone = str(phone_str)
    whatsapp = f"https://wa.me/55{digits}" if len(digits) in [10, 11] else ""
    return formatted_phone, whatsapp

def extract_address_from_item(item, company_name="", credit_counter=None):
    addr = item.get("address") or item.get("formattedAddress") or item.get("vicinity") or item.get("street") or ""
    
    if isinstance(addr, dict):
        street = addr.get("street", "")
        city = addr.get("city", "")
        state = addr.get("state", "")
        addr = f"{street}, {city} - {state}".strip(", -")
    elif isinstance(addr, list):
        addr = ", ".join([str(x) for x in addr])

    if not addr and company_name:
        try:
            url = "https://google.serper.dev/search"
            payload = {"q": f"{company_name} endereco localizacao", "gl": "br", "hl": "pt-br", "num": 2}
            headers = {'X-API-KEY': st.session_state['serper_api_key'], 'Content-Type': 'application/json'}
            if credit_counter is not None:
                credit_counter[0] += 1
            res = requests.post(url, headers=headers, json=payload, timeout=3)
            if res.status_code == 200:
                data = res.json()
                for og in data.get("organic", []):
                    snippet = og.get("snippet", "")
                    match = re.search(r'(Rua|R\.|Avenida|Av\.|Praça|Alameda|Rodovia)[^,\.]+,[^\.]+', snippet, re.IGNORECASE)
                    if match:
                        addr = match.group(0)
                        break
        except Exception:
            pass

    return str(addr).strip()

def fetch_cnpj_and_partners(company_name, city_or_address="", credit_counter=None):
    cnpj_clean = ""
    razao_social = ""
    socios_names = []
    
    location_hint = city_or_address.split("-")[0].strip() if city_or_address else ""
    query = f"{company_name} {location_hint} cnpj brasilapi"
    url = "https://google.serper.dev/search"
    payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
    headers = {'X-API-KEY': st.session_state['serper_api_key'], 'Content-Type': 'application/json'}
    
    try:
        if credit_counter is not None:
            credit_counter[0] += 1
        res = requests.post(url, headers=headers, json=payload, timeout=4)
        if res.status_code == 200:
            data = res.json()
            text_block = ""
            for item in data.get("organic", []):
                text_block += " " + item.get("snippet", "") + " " + item.get("title", "")
            
            cnpjs = re.findall(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b', text_block)
            if cnpjs:
                cnpj_clean = re.sub(r'\D', '', cnpjs[0])
    except Exception:
        pass

    if cnpj_clean:
        try:
            api_url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_clean}"
            r = requests.get(api_url, timeout=4)
            if r.status_code == 200:
                cnpj_data = r.json()
                razao_social = cnpj_data.get("razao_social", "")
                qsa = cnpj_data.get("qsa", [])
                for socio in qsa:
                    nome = socio.get("nome_socio", "")
                    qualificacao = socio.get("qualificacao_socio", "")
                    if nome:
                        socios_names.append(f"{nome} ({qualificacao})" if qualificacao else nome)
        except Exception:
            pass

    cnpj_formatted = f"{cnpj_clean[:2]}.{cnpj_clean[2:5]}.{cnpj_clean[5:8]}/{cnpj_clean[8:12]}-{cnpj_clean[12:]}" if len(cnpj_clean) == 14 else ""
    socios_str = ", ".join(socios_names) if socios_names else ""
    return cnpj_formatted, razao_social, socios_str

def fallback_search_phone_email(company_name, address, api_key, credit_counter=None):
    phone, email, social = "", "", ""
    query = f"{company_name} {address} telefone contato email"
    url = "https://google.serper.dev/search"
    payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    
    try:
        if credit_counter is not None:
            credit_counter[0] += 1
        res = requests.post(url, headers=headers, json=payload, timeout=5)
        if res.status_code == 200:
            data = res.json()
            text_block = ""
            
            for item in data.get("organic", []):
                text_block += " " + item.get("title", "") + " " + item.get("snippet", "")
                link = item.get("link", "")
                if ("instagram.com" in link or "facebook.com" in link) and not social:
                    social = link

            phone_matches = re.findall(r'(?:\(?\d{2}\)?\s*)?(?:9?\d{4}[-\s]?\d{4})', text_block)
            if phone_matches:
                for match in phone_matches:
                    clean_m = re.sub(r'\D', '', match)
                    if len(clean_m) in [10, 11]:
                        phone = match
                        break

            email_matches = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text_block)
            if email_matches:
                valid_e = [e for e in email_matches if not e.endswith(('.png', '.jpg', '.webp', '.js', '.css'))]
                if valid_e:
                    email = valid_e[0]
    except Exception:
        pass

    return phone, email, social

def _fetch_page_html(url, timeout=6):
    """
    Busca o HTML de UMA página, tentando primeiro Scrapling (impersonation
    de navegador real) e caindo pro requests puro se falhar/indisponível.
    Retorna (texto_html, metodo) — metodo é "Scrapling", "Requests" ou "".
    """
    text, metodo = "", ""

    if SCRAPLING_AVAILABLE:
        try:
            page = Fetcher.get(url, impersonate='chrome', stealthy_headers=True, timeout=timeout)
            if page.status == 200:
                text, metodo = page.html_content, "Scrapling"
        except Exception:
            pass

    if not text:
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            response = requests.get(url, headers=headers, timeout=min(timeout, 5))
            if response.status_code == 200:
                text, metodo = response.text, "Requests"
        except Exception:
            pass

    return text, metodo

def _extract_email_social(text):
    """Aplica os regex de e-mail/rede social num HTML já baixado."""
    email, social = "", ""
    if not text:
        return email, social
    emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
    if emails:
        valid_emails = [e for e in emails if not e.endswith(('.png', '.jpg', '.webp', '.js', '.css', '.svg'))]
        if valid_emails:
            email = valid_emails[0]
    socials = re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[a-zA-Z0-9_.-]+', text)
    if socials:
        social = socials[0]
    return email, social

# Páginas comuns onde PMEs brasileiras costumam publicar e-mail/telefone
# quando não está no rodapé da home. Testadas em ordem, para na primeira
# que resolver o e-mail (evita bater em N páginas à toa).
CAMINHOS_CONTATO_COMUNS = ["/contato", "/contato-nos", "/fale-conosco", "/contact", "/contact-us"]

def scrape_website_details(website_url):
    """
    Extrai e-mail e rede social do site do lead.

    Tenta a home primeiro. Se a home responder mas não tiver e-mail (comum:
    o e-mail costuma morar numa página /contato, não no rodapé), tenta até
    3 caminhos comuns de página de contato antes de desistir. Cada tentativa
    usa Scrapling com fallback pra requests, então bloqueio por Cloudflare/WAF
    é contornado em qualquer uma das páginas testadas, não só na home.

    Retorna (email, social, metodo) - "metodo" indica qual caminho
    efetivamente conseguiu os dados, incluindo se veio da home ou de uma
    página de contato secundária (ex: "Scrapling", "Scrapling + /contato").
    """
    email, social = "", ""
    metodo = ""
    if not website_url or not str(website_url).startswith("http"):
        return email, social, metodo

    text, metodo = _fetch_page_html(website_url)
    if not text:
        return email, social, ""  # nem a home abriu — nenhum bypass funcionou

    email, social = _extract_email_social(text)

    # Home abriu, mas sem e-mail visível → tenta páginas de contato comuns
    if not email:
        base = website_url.rstrip('/')
        for caminho in CAMINHOS_CONTATO_COMUNS:
            sub_text, sub_metodo = _fetch_page_html(base + caminho, timeout=5)
            if sub_text:
                sub_email, sub_social = _extract_email_social(sub_text)
                if sub_email:
                    email = sub_email
                    metodo = f"{metodo} + {caminho}"
                if not social and sub_social:
                    social = sub_social
                if email:
                    break

    return email, social, metodo

def hunter_domain_search(website_url, api_key):
    """
    Terceira camada de enriquecimento (só chamada quando site + fallback
    Serper não acharam e-mail). Usa o endpoint Domain Search do Hunter.io,
    que busca e-mails associados ao domínio via padrões conhecidos e
    indexação própria — não depende do HTML da página em si, então
    funciona mesmo quando o site nunca publicou e-mail em texto plano.

    Retorna o e-mail de maior confiança encontrado, ou "" se não achar
    nada ou a chave/cota estiver indisponível.
    """
    if not api_key or not website_url:
        return ""

    try:
        domain = urlparse(website_url).netloc.replace("www.", "")
        if not domain:
            return ""

        url = "https://api.hunter.io/v2/domain-search"
        params = {"domain": domain, "api_key": api_key, "limit": 3}
        res = requests.get(url, params=params, timeout=6)

        if res.status_code == 200:
            data = res.json()
            emails = data.get("data", {}).get("emails", [])
            if emails:
                # Prioriza e-mails de tipo "generic" (contato@, info@) sobre
                # pessoais, e ordena pela confiança que o Hunter atribui
                emails_sorted = sorted(emails, key=lambda e: e.get("confidence", 0), reverse=True)
                return emails_sorted[0].get("value", "")
    except Exception:
        pass

    return ""

def create_excel_report(df, filename="leads_extraidos.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads B2B"
    ws.views.sheetView[0].showGridLines = True
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10, color="000000")
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    
    thin_border = Side(style='thin', color='D9D9D9')
    cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    columns = list(df.columns)
    ws.append(columns)
    
    for col_num in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = cell_border
    ws.row_dimensions[1].height = 26
    
    for r_idx, row in df.iterrows():
        row_num = r_idx + 2
        fill = zebra_fill if r_idx % 2 == 1 else white_fill
        
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=row_num, column=c_idx)
            cell.fill = fill
            cell.border = cell_border
            cell.font = data_font
            
            col_name = columns[c_idx - 1]
            val_str = "" if pd.isna(val) or val is None else str(val)
            
            if col_name == "Link Google Maps" and val_str and val_str.startswith("http"):
                cell.value = f'=HYPERLINK("{val_str}", "Ver no Google Maps")'
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Whatsapp" and val_str and val_str.startswith("http"):
                cell.value = f'=HYPERLINK("{val_str}", "Abrir WhatsApp")'
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Redes Sociais" and val_str and val_str.startswith("http"):
                cell.value = f'=HYPERLINK("{val_str}", "Acessar Perfil")'
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name in ["Status", "Progressão", "Tem Website", "Nota Google", "Total Avaliações", "CNPJ", "Fonte do Contato"]:
                cell.value = val_str
                cell.alignment = align_center
            else:
                cell.value = val_str
                cell.alignment = align_left
                
        ws.row_dimensions[row_num].height = 20
        
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
        
    max_row = len(df) + 50

    if "Status" in columns:
        status_col_idx = columns.index("Status") + 1
        col_letter = get_column_letter(status_col_idx)
        dv_status = DataValidation(type="list", formula1='"A Fazer, Em Andamento, Concluído, Cancelado"', allow_blank=True)
        ws.add_data_validation(dv_status)
        dv_status.add(f"{col_letter}2:{col_letter}{max_row}")

    if "Progressão" in columns:
        progress_col_idx = columns.index("Progressão") + 1
        col_letter = get_column_letter(progress_col_idx)
        dv_progress = DataValidation(type="list", formula1='"1º Contato, Em Negociação, Proposta Enviada, Fechado, Perdido"', allow_blank=True)
        ws.add_data_validation(dv_progress)
        dv_progress.add(f"{col_letter}2:{col_letter}{max_row}")

    ws.freeze_panes = 'A2'
    wb.save(filename)
    return filename

def run_places_extraction(query, api_key, max_results=10, email_opt=True, redes_opt=True, socios_opt=True, hunter_opt=False, hunter_key=""):
    extracted_data = []
    credit_counter = [1]  # já conta a própria busca Places como 1 crédito
    
    url = "https://google.serper.dev/places"
    payload = {"q": query, "gl": "br", "hl": "pt-br"}
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    
    try:
        # A busca Places pode demorar mais que o normal em consultas muito
        # específicas geograficamente (bairro + categoria) sem que isso
        # signifique "poucos resultados" — é só lentidão pontual do lado da
        # Serper. Por isso, em vez de falhar direto no primeiro timeout,
        # tenta de novo uma vez com um timeout maior antes de desistir.
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=10)
        except requests.exceptions.Timeout:
            st.warning("⏳ A busca demorou mais que o normal, tentando de novo com mais tempo...")
            credit_counter[0] += 1  # a tentativa anterior pode ter consumido crédito mesmo sem resposta
            response = requests.post(url, headers=headers, json=payload, timeout=25)
        
        if response.status_code == 402 or "Out of credits" in response.text:
            st.error("❌ Os créditos da chave Serper atual ACABARAM! Altere a chave na barra lateral ou crie uma nova conta no Serper.")
            return pd.DataFrame(), credit_counter[0]
            
        data = response.json()
        
        places = data.get("places", [])
        for item in places[:max_results]:
            company_name = item.get("title", "")
            
            address = extract_address_from_item(item, company_name, credit_counter)
            
            phone_raw = item.get("phoneNumber") or item.get("phone") or ""
            category = item.get("category", "Comércio Local / Empresa")
            rating = item.get("rating", "")
            rating_count = item.get("ratingCount", "")
            website_url = item.get("website", "")
            
            latitude = item.get("latitude", "")
            longitude = item.get("longitude", "")
            if latitude and longitude:
                gmaps_link = f"https://www.google.com/maps/search/?api=1&query={latitude},{longitude}"
            else:
                gmaps_link = f"https://www.google.com/maps/search/{quote_plus(company_name + ' ' + address)}"
            
            real_email = ""
            real_social = ""
            cnpj = ""
            razao_social = ""
            socios = ""
            site_metodo = ""
            fallback_contribuiu = False
            
            if website_url and (email_opt or redes_opt):
                real_email, real_social, site_metodo = scrape_website_details(website_url)
            
            if not phone_raw or not real_email:
                fb_phone, fb_email, fb_social = fallback_search_phone_email(company_name, address, api_key, credit_counter)
                if not phone_raw and fb_phone:
                    phone_raw = fb_phone
                if not real_email and fb_email:
                    real_email = fb_email
                    fallback_contribuiu = True
                if not real_social and fb_social:
                    real_social = fb_social
                    fallback_contribuiu = True

            # --- 3ª camada: Hunter.io, só entra se ainda faltar e-mail ---
            hunter_contribuiu = False
            if hunter_opt and hunter_key and not real_email and website_url:
                hunter_email = hunter_domain_search(website_url, hunter_key)
                if hunter_email:
                    real_email = hunter_email
                    hunter_contribuiu = True

            # --- Monta a rastreabilidade de onde veio o e-mail/rede social ---
            fonte_partes = []
            if site_metodo:
                fonte_partes.append(f"Site Direto ({site_metodo})")
            if fallback_contribuiu:
                fonte_partes.append("Busca Fallback (Serper)")
            if hunter_contribuiu:
                fonte_partes.append("Hunter.io")
            fonte_contato = " + ".join(fonte_partes) if fonte_partes else "Não encontrado"

            if socios_opt:
                cnpj, razao_social, socios = fetch_cnpj_and_partners(company_name, address, credit_counter)

            formatted_phone, wa_link = clean_and_format_phone(phone_raw)

            extracted_data.append({
                "Prompt": query,
                "Nome da Empresa": company_name,
                "Razão Social": razao_social,
                "CNPJ": cnpj,
                "Sócios / Decisores": socios if socios else "Não identificado",
                "Categoria": category,
                "Endereço": address,
                "Telefone": formatted_phone,
                "Whatsapp": wa_link,
                "Email": real_email,
                "Redes Sociais": real_social,
                "Fonte do Contato": fonte_contato,
                "Nota Google": rating,
                "Total Avaliações": rating_count,
                "Status": "A Fazer",
                "Progressão": "1º Contato",
                "Tem Website": "Sim" if website_url else "Não",
                "Link Google Maps": gmaps_link,
                "Observações": f"Site: {website_url}" if website_url else "Sem site oficial"
            })
            
    except requests.exceptions.Timeout:
        st.error("❌ A Serper não respondeu mesmo após a segunda tentativa (25s). Isso costuma ser instabilidade pontual do lado deles — tente rodar a busca de novo em alguns instantes, ou torne o termo um pouco mais amplo (ex: tire o bairro) se persistir.")
    except Exception as e:
        st.error(f"Erro na requisição: {e}")

    return pd.DataFrame(extracted_data), credit_counter[0]

if btn_extrair:
    active_key = st.session_state['serper_api_key']
    with st.spinner(f"Extraindo leads e analisando dados para '{termo_busca}'..."):
        df_leads, creditos_usados = run_places_extraction(
            termo_busca, active_key, qtd_resultados, 
            enriquecer_emails, enriquecer_redes, buscar_socios,
            usar_hunter, st.session_state['hunter_api_key']
        )
        
        if not df_leads.empty:
            filename = "leads_extraidos.xlsx"
            create_excel_report(df_leads, filename)
            
            st.session_state['df_leads'] = df_leads
            st.session_state['filename'] = filename
            st.session_state['creditos_usados'] = creditos_usados
            st.success(f"✅ Sucesso! {len(df_leads)} empresas extraídas com sucesso. (≈{creditos_usados} créditos Serper usados nesta extração)")

# --- APRESENTAÇÃO DOS RESULTADOS E GERADOR DE SCRIPTS ---
if 'df_leads' in st.session_state and not st.session_state['df_leads'].empty:
    df_leads = st.session_state['df_leads']
    filename = st.session_state['filename']
    
    tab1, tab2 = st.tabs(["📋 Tabela de Leads", "💬 Gerador de Script de Vendas"])
    
    with tab1:
        st.subheader("📋 Prévia dos Resultados Reais")
        st.dataframe(df_leads, use_container_width=True)
        
        with open(filename, "rb") as file:
            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    with tab2:
        st.subheader("✍️ Gerador de Abordagem Comercial (Copywriting)")
        
        empresa_selecionada = st.selectbox(
            "Selecione uma empresa da lista para gerar o script:",
            options=df_leads["Nome da Empresa"].tolist()
        )
        
        lead_info = df_leads[df_leads["Nome da Empresa"] == empresa_selecionada].iloc[0]
        
        socio_nome = lead_info["Sócios / Decisores"].split("(")[0].strip() if lead_info["Sócios / Decisores"] != "Não identificado" else "Responsável"
        categoria = lead_info["Categoria"]
        endereco = lead_info["Endereço"]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📱 Script para WhatsApp (Mensagem Curta)")
            script_wa = f"""Olá, {socio_nome}! Tudo bem?

Vi a {empresa_selecionada} aqui no Google e notei que vocês são referência em {categoria} na região.

Estou entrando em contato pois ajudamos empresas do seu segmento a aumentarem o volume de clientes e otimizarem a presença digital.

Você teria 5 minutos nesta semana para trocarmos uma ideia rápida sobre como podemos ajudar a {empresa_selecionada}?

Abraços!"""
            
            st.text_area("Cópia Direta WhatsApp:", value=script_wa, height=220)
            
        with col2:
            st.markdown("### ✉️ E-mail de Apresentação (Cold Mail)")
            script_email = f"""Assunto: Oportunidade de crescimento para {empresa_selecionada}

Olá, {socio_nome}, espero que este e-mail o encontre bem.

Meu nome é [Seu Nome] e acompanho o trabalho de empresas do setor de {categoria} na região.

Analisando a presença digital da {empresa_selecionada}, identifiquei algumas oportunidades claras para expandir a captação de novos clientes qualificados todos os meses.

Desenvolvemos estratégias sob medida que ajudam empresas como a sua a se destacarem e converterem mais oportunidades.

Podemos agendar uma breve conversa de 10 minutos na próxima terça-feira às 14h para eu te apresentar esse diagnóstico gratuitamente?

Atenciosamente,
[Seu Nome] | [Sua Empresa]"""
            
            st.text_area("Cópia Direta E-mail:", value=script_email, height=220)

# --- RODAPÉ COM GERENCIAMENTO DE CONTA E CRÉDITOS ---
st.markdown("---")

# Nota técnica: o Serper.dev não expõe um endpoint público de saldo/créditos
# (confirmado na documentação deles — só têm search/images/videos/places/maps/
# reviews/news/shopping/lens/patents/autocomplete/scrape). Por isso, em vez de
# tentar consultar um saldo que a API não fornece, mostramos o consumo real
# contado pelo próprio app na última extração — que é 100% confiável.
creditos_usados_display = (
    f"**≈{st.session_state['creditos_usados']}** créditos usados na última extração"
    if 'creditos_usados' in st.session_state
    else "*Rode uma extração para ver o consumo estimado aqui*"
)

col_f1, col_f2 = st.columns([3, 1])

with col_f1:
    st.markdown(f"""
    **📊 Consumo Serper (estimado pelo app):**  
    {creditos_usados_display}  
    *Chave Ativa:* `{st.session_state['serper_api_key'][:8]}...{st.session_state['serper_api_key'][-4:]}`  
    *Nota: o Serper não tem endpoint público de saldo, então não dá pra consultar o total restante — só contamos quanto ESTE app gastou.*
    """)

    if st.session_state['hunter_api_key']:
        hunter_available, hunter_used = get_hunter_credits(st.session_state['hunter_api_key'])
        if hunter_available is not None:
            hunter_display = f"**{hunter_available}** buscas restantes (usadas: {hunter_used})"
        else:
            hunter_display = "⚠️ *Não foi possível consultar*"
        st.markdown(f"""
        **📊 Status da Conta Hunter.io:**  
        Saldo Atual: {hunter_display}
        """)

with col_f2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.link_button("🔗 Painel Serper.dev", "https://serper.dev/dashboard", use_container_width=True)
    if st.session_state['hunter_api_key']:
        st.link_button("🔗 Painel Hunter.io", "https://hunter.io/api-keys", use_container_width=True)
