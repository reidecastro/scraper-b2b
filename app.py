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

# --- NOVO: Scrapling (impersonation TLS) - fura Cloudflare/WAF ---
try:
    from scrapling.fetchers import Fetcher
    SCRAPLING_AVAILABLE = True
except ImportError:
    SCRAPLING_AVAILABLE = False

# --- SEGURANCA V4 ---
DEFAULT_KEY = "" # REMOVIDO POR SEGURANCA - use st.secrets
# KEY_FILE removido - nao funciona no Streamlit Cloud

def get_api_key():
    if "SERPER_API_KEY" in st.secrets:
        return st.secrets["SERPER_API_KEY"]
    if 'serper_api_key' in st.session_state and st.session_state['serper_api_key']:
        return st.session_state['serper_api_key']
    return ""

def load_saved_key():
    # V4 - Seguro: usa st.secrets
    if "SERPER_API_KEY" in st.secrets:
        return st.secrets["SERPER_API_KEY"]
    return ""

def save_key(key_str):
    # V4 - Nao salva em disco no Cloud, apenas em memoria
    try:
        st.session_state["serper_api_key"] = key_str.strip()
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")

if 'serper_api_key' not in st.session_state:
    st.session_state['serper_api_key'] = get_api_key()

st.set_page_config(
    page_title="Gerador de Leads B2B - Prospecção Avançada",
    page_icon="🎯",
    layout="wide"
)

st.markdown("""
<style>
   .main-header { font-size: 2.2rem; color: #60A5FA; font-weight: 700; margin-bottom: 0.5rem; }
   .sub-header { font-size: 1.1rem; color: #D1D5DB; margin-bottom: 1.5rem; }
   .stButton>button { background-color: #2563EB; color: white; border-radius: 6px; padding: 0.5rem 1.5rem; font-weight: 600; border: none; }
   .stButton>button:hover { background-color: #1D4ED8; color: white; }
</style>
""", unsafe_allow_html=True)

def get_serper_credits(api_key):
    if not api_key:
        return None
    try:
        headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
        res = requests.post("https://google.serper.dev/credits", headers=headers, timeout=4)
        if res.status_code == 200:
            return res.json().get("credits")
    except Exception:
        pass
    return None

with st.sidebar:
    if os.path.exists("wolfologo.png"):
        st.image("wolfologo.png", use_container_width=True)
    st.header("🔑 Configuração da API")
    st.caption("Configure em Settings > Secrets no Streamlit Cloud")
    current_key = st.session_state['serper_api_key']
    new_key = st.text_input("Chave Serper API", value=current_key, type="password")
    if new_key!= current_key:
        st.session_state['serper_api_key'] = new_key
        save_key(new_key)
        st.success("Chave atualizada!")
    st.markdown("---")
    st.header("⚙️ Configurações da Busca")
    termo_busca = st.text_input("Termo de Busca e Bairro", value="Pizzarias Campinas SP Bairro Castelo")
    qtd_resultados = st.number_input("Quantidade de Resultados", min_value=1, max_value=20, value=10, step=1)
    st.markdown("---")
    st.header("🔍 Opções de Enriquecimento")
    enriquecer_emails = st.checkbox("Buscar E-mails nos Sites", value=True)
    enriquecer_redes = st.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)
    buscar_socios = st.checkbox("Buscar CNPJ & Sócios (BrasilAPI)", value=True)
    st.markdown("<br>", unsafe_allow_html=True)
    btn_extrair = st.button("🚀 Iniciar Extração de Leads", use_container_width=True)

st.markdown('<div class="main-header">🎯 Gerador de Leads B2B - Prospecção Avançada</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Extração de dados + CRM B2B Custo Zero</div>', unsafe_allow_html=True)

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

def extract_address_from_item(item, company_name=""):
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
    def fetch_cnpj_and_partners(company_name, city_or_address=""):
    cnpj_clean = ""
    razao_social = ""
    socios_names = []
    location_hint = city_or_address.split("-")[0].strip() if city_or_address else ""
    query = f"{company_name} {location_hint} cnpj brasilapi"
    url = "https://google.serper.dev/search"
    payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
    headers = {'X-API-KEY': st.session_state['serper_api_key'], 'Content-Type': 'application/json'}
    try:
        res = requests.post(url, headers=headers, json=payload, timeout=4)
        if res.status_code == 200:
            data = res.json()
            text_block = ""
            for item in data.get("organic", []):
                text_block += " " + item.get("snippet", "") + " " + item.get("title", "")
            cnpjs = re.findall(r'\b\d{2}\.?\\d{3}\.?\\d{3}/?\d{4}-?\d{2}\b', text_block)
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

def fallback_search_phone_email(company_name, address, api_key):
    phone, email, social = "", "", ""
    query = f"{company_name} {address} telefone contato email"
    url = "https://google.serper.dev/search"
    payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    try:
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

def decode_cf_email(encoded):
    try:
        r = int(encoded[:2], 16)
        email = ''.join([chr(int(encoded[i:i+2], 16) ^ r) for i in range(2, len(encoded), 2)])
        return email
    except:
        return ""

def scrape_website_details(website_url):
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
            except Exception:
                text = ""
        if not text:
            try:
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                response = requests.get(url_tentativa, headers=headers, timeout=4)
                if response.status_code == 200:
                    text = response.text
                    metodo = "Requests"
            except Exception:
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
            mailtos = re.findall(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})', text_total, re.IGNORECASE)
            if mailtos:
                email = mailtos[0]
        if not email:
            emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}', text_total)
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
        socials = re.findall(r'https?://(?:www\.)?(?:instagram\\.com|facebook\\.com)/[a-zA-Z0-9_.-]+', text_total)
        if socials:
            social = socials[0]
    return email, social, metodo

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

def run_places_extraction(query, api_key, max_results=10, email_opt=True, redes_opt=True, socios_opt=True):
    extracted_data = []
    url = "https://google.serper.dev/places"
    payload = {"q": query, "gl": "br", "hl": "pt-br"}
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        if response.status_code == 402 or "Out of credits" in response.text:
            st.error("❌ Os créditos da chave Serper atual ACABARAM!")
            return pd.DataFrame()
        data = response.json()
        places = data.get("places", [])
        for item in places[:max_results]:
            company_name = item.get("title", "")
            address = extract_address_from_item(item, company_name)
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
                fb_phone, fb_email, fb_social = fallback_search_phone_email(company_name, address, api_key)
                if not phone_raw and fb_phone:
                    phone_raw = fb_phone
                if not real_email and fb_email:
                    real_email = fb_email
                    fallback_contribuiu = True
                if not real_social and fb_social:
                    real_social = fb_social
                    fallback_contribuiu = True
            fonte_partes = []
            if site_metodo:
                fonte_partes.append(f"Site Direto ({site_metodo})")
            if fallback_contribuiu:
                fonte_partes.append("Busca Fallback (Serper)")
            fonte_contato = " + ".join(fonte_partes) if fonte_partes else "Não encontrado"
            if socios_opt:
                cnpj, razao_social, socios = fetch_cnpj_and_partners(company_name, address)
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
    except Exception as e:
        st.error(f"Erro na requisição: {e}")
    return pd.DataFrame(extracted_data)

if btn_extrair:
    active_key = st.session_state['serper_api_key']
    if not active_key:
        st.error("❌ Configure sua chave Serper API na barra lateral ou em st.secrets")
    else:
        with st.spinner(f"Extraindo leads e analisando dados para '{termo_busca}'..."):
            df_leads = run_places_extraction(termo_busca, active_key, qtd_resultados, enriquecer_emails, enriquecer_redes, buscar_socios)
            if not df_leads.empty:
                filename = "leads_extraidos.xlsx"
                create_excel_report(df_leads, filename)
                st.session_state['df_leads'] = df_leads
                st.session_state['filename'] = filename
                st.success(f"✅ Sucesso! {len(df_leads)} empresas extraídas com sucesso.")

if 'df_leads' in st.session_state and not st.session_state['df_leads'].empty:
    df_leads = st.session_state['df_leads']
    filename = st.session_state['filename']
    tab1, tab2 = st.tabs(["📋 Tabela de Leads", "💬 Gerador de Script de Vendas"])
    with tab1:
        st.subheader("📋 Prévia dos Resultados Reais")
        st.dataframe(df_leads, use_container_width=True)
        with open(filename, "rb") as file:
            st.download_button(label="📥 Baixar Planilha Excel (.xlsx)", data=file, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tab2:
        st.subheader("✍️ Gerador de Abordagem Comercial (Copywriting)")
        empresa_selecionada = st.selectbox("Selecione uma empresa da lista para gerar o script:", options=df_leads["Nome da Empresa"].tolist())
        lead_info = df_leads[df_leads["Nome da Empresa"] == empresa_selecionada].iloc[0]
        socio_nome = lead_info["Sócios / Decisores"].split("(")[0].strip() if lead_info["Sócios / Decisores"]!= "Não identificado" else "Responsável"
        categoria = lead_info["Categoria"]
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 📱 Script para WhatsApp (Mensagem Curta)")
            script_wa = f"Olá, {socio_nome}! Tudo bem?\n\nVi a {empresa_selecionada} aqui no Google e notei que vocês são referência em {categoria} na região.\n\nVocê teria 5 minutos nesta semana?"
            st.text_area("Cópia Direta WhatsApp:", value=script_wa, height=220)
        with col2:
            st.markdown("### ✉️ E-mail de Apresentação (Cold Mail)")
            script_email = f"Assunto: Oportunidade de crescimento para {empresa_selecionada}\n\nOlá, {socio_nome},\n\nAnalisando a presença digital da {empresa_selecionada}, identifiquei oportunidades claras para expandir a captação.\n\nPodemos agendar uma breve conversa de 10 minutos?"
            st.text_area("Cópia Direta E-mail:", value=script_email, height=220)

st.markdown("---")
credits_val = get_serper_credits(st.session_state['serper_api_key'])
credits_display = f"**{credits_val:,}** créditos" if credits_val is not None else "⚠️ *Não foi possível consultar*"
col_f1, col_f2 = st.columns([3, 1])
with col_f1:
    key = st.session_state['serper_api_key']
    masked = f"{key[:8]}...{key[-4:]}" if len(key) > 12 else "Não configurada"
    st.markdown(f"**📊 Status da Conta Serper:** \nSaldo Atual: {credits_display} \n*Chave Ativa:* `{masked}`")
with col_f2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.link_button("🔗 Painel Serper.dev", "https://serper.dev/dashboard", use_container_width=True)
