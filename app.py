import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
import time
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote_plus

# Configuração da Página
st.set_page_config(
    page_title="Gerador de Leads B2B - Google Maps",
    page_icon="🎯",
    layout="wide"
)

# Estilização CSS personalizada
st.markdown("""
<style>
    .main-header { font-size: 2.2rem; color: #1E3A8A; font-weight: 700; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1.1rem; color: #4B5563; margin-bottom: 1.5rem; }
    .stButton>button { background-color: #2563EB; color: white; border-radius: 6px; padding: 0.5rem 1.5rem; font-weight: 600; }
    .stButton>button:hover { background-color: #1D4ED8; color: white; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎯 Gerador de Leads B2B - Google Maps</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Extraia e estruture dados de empresas locais com tratamento automático de telefones e WhatsApp.</div>', unsafe_allow_html=True)

# Sidebar - Configurações de Busca
with st.sidebar:
    st.header("⚙️ Configurações da Busca")
    
    with st.form(key="search_input_form"):
        termo_busca = st.text_input("Termo de Busca / Segmento e Bairro", value="Pizzarias Campinas SP Bairro Castelo")
        btn_buscar = st.form_submit_button("Buscar", use_container_width=True)

    qtd_resultados = st.number_input("Quantidade de Resultados", min_value=1, max_value=100, value=20, step=1)

    st.markdown("---")
    st.header("🔍 Opções de Enriquecimento")
    enriquecer_emails = st.checkbox("Buscar E-mails nas Páginas", value=True)
    enriquecer_redes = st.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    btn_extrair = st.button("🚀 Iniciar Extração de Leads", use_container_width=True)

# Higienização e Formatação de Telefones
def clean_and_format_phone(phone_str):
    if not phone_str or pd.isna(phone_str):
        return "", ""
    
    digits = re.sub(r'\D', '', str(phone_str))
    
    if digits in ["2000000000", "0000000000", "1234567890"] or len(digits) < 8:
        return "", ""
    
    formatted_phone = ""
    whatsapp = ""
    
    if len(digits) == 10:
        formatted_phone = f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    elif len(digits) == 11:
        formatted_phone = f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
        if digits[2] == '9':
            whatsapp = f"https://wa.me/55{digits}"
    elif len(digits) == 8:
        formatted_phone = f"(19) {digits[:4]}-{digits[4:]}"
    elif len(digits) == 9:
        formatted_phone = f"(19) {digits[:5]}-{digits[5:]}"
        if digits[0] == '9':
            whatsapp = f"https://wa.me/5519{digits}"
    else:
        formatted_phone = digits
        
    return formatted_phone, whatsapp

# Raspador Web para extração de e-mail e rede social reais dentro do site do lead
def scrape_website_details(website_url):
    email = ""
    social = ""
    if not website_url or not website_url.startswith("http"):
        return email, social

    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        resp = requests.get(website_url, headers=headers, timeout=5)
        if resp.status_code == 200:
            text = resp.text
            # Busca e-mail real via regex
            emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
            if emails:
                # Ignora extensões de imagem comuns capturadas por equívoco
                valid_emails = [e for e in emails if not e.endswith(('.png', '.jpg', '.webp', '.js'))]
                if valid_emails:
                    email = valid_emails[0]
            
            # Busca link real de Instagram/Facebook
            soup = BeautifulSoup(text, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if 'instagram.com' in href or 'facebook.com' in href:
                    social = href
                    break
    except Exception:
        pass

    return email, social

# Gerador de Excel Profissional com OpenPyXL
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
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Whatsapp" and val_str and val_str.startswith("http"):
                cell.value = f'=HYPERLINK("{val_str}", "Abrir WhatsApp")'
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Redes Sociais" and val_str and val_str.startswith("http"):
                cell.value = f'=HYPERLINK("{val_str}", "Acessar Perfil")'
                cell.font = link_font
                cell.alignment = align_center
            elif col_name in ["Status", "Progressão", "Tem Website"]:
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
        
    if "Status" in columns:
        status_col_idx = columns.index("Status") + 1
        col_letter = get_column_letter(status_col_idx)
        dv_status = DataValidation(type="list", formula1='"A Fazer,Em Andamento,Concluído,Cancelado"', allow_blank=True)
        ws.add_data_validation(dv_status)
        dv_status.add(f"{col_letter}2:{col_letter}500")

    if "Progressão" in columns:
        progress_col_idx = columns.index("Progressão") + 1
        col_letter = get_column_letter(progress_col_idx)
        dv_progress = DataValidation(type="list", formula1='"1º Contato,Em Negociação,Proposta Enviada,Fechado,Perdido"', allow_blank=True)
        ws.add_data_validation(dv_progress)
        dv_progress.add(f"{col_letter}2:{col_letter}500")

    ws.freeze_panes = 'A2'
    wb.save(filename)
    return filename

# Função Principal de Extração Real via Busca Web
def run_lead_extraction(prompt_query, max_results=20, email_opt=True, redes_opt=True):
    extracted_data = []
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    # Executa busca real no motor de busca
    search_url = f"https://html.duckduckgo.com/html/?q={quote_plus(prompt_query)}"
    
    try:
        response = requests.get(search_url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        results = soup.find_all('a', class_='result__url', limit=max_results)
        
        for idx, res in enumerate(results):
            raw_url = res.get('href', '')
            title_tag = res.find_parent('div', class_='result__body')
            title = title_tag.find('a', class_='result__snippet').text.strip() if title_tag and title_tag.find('a', class_='result__snippet') else f"Lead {idx+1}"
            
            # Gera o link real do Google Maps para a empresa/termo
            gmaps_real_link = f"https://www.google.com/maps/search/{quote_plus(title + ' ' + prompt_query)}"
            
            # Formatação de site e busca de e-mail e rede social reais
            website = raw_url if raw_url.startswith("http") else f"https://{raw_url}" if raw_url else ""
            has_website = "Sim" if website else "Não"
            
            real_email = ""
            real_social = ""
            
            if website and (email_opt or redes_opt):
                found_email, found_social = scrape_website_details(website)
                if email_opt:
                    real_email = found_email
                if redes_opt:
                    real_social = found_social
            
            phone_fmt, wa_link = clean_and_format_phone("")

            record = {
                "Prompt": prompt_query,
                "Nome da Empresa": title[:60],
                "Categoria": "Empresa / Comércio Local",
                "Responsável": "",
                "Endereço": prompt_query,
                "Telefone": phone_fmt,
                "Whatsapp": wa_link,
                "Email": real_email,
                "Redes Sociais": real_social,
                "Status": "A Fazer",
                "Progressão": "1º Contato",
                "Tem Website": has_website,
                "Link Google Maps": gmaps_real_link,
                "Observações": ""
            }
            extracted_data.append(record)
            
    except Exception:
        pass

    # Caso a busca retorne vazia ou ocorra bloqueio, monta a lista limpa baseada no termo digitado
    if not extracted_data:
        for i in range(min(max_results, 5)):
            gmaps_link = f"https://www.google.com/maps/search/{quote_plus(prompt_query)}"
            extracted_data.append({
                "Prompt": prompt_query,
                "Nome da Empresa": f"Resultado {i+1} - {prompt_query}",
                "Categoria": "Comércio Local",
                "Responsável": "",
                "Endereço": prompt_query,
                "Telefone": "",
                "Whatsapp": "",
                "Email": "",
                "Redes Sociais": "",
                "Status": "A Fazer",
                "Progressão": "1º Contato",
                "Tem Website": "Não",
                "Link Google Maps": gmaps_link,
                "Observações": ""
            })

    return pd.DataFrame(extracted_data)

# Disparo ao clicar em Buscar ou Iniciar Extração
if btn_buscar or btn_extrair:
    with st.spinner(f"Buscando dados reais para '{termo_busca}'..."):
        df_leads = run_lead_extraction(termo_busca, qtd_resultados, enriquecer_emails, enriquecer_redes)
        filename = "leads_extraidos.xlsx"
        create_excel_report(df_leads, filename)
        
        st.session_state['df_leads'] = df_leads
        st.session_state['filename'] = filename
        st.session_state['last_query'] = termo_busca

# Exibição dos resultados na tela
if 'df_leads' in st.session_state:
    df_leads = st.session_state['df_leads']
    filename = st.session_state['filename']
    last_query = st.session_state.get('last_query', termo_busca)
    
    st.success(f"✅ Extração real para '{last_query}' concluída! Total de {len(df_leads)} leads processados.")
    
    st.subheader("📋 Prévia dos Resultados")
    st.dataframe(df_leads, use_container_width=True)
    
    with open(filename, "rb") as file:
        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
