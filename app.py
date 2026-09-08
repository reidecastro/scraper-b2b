import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
import requests
from urllib.parse import quote_plus
from googlesearch import search

# Configuração da Página
st.set_page_config(
    page_title="Gerador de Leads B2B - Busca Real",
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

st.markdown('<div class="main-header">🎯 Gerador de Leads B2B - Dados Reais</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Extraia e estruture dados reais de empresas locais sem uso de simulações ou custos.</div>', unsafe_allow_html=True)

# Sidebar - Configurações de Busca
with st.sidebar:
    st.header("⚙️ Configurações da Busca")
    
    termo_busca = st.text_input("Termo de Busca e Bairro", value="Pizzarias Campinas SP Bairro Castelo")
    qtd_resultados = st.number_input("Quantidade de Resultados", min_value=1, max_value=20, value=10, step=1)

    st.markdown("---")
    st.header("🔍 Opções de Enriquecimento")
    enriquecer_emails = st.checkbox("Buscar E-mails nos Sites", value=True)
    enriquecer_redes = st.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    btn_extrair = st.button("🚀 Iniciar Extração de Leads", use_container_width=True)

# Higienização e Formatação de Telefones
def clean_and_format_phone(phone_str):
    if not phone_str or pd.isna(phone_str):
        return "", ""
    digits = re.sub(r'\D', '', str(phone_str))
    if len(digits) < 8:
        return "", ""
    
    formatted_phone = phone_str
    whatsapp = f"https://wa.me/55{digits}" if len(digits) in [10, 11] else ""
    return formatted_phone, whatsapp

# Scraping leve dentro dos sites reais encontrados
def scrape_website_details(website_url):
    email = ""
    social = ""
    if not website_url or not str(website_url).startswith("http"):
        return email, social

    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(website_url, headers=headers, timeout=4)
        if response.status_code == 200:
            text = response.text
            
            # E-mail real via Regex
            emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
            if emails:
                valid_emails = [e for e in emails if not e.endswith(('.png', '.jpg', '.webp', '.js', '.css', '.svg'))]
                if valid_emails:
                    email = valid_emails[0]
            
            # Redes Sociais reais via Regex
            socials = re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[a-zA-Z0-9_.-]+', text)
            if socials:
                social = socials[0]
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

# Extração de Dados Reais
def run_real_extraction(query, max_results=10, email_opt=True, redes_opt=True):
    extracted_data = []
    
    try:
        # Busca resultados de sites reais no Google usando a biblioteca googlesearch
        urls = list(search(query, num_results=max_results, lang="pt"))
        
        for url in urls:
            # Filtra agregadores irrelevantes para focar em sites reais de empresas
            if any(domain in url for domain in ["tripadvisor", "ifood", "nhandeara", "facebook.com", "instagram.com"]):
                continue
            
            # Extrai o nome limpo do domínio
            domain_name = url.split("//")[-1].split("/")[0].replace("www.", "")
            company_name = domain_name.split(".")[0].capitalize()
            
            # Link oficial de busca direta no Google Maps para o nome da empresa
            gmaps_link = f"https://www.google.com/maps/search/{quote_plus(company_name + ' ' + query)}"
            
            real_email = ""
            real_social = ""
            
            if email_opt or redes_opt:
                found_email, found_social = scrape_website_details(url)
                if email_opt:
                    real_email = found_email
                if redes_opt:
                    real_social = found_social

            record = {
                "Prompt": query,
                "Nome da Empresa": company_name,
                "Categoria": "Comércio Local / Empresa",
                "Responsável": "",
                "Endereço": query,
                "Telefone": "",
                "Whatsapp": "",
                "Email": real_email,
                "Redes Sociais": real_social,
                "Status": "A Fazer",
                "Progressão": "1º Contato",
                "Tem Website": "Sim",
                "Link Google Maps": gmaps_link,
                "Observações": f"Site encontrado: {url}"
            }
            extracted_data.append(record)
            
    except Exception:
        pass

    return pd.DataFrame(extracted_data)

# Execução do Botão
if btn_extrair:
    with st.spinner(f"Extraindo empresas reais para '{termo_busca}'..."):
        df_leads = run_real_extraction(termo_busca, qtd_resultados, enriquecer_emails, enriquecer_redes)
        
        if not df_leads.empty:
            filename = "leads_extraidos.xlsx"
            create_excel_report(df_leads, filename)
            
            st.session_state['df_leads'] = df_leads
            st.session_state['filename'] = filename
            st.success(f"✅ Sucesso! Extraídos {len(df_leads)} leads reais.")
        else:
            st.error("Nenhum resultado foi encontrado para o termo pesquisado. Tente reformular a busca.")

# Exibição e Download
if 'df_leads' in st.session_state:
    df_leads = st.session_state['df_leads']
    filename = st.session_state['filename']
    
    st.subheader("📋 Prévia dos Resultados Reais")
    st.dataframe(df_leads, use_container_width=True)
    
    with open(filename, "rb") as file:
        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
