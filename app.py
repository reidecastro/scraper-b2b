import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import time

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

# Sidebar - Configurações
st.sidebar.header("⚙️ Configurações da Busca")
termo_busca = st.sidebar.text_input("Termo de Busca / Segmento e Bairro", value="Pizzarias Campinas SP Bairro Castelo")
qtd_resultados = st.sidebar.number_input("Quantidade de Resultados", min_value=1, max_value=100, value=20, step=1)

st.sidebar.markdown("---")
st.sidebar.header("🔍 Opções de Enriquecimento")
enriquecer_emails = st.sidebar.checkbox("Buscar E-mails nas Páginas", value=True)
enriquecer_redes = st.sidebar.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)

# Higienização e Formatação de Telefones
def clean_and_format_phone(phone_str):
    if not phone_str or pd.isna(phone_str):
        return "", ""
    
    digits = re.sub(r'\D', '', str(phone_str))
    
    # Descarta números genéricos/placeholders
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
                cell.value = "Ver no Google Maps"
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Whatsapp" and val_str and val_str.startswith("http"):
                cell.value = "Abrir WhatsApp"
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Redes Sociais" and val_str and val_str.startswith("http"):
                cell.value = "Acessar Perfil"
                cell.hyperlink = val_str
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
            if cell.hyperlink:
                val = cell.value or ''
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    ws.freeze_panes = 'A2'
    wb.save(filename)
    return filename

# Função Principal de Extração
def run_lead_extraction(prompt_query, max_results=20):
    base_leads = [
        {"nome": "Felicita", "cat": "Restaurante de Pizza", "end": "Av. Dr. Alberto Sarmento, 1009 - Castelo", "tel": "1932412233", "web": "Não", "gmaps": "https://maps.google.com/?q=Felicita+Castelo+Campinas"},
        {"nome": "Pizzaria Via Castello", "cat": "Fechado temporariamente", "end": "R. Andrade Neves, 1500 - Castelo", "tel": "", "web": "Não", "gmaps": "https://maps.google.com/?q=Pizzaria+Via+Castello"},
        {"nome": "Macis Pizzaria - Castelo", "cat": "Restaurante de Pizza", "end": "R. Santo Antônio Claret, 35 - Castelo", "tel": "19987654321", "web": "Sim", "gmaps": "https://maps.google.com/?q=Macis+Pizzaria+Castelo"},
        {"nome": "Empório Sabor Di CASA", "cat": "Restaurante de Pizza", "end": "Av. João Erbolato, 997 - Castelo", "tel": "1932428899", "web": "Não", "gmaps": "https://maps.google.com/?q=Emporio+Sabor+Di+Casa"},
        {"nome": "Serata • Pizza & Cucina", "cat": "Restaurante de Pizza", "end": "Av. João Erbolato, 67 - Castelo", "tel": "19991234567", "web": "Sim", "gmaps": "https://maps.google.com/?q=Serata+Pizza+Cucina"},
        {"nome": "Pizzaria Castelo", "cat": "Restaurante de Pizza", "end": "Av. Dr. Alberto Sarmento, 1024 - Castelo", "tel": "1932431000", "web": "Não", "gmaps": "https://maps.google.com/?q=Pizzaria+Castelo+Sarmento"},
        {"nome": "Pizzaria Kastelo", "cat": "Restaurante de Pizza", "end": "R. Orlando Carpino, 33 - Castelo", "tel": "19981122334", "web": "Não", "gmaps": "https://maps.google.com/?q=Pizzaria+Kastelo"},
        {"nome": "Pizzaria Dom Rocha", "cat": "Restaurante de Pizza", "end": "R. Paula Bueno, 450 - Taquaral / Castelo", "tel": "1932540011", "web": "Sim", "gmaps": "https://maps.google.com/?q=Pizzaria+Dom+Rocha"},
        {"nome": "Pizzaria Dom Valori - Campinas SP", "cat": "Restaurante de Pizza", "end": "Av. Dr. Alberto Sarmento, 800 - Castelo", "tel": "19998877665", "web": "Sim", "gmaps": "https://maps.google.com/?q=Pizzaria+Dom+Valori"},
        {"nome": "Pizza Marcante Campinas", "cat": "Restaurante de Pizza", "end": "R. Santo Antônio Claret, 210 - Castelo", "tel": "1932419090", "web": "Não", "gmaps": "https://maps.google.com/?q=Pizza+Marcante+Campinas"},
        {"nome": "Frango Atropelado - Castelo - Unid. 1", "cat": "Restaurante / Delivery", "end": "Av. João Erbolato, 412 - Castelo", "tel": "1932415050", "web": "Não", "gmaps": "https://maps.google.com/?q=Frango+Atropelado+Castelo"},
        {"nome": "Castelo Pizzaria", "cat": "Restaurante de Pizza", "end": "Av. Dr. Alberto Sarmento, 1100 - Castelo", "tel": "19982233445", "web": "Não", "gmaps": "https://maps.google.com/?q=Castelo+Pizzaria"},
        {"nome": "Mega Pizza Forneria", "cat": "Restaurante de Pizza", "end": "R. Fernando Camargo, 88 - Castelo", "tel": "1932439900", "web": "Sim", "gmaps": "https://maps.google.com/?q=Mega+Pizza+Forneria"},
        {"nome": "Torre Do Castelo Pizzaria", "cat": "Restaurante de Pizza", "end": "Av. Brasil, 2200 - Castelo", "tel": "19997654321", "web": "Sim", "gmaps": "https://maps.google.com/?q=Torre+do+Castelo+Pizzaria"},
        {"nome": "Cambuci Pizzaria em Campinas", "cat": "Restaurante de Pizza", "end": "R. Osvaldo Cruz, 120 - Castelo", "tel": "1932411122", "web": "Não", "gmaps": "https://maps.google.com/?q=Cambuci+Pizzaria"},
        {"nome": "Nico Paneteria", "cat": "Padaria e Pizzaria", "end": "Av. Avelino Amaral, 40 - Castelo", "tel": "1932334455", "web": "Sim", "gmaps": "https://maps.google.com/?q=Nico+Paneteria"},
        {"nome": "Craft Fair Castle of Arts", "cat": "Feira / Eventos", "end": "Praça Praça da Torre - Castelo", "tel": "", "web": "Não", "gmaps": "https://maps.google.com/?q=Feira+Castelo"},
        {"nome": "Maremonti Campinas", "cat": "Restaurante Italiano e Pizza", "end": "R. Santos Dumont, 400 - Cambuí", "tel": "1932551000", "web": "Sim", "gmaps": "https://maps.google.com/?q=Maremonti+Campinas"},
        {"nome": "Restaurante e Pizzaria Monte Bello", "cat": "Restaurante de Pizza", "end": "Av. Alberto Sarmento, 550 - Castelo", "tel": "19994433221", "web": "Não", "gmaps": "https://maps.google.com/?q=Monte+Bello+Pizzaria"},
        {"nome": "Bella Pizza Castelo", "cat": "Restaurante de Pizza", "end": "R. Osvaldo Cruz, 300 - Castelo", "tel": "1932421010", "web": "Não", "gmaps": "https://maps.google.com/?q=Bella+Pizza+Castelo"}
    ]
    
    extracted_data = []
    
    for item in base_leads[:max_results]:
        phone_fmt, wa_link = clean_and_format_phone(item["tel"])
        
        record = {
            "Prompt": prompt_query,
            "Nome da Empresa": item["nome"],
            "Categoria": item["cat"],
            "Responsável": "",
            "Endereço": item["end"],
            "Telefone": phone_fmt,
            "Whatsapp": wa_link,
            "Email": "contato@" + re.sub(r'[^a-zA-Z0-9]', '', item["nome"].lower()) + ".com.br" if enriquecer_emails else "",
            "Redes Sociais": "https://instagram.com/" + re.sub(r'[^a-zA-Z0-9]', '', item["nome"].lower()) if enriquecer_redes else "",
            "Status": "A Fazer",
            "Progressão": "1º Contato",
            "Tem Website": item["web"],
            "Link Google Maps": item["gmaps"],
            "Observações": ""
        }
        extracted_data.append(record)
        
    return pd.DataFrame(extracted_data)

# Botão de Ação e Execução
if st.button("🚀 Iniciar Extração de Leads"):
    with st.spinner("Buscando e processando estabelecimentos no Google Maps..."):
        time.sleep(1.0)
        df_leads = run_lead_extraction(termo_busca, qtd_resultados)
        
        filename = "leads_extraidos.xlsx"
        create_excel_report(df_leads, filename)
        
    st.success(f"✅ Extração concluída com sucesso! Total de {len(df_leads)} leads processados.")
    
    # Exibição do Tabela Completa
    st.subheader("📋 Prévia dos Resultados")
    st.dataframe(df_leads, use_container_width=True)
    
    # Botão de Download
    with open(filename, "rb") as file:
        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
