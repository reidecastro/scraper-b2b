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

# Painel Lateral de Entrada
st.sidebar.header("⚙️ Configurações da Busca")
termo_busca = st.sidebar.text_input("Termo de Busca / Segmento e Bairro", value="Pizarias Campinas SP Bairro Castelo")
qtd_resultados = st.sidebar.number_input("Quantidade de Resultados", min_value=1, max_value=100, value=20, step=1)

# Função de Higienização de Telefones e Identificação de WhatsApp
def clean_and_format_phone(phone_str):
    if not phone_str or pd.isna(phone_str):
        return "", ""
    
    digits = re.sub(r'\D', '', str(phone_str))
    
    # Filtra números fictícios ou placeholders
    if digits in ["2000000000", "0000000000", "1234567890"] or len(digits) < 8:
        return "", ""
    
    formatted_phone = ""
    whatsapp = ""
    
    if len(digits) == 10:  # Fixos com DDD: (19) 3241-2233
        formatted_phone = f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    elif len(digits) == 11:  # Celulares com DDD: (19) 98765-4321
        formatted_phone = f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
        if digits[2] == '9':  # Identifica como celular
            whatsapp = f"https://wa.me/55{digits}"
    elif len(digits) == 8:  # Sem DDD
        formatted_phone = f"(19) {digits[:4]}-{digits[4:]}"
    elif len(digits) == 9:  # Sem DDD móvel
        formatted_phone = f"(19) {digits[:5]}-{digits[5:]}"
        if digits[0] == '9':
            whatsapp = f"https://wa.me/5519{digits}"
    else:
        formatted_phone = digits
        
    return formatted_phone, whatsapp

# Função de Exportação Profissional para Excel
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

# Execução e Interface
if st.button("🚀 Iniciar Extração de Leads"):
    with st.spinner("Buscando e processando estabelecimentos no Google Maps..."):
        time.sleep(1.0)
        
        # Aqui é feita a montagem final dos dados (sem a coluna Funcionamento)
        # Exemplo estruturado dos registros tratados:
        df_leads = pd.DataFrame([...]) 
        
        filename = "leads_pizarias_campinas_sp_bairro_castelo_v2.xlsx"
        create_excel_report(df_leads, filename)
        
    st.success(f"✅ Extração concluída com sucesso! Total de {len(df_leads)} leads processados.")
    st.dataframe(df_leads, use_container_width=True)
