import streamlit as st
import pandas as pd
from playwright.sync_api import sync_playwright
import time
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Scraper B2B Cloud", page_icon="📍", layout="wide")

st.title("📍 Scraper B2B de Empresas - Google Maps")
st.markdown("Busque leads B2B por nicho/região e baixe a planilha formatada e pronta para prospecção.")

# --- ENTRADA DE DADOS DO USUÁRIO ---
col1, col2 = st.columns([2, 1])

with col1:
    keyword = st.text_input("Palavra-chave (Nicho + Região):", placeholder="Ex: Restaurantes em Campinas SP")

with col2:
    max_results = st.number_input("Número máximo de resultados:", min_value=5, max_value=100, value=20, step=5)

deep_scrape = st.checkbox("🔍 Enriquecer dados acessando o Website dos leads (Extrai E-mail e Redes Sociais se disponível)", value=True)

# --- FUNÇÃO DE EXPORTAÇÃO EXCEL FORMATADA ---
def export_to_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads B2B"

    # Escreve os cabeçalhos
    headers = list(df.columns)
    ws.append(headers)

    # Estilo da Linha 1 (Cabeçalho): Fonte 12, Negrito, Centralizado, Fundo Azul Escuro
    header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    # Insere as linhas de dados
    for row in df.itertuples(index=False):
        ws.append(list(row))

    # Regras de Dropdown (Validação de Dados)
    dv_status = DataValidation(type="list", formula1='"A Fazer,Ativo,Inativo"', allow_blank=True)
    dv_prog = DataValidation(type="list", formula1='"1º Contato,2º Contato,1ª Reunião,Orçamento,Contrato Finalizado"', allow_blank=True)

    ws.add_data_validation(dv_status)
    ws.add_data_validation(dv_prog)

    max_row = ws.max_row
    if max_row > 1:
        # Coluna I (Status) e Coluna J (Progressão)
        dv_status.add(f"I2:I{max_row}")
        dv_prog.add(f"J2:J{max_row}")

    # Formatação de Links Clicáveis do Google Maps e Largura das Colunas
    for row_idx in range(2, max_row + 1):
        map_cell = ws.cell(row=row_idx, column=13)  # Coluna M: Link Google Maps
        if map_cell.value and str(map_cell.value).startswith("http"):
            url = str(map_cell.value)
            map_cell.value = "Ver no Google Maps"
            map_cell.hyperlink = url
            map_cell.font = Font(color="0000FF", underline="single")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 45)

    ws.row_dimensions[1].height = 28
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# --- HELPER: EXTRAÇÃO DE CONTATOS NO WEBSITE ---
def extract_contacts_from_website(context, website_url):
    emails = set()
    socials = set()
    
    try:
        page = context.new_page()
        page.goto(website_url, timeout=10000, wait_until="domcontentloaded")
        time.sleep(1)
        
        # Pega todo o HTML/texto da página
        html_content = page.content()
        
        # Expressão regular para achar e-mails válidos
        found_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', html_content)
        for e in found_emails:
            if not any(ext in e.lower() for ext in ['.png', '.jpg', '.jpeg', '.webp', '.svg', 'wixpress', 'sentry']):
                emails.add(e.lower())

        # Procura links de redes sociais
        links = page.query_selector_all('a[href]')
        for link in links:
            href = link.get_attribute('href') or ""
            if any(domain in href.lower() for domain in ['instagram.com', 'facebook.com', 'linkedin.com', 'twitter.com', 'x.com']):
                # Remove parâmetros de rastreio inúteis do link
                clean_link = href.split('?')[0].rstrip('/')
                if len(clean_link) > 15:
                    socials.add(clean_link)

        page.close()
    except Exception:
        pass
        
    email_str = ", ".join(list(emails)[:2]) if emails else ""
    social_str = ", ".join(list(socials)[:3]) if socials else ""
    return email_str, social_str

# --- FUNÇÃO DE SCRAPING NA NUVEM ---
def scrape_maps_cloud(keyword: str, max_results: int, token: str, do_deep_scrape: bool):
    results = []
    wss_url = f"wss://chrome.browserless.io?token={token}"
    
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(wss_url)
        context = browser.new_context()
        page = context.new_page()
        
        search_url = f"https://www.google.com/maps/search/{keyword.replace(' ', '+')}"
        page.goto(search_url, wait_until="domcontentloaded")
        
        try:
            page.wait_for_selector('div[role="feed"]', timeout=12000)
        except Exception:
            browser.close()
            return pd.DataFrame()

        scrollable_div = 'div[role="feed"]'
        previous_count = 0
        
        while len(results) < max_results:
            cards = page.query_selector_all('div[role="article"]')
            
            if len(cards) == previous_count and len(cards) > 0:
                time.sleep(2)
                cards = page.query_selector_all('div[role="article"]')
                if len(cards) == previous_count:
                    break
                    
            previous_count = len(cards)
            page.evaluate('(selector) => { const el = document.querySelector(selector); if (el) el.scrollBy(0, 1000); }', scrollable_div)
            time.sleep(1.5)

        cards = page.query_selector_all('div[role="article"]')[:max_results]
        
        for card in cards:
            try:
                # 1. Nome da Empresa
                nome = card.get_attribute("aria-label")
                if not nome:
                    title_elem = card.query_selector('.qBF1Pd, div.fontHeadlineSmall')
                    nome = title_elem.inner_text().strip() if title_elem else "N/A"

                # 2. Link do Google Maps
                link_elem = card.query_selector('a[href*="/maps/place/"]')
                if not link_elem:
                    link_elem = card.query_selector('a')
                url_maps = link_elem.get_attribute('href') if link_elem else ""

                # 3. Linhas internas do card
                lines = [line.strip() for line in card.inner_text().split("\n") if line.strip()]
                
                if nome == "N/A" and len(lines) > 0:
                    nome = lines[0]

                telefone = ""
                funcionamento = ""
                endereco = ""
                
                for line in lines:
                    if re.search(r'\(?\d{2}\)?\s*9?\d{4}[-\s]?\d{4}', line):
                        telefone = line
                    elif any(w in line.lower() for w in ["aberto", "fechado", "fecha às", "abre às", "24 horas"]):
                        funcionamento = line
                    elif any(w in line.lower() for w in ["rua", "r.", "av.", "avenida", "alameda", "praça", "bairro", "jd.", "jardim", "sp", "campinas"]):
                        if line != nome and not re.search(r'^\d[\.,]\d\s*\(\d+\)', line):
                            endereco = line

                if not endereco:
                    content_lines = [
                        l for l in lines 
                        if l != nome 
                        and not re.search(r'^\d[\.,]\d\s*\(\d+\)', l) 
                        and not any(w in l.lower() for w in ["aberto", "fechado", "fecha", "abre"]) 
                        and l != telefone
                    ]
                    if len(content_lines) > 1:
                        endereco = content_lines[1]
                    elif len(content_lines) > 0:
                        endereco = content_lines[0]

                whatsapp = telefone if ("9" in telefone and len(re.sub(r'\D', '', telefone)) >= 10) else ""
                
                # Checa presença de Website (Coluna L)
                website_elem = card.query_selector('a[href*="http"]:not([href*="google.com"])')
                website_url = website_elem.get_attribute('href') if website_elem else ""
                tem_website = "Sim" if website_url else "Não"

                # Se Tem Website == Sim e a opção de enriquecimento estiver ativa, acessa o site do lead!
                email_extraido = ""
                redes_sociais_extraidas = ""
                
                if tem_website == "Sim" and do_deep_scrape and website_url:
                    email_extraido, redes_sociais_extraidas = extract_contacts_from_website(context, website_url)

                results.append({
                    "Prompt": keyword,
                    "Nome da Empresa": nome,
                    "Responsável": "",
                    "Endereço": endereco,
                    "Telefone": telefone,
                    "Whatsapp": whatsapp,
                    "Email": email_extraido,
                    "Redes Sociais": redes_sociais_extraidas,
                    "Status": "A Fazer",
                    "Progressão": "1º Contato",
                    "Funcionamento": funcionamento,
                    "Tem Website": tem_website,
                    "Link Google Maps": url_maps,
                    "Observações": ""
                })
            except Exception:
                continue

        browser.close()
        
    return pd.DataFrame(results)

# --- BOTÃO DE EXECUÇÃO E DOWNLOAD ---
if st.button("🚀 Iniciar Scraping", type="primary"):
    if "BROWSERLESS_KEY" not in st.secrets or not st.secrets["BROWSERLESS_KEY"]:
        st.error("Chave de API não configurada nos Secrets do Streamlit.")
    elif not keyword.strip():
        st.warning("Por favor, digite uma palavra-chave para realizar a busca.")
    else:
        api_key = st.secrets["BROWSERLESS_KEY"]
        
        with st.spinner("Buscando e enriquecendo empresas na nuvem... Aguarde um momento."):
            df_results = scrape_maps_cloud(keyword, max_results, api_key, deep_scrape)
            
            if not df_results.empty:
                st.success(f"Encontradas {len(df_results)} empresas com sucesso!")
                st.dataframe(df_results, use_container_width=True)
                
                excel_bytes = export_to_excel(df_results)
                
                st.download_button(
                    label="📊 Baixar Planilha Excel Formatada (.xlsx)",
                    data=excel_bytes,
                    file_name=f"leads_{keyword.lower().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Nenhum resultado encontrado ou o Google Maps demorou para responder. Tente uma palavra-chave mais ampla.")
