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
    keyword = st.text_input("Palavra-chave (Nicho + Região):", placeholder="Ex: Restaurantes no Cambuí Campinas SP")

with col2:
    max_results = st.number_input("Número máximo de resultados:", min_value=5, max_value=100, value=20, step=5)

deep_scrape = st.checkbox("🔍 Enriquecer dados acessando o Website dos leads (Extrai E-mail e Redes Sociais se disponível)", value=True)

# --- FUNÇÃO DE EXPORTAÇÃO EXCEL FORMATADA ---
def export_to_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads B2B"

    headers = list(df.columns)
    ws.append(headers)

    # Estilo da Linha 1 (Cabeçalho): Fonte 12, Negrito, Centralizado, Fundo PRETO
    header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    for row in df.itertuples(index=False):
        ws.append(list(row))

    # Dropdowns de Validação
    dv_status = DataValidation(type="list", formula1='"A Fazer,Ativo,Inativo"', allow_blank=True)
    dv_prog = DataValidation(type="list", formula1='"1º Contato,2º Contato,1ª Reunião,Orçamento,Contrato Finalizado"', allow_blank=True)

    ws.add_data_validation(dv_status)
    ws.add_data_validation(dv_prog)

    max_row = ws.max_row
    if max_row > 1:
        # Coluna J (Status) e Coluna K (Progressão) com a inclusão de Categoria
        dv_status.add(f"J2:J{max_row}")
        dv_prog.add(f"K2:K{max_row}")

    # Formatação de Link do Google Maps
    for row_idx in range(2, max_row + 1):
        map_cell = ws.cell(row=row_idx, column=14)  # Coluna N: Link Google Maps
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

# --- EXTRAÇÃO PROFUNDA DE SITE ---
def extract_contacts_from_website(context, website_url):
    emails = set()
    socials = set()
    
    try:
        page = context.new_page()
        # Define User-Agent de navegador comum para evitar bloqueios em sites de restaurantes
        page.set_extra_http_headers({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        })
        
        page.goto(website_url, timeout=12000, wait_until="domcontentloaded")
        time.sleep(1.5)
        
        html_content = page.content()
        
        # Expressão regular para e-mails
        found_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', html_content)
        for e in found_emails:
            if not any(ext in e.lower() for ext in ['.png', '.jpg', '.jpeg', '.webp', '.svg', 'wixpress', 'sentry', 'domain']):
                emails.add(e.lower())

        # Procura links sociais no rodapé/cabeçalho
        links = page.query_selector_all('a[href]')
        for link in links:
            href = link.get_attribute('href') or ""
            if any(domain in href.lower() for domain in ['instagram.com', 'facebook.com', 'linkedin.com', 'twitter.com']):
                clean_link = href.split('?')[0].rstrip('/')
                if len(clean_link) > 15 and not clean_link.endswith(('instagram.com', 'facebook.com', 'linkedin.com')):
                    socials.add(clean_link)

        page.close()
    except Exception:
        pass
        
    email_str = ", ".join(list(emails)[:2]) if emails else ""
    social_str = ", ".join(list(socials)[:3]) if socials else ""
    return email_str, social_str

# --- SCRAPING AVANÇADO DO GOOGLE MAPS ---
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
            page.wait_for_selector('div[role="feed"]', timeout=15000)
        except Exception:
            browser.close()
            return pd.DataFrame()

        scrollable_div = 'div[role="feed"]'
        previous_count = 0
        
        # Rolagem para carregar a quantidade desejada de cards
        while len(results) < max_results:
            cards = page.query_selector_all('div[role="article"]')
            if len(cards) >= max_results or (len(cards) == previous_count and len(cards) > 0):
                break
            previous_count = len(cards)
            page.evaluate('(selector) => { const el = document.querySelector(selector); if (el) el.scrollBy(0, 1000); }', scrollable_div)
            time.sleep(1.5)

        cards = page.query_selector_all('div[role="article"]')[:max_results]
        
        for card in cards:
            try:
                # Clica no card para abrir o painel detalhado do restaurante
                card.click()
                time.sleep(1.2)
                
                # 1. Nome da Empresa
                nome = card.get_attribute("aria-label")
                if not nome:
                    title_elem = card.query_selector('.qBF1Pd, div.fontHeadlineSmall')
                    nome = title_elem.inner_text().strip() if title_elem else "N/A"

                # 2. Link do Google Maps
                link_elem = card.query_selector('a[href*="/maps/place/"]')
                url_maps = link_elem.get_attribute('href') if link_elem else ""

                # 3. Tratamento e separação da Categoria e Endereço
                lines = [line.strip() for line in card.inner_text().split("\n") if line.strip()]
                categoria = ""
                endereco = ""
                telefone = ""
                funcionamento = ""

                for line in lines:
                    # Se contém o separador '·' do Google Maps (ex: "Restaurant · R. dos Bandeirantes, 66")
                    if '·' in line:
                        parts = [p.strip() for p in line.split('·') if p.strip()]
                        if len(parts) >= 2:
                            categoria = parts[0]
                            # A última parte costuma ser o endereço/rua
                            endereco = parts[-1]
                        elif len(parts) == 1:
                            categoria = parts[0]
                    elif re.search(r'\(?\d{2}\)?\s*9?\d{4}[-\s]?\d{4}', line):
                        telefone = line
                    elif any(w in line.lower() for w in ["aberto", "fechado", "fecha", "abre", "24 horas"]):
                        funcionamento = line
                    elif any(w in line.lower() for w in ["rua", "r.", "av.", "avenida", "alameda", "praça", "bairro", "jd.", "jardim", "sp", "campinas"]):
                        if line != nome and not re.search(r'^\d[\.,]\d', line):
                            endereco = line

                # 4. Captura do Website (procura no card e no painel de detalhes aberto)
                website_url = ""
                
                # Tentativa A: Botão oficial de Website do painel do Google Maps
                web_btn = page.query_selector('a[data-item-id="authority"], a[aria-label*="website"], a[aria-label*="Website"]')
                if web_btn:
                    website_url = web_btn.get_attribute('href') or ""

                # Tentativa B: Links de "Pedir Online" / Reserva / Menu se não achar site direto
                if not website_url:
                    order_btn = page.query_selector('a[href*="http"]:not([href*="google.com"]):not([href*="ggpht"])')
                    if order_btn:
                        website_url = order_btn.get_attribute('href') or ""

                tem_website = "Sim" if website_url else "Não"
                whatsapp = telefone if ("9" in telefone and len(re.sub(r'\D', '', telefone)) >= 10) else ""

                # 5. Acessa o site para extrair E-mail e Redes Sociais
                email_extraido = ""
                redes_sociais_extraidas = ""
                
                if tem_website == "Sim" and do_deep_scrape and website_url:
                    email_extraido, redes_sociais_extraidas = extract_contacts_from_website(context, website_url)

                results.append({
                    "Prompt": keyword,
                    "Nome da Empresa": nome,
                    "Categoria": categoria,
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
