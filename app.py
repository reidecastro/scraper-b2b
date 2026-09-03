import streamlit as st
import pandas as pd
from playwright.sync_api import sync_playwright
import time
import re
import io
import urllib.parse
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Scraper B2B Cloud - Google Meu Negócio & Maps", page_icon="📍", layout="wide")

st.title("📍 Scraper B2B - Google Meu Negócio, Google Maps e Web")
st.markdown("Busque leads B2B extraindo dados estruturados do **Google Meu Negócio (Business Profile)**, Google Maps, Websites e Diretórios Comerciais.")

# --- ENTRADA DE DADOS ---
col1, col2 = st.columns([2, 1])

with col1:
    keyword = st.text_input("Palavra-chave (Nicho + Região):", placeholder="Ex: Pizzarias no Cambui Campinas SP")

with col2:
    max_results = st.number_input("Número máximo de resultados:", min_value=5, max_value=100, value=20, step=5)

deep_search = st.checkbox("🔍 Enriquecer via Google Meu Negócio & Websites (Captura e-mails e redes sociais)", value=True)

# --- FUNÇÃO DE EXPORTAÇÃO EXCEL ---
def export_to_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads B2B"

    headers = list(df.columns)
    ws.append(headers)

    # Cabeçalho: Fundo Preto (#000000) e Texto Branco (#FFFFFF)
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

    # Validação de Dados nas Colunas Status e Progressão
    dv_status = DataValidation(type="list", formula1='"A Fazer,Ativo,Inativo"', allow_blank=True)
    dv_prog = DataValidation(type="list", formula1='"1º Contato,2º Contato,1ª Reunião,Orçamento,Contrato Finalizado"', allow_blank=True)

    ws.add_data_validation(dv_status)
    ws.add_data_validation(dv_prog)

    max_row = ws.max_row
    if max_row > 1:
        dv_status.add(f"J2:J{max_row}")
        dv_prog.add(f"K2:K{max_row}")

    # Link do Google Maps na Coluna N
    for row_idx in range(2, max_row + 1):
        map_cell = ws.cell(row=row_idx, column=14)
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

# --- BUSCA NO PERFIL DO GOOGLE MEU NEGÓCIO (GOOGLE BUSINESS PROFILE) ---
def scrape_google_business_profile(context, empresa_nome, regiao, site_oficial=""):
    emails = set()
    socials = set()
    telefone_gmb = ""
    
    # Consulta direta ao Google Search para forçar a abertura do painel do Google Meu Negócio
    query = f"{empresa_nome} {regiao}"
    search_url = f"https://www.google.com/search?q={urllib.parse.quote(query)}"
    
    try:
        page = context.new_page()
        page.set_extra_http_headers({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        })
        
        page.goto(search_url, timeout=12000, wait_until="domcontentloaded")
        time.sleep(1)
        search_html = page.content()
        
        # 1. Extrai telefone do Painel do Google Meu Negócio
        phone_match = re.search(r'\(?\d{2}\)?\s*9?\d{4}[-\s]?\d{4}', search_html)
        if phone_match:
            telefone_gmb = phone_match.group(0)

        # 2. Extrai e-mails da página do Google Meu Negócio e resultados indexados
        found_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', search_html)
        for e in found_emails:
            if not any(bad in e.lower() for bad in ['google', 'wix', 'sentry', 'domain', 'schema', 'example', '.png', '.jpg', 'png@']):
                emails.add(e.lower())

        # 3. Se tiver site cadastrado no perfil, acessa o site do restaurante para capturar e-mail e Instagram/Facebook
        if site_oficial and site_oficial.startswith("http"):
            try:
                page.goto(site_oficial, timeout=10000, wait_until="domcontentloaded")
                time.sleep(1)
                site_html = page.content()
                
                site_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', site_html)
                for e in site_emails:
                    if not any(bad in e.lower() for bad in ['wix', 'sentry', 'domain', 'schema', 'example', '.png', '.jpg', 'png@']):
                        emails.add(e.lower())
                
                links = page.query_selector_all('a[href]')
                for link in links:
                    href = link.get_attribute('href') or ""
                    if any(d in href.lower() for d in ['instagram.com', 'facebook.com', 'linkedin.com']):
                        clean_link = href.split('?')[0].rstrip('/')
                        if len(clean_link) > 15 and not clean_link.endswith(('instagram.com', 'facebook.com', 'linkedin.com')):
                            socials.add(clean_link)
            except Exception:
                pass

        page.close()
    except Exception:
        pass

    email_str = ", ".join(list(emails)[:2]) if emails else ""
    social_str = ", ".join(list(socials)[:3]) if socials else ""
    
    return email_str, social_str, telefone_gmb

# --- SCRAPER PRINCIPAL MULTI-FONTE ---
def scrape_multi_source(keyword: str, max_results: int, token: str, do_enrich: bool):
    results = []
    wss_url = f"wss://chrome.browserless.io?token={token}"
    
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(wss_url)
        context = browser.new_context()
        page = context.new_page()
        
        maps_url = f"https://www.google.com/maps/search/{urllib.parse.quote(keyword)}"
        page.goto(maps_url, wait_until="domcontentloaded")
        
        try:
            page.wait_for_selector('div[role="feed"]', timeout=15000)
        except Exception:
            browser.close()
            return pd.DataFrame()

        scrollable_div = 'div[role="feed"]'
        previous_count = 0
        
        # Rolagem para carregar todos os itens do feed
        while True:
            cards = page.query_selector_all('div[role="article"]')
            if len(cards) >= max_results or len(cards) == previous_count:
                break
            previous_count = len(cards)
            page.evaluate('(selector) => { const el = document.querySelector(selector); if (el) el.scrollBy(0, 2500); }', scrollable_div)
            time.sleep(1.8)

        cards = page.query_selector_all('div[role="article"]')[:max_results]
        
        for card in cards:
            try:
                raw_text = card.inner_text()
                lines = [l.strip() for l in raw_text.split('\n') if l.strip()]

                # 1. Nome da Empresa
                nome = card.get_attribute("aria-label")
                if not nome and lines:
                    nome = lines[0]

                # 2. Link do Google Maps / Google Meu Negócio
                link_elem = card.query_selector('a[href*="/maps/place/"]')
                url_maps = link_elem.get_attribute('href') if link_elem else ""

                # 3. Telefone ISOLADO via Regex
                phone_match = re.search(r'\(?\d{2}\)?\s*9?\d{4}[-\s]?\d{4}', raw_text)
                telefone = phone_match.group(0) if phone_match else ""

                # 4. Endereço e Categoria LIMPOS
                categoria = ""
                endereco = ""
                funcionamento = ""

                for line in lines:
                    if any(w in line.lower() for w in ["aberto", "fechado", "fecha", "abre", "24 horas"]):
                        funcionamento = line
                    elif '·' in line:
                        parts = [p.strip() for p in line.split('·') if p.strip()]
                        for part in parts:
                            if re.search(r'\d', part) and any(kw in part.lower() for kw in ["rua", "r.", "av.", "avenida", "alameda", "praça", "pç.", "jardim", "cambuí", "castelo", "campinas", "sp"]):
                                endereco = part
                            elif not re.search(r'^\d[\.,]\d', part) and part != telefone and not categoria:
                                categoria = part

                # Fallback de Endereço se não achar pelo ponto '·'
                if not endereco:
                    for line in lines:
                        if line != nome and line != telefone and line != funcionamento:
                            if any(kw in line.lower() for kw in ["rua", "r.", "av.", "avenida", "alameda", "praça", "jardim", "jd.", "campinas"]):
                                endereco = line
                                break

                # 5. Captura do Website do Perfil
                website_url = ""
                all_links = card.query_selector_all('a[href]')
                for l in all_links:
                    h = l.get_attribute('href') or ""
                    if h.startswith('http') and not any(g in h for g in ['google.com', 'google.com.br', 'ggpht.com']):
                        website_url = h
                        break

                tem_website = "Sim" if website_url else "Não"
                whatsapp = telefone if ("9" in telefone and len(re.sub(r'\D', '', telefone)) >= 10) else ""

                # 6. ENRIQUECIMENTO VIA GOOGLE MEU NEGÓCIO & WEB
                email_extraido = ""
                redes_extraidas = ""
                
                if do_enrich:
                    email_extraido, redes_extraidas, telefone_gmb = scrape_google_business_profile(context, nome, keyword, website_url)
                    if not telefone and telefone_gmb:
                        telefone = telefone_gmb
                        if "9" in telefone and len(re.sub(r'\D', '', telefone)) >= 10:
                            whatsapp = telefone

                results.append({
                    "Prompt": keyword,
                    "Nome da Empresa": nome,
                    "Categoria": categoria if categoria else "Restaurante/Serviços",
                    "Responsável": "",
                    "Endereço": endereco,
                    "Telefone": telefone,
                    "Whatsapp": whatsapp,
                    "Email": email_extraido,
                    "Redes Sociais": redes_extraidas,
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

# --- EXECUÇÃO STREAMLIT ---
if st.button("🚀 Iniciar Scraping", type="primary"):
    if "BROWSERLESS_KEY" not in st.secrets or not st.secrets["BROWSERLESS_KEY"]:
        st.error("Chave BROWSERLESS_KEY não configurada nos Secrets do Streamlit.")
    elif not keyword.strip():
        st.warning("Por favor, digite uma palavra-chave para buscar.")
    else:
        api_key = st.secrets["BROWSERLESS_KEY"]
        
        with st.spinner("Extraindo dados do Google Meu Negócio, Maps e Web... Aguarde um instante."):
            df_results = scrape_multi_source(keyword, max_results, api_key, deep_search)
            
            if not df_results.empty:
                st.success(f"Excelente! {len(df_results)} leads extraídos com sucesso.")
                st.dataframe(df_results, use_container_width=True)
                
                excel_bytes = export_to_excel(df_results)
                
                st.download_button(
                    label="📊 Baixar Planilha Excel Formatada (.xlsx)",
                    data=excel_bytes,
                    file_name=f"leads_{keyword.lower().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Nenhum resultado encontrado. Tente ajustar os termos da pesquisa.")
