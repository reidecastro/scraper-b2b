import streamlit as st
import pandas as pd
from playwright.sync_api import sync_playwright
import time
import re

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Scraper B2B Cloud", page_icon="📍", layout="wide")

st.title("📍 Scraper B2B de Empresas - Google Maps")
st.markdown("Busque leads B2B por nicho/região e baixe o relatório completo diretamente em formato CSV.")

# --- ENTRADA DE DADOS DO USUÁRIO ---
col1, col2 = st.columns([2, 1])

with col1:
    keyword = st.text_input("Palavra-chave (Nicho + Região):", placeholder="Ex: Restaurantes em Campinas SP")

with col2:
    max_results = st.number_input("Número máximo de resultados:", min_value=5, max_value=100, value=20, step=5)

# --- FUNÇÃO DE SCRAPING NA NUVEM ---
def scrape_maps_cloud(keyword: str, max_results: int, token: str):
    results = []
    wss_url = f"wss://chrome.browserless.io?token={token}"
    
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(wss_url)
        page = browser.new_page()
        
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
                # 1. Captura o Nome da Empresa (tentativa por classe de título ou aria-label)
                nome = card.get_attribute("aria-label")
                if not nome:
                    title_elem = card.query_selector('.qBF1Pd, div.fontHeadlineSmall')
                    nome = title_elem.inner_text().strip() if title_elem else "N/A"

                # 2. Captura a URL do Google Maps
                link_elem = card.query_selector('a[href*="/maps/place/"]')
                if not link_elem:
                    link_elem = card.query_selector('a')
                url_maps = link_elem.get_attribute('href') if link_elem else "N/A"

                # 3. Organização das linhas de texto do card
                lines = [line.strip() for line in card.inner_text().split("\n") if line.strip()]
                
                if nome == "N/A" and len(lines) > 0:
                    nome = lines[0]

                # Identifica a nota/avaliação (ex: 4.4(42) ou 5.0(28))
                avaliacao = "N/A"
                for line in lines:
                    if re.search(r'^\d[\.,]\d\s*\(\d+\)', line):
                        avaliacao = line
                        break

                # Filtra as linhas restantes para identificar Categoria e Endereço
                content_lines = [
                    l for l in lines 
                    if l != nome and l != avaliacao and not l.startswith("Aberto") and not l.startswith("Fechado")
                ]
                
                categoria = content_lines[0] if len(content_lines) > 0 else "N/A"
                endereco = content_lines[1] if len(content_lines) > 1 else "N/A"
                
                results.append({
                    "Keyword": keyword,
                    "Nome da Empresa": nome,
                    "Avaliação": avaliacao,
                    "Categoria": categoria,
                    "Endereço/Detalhes": endereco,
                    "URL Google Maps": url_maps
                })
            except Exception:
                continue

        browser.close()
        
    return pd.DataFrame(results)

# --- BOTÃO DE EXECUÇÃO ---
if st.button("🚀 Iniciar Scraping", type="primary"):
    if "BROWSERLESS_KEY" not in st.secrets or not st.secrets["BROWSERLESS_KEY"]:
        st.error("Chave de API não configurada nos Secrets do Streamlit.")
    elif not keyword.strip():
        st.warning("Por favor, digite uma palavra-chave para realizar a busca.")
    else:
        api_key = st.secrets["BROWSERLESS_KEY"]
        
        with st.spinner("Buscando empresas na nuvem... Aguarde um momento."):
            df_results = scrape_maps_cloud(keyword, max_results, api_key)
            
            if not df_results.empty:
                st.success(f"Encontradas {len(df_results)} empresas com sucesso!")
                st.dataframe(df_results, use_container_width=True)
                
                csv_data = df_results.to_csv(index=False, encoding="utf-8-sig")
                
                st.download_button(
                    label="📥 Baixar Resultados em CSV",
                    data=csv_data,
                    file_name=f"leads_{keyword.lower().replace(' ', '_')}.csv",
                    mime="text/csv"
                )
            else:
                st.error("Nenhum resultado encontrado ou o Google Maps demorou para responder. Tente uma palavra-chave mais ampla.")
