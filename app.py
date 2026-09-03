import streamlit as st
import pandas as pd
from playwright.sync_api import sync_playwright
import time

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
    # Conexão remota com o navegador em nuvem do Browserless.io
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
            
            # Condição de parada se não houver mais novos resultados
            if len(cards) == previous_count and len(cards) > 0:
                time.sleep(2)
                cards = page.query_selector_all('div[role="article"]')
                if len(cards) == previous_count:
                    break
                    
            previous_count = len(cards)
            page.evaluate(f'document.querySelector("{scrollable_div}").scrollBy(0, 1000)')
            time.sleep(1.5)

        cards = page.query_selector_all('div[role="article"]')[:max_results]
        
        for card in cards:
            try:
                nome = card.get_attribute("aria-label") or "N/A"
                link_elem = card.query_selector('a')
                url_maps = link_elem.get_attribute('href') if link_elem else "N/A"
                
                text_content = card.inner_text().split("\n")
                categoria = text_content[1] if len(text_content) > 1 else "N/A"
                detalhes = text_content[2] if len(text_content) > 2 else "N/A"
                
                results.append({
                    "Keyword": keyword,
                    "Nome da Empresa": nome,
                    "Categoria": categoria,
                    "Detalhes/Endereço": detalhes,
                    "URL Google Maps": url_maps
                })
            except Exception:
                continue

        browser.close()
        
    return pd.DataFrame(results)

# --- BOTÃO DE EXECUÇÃO ---
if st.button("🚀 Iniciar Scraping", type="primary"):
    # Busca a chave salva de forma segura nos Secrets do Streamlit Cloud
    if "BROWSERLESS_KEY" not in st.secrets or not st.secrets["BROWSERLESS_KEY"]:
        st.error("Chave de API não configurada nos Secrets do Streamlit. Verifique as configurações da conta.")
    elif not keyword.strip():
        st.warning("Por favor, digite uma palavra-chave para realizar a busca.")
    else:
        api_key = st.secrets["BROWSERLESS_KEY"]
        
        with st.spinner("Buscando empresas na nuvem... Aguarde um momento."):
            df_results = scrape_maps_cloud(keyword, max_results, api_key)
            
            if not df_results.empty:
                st.success(f"Encontradas {len(df_results)} empresas com sucesso!")
                
                # Exibição dos resultados em tabela na tela
                st.dataframe(df_results, use_container_width=True)
                
                # Conversão para CSV codificado para Excel
                csv_data = df_results.to_csv(index=False, encoding="utf-8-sig")
                
                # Botão de download do CSV
                st.download_button(
                    label="📥 Baixar Resultados em CSV",
                    data=csv_data,
                    file_name=f"leads_{keyword.lower().replace(' ', '_')}.csv",
                    mime="text/csv"
                )
            else:
                st.error("Não foi possível coletar os dados. Verifique se a palavra-chave é válida ou se o limite de créditos do Browserless foi atingido.")
