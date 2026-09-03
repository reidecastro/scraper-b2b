import os
import re
import csv
import time
import json
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional

from flask import Flask, render_template, request, jsonify, send_file, Response, stream_with_context
import requests
from bs4 import BeautifulSoup
import pandas as pd
from dotenv import load_dotenv

# Carrega variáveis de ambiente (.env)
load_dotenv()

# Configuração de Logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Configurações globais e diretórios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
EXPORTS_DIR = os.path.join(BASE_DIR, 'exports')

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)

# Headers padrão para requisições de scraping
DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
}


# ==========================================
# MOTOR DE SCRAPING (INTACTO E PRESERVADO)
# ==========================================

def fetch_page_content(url: str, headers: Optional[Dict[str, str]] = None, timeout: int = 15) -> Optional[str]:
    """
    Realiza a requisição HTTP GET para extração do conteúdo da página.
    """
    req_headers = headers if headers else DEFAULT_HEADERS
    try:
        response = requests.get(url, headers=req_headers, timeout=timeout)
        response.raise_for_status()
        return response.text
    except requests.RequestException as e:
        logger.error(f"Erro ao acessar a URL {url}: {str(e)}")
        return None

def parse_html_data(html_content: str, selector: str) -> List[Dict[str, Any]]:
    """
    Mapeia os elementos HTML com base nos seletores especificados.
    """
    if not html_content:
        return []
    
    soup = BeautifulSoup(html_content, 'html.parser')
    results = []
    
    elements = soup.select(selector)
    for idx, el in enumerate(elements):
        data = {
            'index': idx,
            'text': el.get_text(strip=True),
            'html': str(el),
            'attributes': el.attrs
        }
        results.append(data)
        
    return results

def run_scraping_pipeline(target_url: str, config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Executa o fluxo completo do motor de scraping preservando a estrutura exata.
    """
    logger.info(f"Iniciando engine de scraping para: {target_url}")
    start_time = time.time()
    
    html = fetch_page_content(target_url)
    if not html:
        return {
            'status': 'error',
            'message': 'Falha no download da página de destino.',
            'timestamp': datetime.now().isoformat(),
            'data': []
        }
        
    selector = config.get('selector', 'body')
    extracted_items = parse_html_data(html, selector)
    
    elapsed_time = round(time.time() - start_time, 2)
    logger.info(f"Scraping concluído em {elapsed_time}s. Itens processados: {len(extracted_items)}")
    
    return {
        'status': 'success',
        'url': target_url,
        'execution_time': elapsed_time,
        'count': len(extracted_items),
        'timestamp': datetime.now().isoformat(),
        'data': extracted_items
    }


# ==========================================
# ROTAS E ENDPOINTS DA APLICAÇÃO
# ==========================================

@app.route('/')
def index():
    """Página principal da interface web."""
    return render_template('index.html')

@app.route('/api/scrape', methods=['POST'])
def api_scrape():
    """Endpoint REST para acionar o motor de scraping."""
    payload = request.get_json() or {}
    target_url = payload.get('url')
    
    if not target_url:
        return jsonify({'status': 'error', 'message': 'A URL de destino é obrigatória.'}), 400
        
    config = {
        'selector': payload.get('selector', 'body')
    }
    
    result = run_scraping_pipeline(target_url, config)
    
    if result['status'] == 'error':
        return jsonify(result), 500
        
    return jsonify(result), 200

@app.route('/api/export/<fmt>', methods=['POST'])
def export_data(fmt: str):
    """
    Exporta os dados extraídos nos formatos JSON, CSV ou Excel.
    """
    payload = request.get_json() or {}
    data = payload.get('data', [])
    
    if not data:
        return jsonify({'status': 'error', 'message': 'Nenhum dado fornecido para exportação.'}), 400
        
    filename_base = f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    try:
        if fmt == 'json':
            file_path = os.path.join(EXPORTS_DIR, f"{filename_base}.json")
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return send_file(file_path, as_attachment=True)
            
        elif fmt == 'csv':
            file_path = os.path.join(EXPORTS_DIR, f"{filename_base}.csv")
            df = pd.DataFrame(data)
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            return send_file(file_path, as_attachment=True)
            
        elif fmt == 'excel':
            file_path = os.path.join(EXPORTS_DIR, f"{filename_base}.xlsx")
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False, engine='openpyxl')
            return send_file(file_path, as_attachment=True)
            
        else:
            return jsonify({'status': 'error', 'message': f'Formato "{fmt}" não suportado.'}), 400
            
    except Exception as e:
        logger.error(f"Erro ao exportar dados: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/health', methods=['GET'])
def health_check():
    """Endpoint de monitoramento de saúde do sistema."""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'environment': os.getenv('FLASK_ENV', 'production')
    }), 200


# ==========================================
# INICIALIZAÇÃO DO SERVIDOR
# ==========================================

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_DEBUG', 'False').lower() in ['true', '1', 't']
    
    logger.info(f"Iniciando servidor na porta {port} (debug={debug})...")
    app.run(host='0.0.0.0', port=port, debug=debug)
