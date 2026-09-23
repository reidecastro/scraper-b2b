import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
import requests
from urllib.parse import quote_plus, urlparse
from datetime import datetime, timezone
import os

# ==============================================================================
# SCRAPER B2B + CRM (SUPABASE) + LOGIN GOOGLE + PAINEL ADMIN
# ==============================================================================

# --- Integração com Scrapling (impersonation de TLS/browser real) ---
# Requisito: pip install "scrapling[fetchers]"
# Não precisa rodar "scrapling install" (que baixa browsers) para o uso abaixo,
# pois o Fetcher.get() é puro HTTP com impersonation, sem abrir navegador.
try:
    from scrapling.fetchers import Fetcher
    SCRAPLING_AVAILABLE = True
except ImportError:
    SCRAPLING_AVAILABLE = False

def get_supabase_client():
    """
    Cliente Supabase usando a service_role key — só o backend do Streamlit
    ve essa chave (via Secrets), nunca o navegador do usuário nem o GitHub.
    service_role ignora RLS por padrão, então o controle de quem pode ver/
    editar o que e feito aqui no Python (checando o role do usuário logado),
    não por policy do Postgres.
    """
    try:
        from supabase import create_client
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_SERVICE_KEY"]
        return create_client(url, key)
    except Exception:
        return None

KEY_FILE = ".serper_key"
HUNTER_KEY_FILE = ".hunter_key"

def get_default_key(secret_name):
    """
    Busca a chave (Serper/Hunter) nesta ordem:
    1. Tabela api_keys no Supabase - editável pelo admin na aba Admin,
       sem precisar tocar no GitHub nem nos Secrets do Streamlit quando a
       chave expira ou e trocada.
    2. st.secrets - fallback inicial/bootstrap (configurado no Streamlit Cloud).
    3. String vazia, se nenhum dos dois tiver.
    """
    client = get_supabase_client()
    if client is not None:
        try:
            res = client.table("api_keys").select("key_value").eq("key_name", secret_name).execute()
            if res.data:
                return res.data[0]["key_value"]
        except Exception:
            pass
    try:
        return st.secrets[secret_name]
    except Exception:
        return ""

# Funções genéricas para carregar/salvar qualquer chave localmente (fallback
# secundário, útil rodando local sem Supabase configurado ainda)
def load_saved_key(key_file, secret_name):
    if os.path.exists(key_file):
        try:
            with open(key_file, "r") as f:
                key = f.read().strip()
                if key:
                    return key
        except Exception:
            pass
    return get_default_key(secret_name)

def save_key(key_file, key_str):
    try:
        with open(key_file, "w") as f:
            f.write(key_str.strip())
    except Exception as e:
        st.error(f"Erro ao salvar a chave: {e}")

st.set_page_config(
    page_title="Gerador de Leads B2B - Prospecção Avançada",
    page_icon="🎯",
    layout="wide"
)

# Estilização para Dark Mode
st.markdown("""
<style>
    .main-header { font-size: 2.2rem; color: #60A5FA; font-weight: 700; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1.1rem; color: #D1D5DB; margin-bottom: 1.5rem; }
    .stButton>button { background-color: #2563EB; color: white; border-radius: 6px; padding: 0.5rem 1.5rem; font-weight: 600; border: none; }
    .stButton>button:hover { background-color: #1D4ED8; color: white; }
    .footer-box {
        background-color: #1E293B;
        border-top: 2px solid #3B82F6;
        padding: 15px;
        border-radius: 8px;
        margin-top: 40px;
        color: #F3F4F6;
    }
</style>
""", unsafe_allow_html=True)

# --- GATE DE LOGIN (Google OIDC nativo do Streamlit) ---
# Requer o bloco [auth] configurado em Secrets (client_id, client_secret,
# redirect_uri, cookie_secret, server_metadata_url do Google).
if not st.user.is_logged_in:
    st.markdown('<div class="main-header">🐺 WOLF Terceirizações</div>', unsafe_allow_html=True)
    st.markdown("Faça login com sua conta Google para acessar o CRM.")
    st.button("🔐 Entrar com Google", on_click=st.login)
    st.stop()

def get_or_create_profile(email, nome):
    """
    Busca o perfil (role) do usuário logado. Se é o primeiro login dele,
    cria o registro com role padrão 'user' — exceto reidecastro@gmail.com,
    que já nasce 'admin' (seed feito direto no banco). Atualiza last_login
    a cada acesso.
    """
    client = get_supabase_client()
    if client is None:
        return "user"
    try:
        res = client.table("profiles").select("role").eq("email", email).execute()
        agora = datetime.now(timezone.utc).isoformat()
        if res.data:
            client.table("profiles").update({"last_login": agora}).eq("email", email).execute()
            return res.data[0]["role"]
        else:
            client.table("profiles").insert({
                "email": email, "nome": nome, "role": "user", "last_login": agora
            }).execute()
            return "user"
    except Exception:
        return "user"

if 'user_role' not in st.session_state:
    st.session_state['user_email'] = st.user.email
    st.session_state['user_name'] = getattr(st.user, "name", st.user.email)
    st.session_state['user_role'] = get_or_create_profile(st.user.email, st.session_state['user_name'])

# Inicializa o estado das chaves API (agora priorizando o banco, ver get_default_key)
if 'serper_api_key' not in st.session_state:
    st.session_state['serper_api_key'] = load_saved_key(KEY_FILE, "SERPER_API_KEY")
if 'hunter_api_key' not in st.session_state:
    st.session_state['hunter_api_key'] = load_saved_key(HUNTER_KEY_FILE, "HUNTER_API_KEY")

# Função para checar créditos restantes na API Serper (Usando POST)
# Nota: removida a função get_serper_credits() que existia aqui — o endpoint
# "https://google.serper.dev/credits" nunca foi um endpoint real e documentado
# da API (confirmado: a lista oficial de endpoints do Serper cobre search,
# images, videos, places, maps, reviews, news, shopping, lens, patents,
# autocomplete e scrape — sem nada de conta/créditos). Por isso a consulta
# sempre falhava, com qualquer chave. Ver rodapé do app para o substituto:
# contagem real de créditos gastos por extração, calculada pelo próprio app.

# Função para checar buscas restantes na API Hunter.io
def get_hunter_credits(api_key):
    if not api_key:
        return None
    try:
        res = requests.get(f"https://api.hunter.io/v2/account?api_key={api_key}", timeout=4)
        if res.status_code == 200:
            data = res.json()
            searches = data.get("data", {}).get("requests", {}).get("searches", {})
            available = searches.get("available")
            used = searches.get("used")
            if available is not None:
                return available, used
    except Exception:
        pass
    return None, None

# --- BARRA LATERAL ---
with st.sidebar:
    if os.path.exists("wolflogo.png"):
        st.image("wolflogo.png", use_container_width=True)

    # --- Usuário logado + logout ---
    role_badge = "👑 Admin" if st.session_state['user_role'] == 'admin' else "👤 Usuário"
    st.markdown(f"**{st.session_state['user_name']}**  \n{role_badge}")
    st.button("🚪 Sair", on_click=st.logout, use_container_width=True)
    st.caption("🔑 Chaves de API gerenciadas pelo admin (aba ⚙️ Admin) — não editáveis aqui.")

    st.markdown("---")
    st.header("⚙️ Configurações da Busca")
    termo_busca = st.text_input("Termo de Busca e Bairro", value="Pizzarias Campinas SP Bairro Castelo")
    qtd_resultados = st.number_input("Quantidade de Resultados", min_value=1, max_value=20, value=10, step=1)

    st.markdown("---")
    st.header("🔍 Opções de Enriquecimento")
    enriquecer_emails = st.checkbox("Buscar E-mails nos Sites", value=True)
    enriquecer_redes = st.checkbox("Buscar Redes Sociais (Instagram/FB)", value=True)
    buscar_socios = st.checkbox("Buscar CNPJ & Sócios (BrasilAPI)", value=True)
    usar_hunter = st.checkbox(
        "Completar E-mail via Hunter.io (só quando faltar)",
        value=False,
        help="Consome cota do Hunter (25 buscas/mês no plano free) — só é chamado pros leads que ainda estão sem e-mail depois do site e da busca Serper. Desligado por padrão pra você não gastar cota sem querer."
    )

    # --- Estimativa de créditos Serper antes de rodar ---
    # Mínimo: só a busca Places (1 crédito, cobre o lote inteiro).
    # Máximo: Places + 1 busca fallback por lead (se faltar contato) +
    # 1 busca de CNPJ por lead (se a opção estiver marcada).
    est_min = 1
    est_max = 1 + qtd_resultados + (qtd_resultados if buscar_socios else 0)
    st.caption(f"💰 Estimativa de custo Serper: **{est_min} a {est_max} créditos** nesta extração (varia conforme quantos leads já vêm com contato direto do Google Places).")
    
    st.markdown("<br>", unsafe_allow_html=True)
    btn_extrair = st.button("🚀 Iniciar Extração de Leads", use_container_width=True)

# --- CABEÇALHO PRINCIPAL ---
st.markdown('<div class="main-header">🎯 Gerador de Leads B2B - Prospecção Avançada</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Extração de dados de estabelecimentos, empresas e profissionais.</div>', unsafe_allow_html=True)

def clean_and_format_phone(phone_str):
    if not phone_str or pd.isna(phone_str):
        return "", ""
    digits = re.sub(r'\D', '', str(phone_str))
    
    if len(digits) > 11 and digits.startswith("55"):
        digits = digits[2:]
        
    if len(digits) < 8:
        return str(phone_str), ""
    
    formatted_phone = str(phone_str)
    whatsapp = f"https://wa.me/55{digits}" if len(digits) in [10, 11] else ""
    return formatted_phone, whatsapp

def _get_pool_keys(provider):
    """Chaves ativas do pool pra um provider ('serper' ou 'hunter'), menos
    usada recentemente primeiro (round-robin simples)."""
    client = get_supabase_client()
    if client is None:
        return []
    try:
        res = (client.table("api_pool")
               .select("*")
               .eq("provider", provider)
               .eq("status", "active")
               .order("last_used_at", desc=False)
               .execute())
        return res.data
    except Exception:
        return []

def _marcar_pool_key_usada(pool_id):
    client = get_supabase_client()
    if client is None or pool_id is None:
        return
    try:
        client.table("api_pool").update({
            "last_used_at": datetime.now(timezone.utc).isoformat()
        }).eq("id", pool_id).execute()
    except Exception:
        pass

def _marcar_pool_key_esgotada(pool_id, motivo):
    client = get_supabase_client()
    if client is None or pool_id is None:
        return
    try:
        client.table("api_pool").update({
            "status": "exhausted",
            "last_error": motivo,
        }).eq("id", pool_id).execute()
    except Exception:
        pass

def _capturar_headers_debug(provider, headers_resposta):
    """Guarda um snapshot dos headers da última resposta em session_state,
    só pra investigação manual (visível no painel Admin > Debug). Não é
    usado por nenhuma lógica de negócio."""
    try:
        st.session_state[f'_debug_headers_{provider}'] = dict(headers_resposta)
    except Exception:
        pass

def serper_post(url, payload, credit_counter=None, timeout=10):
    """
    POST autenticado na Serper com hot-swap automático de chave: se o pool
    (tabela api_pool) tiver chaves cadastradas, tenta cada uma em ordem até
    conseguir. Uma chave é marcada como 'exhausted' e pulada nas próximas
    tentativas quando responde 401/403/429 ou "Out of credits". Timeout
    tenta de novo com o dobro do tempo NA MESMA chave antes de trocar (rede
    lenta não é motivo pra descartar uma chave boa). Sem pool configurado,
    cai pra chave única legada (session_state/secrets), mantendo tudo
    funcionando como antes. Retorna o Response ou None se tudo falhar.
    """
    tentativas = list(_get_pool_keys("serper"))
    usando_pool = bool(tentativas)

    if not tentativas:
        chave_unica = st.session_state.get('serper_api_key', '')
        tentativas = [{"id": None, "api_key": chave_unica}] if chave_unica else []

    for chave_info in tentativas:
        api_key = chave_info.get("api_key")
        if not api_key:
            continue
        headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}

        for tentativa_timeout in (timeout, timeout * 2):
            try:
                if credit_counter is not None:
                    credit_counter[0] += 1
                res = requests.post(url, headers=headers, json=payload, timeout=tentativa_timeout)
                _capturar_headers_debug("serper", res.headers)

                if res.status_code in (401, 403, 429) or "Out of credits" in res.text:
                    if usando_pool:
                        _marcar_pool_key_esgotada(chave_info["id"], f"HTTP {res.status_code}")
                    break  # essa chave não serve — troca pra próxima do pool

                if usando_pool:
                    _marcar_pool_key_usada(chave_info["id"])
                return res
            except requests.exceptions.Timeout:
                continue  # mesma chave, timeout maior
            except requests.exceptions.RequestException:
                break  # erro de conexão — troca de chave

    return None

def hunter_get(url, params, timeout=6):
    """Equivalente ao serper_post, mas pra GET autenticado na Hunter.io
    (sem retry de timeout, já que domain-search é rápido por natureza)."""
    tentativas = list(_get_pool_keys("hunter"))
    usando_pool = bool(tentativas)

    if not tentativas:
        chave_unica = st.session_state.get('hunter_api_key', '')
        tentativas = [{"id": None, "api_key": chave_unica}] if chave_unica else []

    for chave_info in tentativas:
        api_key = chave_info.get("api_key")
        if not api_key:
            continue
        p = dict(params)
        p["api_key"] = api_key
        try:
            res = requests.get(url, params=p, timeout=timeout)
            _capturar_headers_debug("hunter", res.headers)

            if res.status_code in (401, 403, 429):
                if usando_pool:
                    _marcar_pool_key_esgotada(chave_info["id"], f"HTTP {res.status_code}")
                continue

            if usando_pool:
                _marcar_pool_key_usada(chave_info["id"])
            return res
        except requests.exceptions.RequestException:
            continue

    return None

def extract_address_from_item(item, company_name="", credit_counter=None):
    addr = item.get("address") or item.get("formattedAddress") or item.get("vicinity") or item.get("street") or ""
    
    if isinstance(addr, dict):
        street = addr.get("street", "")
        city = addr.get("city", "")
        state = addr.get("state", "")
        addr = f"{street}, {city} - {state}".strip(", -")
    elif isinstance(addr, list):
        addr = ", ".join([str(x) for x in addr])

    if not addr and company_name:
        try:
            url = "https://google.serper.dev/search"
            payload = {"q": f"{company_name} endereco localizacao", "gl": "br", "hl": "pt-br", "num": 2}
            res = serper_post(url, payload, credit_counter, timeout=3)
            if res is not None and res.status_code == 200:
                data = res.json()
                for og in data.get("organic", []):
                    snippet = og.get("snippet", "")
                    match = re.search(r'(Rua|R\.|Avenida|Av\.|Praça|Alameda|Rodovia)[^,\.]+,[^\.]+', snippet, re.IGNORECASE)
                    if match:
                        addr = match.group(0)
                        break
        except Exception:
            pass

    return str(addr).strip()

def fetch_cnpj_and_partners(company_name, city_or_address="", credit_counter=None):
    cnpj_clean = ""
    razao_social = ""
    socios_names = []
    
    location_hint = city_or_address.split("-")[0].strip() if city_or_address else ""
    query = f"{company_name} {location_hint} cnpj brasilapi"
    url = "https://google.serper.dev/search"
    payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
    
    try:
        res = serper_post(url, payload, credit_counter, timeout=4)
        if res is not None and res.status_code == 200:
            data = res.json()
            text_block = ""
            for item in data.get("organic", []):
                text_block += " " + item.get("snippet", "") + " " + item.get("title", "")
            
            cnpjs = re.findall(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b', text_block)
            if cnpjs:
                cnpj_clean = re.sub(r'\D', '', cnpjs[0])
    except Exception:
        pass

    if cnpj_clean:
        try:
            api_url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_clean}"
            r = requests.get(api_url, timeout=4)
            if r.status_code == 200:
                cnpj_data = r.json()
                razao_social = cnpj_data.get("razao_social", "")
                qsa = cnpj_data.get("qsa", [])
                for socio in qsa:
                    nome = socio.get("nome_socio", "")
                    qualificacao = socio.get("qualificacao_socio", "")
                    if nome:
                        socios_names.append(f"{nome} ({qualificacao})" if qualificacao else nome)
        except Exception:
            pass

    cnpj_formatted = f"{cnpj_clean[:2]}.{cnpj_clean[2:5]}.{cnpj_clean[5:8]}/{cnpj_clean[8:12]}-{cnpj_clean[12:]}" if len(cnpj_clean) == 14 else ""
    socios_str = ", ".join(socios_names) if socios_names else ""
    return cnpj_formatted, razao_social, socios_str

def fallback_search_phone_email(company_name, address, api_key, credit_counter=None):
    phone, email, social = "", "", ""
    query = f"{company_name} {address} telefone contato email"
    url = "https://google.serper.dev/search"
    payload = {"q": query, "gl": "br", "hl": "pt-br", "num": 3}
    
    try:
        res = serper_post(url, payload, credit_counter, timeout=5)
        if res is not None and res.status_code == 200:
            data = res.json()
            text_block = ""
            
            for item in data.get("organic", []):
                text_block += " " + item.get("title", "") + " " + item.get("snippet", "")
                link = item.get("link", "")
                if ("instagram.com" in link or "facebook.com" in link) and not social:
                    social = link

            phone_matches = re.findall(r'(?:\(?\d{2}\)?\s*)?(?:9?\d{4}[-\s]?\d{4})', text_block)
            if phone_matches:
                for match in phone_matches:
                    clean_m = re.sub(r'\D', '', match)
                    if len(clean_m) in [10, 11]:
                        phone = match
                        break

            email_matches = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text_block)
            if email_matches:
                valid_e = [e for e in email_matches if not e.endswith(('.png', '.jpg', '.webp', '.js', '.css'))]
                if valid_e:
                    email = valid_e[0]
    except Exception:
        pass

    return phone, email, social

def _fetch_page_html(url, timeout=6):
    """
    Busca o HTML de UMA página, tentando primeiro Scrapling (impersonation
    de navegador real) e caindo pro requests puro se falhar/indisponível.
    Retorna (texto_html, metodo) — metodo é "Scrapling", "Requests" ou "".
    """
    text, metodo = "", ""

    if SCRAPLING_AVAILABLE:
        try:
            page = Fetcher.get(url, impersonate='chrome', stealthy_headers=True, timeout=timeout)
            if page.status == 200:
                text, metodo = page.html_content, "Scrapling"
        except Exception:
            pass

    if not text:
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            response = requests.get(url, headers=headers, timeout=min(timeout, 5))
            if response.status_code == 200:
                text, metodo = response.text, "Requests"
        except Exception:
            pass

    return text, metodo

def _extract_email_social(text):
    """Aplica os regex de e-mail/rede social num HTML já baixado."""
    email, social = "", ""
    if not text:
        return email, social
    emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
    if emails:
        valid_emails = [e for e in emails if not e.endswith(('.png', '.jpg', '.webp', '.js', '.css', '.svg'))]
        if valid_emails:
            email = valid_emails[0]
    socials = re.findall(r'https?://(?:www\.)?(?:instagram\.com|facebook\.com)/[a-zA-Z0-9_.-]+', text)
    if socials:
        social = socials[0]
    return email, social

# Páginas comuns onde PMEs brasileiras costumam publicar e-mail/telefone
# quando não está no rodapé da home. Testadas em ordem, para na primeira
# que resolver o e-mail (evita bater em N páginas à toa).
CAMINHOS_CONTATO_COMUNS = ["/contato", "/contato-nos", "/fale-conosco", "/contact", "/contact-us"]

def scrape_website_details(website_url):
    """
    Extrai e-mail e rede social do site do lead.

    Tenta a home primeiro. Se a home responder mas não tiver e-mail (comum:
    o e-mail costuma morar numa página /contato, não no rodapé), tenta até
    3 caminhos comuns de página de contato antes de desistir. Cada tentativa
    usa Scrapling com fallback pra requests, então bloqueio por Cloudflare/WAF
    é contornado em qualquer uma das páginas testadas, não só na home.

    Retorna (email, social, metodo) - "metodo" indica qual caminho
    efetivamente conseguiu os dados, incluindo se veio da home ou de uma
    página de contato secundária (ex: "Scrapling", "Scrapling + /contato").
    """
    email, social = "", ""
    metodo = ""
    if not website_url or not str(website_url).startswith("http"):
        return email, social, metodo

    text, metodo = _fetch_page_html(website_url)
    if not text:
        return email, social, ""  # nem a home abriu — nenhum bypass funcionou

    email, social = _extract_email_social(text)

    # Home abriu, mas sem e-mail visível → tenta páginas de contato comuns
    if not email:
        base = website_url.rstrip('/')
        for caminho in CAMINHOS_CONTATO_COMUNS:
            sub_text, sub_metodo = _fetch_page_html(base + caminho, timeout=5)
            if sub_text:
                sub_email, sub_social = _extract_email_social(sub_text)
                if sub_email:
                    email = sub_email
                    metodo = f"{metodo} + {caminho}"
                if not social and sub_social:
                    social = sub_social
                if email:
                    break

    return email, social, metodo

def hunter_domain_search(website_url, api_key):
    """
    Terceira camada de enriquecimento (só chamada quando site + fallback
    Serper não acharam e-mail). Usa o endpoint Domain Search do Hunter.io,
    que busca e-mails associados ao domínio via padrões conhecidos e
    indexação própria — não depende do HTML da página em si, então
    funciona mesmo quando o site nunca publicou e-mail em texto plano.

    Passa por hunter_get, que faz hot-swap automático entre as chaves do
    pool (aba Admin) se a atual estiver esgotada/sem autorização — o
    parâmetro api_key vira só um fallback legado, sem pool configurado.

    Retorna o e-mail de maior confiança encontrado, ou "" se não achar
    nada ou nenhuma chave disponível conseguir responder.
    """
    if not website_url:
        return ""

    try:
        domain = urlparse(website_url).netloc.replace("www.", "")
        if not domain:
            return ""

        url = "https://api.hunter.io/v2/domain-search"
        params = {"domain": domain, "limit": 3}
        res = hunter_get(url, params, timeout=6)

        if res is not None and res.status_code == 200:
            data = res.json()
            emails = data.get("data", {}).get("emails", [])
            if emails:
                # Prioriza e-mails de tipo "generic" (contato@, info@) sobre
                # pessoais, e ordena pela confiança que o Hunter atribui
                emails_sorted = sorted(emails, key=lambda e: e.get("confidence", 0), reverse=True)
                return emails_sorted[0].get("value", "")
    except Exception:
        pass

    return ""

def calcular_score(row):
    """
    Score simples de 0-100 pra priorizar leads no CRM. Pontua presença de
    canais de contato e qualidade da avaliação no Google — não é machine
    learning, é uma heurística direta e fácil de ajustar depois.
    """
    score = 40
    if row.get("Telefone"): score += 10
    if row.get("Whatsapp"): score += 10
    if row.get("Email"): score += 15
    if row.get("Redes Sociais"): score += 5
    if row.get("Website"): score += 10
    try:
        nota = float(row.get("Nota Google") or 0)
        if nota >= 4.5: score += 10
        elif nota >= 4.0: score += 5
    except (ValueError, TypeError):
        pass
    return max(0, min(score, 100))

def _valor_ou_none(v):
    """Normaliza valores vazios/NaN do pandas para None (NULL no Postgres)."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s and s.lower() != "nan" else None

def salvar_leads_supabase(df_leads):
    """
    Insere os leads extraídos no CRM (tabela leads). Dedupe: pula o lead se
    já existir por CNPJ, ou por nome+endereço quando não há CNPJ — assim
    rodar a mesma busca de novo não duplica nem reseta o pipeline/notas que
    você já tenha preenchido manualmente pra esse lead.
    Retorna (novos, ja_existentes, erro).
    """
    client = get_supabase_client()
    if client is None:
        return 0, 0, "Supabase não configurado (verifique SUPABASE_URL e SUPABASE_SERVICE_KEY nos Secrets)."

    novos, existentes = 0, 0

    for _, row in df_leads.iterrows():
        cnpj = _valor_ou_none(row.get("CNPJ"))
        nome = row.get("Nome da Empresa") or ""
        endereco = row.get("Endereço") or ""

        try:
            if cnpj:
                existing = client.table("leads").select("id").eq("cnpj", cnpj).execute()
            else:
                existing = client.table("leads").select("id").ilike("nome_empresa", nome).ilike("endereco", endereco).execute()

            if existing.data:
                existentes += 1
                continue

            nota_google = row.get("Nota Google")
            total_avaliacoes = row.get("Total Avaliações")

            client.table("leads").insert({
                "nome_empresa": nome,
                "razao_social": _valor_ou_none(row.get("Razão Social")),
                "cnpj": cnpj,
                "socios": _valor_ou_none(row.get("Sócios / Decisores")),
                "categoria": _valor_ou_none(row.get("Categoria")),
                "endereco": _valor_ou_none(endereco),
                "telefone": _valor_ou_none(row.get("Telefone")),
                "whatsapp": _valor_ou_none(row.get("Whatsapp")),
                "email": _valor_ou_none(row.get("Email")),
                "redes_sociais": _valor_ou_none(row.get("Redes Sociais")),
                "website": _valor_ou_none(row.get("Website")),
                "gmaps_link": _valor_ou_none(row.get("Link Google Maps")),
                "nota_google": float(nota_google) if _valor_ou_none(nota_google) else None,
                "total_avaliacoes": int(total_avaliacoes) if _valor_ou_none(total_avaliacoes) else None,
                "origem_busca": _valor_ou_none(row.get("Prompt")),
                "fonte_contato": _valor_ou_none(row.get("Fonte do Contato")),
                "score": calcular_score(row),
            }).execute()
            novos += 1
        except Exception:
            continue

    return novos, existentes, None

def carregar_leads_crm():
    """Lê todos os leads salvos no CRM, mais recentes primeiro."""
    client = get_supabase_client()
    if client is None:
        return None
    try:
        res = client.table("leads").select("*").order("created_at", desc=True).execute()
        return res.data
    except Exception:
        return None

def atualizar_lead(lead_id, updates):
    client = get_supabase_client()
    if client is None:
        return False
    try:
        client.table("leads").update(updates).eq("id", lead_id).execute()
        return True
    except Exception:
        return False

# --- Funções exclusivas do painel Admin ---
def listar_api_keys_admin():
    client = get_supabase_client()
    if client is None:
        return []
    try:
        return client.table("api_keys").select("*").execute().data
    except Exception:
        return []

def salvar_api_key_admin(key_name, key_value, admin_email):
    client = get_supabase_client()
    if client is None:
        return False
    try:
        client.table("api_keys").upsert({
            "key_name": key_name,
            "key_value": key_value,
            "updated_by": admin_email,
        }).execute()
        return True
    except Exception:
        return False

# --- Gerenciamento do Pool de Chaves (rotação automática) ---
def listar_pool_admin(provider=None):
    """Todas as chaves do pool, opcionalmente filtradas por provider."""
    client = get_supabase_client()
    if client is None:
        return []
    try:
        q = client.table("api_pool").select("*").order("provider").order("created_at")
        if provider:
            q = q.eq("provider", provider)
        return q.execute().data
    except Exception:
        return []

def adicionar_chave_pool(provider, api_key, label=""):
    client = get_supabase_client()
    if client is None:
        return False, "Supabase não configurado."
    try:
        client.table("api_pool").insert({
            "provider": provider,
            "api_key": api_key,
            "label": label or None,
            "status": "active",
        }).execute()
        return True, None
    except Exception as e:
        # Erro mais comum aqui: unique(provider, api_key) — chave duplicada
        return False, str(e)

def atualizar_status_pool(pool_id, novo_status):
    client = get_supabase_client()
    if client is None:
        return False
    try:
        client.table("api_pool").update({"status": novo_status, "last_error": None}).eq("id", pool_id).execute()
        return True
    except Exception:
        return False

def remover_chave_pool(pool_id):
    client = get_supabase_client()
    if client is None:
        return False
    try:
        client.table("api_pool").delete().eq("id", pool_id).execute()
        return True
    except Exception:
        return False

def listar_usuarios_admin():
    client = get_supabase_client()
    if client is None:
        return []
    try:
        return client.table("profiles").select("*").order("created_at").execute().data
    except Exception:
        return []

def atualizar_role_usuario(email, novo_role):
    client = get_supabase_client()
    if client is None:
        return False
    try:
        client.table("profiles").update({"role": novo_role}).eq("email", email).execute()
        return True
    except Exception:
        return False

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
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Whatsapp" and val_str and val_str.startswith("http"):
                cell.value = f'=HYPERLINK("{val_str}", "Abrir WhatsApp")'
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name == "Redes Sociais" and val_str and val_str.startswith("http"):
                cell.value = f'=HYPERLINK("{val_str}", "Acessar Perfil")'
                cell.hyperlink = val_str
                cell.font = link_font
                cell.alignment = align_center
            elif col_name in ["Status", "Progressão", "Tem Website", "Nota Google", "Total Avaliações", "CNPJ", "Fonte do Contato"]:
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
        
    max_row = len(df) + 50

    if "Status" in columns:
        status_col_idx = columns.index("Status") + 1
        col_letter = get_column_letter(status_col_idx)
        dv_status = DataValidation(type="list", formula1='"A Fazer, Em Andamento, Concluído, Cancelado"', allow_blank=True)
        ws.add_data_validation(dv_status)
        dv_status.add(f"{col_letter}2:{col_letter}{max_row}")

    if "Progressão" in columns:
        progress_col_idx = columns.index("Progressão") + 1
        col_letter = get_column_letter(progress_col_idx)
        dv_progress = DataValidation(type="list", formula1='"1º Contato, Em Negociação, Proposta Enviada, Fechado, Perdido"', allow_blank=True)
        ws.add_data_validation(dv_progress)
        dv_progress.add(f"{col_letter}2:{col_letter}{max_row}")

    ws.freeze_panes = 'A2'
    wb.save(filename)
    return filename

def run_places_extraction(query, api_key, max_results=10, email_opt=True, redes_opt=True, socios_opt=True, hunter_opt=False, hunter_key=""):
    extracted_data = []
    credit_counter = [0]
    
    url = "https://google.serper.dev/places"
    payload = {"q": query, "gl": "br", "hl": "pt-br"}
    
    try:
        # serper_post já cuida do retry em timeout (mesma chave, timeout
        # dobrado) e do hot-swap pra próxima chave do pool em caso de
        # esgotamento/erro — nenhuma dessas lógicas precisa mais estar aqui.
        response = serper_post(url, payload, credit_counter, timeout=10)

        if response is None:
            st.error("❌ Não foi possível completar a busca — todas as chaves Serper disponíveis falharam (esgotadas, sem autorização, ou fora do ar). Confira o pool de chaves na aba ⚙️ Admin.")
            return pd.DataFrame(), credit_counter[0]
        
        if response.status_code == 402 or "Out of credits" in response.text:
            st.error("❌ Os créditos da chave Serper atual ACABARAM! Altere a chave na barra lateral ou crie uma nova conta no Serper.")
            return pd.DataFrame(), credit_counter[0]
            
        data = response.json()
        
        places = data.get("places", [])
        for item in places[:max_results]:
            company_name = item.get("title", "")
            
            address = extract_address_from_item(item, company_name, credit_counter)
            
            phone_raw = item.get("phoneNumber") or item.get("phone") or ""
            category = item.get("category", "Comércio Local / Empresa")
            rating = item.get("rating", "")
            rating_count = item.get("ratingCount", "")
            website_url = item.get("website", "")
            
            latitude = item.get("latitude", "")
            longitude = item.get("longitude", "")
            if latitude and longitude:
                gmaps_link = f"https://www.google.com/maps/search/?api=1&query={latitude},{longitude}"
            else:
                gmaps_link = f"https://www.google.com/maps/search/{quote_plus(company_name + ' ' + address)}"
            
            real_email = ""
            real_social = ""
            cnpj = ""
            razao_social = ""
            socios = ""
            site_metodo = ""
            fallback_contribuiu = False
            
            if website_url and (email_opt or redes_opt):
                real_email, real_social, site_metodo = scrape_website_details(website_url)
            
            if not phone_raw or not real_email:
                fb_phone, fb_email, fb_social = fallback_search_phone_email(company_name, address, api_key, credit_counter)
                if not phone_raw and fb_phone:
                    phone_raw = fb_phone
                if not real_email and fb_email:
                    real_email = fb_email
                    fallback_contribuiu = True
                if not real_social and fb_social:
                    real_social = fb_social
                    fallback_contribuiu = True

            # --- 3ª camada: Hunter.io, só entra se ainda faltar e-mail ---
            hunter_contribuiu = False
            if hunter_opt and hunter_key and not real_email and website_url:
                hunter_email = hunter_domain_search(website_url, hunter_key)
                if hunter_email:
                    real_email = hunter_email
                    hunter_contribuiu = True

            # --- Monta a rastreabilidade de onde veio o e-mail/rede social ---
            fonte_partes = []
            if site_metodo:
                fonte_partes.append(f"Site Direto ({site_metodo})")
            if fallback_contribuiu:
                fonte_partes.append("Busca Fallback (Serper)")
            if hunter_contribuiu:
                fonte_partes.append("Hunter.io")
            fonte_contato = " + ".join(fonte_partes) if fonte_partes else "Não encontrado"

            if socios_opt:
                cnpj, razao_social, socios = fetch_cnpj_and_partners(company_name, address, credit_counter)

            formatted_phone, wa_link = clean_and_format_phone(phone_raw)

            extracted_data.append({
                "Prompt": query,
                "Nome da Empresa": company_name,
                "Razão Social": razao_social,
                "CNPJ": cnpj,
                "Sócios / Decisores": socios if socios else "Não identificado",
                "Categoria": category,
                "Endereço": address,
                "Telefone": formatted_phone,
                "Whatsapp": wa_link,
                "Email": real_email,
                "Redes Sociais": real_social,
                "Website": website_url,
                "Fonte do Contato": fonte_contato,
                "Nota Google": rating,
                "Total Avaliações": rating_count,
                "Status": "A Fazer",
                "Progressão": "1º Contato",
                "Tem Website": "Sim" if website_url else "Não",
                "Link Google Maps": gmaps_link,
                "Observações": f"Site: {website_url}" if website_url else "Sem site oficial"
            })
            
    except requests.exceptions.Timeout:
        st.error("❌ A Serper não respondeu a tempo, mesmo com o retry automático. Isso costuma ser instabilidade pontual do lado deles — tente rodar a busca de novo em alguns instantes, ou torne o termo um pouco mais amplo (ex: tire o bairro) se persistir.")
    except Exception as e:
        st.error(f"Erro na requisição: {e}")

    return pd.DataFrame(extracted_data), credit_counter[0]

if btn_extrair:
    active_key = st.session_state['serper_api_key']
    with st.spinner(f"Extraindo leads e analisando dados para '{termo_busca}'..."):
        df_leads, creditos_usados = run_places_extraction(
            termo_busca, active_key, qtd_resultados, 
            enriquecer_emails, enriquecer_redes, buscar_socios,
            usar_hunter, st.session_state['hunter_api_key']
        )
        
        if not df_leads.empty:
            filename = "leads_extraidos.xlsx"
            create_excel_report(df_leads, filename)
            
            st.session_state['df_leads'] = df_leads
            st.session_state['filename'] = filename
            st.session_state['creditos_usados'] = creditos_usados
            st.success(f"✅ Sucesso! {len(df_leads)} empresas extraídas com sucesso. (≈{creditos_usados} créditos Serper usados nesta extração)")

            # --- Salva automaticamente no CRM (com dedupe) ---
            novos, existentes, erro_crm = salvar_leads_supabase(df_leads)
            if erro_crm:
                st.warning(f"⚠️ CRM: {erro_crm}")
            else:
                st.info(f"💾 CRM atualizado: {novos} lead(s) novo(s) salvo(s), {existentes} já existiam (pipeline preservado).")

# --- APRESENTAÇÃO DOS RESULTADOS, SCRIPTS, CRM E ADMIN ---
tab_labels = ["📋 Tabela de Leads", "💬 Gerador de Script de Vendas", "🗂️ Meus Leads (CRM)"]
is_admin = st.session_state['user_role'] == 'admin'
if is_admin:
    tab_labels.append("⚙️ Admin")

tabs = st.tabs(tab_labels)
tab1, tab2, tab_crm = tabs[0], tabs[1], tabs[2]
tab_admin = tabs[3] if is_admin else None

tem_resultado_recente = 'df_leads' in st.session_state and not st.session_state['df_leads'].empty

with tab1:
    if tem_resultado_recente:
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
    else:
        st.info("Rode uma extração na barra lateral para ver os resultados aqui.")

with tab2:
    if tem_resultado_recente:
        df_leads = st.session_state['df_leads']

        st.subheader("✍️ Gerador de Abordagem Comercial (Copywriting)")
        
        empresa_selecionada = st.selectbox(
            "Selecione uma empresa da lista para gerar o script:",
            options=df_leads["Nome da Empresa"].tolist()
        )
        
        lead_info = df_leads[df_leads["Nome da Empresa"] == empresa_selecionada].iloc[0]
        
        socio_nome = lead_info["Sócios / Decisores"].split("(")[0].strip() if lead_info["Sócios / Decisores"] != "Não identificado" else "Responsável"
        categoria = lead_info["Categoria"]
        endereco = lead_info["Endereço"]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📱 Script para WhatsApp (Mensagem Curta)")
            script_wa = f"""Olá, {socio_nome}! Tudo bem?

Vi a {empresa_selecionada} aqui no Google e notei que vocês são referência em {categoria} na região.

Estou entrando em contato pois ajudamos empresas do seu segmento a aumentarem o volume de clientes e otimizarem a presença digital.

Você teria 5 minutos nesta semana para trocarmos uma ideia rápida sobre como podemos ajudar a {empresa_selecionada}?

Abraços!"""
            
            st.text_area("Cópia Direta WhatsApp:", value=script_wa, height=220)
            
        with col2:
            st.markdown("### ✉️ E-mail de Apresentação (Cold Mail)")
            script_email = f"""Assunto: Oportunidade de crescimento para {empresa_selecionada}

Olá, {socio_nome}, espero que este e-mail o encontre bem.

Meu nome é [Seu Nome] e acompanho o trabalho de empresas do setor de {categoria} na região.

Analisando a presença digital da {empresa_selecionada}, identifiquei algumas oportunidades claras para expandir a captação de novos clientes qualificados todos os meses.

Desenvolvemos estratégias sob medida que ajudam empresas como a sua a se destacarem e converterem mais oportunidades.

Podemos agendar uma breve conversa de 10 minutos na próxima terça-feira às 14h para eu te apresentar esse diagnóstico gratuitamente?

Atenciosamente,
[Seu Nome] | [Sua Empresa]"""
            
            st.text_area("Cópia Direta E-mail:", value=script_email, height=220)
    else:
        st.info("Rode uma extração na barra lateral para gerar scripts de vendas.")

with tab_crm:
    st.subheader("🗂️ Meus Leads (CRM)")
    leads_crm = carregar_leads_crm()

    if leads_crm is None:
        st.warning("Supabase não configurado. Peça ao admin para configurar SUPABASE_URL e SUPABASE_SERVICE_KEY nos Secrets.")
    elif not leads_crm:
        st.info("Nenhum lead salvo ainda. Rode uma extração na barra lateral — ela salva automaticamente aqui.")
    else:
        df_crm = pd.DataFrame(leads_crm)

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Leads Salvos", len(df_crm))
        col_m2.metric("Alta Prioridade (score ≥75)", int((df_crm["score"] >= 75).sum()))
        col_m3.metric("Contato Direto", int((df_crm["email"].notna() | df_crm["whatsapp"].notna()).sum()))
        col_m4.metric("Score Médio", round(df_crm["score"].mean(), 1) if len(df_crm) else 0)

        st.markdown("##### Pipeline")
        etapas = ['Novo', 'Contatado', 'Em negociação', 'Proposta', 'Ganho', 'Perdido']
        cols_pipeline = st.columns(len(etapas))
        for col, etapa in zip(cols_pipeline, etapas):
            qtd = int((df_crm["etapa_pipeline"] == etapa).sum())
            col.metric(etapa, qtd)

        st.markdown("---")
        st.markdown("##### Editar leads (etapa, score e nota de contato)")
        st.caption("Dados de contato (telefone/e-mail/etc.) são somente leitura aqui — a extração é a fonte da verdade pra esses campos.")

        colunas_exibir = ["nome_empresa", "categoria", "endereco", "telefone", "whatsapp", "email", "etapa_pipeline", "score", "nota_contato"]
        df_editor = df_crm[colunas_exibir].copy()

        edited = st.data_editor(
            df_editor,
            column_config={
                "nome_empresa": st.column_config.TextColumn("Empresa", disabled=True),
                "categoria": st.column_config.TextColumn("Categoria", disabled=True),
                "endereco": st.column_config.TextColumn("Endereço", disabled=True),
                "telefone": st.column_config.TextColumn("Telefone", disabled=True),
                "whatsapp": st.column_config.TextColumn("Whatsapp", disabled=True),
                "email": st.column_config.TextColumn("E-mail", disabled=True),
                "etapa_pipeline": st.column_config.SelectboxColumn("Etapa", options=etapas),
                "score": st.column_config.NumberColumn("Score", min_value=0, max_value=100),
                "nota_contato": st.column_config.TextColumn("Nota de Contato"),
            },
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="crm_editor"
        )

        if st.button("💾 Salvar alterações no CRM"):
            alteracoes = 0
            for i in range(len(edited)):
                original = df_editor.iloc[i]
                novo = edited.iloc[i]
                mudou = (
                    original["etapa_pipeline"] != novo["etapa_pipeline"]
                    or original["score"] != novo["score"]
                    or (original["nota_contato"] or "") != (novo["nota_contato"] or "")
                )
                if mudou:
                    atualizar_lead(df_crm.iloc[i]["id"], {
                        "etapa_pipeline": novo["etapa_pipeline"],
                        "score": int(novo["score"]),
                        "nota_contato": novo["nota_contato"] or None,
                    })
                    alteracoes += 1
            st.success(f"{alteracoes} lead(s) atualizado(s) com sucesso.")
            st.rerun()

if tab_admin is not None:
    with tab_admin:
        st.subheader("⚙️ Painel Admin")

        admin_tab_pool, admin_tab_keys, admin_tab_users, admin_tab_debug = st.tabs(
            ["🔄 Pool de Chaves", "🔑 Chave Única (legado)", "👥 Usuários", "🐛 Debug"]
        )

        with admin_tab_pool:
            st.markdown("Cadastre **várias** chaves por provider. O sistema tenta a menos usada recentemente primeiro, e troca sozinho (hot-swap) pra próxima quando uma responde 401/403/429 (esgotada/sem autorização).")

            for provider, label in [("serper", "Serper"), ("hunter", "Hunter.io")]:
                st.markdown(f"##### {label}")
                chaves_provider = [k for k in listar_pool_admin() if k["provider"] == provider]

                if not chaves_provider:
                    st.caption(f"Nenhuma chave {label} no pool ainda — sem pool, o app usa a chave única da aba ao lado (compatibilidade).")
                else:
                    for k in chaves_provider:
                        col_a, col_b, col_c, col_d = st.columns([3, 2, 2, 2])
                        apelido = k.get("label") or "(sem apelido)"
                        mascara = k["api_key"][:8] + "..." + k["api_key"][-4:] if len(k["api_key"]) > 12 else k["api_key"]
                        col_a.write(f"**{apelido}**  \n`{mascara}`")
                        status_cor = {"active": "🟢 Ativa", "exhausted": "🔴 Esgotada", "inactive": "⚪ Inativa"}
                        col_b.write(status_cor.get(k["status"], k["status"]))
                        if k.get("last_error"):
                            col_b.caption(f"Último erro: {k['last_error']}")
                        if k["status"] != "active":
                            if col_c.button("Reativar", key=f"reativar_{k['id']}"):
                                atualizar_status_pool(k["id"], "active")
                                st.rerun()
                        else:
                            if col_c.button("Desativar", key=f"desativar_{k['id']}"):
                                atualizar_status_pool(k["id"], "inactive")
                                st.rerun()
                        if col_d.button("Remover", key=f"remover_{k['id']}"):
                            remover_chave_pool(k["id"])
                            st.rerun()

                with st.form(key=f"add_pool_{provider}", clear_on_submit=True):
                    nova_chave = st.text_input(f"Nova chave {label}", type="password", key=f"nova_chave_{provider}")
                    apelido_novo = st.text_input("Apelido (opcional, ex: 'conta 2')", key=f"apelido_{provider}")
                    if st.form_submit_button(f"Adicionar ao pool {label}"):
                        if nova_chave.strip():
                            ok, erro = adicionar_chave_pool(provider, nova_chave.strip(), apelido_novo.strip())
                            if ok:
                                st.success("Chave adicionada!")
                                st.rerun()
                            else:
                                st.error(f"Erro ao adicionar (provavelmente chave duplicada): {erro}")
                st.markdown("---")

        with admin_tab_keys:
            st.markdown("Edite as chaves aqui quando expirarem — não precisa tocar no GitHub nem nos Secrets do Streamlit.")
            chaves_existentes = {k["key_name"]: k["key_value"] for k in listar_api_keys_admin()}

            for nome_chave, label in [("SERPER_API_KEY", "Serper API"), ("HUNTER_API_KEY", "Hunter.io API")]:
                valor_atual = chaves_existentes.get(nome_chave, "")
                novo_valor = st.text_input(f"Chave {label}", value=valor_atual, type="password", key=f"admin_key_{nome_chave}")
                if st.button(f"Salvar {label}", key=f"save_{nome_chave}"):
                    if salvar_api_key_admin(nome_chave, novo_valor, st.session_state['user_email']):
                        st.success(f"{label} atualizada! (vale a partir da próxima sessão de cada usuário)")
                        st.rerun()
                    else:
                        st.error("Erro ao salvar — verifique a conexão com o Supabase.")

        with admin_tab_users:
            st.markdown("Usuários que já logaram no app.")
            usuarios = listar_usuarios_admin()
            if not usuarios:
                st.info("Nenhum usuário cadastrado ainda (além de você).")
            else:
                for u in usuarios:
                    col_a, col_b, col_c = st.columns([3, 2, 2])
                    col_a.write(f"**{u.get('nome') or u['email']}**  \n{u['email']}")
                    eh_voce = u['email'] == 'reidecastro@gmail.com'
                    novo_role = col_b.selectbox(
                        "Role", ["user", "admin"],
                        index=["user", "admin"].index(u["role"]),
                        key=f"role_{u['email']}",
                        disabled=eh_voce,
                        label_visibility="collapsed"
                    )
                    if col_c.button("Atualizar", key=f"upd_{u['email']}", disabled=eh_voce):
                        if atualizar_role_usuario(u['email'], novo_role):
                            st.success("Atualizado!")
                            st.rerun()

        with admin_tab_debug:
            st.markdown("Snapshot dos headers da **última** resposta de cada API nesta sessão — só pra investigar se existe algum campo de saldo/crédito trafegando aí (ex: procurando por algo como `x-remaining-credits`). Não afeta nenhuma lógica do app.")
            for provider, label in [("serper", "Serper"), ("hunter", "Hunter.io")]:
                st.markdown(f"##### {label}")
                headers_debug = st.session_state.get(f'_debug_headers_{provider}')
                if headers_debug:
                    st.json(headers_debug)
                else:
                    st.caption("Nenhuma chamada feita ainda nesta sessão — rode uma extração e volte aqui.")

# --- RODAPÉ COM GERENCIAMENTO DE CONTA E CRÉDITOS ---
st.markdown("---")

# Nota técnica: o Serper.dev não expõe um endpoint público de saldo/créditos
# (confirmado na documentação deles — só têm search/images/videos/places/maps/
# reviews/news/shopping/lens/patents/autocomplete/scrape). Por isso, em vez de
# tentar consultar um saldo que a API não fornece, mostramos o consumo real
# contado pelo próprio app na última extração — que é 100% confiável.
creditos_usados_display = (
    f"**≈{st.session_state['creditos_usados']}** créditos usados na última extração"
    if 'creditos_usados' in st.session_state
    else "*Rode uma extração para ver o consumo estimado aqui*"
)

col_f1, col_f2 = st.columns([3, 1])

with col_f1:
    st.markdown(f"""
    **📊 Consumo Serper (estimado pelo app):**  
    {creditos_usados_display}  
    *Chave Ativa:* `{st.session_state['serper_api_key'][:8]}...{st.session_state['serper_api_key'][-4:]}`  
    *Nota: o Serper não tem endpoint público de saldo, então não dá pra consultar o total restante — só contamos quanto ESTE app gastou.*
    """)

    if st.session_state['hunter_api_key']:
        hunter_available, hunter_used = get_hunter_credits(st.session_state['hunter_api_key'])
        if hunter_available is not None:
            hunter_display = f"**{hunter_available}** buscas restantes (usadas: {hunter_used})"
        else:
            hunter_display = "⚠️ *Não foi possível consultar*"
        st.markdown(f"""
        **📊 Status da Conta Hunter.io:**  
        Saldo Atual: {hunter_display}
        """)

with col_f2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.link_button("🔗 Painel Serper.dev", "https://serper.dev/dashboard", use_container_width=True)
    if st.session_state['hunter_api_key']:
        st.link_button("🔗 Painel Hunter.io", "https://hunter.io/api-keys", use_container_width=True)
