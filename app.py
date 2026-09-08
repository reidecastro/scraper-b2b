def run_places_extraction(query, api_key, max_results=10, email_opt=True, redes_opt=True):
    extracted_data = []
    
    url = "https://google.serper.dev/places"
    payload = {"q": query, "gl": "br", "hl": "pt-br"}
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        data = response.json()
        
        places = data.get("places", [])
        for item in places[:max_results]:
            company_name = item.get("title", "")
            
            # Tenta pegar o endereço em todas as chaves possíveis que a Serper retorna
            address = item.get("address") or item.get("formattedAddress") or item.get("street") or ""
            
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
            
            if website_url:
                if email_opt or redes_opt:
                    real_email, real_social = scrape_website_details(website_url)
            
            if not phone_raw or not real_email:
                fb_phone, fb_email, fb_social = fallback_search_phone_email(company_name, address, api_key)
                if not phone_raw and fb_phone:
                    phone_raw = fb_phone
                if not real_email and fb_email:
                    real_email = fb_email
                if not real_social and fb_social:
                    real_social = fb_social

            formatted_phone, wa_link = clean_and_format_phone(phone_raw)

            extracted_data.append({
                "Prompt": query,
                "Nome da Empresa": company_name,
                "Categoria": category,
                "Responsável": "",
                "Endereço": address,  # CORRIGIDO: usa o endereço puro sem colocar o termo da busca de fallback
                "Telefone": formatted_phone,
                "Whatsapp": wa_link,
                "Email": real_email,
                "Redes Sociais": real_social,
                "Nota Google": rating,
                "Total Avaliações": rating_count,
                "Status": "A Fazer",
                "Progressão": "1º Contato",
                "Tem Website": "Sim" if website_url else "Não",
                "Link Google Maps": gmaps_link,
                "Observações": f"Site: {website_url}" if website_url else "Sem site oficial"
            })
            
    except Exception as e:
        st.error(f"Erro na requisição: {e}")

    return pd.DataFrame(extracted_data)
