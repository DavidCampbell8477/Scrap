import requests
import json
import pandas as pd
from collections import defaultdict

# XHR header and payload
url = "https://www.messe-stuttgart.de/solr/core_de/select"
paylaod = {
    "fi": "*,contacts:[json],av_booths:[json],av_jobs:[json],av_badges:[json],av_exhibitor_object:[json],av_teaser:[json],av_products:[json],av_news:[json],exhibitor_social_media:[json]",
    "q": "*:*",
    "fq": "{!tag=facet}((av_suggest_context:(((fairId__35135) OR (always__35135) OR (fairId__35137) OR (always__35137))) AND type:MkdbExhibitorIndexBundle|||Documents|||Exhibitor))",
    "start": "0",
    "rows": "500",
    "json.facet": "{\"type\":{\"type\":\"terms\",\"field\":\"type\",\"limit\":-1,\"domain\":{\"excludeTags\":\"facet,query,favourites\"}},\"typeReduced\":{\"type\":\"terms\",\"field\":\"type\",\"limit\":-1,\"mincount\":0,\"domain\":{\"filter\":\"((av_suggest_context:(((fairId__35135) OR (always__35135) OR (fairId__35137) OR (always__35137))) AND type:MkdbExhibitorIndexBundle|||Documents|||Exhibitor) OR type:MkdbExhibitorIndexBundle|||Documents|||ExhibitorNews) AND *:*\",\"excludeTags\":\"facet,query,favourites\"}},\"travel_destination\":{\"type\":\"terms\",\"field\":\"av_travel_destinations\",\"limit\":-1,\"sort\":{\"index\":\"asc\"},\"domain\":{\"blockChildren\":\"type:MkdbExhibitorIndexBundle|||Documents|||Exhibitor\",\"filter\":\"(fair_id:35135 OR fair_id:35137)\",\"excludeTags\":\"facet,query,favourites\"}},\"hall\":{\"type\":\"terms\",\"field\":\"av_halls\",\"limit\":-1,\"sort\":{\"index\":\"asc\"},\"domain\":{\"filter\":\"(fair_ids:35135 OR fair_ids:35137) AND -(av_halls:Undefiniert OR av_halls:Undefined) AND type:MkdbExhibitorIndexBundle|||Documents|||Exhibitor\",\"excludeTags\":\"facet,query,favourites\"}},\"country\":{\"type\":\"terms\",\"field\":\"av_country\",\"limit\":-1,\"method\":\"enum\",\"sort\":{\"index\":\"asc\"},\"domain\":{\"excludeTags\":\"facet,query,favourites\"}},\"fairEvent\":{\"type\":\"terms\",\"field\":\"fair_ids\",\"limit\":-1,\"sort\":{\"index\":\"asc\"},\"domain\":{\"excludeTags\":\"facet,query,favourites\"}}}"
}
headers = {
    "accept": "application/json, text/plain, */*",
    "accept-encoding": "gzip, deflate, br, zstd",
    "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
    "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
    "origin": "https://www.messe-stuttgart.de",
    "priority": "u=1, i",
    "referer": "https://www.messe-stuttgart.de/eltefa/ausstellung/unternehmen-produkte/ausstellerverzeichnis/",
    "sec-ch-ua": '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
    }   

# Get data from Solr
response = requests.post(url, params=paylaod, headers=headers)
if response.status_code == 200: 
    data = response.json()
    if 'response' in data:
        response_data = data['response']
    else:
        print("Key 'response' not found in the JSON response")
else:
    print(f"Request failed with status code: {response.status_code}")

# Extract information from response
companies_data = []
contacts_data = []
social_media_data = defaultdict(list)

# Safely parse data that might be a JSON string or list/dict
def safe_parse(data):
    if isinstance(data, str):
        try:
            return json.loads(data)
        except json.JSONDecodeError:
            return []
    elif isinstance(data, list):
        return data
    return []
# Collect company info, social media and Contacts
for doc in response_data['docs']:
    company_name = doc.get('title', '')
    company = {
        'Company Name': doc.get('title', ''),
        'Country': doc.get('av_country', ''),
        'City': doc.get('av_city', ''),
        'Website': doc.get('av_url', ''),
        'Company Email': doc.get('av_email', ''),
    }
    companies_data.append(company)
    # Collect social media links
    # Handle exhibitor_social_media field (could be string or list)
    social_media = safe_parse(doc.get('exhibitor_social_media', []))
    for sm in social_media:
        sm = safe_parse(sm)
        link = sm.get('link', '')
        social_media_data[company_name].append(link)    
    # Collect contacts
    contacts = safe_parse(doc.get('contacts', []))
    for contact in contacts:
        if isinstance(contact, dict):
            contacts_data.append({
                "Company Name": company_name,
                "Contact Name": contact.get('name', ''),
                "Responsibility": contact.get('responsibility', ''),
                "Contact Email": contact.get('email', ''),
                "Phone": contact.get('phone', '')
            })
# Add social media to company data
for company in companies_data:
    name = company["Company Name"]
    company["Social Media Links"] = ", ".join(social_media_data.get(name, []))
    
# Create a DataFrame and save to Excel
df_companies = pd.DataFrame(companies_data)
df_contacts = pd.DataFrame(contacts_data)

# Save to Excel with multiple sheets
with pd.ExcelWriter('exhibitor_data.xlsx') as writer:
    df_companies.to_excel(writer, sheet_name='Companies', index=False)
    df_contacts.to_excel(writer, sheet_name='Contacts', index=False)
