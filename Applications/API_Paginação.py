import requests
import json

page_size = 3000
base_url = ""
tabela = "sn_customerservice_customers_billing_info"
coluna_contar = "?soldpdct_name"
colunas = "soldpdct_name"
url = base_url + tabela

url_contar = url + coluna_contar
headers = {"Accept": "application/json", "Content-Type": "application/json"}
api_info = requests.get(url_contar, headers=headers).json()
contar = api_info["result"]
total_records = 100000  # List.Count(contar)
num_pages = total_records // page_size + (1 if total_records % page_size > 0 else 0)  # Número total de páginas

incidents = []
for page_num in range(1, num_pages + 1):
    offset = (page_num - 1) * page_size
    url_loop = f"{url}?sysparm_limit={page_size}&sysparm_offset={offset}&sysparm_fields={colunas}"
    response = requests.get(url_loop, headers=headers).json()["result"]
    if page_num == num_pages:
        break
    incidents.extend(response)

incidents_table = pd.DataFrame(incidents)
incidents_table = incidents_table[incidents_table["soldpdct_name"].notnull() & (incidents_table["soldpdct_name"] != "")]

print(incidents_table)
