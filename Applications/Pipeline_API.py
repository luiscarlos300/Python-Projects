import requests
import pandas as pd

base_url = ""
usuario = ''
senha = ''

def chamar_api_com_autenticacao(usuario, senha, url, page_size=3000, colunas="", query=""):
    try:
        auth = (usuario, senha)

        api_info = requests.get(url, auth=auth, headers={"Accept": "application/json", "Content-Type": "application/json"}).json()
        contar = api_info.get('result')

        total_records = 45000
        num_pages = -(-total_records // page_size)

        incidents = []

        for page_num in range(1, num_pages + 1):
            offset = (page_num - 1) * page_size
            url_loop = f"{url}?sysparm_limit={page_size}&sysparm_offset={offset}&sysparm_fields={colunas}&sysparm_query={query}"

            source = requests.get(url_loop, auth=auth, headers={"Accept": "application/json", "Content-Type": "application/json"}).json()

            incidents.extend(source.get('result', []))

        return incidents

    except requests.exceptions.RequestException as e:
        print(f'Erro na requisição: {str(e)}')

# Chamando API cmdb_ci
tabela_cmdb_ci = "cmdb_ci"
colunas_cmdb_ci = "install_date,sys_id"
query_cmdb_ci = "install_date!=null"
url_cmdb_ci = f"{base_url}{tabela_cmdb_ci}"
resultado_cmdb_ci = chamar_api_com_autenticacao(usuario, senha, url_cmdb_ci, page_size=3000, colunas=colunas_cmdb_ci)
df_cmdb_ci = pd.DataFrame(resultado_cmdb_ci)
df_cmdb_ci = df_cmdb_ci.rename(columns ={"sys_id": "sys_id_cmdb", "install_date": "Instalado"})

# Chamando API sn_install_base_item
tabela_sn_install_base_item = "sn_install_base_item"
colunas_sn_install_base_item = "number,u_id_instalacao,install_date,u_cancellation_date,configuration_item,name,sys_id,u_glide_date_1,u_glide_date_2,u_choice_1,u_notes"
query_sn_install_base_item = ""
url_sn_install_base_item = f"{base_url}{tabela_sn_install_base_item}"
resultado_sn_install_base_item = chamar_api_com_autenticacao(usuario, senha, url_sn_install_base_item, page_size=3000, colunas=colunas_sn_install_base_item)
df_sn_install_base_item = pd.DataFrame(resultado_sn_install_base_item)
df_sn_install_base_item = df_sn_install_base_item.rename(columns ={"sys_id": "sys_id_sn",
                                                                   "number": "Número",
                                                                   "u_id_instalacao": "ID de Instalação",
                                                                   "u_cancellation_date": "Data de Cancelamento",
                                                                   "name": "Item Base Instalada",
                                                                   "u_glide_date_1": "Data do Inicio do Bloqueio",
                                                                   "u_glide_date_2": "Data do Fim do Bloqueio",
                                                                   "install_date": "Data de instalação",
                                                                   "u_choice_1": "Motivo do Cancelamento",
                                                                   "u_notes": "Anotações"})
df_sn_install_base_item['value_sn'] = df_sn_install_base_item['configuration_item'].apply(lambda x: x.get('value') if isinstance(x, dict) else None)

# Merge cmdb_ci e sn_install_base_item p/ data instalação
df_sn_cmdb = df_sn_install_base_item.merge(df_cmdb_ci, how="left", right_on="sys_id_cmdb", left_on="value_sn")

# Chamando API customer_account
tabela_customer_account = "customer_account"
colunas_customer_account = "name," \
                           "sys_id," \
                           "zip," \
                           "u_cnpj_new," \
                           "u_segment," \
                           "state," \
                           "city," \
                           "street," \
                           "email"
query_customer_account = ""
url_customer_account = f"{base_url}{tabela_customer_account}"
resultado_customer_account = chamar_api_com_autenticacao(usuario, senha, url_customer_account, page_size=3000, colunas=colunas_customer_account, query= query_customer_account)
df_customer_account = pd.DataFrame(resultado_customer_account)
df_customer_account = df_customer_account.rename( columns={"sys_id": "sys_id_customer_account",
                                                           "name": "Nome",
                                                           "zip": "CEP",
                                                           "u_cnpj_new": "CNPJ",
                                                           "u_segment": "Segmento",
                                                           "state": "Estado/província",
                                                           "city": "Cidade",
                                                           "street": "Rua",
                                                           "email": "E-mail"})

# Chamando API csm_consumer
tabela_csm_consumer = "csm_consumer"
colunas_csm_consumer = "name," \
                       "sys_id," \
                       "zip," \
                       "u_cpf_new," \
                       "email," \
                       "state," \
                       "city," \
                       "street, " \
                       "mobile_phone"
query_csm_consumer = ""
url_csm_consumer = f"{base_url}{tabela_csm_consumer}"
resultado_csm_consumer = chamar_api_com_autenticacao(usuario, senha, url_csm_consumer, page_size=3000, colunas=colunas_csm_consumer, query= query_csm_consumer)
df_csm_consumer = pd.DataFrame(resultado_csm_consumer)
df_csm_consumer["Segmento"] = []
df_csm_consumer = df_csm_consumer.rename( columns={"sys_id": "sys_id_csm_consumer",
                                                   "name": "Nome",
                                                   "zip": "CEP",
                                                   "u_cpf_new": "CPF",
                                                   "email": "E-mail",
                                                   "state": "Estado/província",
                                                   "city": "Cidade",
                                                   "street": "Rua",
                                                   "mobile_phone": "Telefone celular"})

# Chamando API sn_customerservice_customers_billing_info
tabela_sn_billing = "sn_customerservice_customers_billing_info"
colunas_sn_billing = "soldpdct_name," \
                     "custbill_u_install_base_item," \
                     "soldpdct_u_network," \
                     "custbill_u_plan_value," \
                     "custbill_u_account," \
                     "custbill_u_cliente_suspenso_temporariamente," \
                     "custbill_u_inicio_faturamento," \
                     "custbill_u_billing_email," \
                     "soldpdct_consumer, " \
                     "custbill_u_due_date"
query_sn_billing = ""
url_sn_billing = f"{base_url}{tabela_sn_billing}"
resultado_sn_billing = chamar_api_com_autenticacao(usuario, senha, url_sn_billing, page_size=3000, colunas=colunas_sn_billing, query= query_sn_billing)
df_sn_billing = pd.DataFrame(resultado_sn_billing)
df_sn_billing = df_sn_billing.rename(columns={"soldpdct_name": "Nome Produto",
                                              "soldpdct_u_network": "Rede utilizada",
                                              "custbill_u_plan_value": "Valor cobrado para o cliente",
                                              "custbill_u_cliente_suspenso_temporariamente": "Cliente Suspenso Temporariamente",
                                              "custbill_u_inicio_faturamento": "Início do Faturamento",
                                              "custbill_u_billing_email": "E-mail Faturamento",
                                              "custbill_u_due_date": "Vencimento"})

df_sn_billing["value_sn_billing"] = df_sn_billing["custbill_u_install_base_item"].apply(lambda x: x.get("value") if isinstance(x, dict) else None)
df_sn_billing["value_customer_billing"] = df_sn_billing["custbill_u_account"].apply(lambda x: x.get("value") if isinstance(x, dict) else None)
df_sn_billing["value_consumer_billing"] = df_sn_billing["soldpdct_consumer"].apply(lambda x: x.get("value") if isinstance(x, dict) else None)

# Merge sn_customerservice_customers_billing_info c/ df_sn_cmdb
df_sn_billing_sn_install = df_sn_billing.merge(df_sn_cmdb, how="left", right_on="sys_id_sn", left_on="value_sn_billing")

# Merge sn_customerservice_customers_billing_info c/ csm_consumer
df_billing_consumer = df_sn_billing_sn_install.merge(df_csm_consumer, how="left", right_on="sys_id_csm_consumer", left_on="value_consumer_billing")
df_billing_consumer["len"] = df_billing_consumer["CPF"].str.len()
df_billing_consumer_filtered = df_billing_consumer[df_billing_consumer["len"] < 18]
df_billing_consumer_filtered = df_billing_consumer_filtered.drop(columns={"custbill_u_install_base_item",
                                                                 "custbill_u_account",
                                                                 "soldpdct_consumer",
                                                                 "value_sn_billing",
                                                                 "value_customer_billing",
                                                                 "value_consumer_billing",
                                                                 "sys_id_sn",
                                                                 "configuration_item",
                                                                 "value_sn",
                                                                 "sys_id_cmdb",
                                                                 "len"})

# Merge sn_customerservice_customers_billing_info c/ customer_account
df_billing_customer = df_sn_billing_sn_install.merge(df_customer_account, how="left", right_on="sys_id_customer_account", left_on="value_customer_billing")
df_billing_customer["len"] = df_billing_customer["CNPJ"].str.len() #apply(lambda x: len(x))
df_billing_customer_filtered = df_billing_customer[df_billing_customer['len']>13]
df_billing_customer_filtered = df_billing_customer_filtered.drop(columns={"custbill_u_install_base_item",
                                                                 "custbill_u_account",
                                                                 "soldpdct_consumer",
                                                                 "value_sn_billing",
                                                                 "value_customer_billing",
                                                                 "value_consumer_billing",
                                                                 "sys_id_sn",
                                                                 "sys_id_customer_account",
                                                                 "sys_id_csm_consumer"
                                                                 "configuration_item",
                                                                 "value_sn",
                                                                 "sys_id_cmdb",
                                                                 "len"})

df_billing_consumer_filtered.to_excel("f_billing_info_cpf.xlsx", index=False)
df_billing_customer_filtered.to_excel("f_billing_info_cnpj.xlsx", index=False)