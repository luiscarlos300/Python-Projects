import pandas as pd
import requests

base_url = ''
usuario = ''
senha = 

def chamar_api_com_autenticacao(usuario, senha, url, page_size=3000, colunas="", query="", total_records = 40000):
    try:
        auth = (usuario, senha)

        api_info = requests.get(url, auth=auth, headers={"Accept": "application/json", "Content-Type": "application/json"}).json()

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

# Chamando API sys_user
tabela_sys_user = "sys_user"
colunas_sys_user = "sys_id, name"
query_sys_user = ""
url_sys_user = f"{base_url}{tabela_sys_user}"
resultado_sys_user = chamar_api_com_autenticacao(usuario, senha, url_sys_user, page_size=3000, colunas=colunas_sys_user)
df_sys_user_opened = pd.DataFrame(resultado_sys_user)
df_sys_user_closed = pd.DataFrame(resultado_sys_user)
df_sys_user_opened = df_sys_user_opened.rename(columns={"sys_id": "sys_id_sys_user_opened",
                                          "name": "Aberto por"})
df_sys_user_closed = df_sys_user_closed.rename(columns={"sys_id": "sys_id_sys_user_closed",
                                          "name": "Encerrado por"})
# Chamando API sys_user_group
tabela_sys_user_group = "sys_user_group"
colunas_sys_user_group = "sys_id,name"
query_sys_user_group = ""
url_sys_user_group = f"{base_url}{tabela_sys_user_group}"
resultado_sys_user_group = chamar_api_com_autenticacao(usuario, senha, url_sys_user_group, page_size=3000, colunas=colunas_sys_user_group)
df_sys_user_group = pd.DataFrame(resultado_sys_user_group)
df_sys_user_group = df_sys_user_group.rename(columns={"sys_id": "sys_id_user_group",
                                                      "name": "Grupo de atribuição"})

# Chamando API csm_consumer
tabela_csm_consumer = "csm_consumer"
colunas_csm_consumer = "name," \
                       "sys_id"
query_csm_consumer = ""
url_csm_consumer = f"{base_url}{tabela_csm_consumer}"
resultado_csm_consumer = chamar_api_com_autenticacao(usuario, senha, url_csm_consumer, page_size=3000, colunas=colunas_csm_consumer, query= query_csm_consumer)
df_csm_consumer = pd.DataFrame(resultado_csm_consumer)
df_csm_consumer = df_csm_consumer.rename( columns={"sys_id": "sys_id_csm_consumer",
                                                   "name": "Nome do Cliente Pessoa Física"})

df_csm_consumer_client = pd.DataFrame(resultado_csm_consumer).rename( columns={"sys_id": "sys_id_csm_consumer",
                                                   "name": "Cliente"})

# Chamando API customer_account
tabela_customer_account = "customer_account"
colunas_customer_account = "name," \
                           "sys_id"
query_customer_account = ""
url_customer_account = f"{base_url}{tabela_customer_account}"
resultado_customer_account = chamar_api_com_autenticacao(usuario, senha, url_customer_account, page_size=3000, colunas=colunas_customer_account, query= query_customer_account)
df_customer_account = pd.DataFrame(resultado_customer_account)
df_customer_account = df_customer_account.rename( columns={"sys_id": "sys_id_customer_account",
                                                           "name": "Nome do Cliente Corporativo"})

try: # Chamando API x_eqte_call_center_financial_call
    tabela_financial = "x_eqte_call_center_financial_call"
    colunas_financial = "number,opened_at,opened_by, state, closed_by, closed_at, reason, consumer, account, assignment_group, description"
    query_financial = ""
    url_financial = f"{base_url}{tabela_financial}"
    resultado_financial = chamar_api_com_autenticacao(usuario, senha, url_financial, page_size=200, colunas=colunas_financial, total_records=6000)
    df_financial = pd.DataFrame(resultado_financial)
    df_financial["value_assignment_group"] = df_financial["assignment_group"].apply(lambda x: x.get('value') if isinstance(x, dict) else None)
    df_financial_group = df_financial.merge(df_sys_user_group,how="left", right_on="sys_id_user_group", left_on="value_assignment_group")
    df_financial_group["value_user_opened"] = df_financial_group["opened_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_financial_group_opened = df_financial_group.merge(df_sys_user_opened, how="left", right_on="sys_id_sys_user_opened", left_on="value_user_opened")
    df_financial_group_opened["value_user_closed"] = df_financial_group_opened["closed_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_financial_group_opened_closed = df_financial_group_opened.merge(df_sys_user_closed, how="left", right_on="sys_id_sys_user_closed", left_on="value_user_closed")
    df_financial_group_opened_closed["value_csm"] = df_financial_group_opened_closed["consumer"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_financial_csm = df_financial_group_opened_closed.merge(df_csm_consumer, how="left", right_on="sys_id_csm_consumer", left_on="value_csm")
    df_financial_csm["value_customer"] = df_financial_csm["account"].apply(lambda x: x.get("value") if isinstance(x, dict) else None)
    df_financial_account = df_financial_csm.merge(df_customer_account, how="left", right_on="sys_id_customer_account", left_on="value_customer")
    df_financial_account = df_financial_account.drop(columns={"value_customer",
                                                          "value_csm",
                                                          "value_user_closed",
                                                          "value_user_opened",
                                                          "value_assignment_group",
                                                          "sys_id_customer_account",
                                                          "sys_id_csm_consumer",
                                                          "sys_id_sys_user_closed",
                                                          "sys_id_sys_user_opened",
                                                          "sys_id_user_group",
                                                          "assignment_group",
                                                          "opened_by",
                                                          "consumer",
                                                          "account",
                                                          "closed_by"})

    reason_replace_fin = ["dissatisfaction_product_quality",
                          "registration_update",
                          "unfeasible_change_address",
                          "address_change",
                          "unknown_installation",
                          "financial_problems",
                          "local_point_change",
                          "dissatisfaction_sales",
                          "technical_complaints",
                          "change_ownership",
                          "speed_change",
                          "sales_complaint",
                          "expiration_change",
                          "invoice_contestation",
                          "unlock_req",
                          "debt_negotiation",
                          "2nd_copy_of_invoice",
                          "slowness_connection_problem_outages",
                          "system_failure",
                          "customer_not_browse"]

    reason_replace_fin_pt = ["Insatisfação com produto/qualidade",
                            "Atualização de cadastro",
                            "Inviabilidade para mudança de endereço",
                            "Mudança de endereço",
                            "Desconhece instalação",
                            "Problemas financeiros",
                            "Mudança local de ponto",
                            "Insatisfação com a venda",
                            "Reclamação técnicas",
                            "Alteração de Titularidade",
                            "Alteração de Velocidade",
                            "Reclamação de vendas",
                            "Alteração de Vencimento",
                            "Contestação de fatura",
                            "Solicitação de Desbloqueio",
                            "Negociação de Débito",
                            "Solicitação de 2ª via fatura",
                            "Lentidão/Problema de conexão/Quedas",
                            "Falha de sistema (falha no sistema API)",
                            "Cliente não navega"]

    df_financial_account["reason"] = df_financial_account["reason"].replace(reason_replace_fin,reason_replace_fin_pt)
    df_financial_account["state"] = df_financial_account["state"].replace(["3", "1", "10", "18", "6", "7"], ["Encerrado", "Novo(a)", "Em aberto", "Aguardando informações", "Resolvido","Cancelada"])
    df_financial_account = df_financial_account.rename(columns={"number": "Número",
                                         "state": "Estado",
                                         "opened_at": "Aberto",
                                         "closed_at": "Encerrado",
                                         "reason": "Motivo",
                                         "description": "Descrição Resumida"})

    df_financial_account.to_excel("",index=False)

except Exception as e:
    print(f"erro financial {e}")

try: # Chamando API x_eqte_call_center_after_sales_call
    tabela_pos_venda = "x_eqte_call_center_after_sales_call"
    colunas_pos_venda = "number,opened_at,opened_by, state, closed_by, closed_at, reason, consumer, account, assignment_group, description"
    query_pos_venda = ""
    url_pos_venda = f"{base_url}{tabela_pos_venda}"
    resultado_pos_venda = chamar_api_com_autenticacao(usuario, senha, url_pos_venda, page_size=1000, colunas=colunas_pos_venda, total_records=6000)
    df_pos_venda = pd.DataFrame(resultado_pos_venda)
    df_pos_venda["value_assignment_group"] = df_pos_venda["assignment_group"].apply( lambda x: x.get('value') if isinstance(x, dict) else None)
    df_pos_venda_group = df_pos_venda.merge(df_sys_user_group,how="left", right_on="sys_id_user_group", left_on="value_assignment_group")
    df_pos_venda_group["value_user_opened"] = df_pos_venda_group["opened_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_pos_venda_group_opened = df_pos_venda_group.merge(df_sys_user_opened, how="left", right_on="sys_id_sys_user_opened", left_on="value_user_opened")
    df_pos_venda_group_opened["value_user_closed"] = df_pos_venda_group_opened["closed_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_pos_venda_group_opened_closed = df_pos_venda_group_opened.merge(df_sys_user_closed, how="left", right_on="sys_id_sys_user_closed", left_on="value_user_closed")
    df_pos_venda_group_opened_closed["value_csm"] = df_pos_venda_group_opened_closed["consumer"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_pos_venda_csm = df_pos_venda_group_opened_closed.merge(df_csm_consumer, how="left", right_on="sys_id_csm_consumer", left_on="value_csm")
    df_pos_venda_csm["value_customer"] = df_pos_venda_csm["account"].apply(lambda x: x.get("value") if isinstance(x, dict) else None)
    df_pos_venda_account = df_pos_venda_csm.merge(df_customer_account, how="left", right_on="sys_id_customer_account", left_on="value_customer")
    df_pos_venda_account = df_pos_venda_account.drop(columns={"value_customer",
                                                          "value_csm",
                                                          "value_user_closed",
                                                          "value_user_opened",
                                                          "value_assignment_group",
                                                          "sys_id_customer_account",
                                                          "sys_id_csm_consumer",
                                                          "sys_id_sys_user_closed",
                                                          "sys_id_sys_user_opened",
                                                          "sys_id_user_group",
                                                          "assignment_group",
                                                          "opened_by",
                                                          "consumer",
                                                          "account",
                                                          "closed_by"})

    reason_replace_pos = ["dissatisfaction_product_quality",
                          "registration_update",
                          "unfeasible_change_address",
                          "address_change",
                          "unknown_installation",
                          "financial_problems",
                          "local_point_change",
                          "dissatisfaction_sales",
                          "technical_complaints",
                          "change_ownership",
                          "speed_change",
                          "sales_complaint",
                          "expiration_change",
                          "invoice_contestation",
                          "unlock_req",
                          "debt_negotiation",
                          "2nd_copy_of_invoice",
                          "slowness_connection_problem_outages",
                          "system_failure",
                          "customer_not_browse"]

    reason_replace_pos_pt = ["Insatisfação com produto/qualidade",
                            "Atualização de cadastro",
                            "Inviabilidade para mudança de endereço",
                            "Mudança de endereço",
                            "Desconhece instalação",
                            "Problemas financeiros",
                            "Mudança local de ponto",
                            "Insatisfação com a venda",
                            "Reclamação técnicas",
                            "Alteração de Titularidade",
                            "Alteração de Velocidade",
                            "Reclamação de vendas",
                            "Alteração de Vencimento",
                            "Contestação de fatura",
                            "Solicitação de Desbloqueio",
                            "Negociação de Débito",
                            "Solicitação de 2ª via fatura",
                            "Lentidão/Problema de conexão/Quedas",
                            "Falha de sistema (falha no sistema API)",
                            "Cliente não navega"]

    df_pos_venda_account["reason"] = df_pos_venda_account["reason"].replace(reason_replace_pos,reason_replace_pos_pt)
    df_pos_venda_account["state"] = df_pos_venda_account["state"].replace(["3","1","10", "18", "6", "7"], ["Encerrado", "Novo(a)", "Em aberto", "Aguardando informações", "Resolvido","Cancelada"])
    df_pos_venda_account = df_pos_venda_account.rename(columns={"number": "Número",
                                         "state": "Estado",
                                         "opened_at": "Aberto",
                                         "closed_at": "Encerrado",
                                         "reason": "Motivo",
                                         "description": "Descrição Resumida"})

    df_pos_venda_account.to_excel("", index=False)

except Exception as e:
    print(f"erro pos venda {e}")

try: # Chamando API x_eqte_call_center_technical_call
    tabela_tecnico = "x_eqte_call_center_technical_call"
    colunas_tecnico = "number,opened_at,opened_by, state, closed_by, closed_at, reason, consumer, account, assignment_group, description"
    query_tecnico = ""
    url_tecnico = f"{base_url}{tabela_tecnico}"
    resultado_tecnico = chamar_api_com_autenticacao(usuario, senha, url_tecnico, page_size=1000, colunas=colunas_tecnico, total_records=6000)
    df_tecnico = pd.DataFrame(resultado_tecnico)
    df_tecnico["value_assignment_group"] = df_tecnico["assignment_group"].apply( lambda x: x.get('value') if isinstance(x, dict) else None)
    df_tecnico_group = df_tecnico.merge(df_sys_user_group,how="left", right_on="sys_id_user_group", left_on="value_assignment_group")
    df_tecnico_group["value_user_opened"] = df_tecnico_group["opened_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_tecnico_group_opened = df_tecnico_group.merge(df_sys_user_opened, how="left", right_on="sys_id_sys_user_opened", left_on="value_user_opened")
    df_tecnico_group_opened["value_user_closed"] = df_tecnico_group_opened["closed_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_tecnico_group_opened_closed = df_tecnico_group_opened.merge(df_sys_user_closed, how="left", right_on="sys_id_sys_user_closed", left_on="value_user_closed")
    df_tecnico_group_opened_closed["value_csm"] = df_tecnico_group_opened_closed["consumer"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_tecnico_csm = df_tecnico_group_opened_closed.merge(df_csm_consumer, how="left", right_on="sys_id_csm_consumer", left_on="value_csm")
    df_tecnico_csm["value_customer"] = df_tecnico_csm["account"].apply(lambda x: x.get("value") if isinstance(x, dict) else None)
    df_tecnico_account = df_tecnico_csm.merge(df_customer_account, how="left", right_on="sys_id_customer_account", left_on="value_customer")
    df_tecnico_account = df_tecnico_account.drop(columns={"value_customer",
                                                          "value_csm",
                                                          "value_user_closed",
                                                          "value_user_opened",
                                                          "value_assignment_group",
                                                          "sys_id_customer_account",
                                                          "sys_id_csm_consumer",
                                                          "sys_id_sys_user_closed",
                                                          "sys_id_sys_user_opened",
                                                          "sys_id_user_group",
                                                          "assignment_group",
                                                          "opened_by",
                                                          "consumer",
                                                          "account",
                                                          "closed_by"})

    reason_replace_tec = ["dissatisfaction_product_quality",
                          "registration_update",
                          "unfeasible_change_address",
                          "address_change",
                          "unknown_installation",
                          "financial_problems",
                          "local_point_change",
                          "dissatisfaction_sales",
                          "technical_complaints",
                          "change_ownership",
                          "speed_change",
                          "sales_complaint",
                          "expiration_change",
                          "invoice_contestation",
                          "unlock_req",
                          "debt_negotiation",
                          "2nd_copy_of_invoice",
                          "slowness_connection_problem_outages",
                          "system_failure",
                          "customer_not_browse"]

    reason_replace_tec_pt = ["Insatisfação com produto/qualidade",
                            "Atualização de cadastro",
                            "Inviabilidade para mudança de endereço",
                            "Mudança de endereço",
                            "Desconhece instalação",
                            "Problemas financeiros",
                            "Mudança local de ponto",
                            "Insatisfação com a venda",
                            "Reclamação técnicas",
                            "Alteração de Titularidade",
                            "Alteração de Velocidade",
                            "Reclamação de vendas",
                            "Alteração de Vencimento",
                            "Contestação de fatura",
                            "Solicitação de Desbloqueio",
                            "Negociação de Débito",
                            "Solicitação de 2ª via fatura",
                            "Lentidão/Problema de conexão/Quedas",
                            "Falha de sistema (falha no sistema API)",
                            "Cliente não navega"]

    df_tecnico_account["reason"] = df_tecnico_account["reason"].replace(reason_replace_tec,reason_replace_tec_pt)
    df_tecnico_account["state"] = df_tecnico_account["state"].replace(["3","1","10", "18", "6", "7"], ["Encerrado", "Novo(a)", "Em aberto", "Aguardando informações", "Resolvido","Cancelada"])
    df_tecnico_account = df_tecnico_account.rename(columns={"number": "Número",
                                         "state": "Estado",
                                         "opened_at": "Aberto",
                                         "closed_at": "Encerrado",
                                         "reason": "Motivo",
                                         "description": "Descrição Resumida"})
    df_tecnico_account.to_excel("",index=False)

except Exception as e:
    print(f"erro tecnico {e}")

try: # Chamando API x_eqte_call_center_service_registration
    tabela_registro = "x_eqte_call_center_service_registration"
    colunas_registro = "number,opened_at,opened_by, state, closed_by, closed_at, reason, consumer, assignment_group, description, name_of_potential_customer, customer_exist"
    query_registro = ""
    url_registro = f"{base_url}{tabela_registro}"
    resultado_registro = chamar_api_com_autenticacao(usuario, senha, url_registro, page_size=2000, colunas=colunas_registro, total_records=12000)
    df_registro = pd.DataFrame(resultado_registro)
    df_registro["value_assignment_group"] = df_registro["assignment_group"].apply( lambda x: x.get('value') if isinstance(x, dict) else None)
    df_registro_group = df_registro.merge(df_sys_user_group,how="left", right_on="sys_id_user_group", left_on="value_assignment_group")
    df_registro_group["value_user_opened"] = df_registro_group["opened_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_registro_group_opened = df_registro_group.merge(df_sys_user_opened, how="left", right_on="sys_id_sys_user_opened", left_on="value_user_opened")
    df_registro_group_opened["value_user_closed"] = df_registro_group_opened["closed_by"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_registro_group_opened_closed = df_registro_group_opened.merge(df_sys_user_closed, how="left", right_on="sys_id_sys_user_closed", left_on="value_user_closed")
    df_registro_group_opened_closed["value_csm"] = df_registro_group_opened_closed["consumer"].apply( lambda x: x.get("value") if isinstance(x, dict) else None)
    df_registro_csm = df_registro_group_opened_closed.merge(df_csm_consumer_client, how="left", right_on="sys_id_csm_consumer", left_on="value_csm")
    df_registro_csm = df_registro_csm.drop(columns={"value_csm",
                                                          "value_user_closed",
                                                          "value_user_opened",
                                                          "value_assignment_group",
                                                          "sys_id_csm_consumer",
                                                          "sys_id_sys_user_closed",
                                                          "sys_id_sys_user_opened",
                                                          "sys_id_user_group",
                                                          "assignment_group",
                                                          "opened_by",
                                                          "consumer",
                                                          "closed_by"})

    reason_replace_reg = ["payments_unlocks",
                          "coverage_availability_services",
                          "remote_settings",
                          "invoices_debts",
                          "dates_intallations_change_addresses",
                          "changing_wi_fi",
                          "send_2nd_invoice",
                          "call_failure",
                          "change_ownership",
                          "scheduled_outage",
                          "negotiation_instalments",
                          "service_plans_bundles",
                          "promotions_discounts"]

    reason_replace_reg_pt = ["Pagamentos e desbloqueios",
                             "Cobertura e disponibilidade de serviços",
                             "Configurações remotas (Reboot/Habilitar/Desabilitar ONT)",
                             "Faturas e débitos",
                             "Datas de instalações/mudanças de endereços",
                             "Alteração de Wi-fi (Habiliar/Desabilitar/Alterar)",
                             "Envio de 2ª via fatura",
                             "Falha na ligação",
                             "Troca de Titularidade",
                             "Interrupção Programada",
                             "Negociação/Parcelamento",
                             "Planos e pacotes de serviços",
                             "Promoçoes e descontos"]

    df_registro_csm["reason"] = df_registro_csm["reason"].replace(reason_replace_reg,reason_replace_reg_pt)
    df_registro_csm["state"] = df_registro_csm["state"].replace(["3","1","10", "18", "6", "7"], ["Encerrado", "Novo(a)", "Em aberto", "Aguardando informações", "Resolvido","Cancelada"])
    df_registro_csm["customer_exist"] = df_registro_csm["customer_exist"].replace(["yes", "no"], ["Sim", "Não"])

    df_registro_csm = df_registro_csm.rename(columns={"number": "Número",
                                         "state": "Estado",
                                         "opened_at": "Aberto",
                                         "closed_at": "Encerrado",
                                         "reason": "Motivo",
                                         "description": "Descrição Resumida",
                                         "name_of_potential_customer": "Nome do possível cliente",
                                         "customer_exist": "Cliente já existe?"})

    df_registro_csm.to_excel("", index=False)

except Exception as e:
    print(f"erro registro {e}")