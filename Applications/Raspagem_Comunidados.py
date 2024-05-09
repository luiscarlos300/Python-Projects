from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
from openpyxl import Workbook
import pandas as pd
from datetime import date
import time
import requests
import re
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def realizar_raspagem_e_enviar_email():
    try:
        data_list = []

        # 1º) Raspagem inicial de incidentes
        try:
            browser = webdriver.Edge()
            browser.get("")
            time.sleep(5)
            username = browser.find_element("name", "Ecom_User_ID")
            password = browser.find_element("name", "Ecom_Password")
            login_button = browser.find_element("id", "submit")
            username.send_keys("")
            password.send_keys("")
            login_button.click()
            time.sleep(5)
            browser.get("")
            time.sleep(25)
            contagem = 1
            while contagem < 2: # Qtde. de páginas comunicados
                contagem += 1
                soup = BeautifulSoup(browser.page_source, 'html.parser')
                rows = soup.find_all('tr', class_='')
                for row in rows:
                    cells = row.find_all('td', class_=['cursor-pointer', 'q-td', 'text-left', 'text-center'])
                    data = [cell.get_text(strip=True) for cell in cells]
                    data_list.append(data)
                wait = WebDriverWait(browser, 30)
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))
                next_button_first = browser.find_element(By.XPATH,"//button[contains(@class, 'q-btn') and .//i[contains(@class, 'q-icon') and text()='keyboard_arrow_right']]")
                time.sleep(8)
                if next_button_first.is_enabled():
                    next_button_first.click()
                    time.sleep(25)
                    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))
        except Exception as e:
            print(f"Erro: {e}")

        columns = ['Identificador', 'Tipo', 'Status', 'Qt', 'UF', 'Municipio', 'Criticidade', 'Data','Data Ult Atualizacao', 'Data Inicio', 'Previsão de Retorno', 'Ticket Associados']
        df = pd.DataFrame(data_list, columns=columns)

        # 2º) Raspagem eq's por incidentes
        try:
            lista_id_comunicado = df['Identificador'].tolist()
            resultado_data = {'Comunicado': [], 'EQ Values': []}
            for id_lista in lista_id_comunicado:
                if id_lista and not any(char in {'E', 'e', 'C', 'c'} for char in id_lista):
                    url = f"https:///portal/#/clientesImpactados/{id_lista}"
                    browser.get(url)
                    time.sleep(5)
                    while True:
                        page_source = browser.page_source
                        soup = BeautifulSoup(page_source, 'html.parser')
                        span_element = soup.find('span', class_='text-weight-medium text-grey-9 text-h6')
                        if span_element:
                            texto_span = span_element.get_text(strip=True)
                            comunicado_valor = texto_span.split('Casas Conectadas Impactadas do Comunicado ')[-1]
                            eq_values = [td.text.strip() for td in soup.find_all('td', class_='q-td') if 'EQ' in td.text]
                            for eq_value in eq_values:
                                resultado_data['Comunicado'].append(comunicado_valor)
                                resultado_data['EQ Values'].append(eq_value)
                            try:
                                next_button = browser.find_element(By.XPATH,"//button[contains(@class, 'q-btn--wrap') and .//i[@class='q-icon notranslate material-icons' and contains(text(), 'keyboard_arrow_right')]]")
                                if next_button.is_displayed() and next_button.is_enabled():
                                    try:
                                        WebDriverWait(browser, 25).until(EC.invisibility_of_element_located((By.CLASS_NAME, "q-loading")))
                                    except TimeoutException:
                                        pass
                                    next_button.click()
                                    time.sleep(7)
                                else:
                                    break
                            except NoSuchElementException:
                                break
            resultado_df = pd.DataFrame(resultado_data)
            time.sleep(8)
        except Exception as e:
            print(f"Erro: {e}")
        finally:
            browser.quit()

        # 3º) Merge dos dataFrames
        try:
            df['Identificador'] = pd.to_numeric(df['Identificador'], errors='coerce')
            resultado_df['Comunicado'] = pd.to_numeric(resultado_df['Comunicado'], errors='coerce')
            merged_df = pd.merge(resultado_df,df, left_on='Comunicado', right_on='Identificador', how='inner')
        except Exception as e:
            print(f"Erro durante o merge: {e}")
        finally:
            browser.quit()

        # 4º) Chamando api servicenow - clientes
        base_url = ""
        usuario = ''
        senha = ''
        def chamar_api_com_autenticacao(usuario, senha, url, page_size=2000, colunas="", query=""):
            try:
                auth = (usuario, senha)
                api_info = requests.get(url, auth=auth, headers={"Accept": "application/json", "Content-Type": "application/json"}).json()
                contar = api_info.get('result')
                total_records = 38000 # Qtde. linhas
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

        # Chamando API sn_install_base_item
        tabela_sn_install_base_item = "sn_install_base_item"
        colunas_sn_install_base_item = "u_id_instalacao,configuration_item,sys_id"
        query_sn_install_base_item = "u_id_instalacao!=null"
        url_sn_install_base_item = f"{base_url}{tabela_sn_install_base_item}"
        resultado_sn_install_base_item = chamar_api_com_autenticacao(usuario, senha, url_sn_install_base_item, page_size=2000, colunas=colunas_sn_install_base_item)
        df_sn_install_base_item = pd.DataFrame(resultado_sn_install_base_item)

        # Chamando API csm_consumer
        tabela_csm_consumer = "csm_consumer"
        colunas_csm_consumer = "name,sys_id,u_cpf_new"
        query_csm_consumer = ""
        url_csm_consumer = f"{base_url}{tabela_csm_consumer}"
        resultado_csm_consumer = chamar_api_com_autenticacao(usuario, senha, url_csm_consumer, page_size=2000, colunas=colunas_csm_consumer, query= query_csm_consumer)
        df_csm_consumer = pd.DataFrame(resultado_csm_consumer)

        # Chamando API sn_customerservice_customers_billing_info
        tabela_billing = "sn_customerservice_customers_billing_info"
        colunas_billing = "custbill_u_install_base_item,soldpdct_consumer"
        query_billing = ""
        url_billing = f"{base_url}{tabela_billing}"
        resultado_billing = chamar_api_com_autenticacao(usuario, senha, url_billing, page_size=2000, colunas=colunas_billing, query= query_billing)
        df_billing = pd.DataFrame(resultado_billing)
        df_billing['value_custbill'] = df_billing['custbill_u_install_base_item'].apply(lambda x : x.get('value') if isinstance(x, dict) else None)
        df_billing['value_soldpdct'] = df_billing['soldpdct_consumer'].apply(lambda x : x.get('value') if isinstance(x, dict) else None)
        df_billing_merged = pd.merge(df_billing, df_sn_install_base_item[['u_id_instalacao', 'sys_id']], left_on='value_custbill', right_on='sys_id', how='left')
        df_billing_merged = pd.merge(df_billing_merged, df_csm_consumer[['sys_id', 'name', 'u_cpf_new']], left_on='value_soldpdct', right_on='sys_id', how='left')
        df_billing_merged.drop(['custbill_u_install_base_item', 'soldpdct_consumer', 'value_custbill', 'value_soldpdct', 'sys_id_x', 'sys_id_y' ], axis=1, inplace=True)

        # 5º) Merge final comunicados, EQ e clientes
        df_final = pd.merge(merged_df, df_billing_merged, left_on='EQ Values', right_on='u_id_instalacao', how='inner')
        df_final['Data Importação'] = date.today().strftime('%Y-%m-%d')
        df_final.to_excel('f_Comunicados_Andamento.xlsx', index=False)

    except Exception as e:
        print(f"Erro: {e}")
        time.sleep(20)  # Espera 20 segundos antes de tentar novamente
        realizar_raspagem_e_enviar_email()  # Chama a função novamente após esperar

realizar_raspagem_e_enviar_email()