import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from datetime import date
import openpyxl
from bs4 import BeautifulSoup
import pandas as pd
import time

data_list = []

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
    time.sleep(15)
    button_filter = browser.find_element(By.XPATH, "//button[contains(@class, 'q-btn--round') and .//i[contains(@class, 'q-icon') and text()='filter_alt']]")
    time.sleep(5)

    # Status
    encerrada_button = browser.find_element(By.XPATH, "//div[contains(@class, 'q-toggle') and @aria-label='Encerrada']")

    # Tipo
    bloqueio_button = browser.find_element(By.XPATH, "//div[contains(@class, 'q-toggle') and @aria-label='Bloqueio total']")
    desbloqueio_button = browser.find_element(By.XPATH, "//div[contains(@class, 'q-toggle') and @aria-label='Desbloqueio']")

    # Aplicar
    button_apply_filters = browser.find_element(By.XPATH, "//button[contains(@class, 'q-btn--outline') and .//span[@class='block' and text()='Aplicar filtros']]")

    button_filter.click()
    encerrada_button.click()
    bloqueio_button.click()
    desbloqueio_button.click()
    button_apply_filters.click()
    time.sleep(5)

    next_button_first = browser.find_element(By.XPATH, "//button[contains(@class, 'q-btn') and .//i[contains(@class, 'q-icon') and text()='keyboard_arrow_right']]")
    contador = 0
    while next_button_first.is_enabled() and contador <= 150:
        # Criar a sopa (BeautifulSoup) para a página atual
        soup = BeautifulSoup(browser.page_source, 'html.parser')

        # Encontrar todas as linhas na tabela
        rows = soup.find_all('tr', class_='')

        # Iterar sobre as linhas e extrair os dados
        for row in rows:
            cells = row.find_all('td', class_=['cursor-pointer', 'q-td', 'text-left', 'text-center'])
            data = [cell.get_text(strip=True) for cell in cells]
            data_list.append(data)

        # Aguardar até que o elemento de carregamento desapareça
        wait = WebDriverWait(browser, 20)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))

        # Clicar no botão "Próxima página"
        next_button_first.click()
        contador += 1

        # Aguardar um tempo para a próxima página carregar completamente
        time.sleep(30)

        # Aguardar até que o elemento de carregamento desapareça novamente
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))

        # Encontrar novamente o botão "Próxima página" para a próxima iteração
        next_button_first = browser.find_element(By.XPATH, "//button[contains(@class, 'q-btn') and .//i[contains(@class, 'q-icon') and text()='keyboard_arrow_right']]")

except Exception as e:
    print(f"Erro: {e}")

columns = ['ID', 'Tipo', 'ID da Casa Conectada', 'Status', 'Ordem relacionada', 'Data/Hora Criação', 'Usuário Criação']
df = pd.DataFrame(data_list,columns=columns)
df['Data Input'] = datetime.datetime.today()
df.to_excel('OS_DBR_Raspagem_Encerrado.xlsx', index=False)