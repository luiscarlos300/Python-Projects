from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import datetime
import random
import openpyxl
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
    time.sleep(5)
    browser.get("")

    # Close message
    close_button = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH,"//button[contains(@class, 'q-btn') and contains(@class, 'q-btn--flat') and contains(@class, 'q-btn--round') and contains(@class, 'q-btn--actionable') and .//i[text()='close']]"))
    )
    close_button.click()

    # Filter
    button_filter = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'q-btn') and .//i[text()='filter_alt']]"))
    )
    button_filter.click()

    # Type
    retirada_button = WebDriverWait(browser,10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'q-toggle') and @aria-label='Retirada']"))
    )
    retirada_button.click()

    # Apply type
    button_apply_filters = WebDriverWait(browser,10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'q-btn--outline') and .//span[@class='block' and text()='Aplicar filtros']]"))
    )
    button_apply_filters.click()
    time.sleep(3)
    button_filter.click()

    # Click dropicon page
    dropdown_icon = WebDriverWait(browser, 15).until(
        EC.presence_of_element_located((By.XPATH,"//div[contains(@class, 'q-field__append')]//i[contains(@class, 'q-select__dropdown-icon') and text()='arrow_drop_down']"))
    )
    dropdown_icon.click()

    # Select 50
    option_50 = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='q-virtual-scroll__content']//div[@class='q-item__label' and text()='50']"))
    )
    option_50.click()
    time.sleep(10)

    next_button_first = WebDriverWait(browser,10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'q-btn') and .//i[contains(@class, 'q-icon') and text()='keyboard_arrow_right']]"))
    )

    num_pages = 2
    pages = 0
    while next_button_first.is_enabled() and pages <= num_pages:
        soup = BeautifulSoup(browser.page_source, 'html.parser')
        rows = soup.find_all('tr', class_='')
        for row in rows:
            cells = row.find_all('td', class_=['cursor-pointer', 'q-td', 'text-left', 'text-center'])
            data = [cell.get_text(strip=True) for cell in cells]
            data_list.append(data)
        wait = WebDriverWait(browser, 20)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))
        next_button_first.click()
        pages += 1
        time.sleep(10)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))
        next_button_first = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH,"//button[contains(@class, 'q-btn') and .//i[contains(@class, 'q-icon') and text()='keyboard_arrow_right']]"))
        )

except Exception as e:
    print(f"Erro encontrado: {e}")
    time.sleep(15)

columns = ['ID', 'Tipo', 'ID da Casa Conectada', 'Status', 'Ordem relacionada', 'Data/Hora Criação', 'Usuário Criação']
df = pd.DataFrame(data_list,columns=columns)
df_filter = df[df['ID'].notnull()]
df_filter['Data/Hora Raspagem'] = datetime.datetime.today()
df_filter.to_excel('OS_Retirada_Raspagem.xlsx', index=False)