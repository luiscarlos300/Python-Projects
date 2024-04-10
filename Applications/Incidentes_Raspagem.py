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
import tkinter as tk
from tkinter import messagebox

data_list = []

def show_success_message():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Sucesso", "Os dados foram salvos com sucesso em 'Incidentes.xlsx'")
    root.destroy()

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
    button_incident = browser.find_element(By.XPATH, "//div[contains(@class, 'q-tab__content') and .//div[@class='q-tab__label' and text()='Acompanhe seu incidente']]")
    button_incident.click()
    time.sleep(25)
    soup = BeautifulSoup(browser.page_source, 'html.parser')
    next_button_first = browser.find_element(By.XPATH,"//button[contains(@class, 'q-btn') and .//i[contains(@class, 'q-icon') and text()='keyboard_arrow_right']]")
    while next_button_first.is_enabled():
        rows = soup.find_all('tr', class_='')
        for row in rows:
            cells = row.find_all('td', class_=['cursor-pointer', 'q-td', 'text-left', 'text-center'])
            data = [cell.get_text(strip=True) for cell in cells]
            data_list.append(data)
        wait = WebDriverWait(browser, 20)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))
        next_button_first.click()
        time.sleep(30)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'q-loading')))
except Exception as e:
    print(f"Erro: {e}")

columns = ['ID do chamado', 'Status', 'Funcionalidade/Serviço', 'Data Criação']
df = pd.DataFrame(data_list,columns=columns)
df['Data Input'] = datetime.datetime.today()
df.to_excel('Incidentes.xlsx', index=False)
show_success_message()