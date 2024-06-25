from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.edge.options import Options
from bs4 import BeautifulSoup
import sys
import paramiko
import time
import win32com.client
import os

#opcoes = Options()
#opcoes.add_argument('--headless')
browser = webdriver.Edge()#options=opcoes)
browser.get('')
username = browser.find_element("name", "Ecom_User_ID")
password = browser.find_element("name", "Ecom_Password")
login_button = browser.find_element("id", "submit")
username.send_keys("")
password.send_keys("")
login_button.click()
time.sleep(10)
browser.get('')

close_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'q-btn') and contains(@class, 'q-btn--flat') and contains(@class, 'q-btn--round') and contains(@class, 'q-btn--actionable') and .//i[text()='close']]")))
close_button.click()

nova_os_tab = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'q-tab__label') and text()='Nova OS']")))
nova_os_tab.click()

dropdown_icon = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'q-field__append q-field__marginal row no-wrap items-center q-anchor--skip')]//i[text()='arrow_drop_down']")))
time.sleep(5) # Para selecionar o dropwdown correto!
dropdown_icon.click()

desbloqueio_option = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='q-item__label' and text()='Desbloqueio']")))
desbloqueio_option.click()

check_box = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'q-checkbox__bg')]")))
check_box.click()

input_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,"//div[@class='row']/label/div[@class='q-field__inner relative-position col self-stretch']/div[@class='q-field__control relative-position row no-wrap text-deep-purple-4']/div[@class='q-field__control-container col relative-position row no-wrap q-anchor--skip']/input[@class='q-field__native q-placeholder']")))
input_button.click()

try:
    if getattr(sys, 'frozen', False):
        sys.argv = [sys.executable] + sys.argv
    host = ''
    port = 22
    username = ''
    password = ''
    local_directory = ''
    remote_directory = '/dados/relatorios'
    def download_latest_file():
        try:
            transport = paramiko.Transport((host, port))
            transport.connect(username=username, password=password)
            sftp = paramiko.SFTPClient.from_transport(transport)
            files = [file for file in sftp.listdir(remote_directory) if "LASTDESBLOQ" in file]
            latest_file = None
            latest_timestamp = None

            for file in files:
                timestamp_str = file.split('_')[2].split('.')[0]
                timestamp = int(timestamp_str)

                if latest_timestamp is None or timestamp > latest_timestamp:
                    latest_file = file
                    latest_timestamp = timestamp

            if latest_file:
                remote_file_path = remote_directory + '/' + latest_file
                local_file_path = local_directory + '/' + latest_file
                sftp.get(remote_file_path, local_file_path)

            sftp.close()
            transport.close()

        except Exception as e:
            print(f"Erro ao baixar arquivos FTP: {e}")

    if __name__ == '__main__':
        download_latest_file()

except Exception as e:
    print(f'Erro pipeline FTP: {e}')


directory_path = r''
files = os.listdir(directory_path)

latest_timestamp = None
latest_file = None

for file in files:
    if file.endswith(".csv"):
        try:
            filename_without_extension = file[:-4]
            parts = filename_without_extension.split('_')
            timestamp_str = parts[-1]
            timestamp = int(timestamp_str)

            if latest_timestamp is None or timestamp > latest_timestamp:
                latest_file = file
                latest_timestamp = timestamp
        except (IndexError, ValueError):
            continue

if latest_file:
    atual = latest_file
    print("Latest file:", latest_file)
    file_path = os.path.join(directory_path, atual)
    input_file = WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class, 'q-uploader') and contains(@class, 'column') and contains(@class, 'no-wrap') and contains(@class, 'full-width') and contains(@class, 'no-shadow') and @style='height: 100px; border: 2px dashed rgb(0, 172, 193);']//div[@class='q-uploader__list scroll']//div[contains(@class, 'fit') and contains(@class, 'cursor-pointer')]//input[@type='file' and contains(@class, 'q-uploader__input') and contains(@class, 'overflow-hidden') and contains(@class, 'absolute-full')]")))
    input_file.send_keys(file_path)
else:
    print("No CSV files found in the directory.")

ok_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'q-btn--flat') and .//span[text()='OK']]")))
ok_button.click()
time.sleep(5)

send_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'q-btn') and .//span[text()='Enviar']]")))
#send_button.click()