from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
import threading

def buscar_grupos(browser, grupos):
    lista = grupos.splitlines()

    counter = 0

    for item in range(len(lista)):
        dropdown = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.ID, "id_grupo"))
        )
        select = Select(dropdown)
        select.select_by_visible_text(lista[item])

        buscar_button = WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(normalize-space(text()), 'Buscar')]")
            )
        )
        buscar_button.click()
        time.sleep(15)
        export_button = WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[@aria-label='Export']")
            )
        )
        export_button.click()
        csv_option = WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[@data-type='csv']")
            )
        )
        csv_option.click()
        time.sleep(10)

        counter += 1

        if counter == len(lista):
            break

    browser.quit()

def baixar_bases(usuario, senha, data_inicial, data_final,lista_funcao, progress):

    progress.start(40)

    messagebox.showinfo("Iniciar", "Operação iniciada, aguarde.")
    options = webdriver.EdgeOptions()
    options.add_argument("--headless=new")
    browser = webdriver.Edge(options=options)
    browser.get('')
    username = browser.find_element('name', 'username')
    password = browser.find_element('name', 'password')
    login_button = browser.find_element(By.CLASS_NAME, 'btn')
    username.send_keys(usuario)
    password.send_keys(senha)
    login_button.click()
    browser.get('')

    date_from_element = browser.find_element(By.ID, 'id_date_from')
    browser.execute_script(
        f"arguments[0].setAttribute('value', '{str(data_inicial) + ' 00:00:00'}')",
        date_from_element
    )

    date_to_element = browser.find_element(By.ID, 'id_date_till')
    browser.execute_script(
        f"arguments[0].setAttribute('value', '{str(data_final) + ' 23:59:59' }')",
        date_to_element
    )

    buscar_grupos(browser=browser, grupos=lista_funcao)

    progress.stop()
    messagebox.showinfo("Finalizado", "Operação finalizada.")

def iniciar_thread(userInput, userInput_key, cal_init, cal_final, texto_grupos, progress):
    threading.Thread(target=lambda: baixar_bases(
        usuario=userInput.get(),
        senha=userInput_key.get(),
        data_inicial=cal_init.get_date(),
        data_final=cal_final.get_date(),
        lista_funcao=texto_grupos.get("1.0", tk.END),
        progress=progress
    ), daemon=True).start()

def criar_interface():
    root = tk.Tk()
    root.title("NOC")
    root.geometry("665x600")
    root.config(bg="sky blue")

    # Configurar o frame da esquerda
    leftFrame = tk.Frame(root, bg="sky blue", height=700, width=300)
    leftFrame.grid(row=0, column=0, padx=10, pady=2, columnspan=2, rowspan=8, sticky="n")
    leftFrame.grid_propagate(False)

    # Título
    title_label = tk.Label(leftFrame, text="BAIXAR RELATÓRIOS", font=("Arial", 12, "bold"), bg="sky blue",
                           anchor="w")
    title_label.grid(row=0, column=0, padx=20, pady=(50, 20), sticky="w")

    # Campo de Login
    user = tk.Label(leftFrame, text="Login", font=("Arial", 11), bg="sky blue", anchor="w")
    user.grid(row=1, column=0, padx=20, pady=(10, 5), sticky="w")

    userInput = tk.Entry(leftFrame, width=30, font=("Arial", 11))
    userInput.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="w")

    # Campo de Senha
    user_key = tk.Label(leftFrame, text="Senha", font=("Arial", 11), bg="sky blue", anchor="w")
    user_key.grid(row=3, column=0, padx=20, pady=(10, 5), sticky="w")

    userInput_key = tk.Entry(leftFrame, width=30, show="*",font=("Arial", 11))
    userInput_key.grid(row=4, column=0, padx=20, pady=(5, 20), sticky="w")

    Grupos = """META_EQTAL-Agência
META_EQTAL-Escritório
META_EQTAL-Operação
META_EQTAL-Outros"""

    texto_label = tk.Label(leftFrame, text="Grupos", font=("Arial", 11), bg="sky blue", anchor="w")
    texto_label.grid(row=5, column=0, padx=20, pady=(5, 5), sticky="w")

    texto_grupos = tk.Text(leftFrame, height=15, width=33)
    texto_grupos.grid(row=6, column=0, columnspan=2, padx=15, pady=(5, 5))
    texto_grupos.insert(tk.END, Grupos)

    # Configurar o frame da direita (se necessário)
    rightFrame = tk.Frame(root, bg="gainsboro", height=595, width=331)
    rightFrame.grid(row=0, column=2, padx=10, pady=2, columnspan=2, rowspan=8, sticky="n")
    rightFrame.grid_propagate(False)

    # Configurar o layout de colunas no frame para centralizar
    rightFrame.grid_columnconfigure(0, weight=1)  # Coluna central
    rightFrame.grid_columnconfigure(1, weight=1)  # Coluna direita

    # Seleção de data inicial
    date_init = tk.Label(rightFrame, text="SELECIONE A DATA INICIAL", font=("Arial", 10), bg="gainsboro")
    date_init.grid(row=0, column=0, padx=10, pady=(20, 5), sticky="n", columnspan=2)

    cal_init = Calendar(rightFrame, selectmode='day', year=2024, month=11, day=28, locale='pt_br')
    cal_init.grid(row=1, column=0, padx=10, pady=(5, 20), sticky="n", columnspan=2)

    # Seleção de data final
    date_final = tk.Label(rightFrame, text="SELECIONE A DATA FINAL", font=("Arial", 10), bg="gainsboro")
    date_final.grid(row=2, column=0, padx=10, pady=(5, 5), sticky="n", columnspan=2)

    cal_final = Calendar(rightFrame, selectmode='day', year=2024, month=11, day=28, locale='pt_br')
    cal_final.grid(row=3, column=0, padx=10, pady=(5, 20), sticky="n", columnspan=2)

    # Botão de iniciar
    init_button = tk.Button(rightFrame, text="INICIAR DOWNLOAD", width=20, height=2,
                            command=lambda: iniciar_thread(userInput, userInput_key, cal_init, cal_final, texto_grupos, progress),
                            bg="sky blue")
    init_button.grid(row=4, column=0, padx=6, pady=(6, 5), sticky="n", columnspan=2)

    # Barra de progresso
    progress = ttk.Progressbar(rightFrame, orient='horizontal', length=300, mode='indeterminate')
    progress.grid(row=5, column=0, padx=6, pady=(6, 5), sticky="n", columnspan=2)
    root.mainloop()

if __name__ == "__main__":
    criar_interface()