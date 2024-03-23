import pandas as pd
import win32com.client as win32
import openpyxl as openpyxl
import getpass
import requests
import tkinter as tk
from tkinter import messagebox

tabela_emails = pd.read_excel('banco_dados.xlsx')
tabela_manifestacoes = pd.read_excel('sc_req_item.xlsx')

def chamar_api_com_autenticacao(usuario, senha, url):
    try:
        auth = (usuario, senha)
        response = requests.get(url, auth=auth)
        if response.status_code == 200:
            dados = response.json()
            return dados
        else:
            print(f'Erro ao chamar a API: {response.status_code}')
    except requests.exceptions.RequestException as e:
        print(f'Erro na requisição: {str(e)}')
        
url_api = ''
usuario = ''
senha = ''
resultado = chamar_api_com_autenticacao(usuario, senha, url_api)
lista_emails_clientes = [item['custbill_u_billing_email'] for item in resultado['result']]
lista_emails_clientes_unicos = list(set(lista_emails_clientes))
Corpo_email = ''' 
        <p>Prezados</p>,

        <p> Corpo do Email </p>

        <p>Qualquer dúvida estou a disposição.</p>
'''
def enviar_emails(emails):
    outlook = win32.Dispatch('outlook.application')

    for destinatario in emails:
        mail = outlook.Createitem(0)
        mail.To = destinatario
        mail.Subject = 'Relatório de Vendas'
        mail.HTMLBody = Corpo_email
        mail.Send()
        
corpo_email_manifestacao = ''' 
        <p>Prezados</p>,

        <p> Corpo do Email </p>

        <p>Qualquer dúvida estou a disposição.</p>
'''

def enviar_emails_manifestacao(emails_manifestacao):
    outlook = win32.Dispatch('outlook.application')

    for destinatario in emails_manifestacao:
        mail = outlook.Createitem(0)
        mail.To = destinatario
        mail.Subject = 'Relatório de Vendas'
        mail.HTMLBody = corpo_email_manifestacao
        mail.Send()
        
emails_api = [
    ""
]
emails_excel = tabela_emails['Destinatários Onboard'].tolist()
emails_novos = [email for email in emails_api if email not in emails_excel]
emails_excel_manifestacao = tabela_manifestacoes['E-mail do Cliente'].tolist()
emails_excel_manifestacao_existentes = tabela_emails['Destinatários Manifestacao'].tolist()
emails_novos_manifestacao = [email for email in emails_excel_manifestacao if email not in emails_excel_manifestacao_existentes]
nome_usuario = getpass.getuser()
nome_usuario = nome_usuario.capitalize()
janela = tk.Tk()
janela.title('AutoMail by LC')
janela.geometry("433x540")
janela.configure(bg='#4373C9')
messagebox.showinfo("AutoMail by LC", "Antes de iniciar os envios logue na sua conta Outlook")
texto_orientacao = tk.Label(janela, text="Bem-vindo(a) ao AutoMail by LC", font=("Arial", 14, 'bold'), bg='#4373C9', fg='white')
texto_orientacao.grid(column=0, row=1, padx=30, pady=20)
contador_cliques = 0

def enviar_email_salvar_planilha():
    enviar_emails(emails_novos)
    global contador_cliques
    contador_cliques +=1
    nome_arquivo = 'banco_dados.xlsx'
    coluna_destinatarios = 'A'
    planilha = openpyxl.load_workbook(nome_arquivo)
    planilha_ativa = planilha.active
    proximo_indice = len(planilha_ativa[coluna_destinatarios]) + 1
    for email in emails_novos:
        planilha_ativa[coluna_destinatarios + str(proximo_indice)] = email
        proximo_indice += 1
    planilha.save(nome_arquivo)
    if emails_novos:
        messagebox.showinfo("AutoMail by LC", "Emails enviados com sucesso")
    else:
        messagebox.showinfo("AutoMail by LC", "Não há novos clientes para enviar emails.")
    if contador_cliques == 2 and emails_novos:
        messagebox.showinfo("AutoMail by LC", "Excesso de envios, fechando o aplicativo")
        janela.destroy()
        
contador_cliques_manifestacoes = 0

def enviar_email_salvar_planilha_manifestacao():
    enviar_emails_manifestacao(emails_novos_manifestacao)
    global contador_cliques_manifestacoes
    contador_cliques_manifestacoes +=1
    nome_arquivo = 'banco_dados.xlsx'
    coluna_destinatarios = 'B'
    planilha = openpyxl.load_workbook(nome_arquivo)
    planilha_ativa = planilha.active
    proximo_indice = len(planilha_ativa[coluna_destinatarios]) + 1
    for email in emails_novos_manifestacao:
        planilha_ativa[coluna_destinatarios + str(proximo_indice)] = email
        proximo_indice += 1
    planilha.save(nome_arquivo)
    if emails_novos_manifestacao:
        messagebox.showinfo("AutoMail by LC", "Emails enviados com sucesso")
    else:
        messagebox.showinfo("AutoMail by LC", "Não há novos clientes para enviar emails.")
    if contador_cliques_manifestacoes == 2 and emails_novos_manifestacao:
        messagebox.showinfo("AutoMail by LC", "Excesso de envios, fechando o aplicativo")
        outra_janela.destroy()
        
def abrir_outra_interface():
    global outra_janela
    janela.withdraw()
    outra_janela = tk.Tk()
    outra_janela.title('AutoMail by LC')
    outra_janela.geometry("433x540")
    outra_janela.configure(bg='#4373C9')
    texto_orientacao_outra = tk.Label(outra_janela, text="Bem-vindo(a) ao AutoMail by LC", font=("Arial", 14, 'bold'), bg='#4373C9', fg='white')
    texto_orientacao_outra.grid(column=0, row=1, padx=30, pady=20)
    texto_inicial_manifestacoes = tk.Label(outra_janela, text="Você está na interface de Envio Manifestações",font=("Arial", 12, 'bold'), bg='#4373C9', fg='white')
    texto_inicial_manifestacoes.grid(column=0, row=2, padx=30, pady=5)
    btn_primario = tk.Button(outra_janela, text="Ir para Envio OnBoard", command=voltar_para_janela_principal,font=("Arial", 10, 'bold'),bg='#4373C9', fg='white')
    btn_primario.grid(column=0, row=3, padx=30, pady=20)
    borda_texto_secundario = tk.Frame(outra_janela, bg='#4373C9', padx=5, pady=5)
    borda_texto_secundario.grid(column=0, row=4, padx=15, pady=5)
    canvas = tk.Canvas(borda_texto_secundario, height=0.001, bg='white')
    canvas.pack(fill='x')
    texto_secundario = tk.Label(outra_janela,text="{}, para iniciar o envio clique no botão a seguir: ".format(nome_usuario),font=("Arial", 12), bg='#4373C9', fg='white')
    texto_secundario.grid(column=0, row=5, padx=30, pady=5)
    btn_enviar_manifestacoes = tk.Button(outra_janela, text="Enviar Emails Manifestação", command=enviar_email_salvar_planilha_manifestacao,font=("Arial", 12, 'bold'), bg='#4373C9', fg='white')
    btn_enviar_manifestacoes.grid(column=0, row=6, padx=30, pady=20)
    texto_terceiro = tk.Label(outra_janela, text="{}, o corpo do email terá o seguinte formato: ".format(nome_usuario),font=("Arial", 12), bg='#4373C9', fg='white')
    texto_terceiro.grid(column=0, row=8, padx=30, pady=10)
    texto_quarto = tk.Label(outra_janela, text=corpo_email_manifestacao, font=("Arial", 12), bg='#4373C9', fg='white')
    texto_quarto.grid(column=0, row=9, padx=30, pady=5)
    
def voltar_para_janela_principal():
    global outra_janela
    outra_janela.destroy()
    janela.deiconify()
    
texto_inicial_onboard = tk.Label(janela, text="Você está na interface de Envio OnBoard",font=("Arial", 12, 'bold'), bg='#4373C9', fg='white')
texto_inicial_onboard.grid(column=0, row=2, padx=30, pady=5)
btn_secundario = tk.Button(janela, text="Ir para Envio Manifestações", command=abrir_outra_interface, font=("Arial", 10, 'bold'),bg='#4373C9', fg='white')
btn_secundario.grid(column=0, row=3, padx=30, pady=20)
borda_texto_inicial = tk.Frame(janela, bg='#4373C9', padx=5, pady=5)
borda_texto_inicial.grid(column=0, row=4, padx=15, pady=5)
canvas = tk.Canvas(borda_texto_inicial, height=0.001, bg='white')
canvas.pack(fill='x')
texto_email = tk.Label(janela, text="{}, para iniciar o envio clique no botão a seguir: ".format(nome_usuario), font=("Arial", 12), bg='#4373C9', fg='white')
texto_email.grid(column=0, row=5, padx=30, pady=5)
btn_enviar = tk.Button(janela, text="Enviar Emails OnBoard", command=enviar_email_salvar_planilha, font=("Arial", 12, 'bold'), bg='#4373C9', fg='white')
btn_enviar.grid(column=0, row=6, padx=30, pady=20)
texto_email = tk.Label(janela, text="{}, o corpo do email terá o seguinte formato: ".format(nome_usuario),font=("Arial", 12), bg='#4373C9', fg='white')
texto_email.grid(column=0, row=8, padx=30, pady=10)
texto_email = tk.Label(janela, text=Corpo_email, font=("Arial", 12), bg='#4373C9', fg='white')
texto_email.grid(column=0, row=9, padx=30, pady=5)
janela.mainloop()