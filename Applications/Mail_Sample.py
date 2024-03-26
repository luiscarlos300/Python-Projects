import pandas as pd
import win32com.client as win32
import tkinter as tk

tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

outlook = win32.Dispatch('outlook.application')
mail = outlook.Createitem(0)
mail.To = ''
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = ''' 
<p>Prezados</p>,

<p>Segue o Relatório de Vendas.</p>

<p>Faturamento:<p/>
{}

<p>Quantidade Vendida:</p>
{}

<p>Ticket Médio:</p>
{}

<p>Qualquer dúvida estou a disposição.</p>
'''.format(faturamento.to_html(),quantidade.to_html(),ticket_medio.to_html())

janela = tk.Tk()
janela.title('Enviar Email')
janela.geometry("200x200")

texto_orientacao = tk.Label(janela, text="Clique em no botão para enviar o email", font = ("Arial",15))
texto_orientacao.grid(column=0, row=1, padx=50, pady =50)

def enviar_click():
    mail.Send()

btn_enviar = tk.Button(janela, text="Enviar", command=enviar_click, font = ("Arial",15))
btn_enviar.grid(column=0, row=5, padx=50, pady =50)

texto = ""

texto_email = tk.Label(janela, text=texto, font = ("Arial",15))
texto_email.grid(column=0, row=6, padx=50, pady =50)
janela.mainloop()