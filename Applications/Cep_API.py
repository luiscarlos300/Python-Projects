import requests
import pandas as pd
import time

def consulta_cep_brasilapi(cep):
    url = f"https://brasilapi.com.br/api/cep/v2/{cep}"
    try:
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        json_content = response.json()
        return json_content
    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição para {url}: {e}")
        return None

ceps = [ "66640030","66117080" ]
resultados = []
ceps_por_requisicao = 2000
ceps_por_planilha = 2000

for i in range(0, len(ceps), ceps_por_requisicao):
    ceps_lote = ceps[i:i + ceps_por_requisicao]
    for cep in ceps_lote:
        resultado = consulta_cep_brasilapi(cep)
        if resultado:
            resultados.append(resultado)

    if (i + ceps_por_requisicao) % ceps_por_planilha == 0:
        df = pd.DataFrame(resultados)
        caminho_planilha = f'C:/Users/Downloads/AnaliseCepBrasilApi_{i // ceps_por_requisicao + 1}.xlsx'
        df.to_excel(caminho_planilha, index=False)
        print(f"Planilha salva em: {caminho_planilha}")
        resultados = []
    time.sleep(2)
if resultados:
    df = pd.DataFrame(resultados)
    caminho_planilha = f'C:/Users/Downloads/AnaliseCepBrasilApi_last.xlsx'
    df.to_excel(caminho_planilha, index=False)
    print(f"Última planilha salva em: {caminho_planilha}")