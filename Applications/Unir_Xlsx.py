import os
import pandas as pd
import warnings

def unir_planilhas(pasta_origem, arquivo_saida_base, max_linhas=1000000):
    # Lista todos os arquivos na pasta de origem
    arquivos = [f for f in os.listdir(pasta_origem) if f.endswith('.xlsx') or f.endswith('.xls')]

    contador_linhas = 0
    numero_arquivo_saida = 1
    writer = None

    # Para cada arquivo, lê cada planilha e processa os dados
    for arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta_origem, arquivo)
        try:
            with warnings.catch_warnings(record=True):
                warnings.simplefilter("ignore")
                xls = pd.ExcelFile(caminho_arquivo)
                for planilha in xls.sheet_names:
                    # Verifica o número de colunas na planilha atual
                    df_temp = pd.read_excel(xls, planilha, nrows=0)
                    num_colunas = min(17, len(df_temp.columns))

                    if num_colunas == 0:
                        continue  # Ignora planilhas vazias

                    df = pd.read_excel(xls, planilha, usecols=range(num_colunas))  # Considera as colunas disponíveis

                    # Se não há writer ativo ou o limite de linhas foi atingido, cria um novo writer
                    if writer is None or contador_linhas + len(df) > max_linhas:
                        if writer is not None:
                            writer.close()
                        nome_arquivo_saida = f"{arquivo_saida_base}_{numero_arquivo_saida}.xlsx"
                        writer = pd.ExcelWriter(nome_arquivo_saida, engine='openpyxl')
                        numero_arquivo_saida += 1
                        contador_linhas = 0

                    # Nomeia a planilha com base no nome do arquivo e da planilha original
                    nome_planilha = f"{os.path.splitext(arquivo)[0]}_{planilha}"
                    if len(nome_planilha) > 31:
                        nome_planilha = nome_planilha[:31]

                    # Escreve a planilha no arquivo de saída atual
                    df.to_excel(writer, sheet_name=nome_planilha, index=False)
                    contador_linhas += len(df)
        except Exception as e:
            print(f"Erro ao ler o arquivo {arquivo}: {e}")

    # Fecha o último writer
    if writer is not None:
        writer.close()


if __name__ == "__main__":
    pasta_origem = r''
    arquivo_saida_base = r''
    unir_planilhas(pasta_origem, arquivo_saida_base)
    print(f"As planilhas foram unidas nos arquivos com base {arquivo_saida_base}")
