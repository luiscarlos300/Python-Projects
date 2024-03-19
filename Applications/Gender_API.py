import pandas as pd
from unidecode import unidecode

def limpar_nome(nome):
    return unidecode(str(nome)).upper()
try:
    df1 = pd.read_excel('')
    df2 = pd.read_excel('')
except Exception as e:
    print(f"Erro ao carregar os arquivos: {e}")
    exit()

df1['Primeiro Nome'] = df1['Primeiro Nome'].apply(limpar_nome)
df2 = df2.rename(columns={'first_name': 'Primeiro Nome'})
colunas_necessarias = ['Primeiro Nome', 'classification']
if not set(colunas_necessarias).issubset(df2.columns):
    print("Colunas necessárias não encontradas na segunda base de dados.")
    exit()
df3 = pd.merge(df1, df2[colunas_necessarias], how='left', on='Primeiro Nome')
df3 = df3.rename(columns={'classification': 'Genero'})
try:
    df3.to_excel('', index=False)
    print("Processo concluído com sucesso!")
except Exception as e:
    print(f"Erro ao salvar a nova base de dados: {e}")
