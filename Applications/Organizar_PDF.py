import pdfplumber
import pandas as pd
import os

# Definir o caminho da pasta contendo os arquivos PDF
pdf_folder_path = 'C:/Users/Desktop/Equipamentos/TIPO 1'  # Ajuste para o caminho correto da pasta

# Inicializar uma lista para armazenar os dados
data2 = []


# Função para verificar e extrair os dados das linhas
def extract_data_from_line(line):
    parts = line.split()
    if len(parts) >= 14:
        try:
            # Extrair as colunas fixas
            codigo = parts[0]
            ncm_sh = parts[-13]
            cst = parts[-12]
            cfop = parts[-11]
            unid = parts[-10]
            qtd = parts[-9]
            vlr_unit = parts[-8]
            vlr_total = parts[-7]
            bc_icms = parts[-6]
            vlr_icms = parts[-5]
            vlr_ipi = parts[-4]
            aliq_icms = parts[-3]
            aliq_ipi = parts[-2]

            # A descrição é o texto entre o código e o NCM/SH
            descricao = ' '.join(parts[1:-13])

            return [codigo, ncm_sh, cst, cfop, unid, qtd, vlr_unit, vlr_total, bc_icms, vlr_icms, vlr_ipi,
                    aliq_icms, aliq_ipi, descricao]
        except Exception as e:
            print(f"Error processing line: {line}\nError: {e}")
            return None
    return None


# Iterar sobre todos os arquivos na pasta
for pdf_file in os.listdir(pdf_folder_path):
    if pdf_file.endswith('.pdf'):
        pdf_path = os.path.join(pdf_folder_path, pdf_file)

        # Abrir o PDF com pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            # Iterar sobre cada página do PDF
            for page in pdf.pages:
                text = page.extract_text()

                # Verificar a estrutura do PDF
                if "CÓDIGO" in text and "DESCRIÇÃO DO PRODUTO/SERVIÇO" in text:
                    # Extrair os dados de interesse
                    lines = text.split('\n')
                    for line in lines:
                        if "CÓDIGO" not in line and "DESCRIÇÃO DO PRODUTO/SERVIÇO" not in line:
                            data = extract_data_from_line(line)
                            if data:
                                data2.append(data)

# Converter a lista em um dataframe
df2 = pd.DataFrame(data2, columns=['CÓDIGO', 'NCM/SH', 'CST', 'CFOP', 'UNID.', 'QTD.',
                                   'VLR. UNIT.', 'VLR. TOTAL', 'BC ICMS', 'VLR. ICMS', 'VLR. IPI', 'ALÍQ. ICMS',
                                   'ALÍQ. IPI','DESCRIÇÃO DO PRODUTO/SERVIÇO'])

# Salvar o dataframe em um arquivo CSV
output_csv_path2 = os.path.join(pdf_folder_path, 'output2.csv')
df2.to_csv(output_csv_path2, index=False)

# Exibir o dataframe
print(df2)
