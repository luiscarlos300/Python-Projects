import os
import fnmatch
import shutil


def find_and_copy_pdfs(root_directory, destination_directory):
    if not os.path.exists(destination_directory):
        os.makedirs(destination_directory)

    for root, dirs, files in os.walk(root_directory):
        for file in fnmatch.filter(files, '*.pdf'):
            source_path = os.path.join(root, file)
            destination_path = os.path.join(destination_directory, file)
            shutil.copy2(source_path, destination_path)
            print(f'Copied {source_path} to {destination_path}')


# Exemplo de uso:
root_directory = '/Faturamento'  # Diretório de origem
destination_directory = 'C:/Users//Desktop/Notas'  # Diretório de destino

find_and_copy_pdfs(root_directory, destination_directory)
