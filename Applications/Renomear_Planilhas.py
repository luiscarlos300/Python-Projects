import os
import random

def rename_excel_files(folder_path):
    files = os.listdir(folder_path)
    excel_files = [f for f in files if f.endswith('.xlsx') or f.endswith('.xlsm')]

    for i, file_name in enumerate(excel_files, start=1):
        num_random = random.random()
        new_file_name = f'{i}{num_random}.xlsx'
        old_file_path = os.path.join(folder_path, file_name)
        new_file_path = os.path.join(folder_path, new_file_name)

        os.rename(old_file_path, new_file_path)
        print(f'Renomeado: {file_name} para {new_file_name}')


folder_path = 'C'
rename_excel_files(folder_path)
