import os
import fnmatch
import shutil
import PyPDF2


def contains_keywords(pdf_path, keywords):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text = page.extract_text()
                if any(keyword in text for keyword in keywords):
                    return True
        return False
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return False


def find_and_move_pdfs_with_keywords(directory, keywords, destination):
    if not os.path.exists(destination):
        os.makedirs(destination)

    for root, dirs, files in os.walk(directory):
        for file in fnmatch.filter(files, '*.pdf'):
            pdf_path = os.path.join(root, file)
            if contains_keywords(pdf_path, keywords):
                dest_path = os.path.join(destination, file)
                shutil.move(pdf_path, dest_path)
                print(f"Moved {pdf_path} to {dest_path}")


# Exemplo de uso:
directory = 'C:/Users//Desktop'  # Diretório onde os PDFs serão verificados
destination = 'C:/Users//Desktop/Licenças'  # Diretório para onde os PDFs serão movidos
keywords = ['Firewall', 'firewall']  # Lista de palavras-chave

find_and_move_pdfs_with_keywords(directory, keywords, destination)
