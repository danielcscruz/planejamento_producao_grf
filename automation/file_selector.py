import os
from InquirerPy import inquirer


def escolher_arquivo_excel():
    xlsx_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

    if not xlsx_files:
        raise FileNotFoundError("Nenhum arquivo .xlsx encontrado no diretório atual.")

    file_choice = inquirer.select(
        message="Escolha um arquivo Excel:",
        choices=xlsx_files
    ).execute()

    print(f"\nVocê escolheu: {file_choice}")
    return file_choice
