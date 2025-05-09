import os
from InquirerPy import inquirer

from automation.core.constants import DEFAULT_EXP_PATH


def escolher_arquivo_excel():
    xlsx_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

    if not xlsx_files:
        raise FileNotFoundError("Nenhum arquivo .xlsx encontrado no diretório atual.")

    # Adiciona a opção de "Voltar" no final da lista
    xlsx_files.append("🔙 Voltar")

    file_choice = inquirer.select(
        message="Escolha um arquivo Excel:",
        choices=xlsx_files
    ).execute()

    if file_choice == "🔙 Voltar":
        print("\n🔙 Retornando ao menu anterior...\n")
        return None  # Ou implemente a lógica para voltar ao menu anterior

    print(f"\nVocê escolheu: {file_choice}")
    return file_choice

def escolher_arquivo_exportar():
    xlsx_files = [f for f in os.listdir(DEFAULT_EXP_PATH) if f.endswith('.xlsx')]

    if not xlsx_files:
        raise FileNotFoundError("Nenhum arquivo .xlsx encontrado no diretório atual.")

    # Adiciona a opção de "Voltar" no final da lista
    xlsx_files.append("🔙 Voltar")

    file_choice = inquirer.select(
        message="Escolha um arquivo Excel:",
        choices=xlsx_files
    ).execute()

    if file_choice == "🔙 Voltar":
        print("\n🔙 Retornando ao menu anterior...\n")
        return None  # Ou implemente a lógica para voltar ao menu anterior

    print(f"\nVocê escolheu: {file_choice}")
    return file_choice