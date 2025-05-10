import pandas as pd
import os
from InquirerPy import inquirer
from datetime import datetime
from automation.ui.table_renderer import processar_tabela

def excluir_pedido(arquivo_path: str):
    arquivo_path = os.path.join('exp', arquivo_path)

    # Usa a função processar_tabela para carregar e formatar o DataFrame
    df_formatado, _ = processar_tabela(arquivo_path)

    # Cria identificadores únicos baseados nas colunas relevantes
    df_formatado["Identificador"] = (
        df_formatado["PEDIDO"].astype(str) + " - " +
        df_formatado["CLIENTE"] + " - " +
        df_formatado["PRODUTO"]
    )
    unique_options = df_formatado["Identificador"].drop_duplicates().tolist()

    # Usuário seleciona qual produção excluir
    choice = inquirer.select(
        message="Selecione a produção para excluir:",
        choices=unique_options,
    ).execute()

    # Filtra os índices correspondentes
    indices_to_remove = df_formatado[df_formatado["Identificador"] == choice].index

    if len(indices_to_remove) == 0:
        print("Nenhuma linha encontrada para essa seleção.")
    else:
        df_formatado.drop(indices_to_remove, inplace=True)

        # Remove a coluna auxiliar antes de salvar
        df_formatado.drop(columns=["Identificador"], inplace=True)

        # Salva o novo arquivo Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"planejamentoproducao_atualizado_{timestamp}.xlsx"
        df_formatado.to_excel(output_file, index=False)

        print(f"Produção removida com sucesso! Arquivo salvo como {output_file}")