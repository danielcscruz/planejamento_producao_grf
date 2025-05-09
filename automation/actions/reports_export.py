import pandas as pd
from openpyxl import load_workbook
import os

from automation.core.constants import SETOR_ORDEM

def gerar_relatorio_arquivo(arquivo_path: str):
    print("Iniciou função de exportar arquivo")
    print(f"Arquivo base: {arquivo_path}")

    # Define o diretório 'exp/' antes do nome do arquivo
    arquivo_full_path = os.path.join('exp', os.path.basename(arquivo_path))
    print(f"Arquivo completo com rota: {arquivo_full_path}")

    # Define os setores a serem processados
    setores = SETOR_ORDEM

    # Define o diretório onde os arquivos serão salvos
    nome_base = os.path.splitext(os.path.basename(arquivo_full_path))[0]  # Ex: 'planejamentoproducaogrf'
    output_dir = os.path.join(os.path.dirname(arquivo_full_path), nome_base)
    os.makedirs(output_dir, exist_ok=True)

    # Carrega o workbook com os valores já calculados (não as fórmulas)
    wb = load_workbook(arquivo_full_path, data_only=True)
    ws = wb.active
    # Lê a planilha base
    # Lê os dados da planilha em forma de lista de listas
    data = [[cell.value for cell in row] for row in ws.iter_rows()]
    df_base = pd.DataFrame(data)

    for setor in setores:
        print(f"Processando setor: {setor}")
        df_setor = df_base.copy()

        # Exclui linhas 3 a 11 (índices 2 a 10)
        df_setor.drop(index=range(2, 11), inplace=True, errors='ignore')

        # Preenche a primeira linha a partir da coluna G (índice 6)
        df_setor.iloc[0, 6:] = df_setor.iloc[0, 6:].fillna("")

        # Preserva título/subtítulo (linhas 0 e 1)
        df_head = df_setor.iloc[1:2]  # Seleciona apenas a linha 2 (índice 1)
        df_body = df_setor.iloc[2:]

        # Aplica o filtro com a nova condicional
        df_filtered = df_body[
            (df_body[6] == setor) &  # Coluna 6 deve ser igual ao setor
            (df_body[0].notna() | df_body[4].notna())  # Coluna A (0) ou E (4) deve estar preenchida
        ]

        # Junta título + linhas filtradas
        df_final = pd.concat([df_head, df_filtered])

        # Define nome e caminho do arquivo de saída
        output_filename = f"planejamento_{setor}.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        # Salva como Excel sem índice e sem cabeçalho
        df_final.to_excel(output_path, index=False, header=False)
        print(f"  ✔ Arquivo salvo: {output_path}")