import os
import openpyxl
import pandas as pd
from openpyxl import Workbook
from datetime import datetime


def criar_novo_plano(df_formatado, inicio_data):
    # Criando nova planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Plano de Produção"

    # Adicionando títulos principais
    colunas = ['PEDIDO', 'ENTREGA', 'CLIENTE', 'PRODUTO', 'QUANTIDADE']
    for i, col in enumerate(colunas, start=1):
        ws.cell(row=1, column=i, value=col)

    # Colunas adicionais
    ws.cell(row=1, column=6, value='RESTA')
    ws.cell(row=1, column=7, value='SETOR')

    # Adicionando os dados a partir da linha 12
    linha_inicio = 12
    for i, row in df_formatado.iterrows():
        for j, col in enumerate(colunas, start=1):
            ws.cell(row=linha_inicio, column=j, value=row[col])
        linha_inicio += 1

    # 🔽 Adicionando calendário
    calendario_path = os.path.join(os.path.dirname(__file__), "..", "data", "_CALENDARIO.csv")
    calendario_path = os.path.abspath(calendario_path)

    if os.path.exists(calendario_path):
        calendario_df = pd.read_csv(calendario_path, sep=None, engine="python")
        for idx, (_, linha) in enumerate(calendario_df.iterrows(), start=8):  # Coluna H = 8
            semana_abreviada = linha["SEMANA"][:3]  # Ex: "Qua" para "Quarta-feira"
            ws.cell(row=1, column=idx, value=semana_abreviada)
            ws.cell(row=2, column=idx, value=linha["FORMATADO"])
    else:
        print(f"⚠️ Arquivo de calendário '{calendario_path}' não encontrado. Parte do calendário será ignorada.")

    # Criando pasta e salvando
    os.makedirs("exp", exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    arquivo_destino = os.path.join("exp", f"Plano_Producao_{timestamp}.xlsx")
    wb.save(arquivo_destino)

    print(f"📄 Novo plano de produção gerado: {arquivo_destino}")
    return arquivo_destino
