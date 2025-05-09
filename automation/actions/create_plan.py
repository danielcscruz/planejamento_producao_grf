from openpyxl import load_workbook
import pandas as pd
from automation.core.production_planner import preencher_producao  

def criar_novo_plano(df_priorizado: pd.DataFrame):
    total = len(df_priorizado)
    arquivo_path = "model/planejamento.xlsx"

    wb = load_workbook(arquivo_path)
    ws = wb.active

    ultimo_dia_list = []
    primeiro_dia_list = []
    delay_list = []


    for index, row in df_priorizado.iterrows():
        linha = 12
        while ws.cell(row=linha, column=1).value:
            linha += 1

        # Pegar valores da tabela
        pedido = row.get("PEDIDO")
        entrega = row.get("ENTREGA")
        cliente = row.get("CLIENTE")
        produto = row.get("PRODUTO")
        quantidade= row.get("QUANTIDADE")
        corte = row.get("TIPO DE CORTE")
        
        # Preenche os dados na planilha
        ws.cell(row=linha, column=1, value=pedido)
        ws.cell(row=linha, column=2, value=entrega)
        ws.cell(row=linha, column=3, value=cliente)
        ws.cell(row=linha, column=4, value=produto)
        ws.cell(row=linha, column=5, value=quantidade)


        if not corte:
            print(f"[Linha {index}] TIPO DE CORTE não especificado. Pulando.")
            ultimo_dia_list.append(None)
            primeiro_dia_list.append(None)
            delay_list.append(None)
            continue

        try:
            salvar = (index == total - 1)  # só salva no último
            print(f"[Linha {index}] Iniciando preenchimento para tipo de corte: {corte}")

            primeiro_dia, ultimo_dia, delay = preencher_producao(
                ws=ws, 
                df_priorizado=df_priorizado, 
                quantidade=quantidade, 
                setor="PCP", 
                linha=linha, 
                calendario_path="data/_CALENDARIO.csv", 
                planilha_path=arquivo_path, 
                workbook=wb, 
                corte=corte, 
                salvar=salvar
            )

            # Padroniza a data para o formato DD/MM/AAAA
            if ultimo_dia:
                ultimo_dia = ultimo_dia.strftime('%d/%m/%Y')
            if primeiro_dia:
                primeiro_dia = primeiro_dia.strftime('%d/%m/%Y')

            primeiro_dia_list.append(primeiro_dia)
            ultimo_dia_list.append(ultimo_dia)
            delay_list.append(delay)
            print(f"[Linha {index}] Preenchimento concluído.\n")
        except Exception as e:
            print(f"[Linha {index}] Erro ao preencher produção: {e}")
            ultimo_dia_list.append(None)
            primeiro_dia_list.append(None)
            delay_list.append(None)

    df_priorizado["PRIMEIRO DIA"] = primeiro_dia_list   
    df_priorizado["ULTIMO DIA"] = ultimo_dia_list
    df_priorizado["DELAY"] = delay_list


    return df_priorizado