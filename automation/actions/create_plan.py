from openpyxl import load_workbook
import pandas as pd
import os
from tabulate import tabulate
from InquirerPy import inquirer

from automation.core.production_planner import preencher_producao  
from automation.core.excel_utils import atualizar_limites_maximos, atualizar_celulas_limite
from automation.core.constants import DEFAULT_CONFIG_PATH
from automation.validators.start_date_validator import validar_data_input
from automation.ui.sector_selector import selecionar_setor_inicio

def criar_novo_plano(df_priorizado: pd.DataFrame):
    total = len(df_priorizado)
    arquivo_path = "model/planejamento.xlsx"

    wb = load_workbook(arquivo_path)
    ws = wb.active

    ultimo_dia_list = []
    primeiro_dia_list = []
    delay_list = []

    # Gera a lista de limites máximos
    max_list = atualizar_limites_maximos(config_path=DEFAULT_CONFIG_PATH)

    # Atualiza as células [E3:E11] na planilha
    atualizar_celulas_limite(ws, max_list)

    while True:
        try:
            inicio_plano_str = input("\n📅 Data de Inicio do Plano (DD/MM/AAAA): ").strip()
            inicio_plano = validar_data_input(inicio_plano_str)
            break  # Sai do loop se a data for válida
        except ValueError as e:
            print(e)  # Exibe a mensagem de erro da exceção



    while True:
        escolher_setor = inquirer.select(
            message="\nVocê deseja escolher por qual setor deseja iniciar o plano?",
            choices=["Não - Iniciar todos por PCP [Padrão]", "Sim - Desejo escolher por qual setor irá iniciar cada produção"],
            default="Não - Iniciar todos por PCP [Padrão]"
        ).execute()
        if escolher_setor == "Não - Iniciar todos por PCP [Padrão]":
            df_priorizado["SETOR"] = "PCP"

        if escolher_setor == "Sim - Desejo escolher por qual setor irá iniciar cada produção":
            setor_lista = selecionar_setor_inicio(df_priorizado)

            df_priorizado["SETOR"] = df_priorizado["PEDIDO"].map(
                lambda pedido: setor_lista[pedido]['setor']
            )
        os.system('cls' if os.name == 'nt' else 'clear')
        print(tabulate(df_priorizado, headers='keys', tablefmt='grid', showindex=False))

        # Pergunta ao usuário se deseja prosseguir ou refazer as escolhas
        confirmar = inquirer.select(
            message="Deseja prosseguir com essas escolhas?",
            choices =["Sim - Prosseguir","Não - Refazer as minhas escolhas"],
            default="Sim - Prosseguir"
        ).execute()

        if confirmar == "Sim - Prosseguir":
            break  # Sai do loop se o usuário confirmar
        else:
            print("Refazendo as escolhas para os setores...\n")



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
        setor=row.get("SETOR")
        
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
                setor=setor, 
                linha=linha, 
                calendario_path="data/_CALENDARIO.csv", 
                planilha_path=arquivo_path, 
                workbook=wb,
                data_inicio=inicio_plano,
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