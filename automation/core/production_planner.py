"""
Módulo principal para planejamento de produção.
"""

from datetime import datetime, timedelta
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from InquirerPy import inquirer
import csv

from automation.core.constants import SETOR_ORDEM, DEFAULT_CALENDARIO_PATH, DEFAULT_CONFIG_PATH
from automation.core.calendar_utils import obter_proximos_dias_uteis
from automation.core.excel_utils import encontrar_coluna_por_data, calcular_producao_planejada, obter_limite_producao
from automation.core.file_utils import salvar_nova_versao


def obter_delta_dias_estampa(config_path=DEFAULT_CONFIG_PATH):
    """
    Obtém o valor de DELTA_DIAS_ESTAMPA do arquivo de configuração CSV.

    Args:
        config_path (str): Caminho para o arquivo de configuração CSV.

    Returns:
        int: Valor de DELTA_DIAS_ESTAMPA, ou 5 como padrão se não encontrado.
    """
    try:
        with open(config_path, 'r') as config_file:
            reader = csv.DictReader(config_file)
            for row in reader:
                if row['PARAMETRO'] == 'DELTA_DIAS_ESTAMPA':
                    return int(row['VALOR'])  # Converte o valor para inteiro
    except FileNotFoundError:
        print(f"Arquivo de configuração não encontrado: {config_path}. Usando valor padrão (5).")
    except ValueError:
        print(f"Erro ao converter o valor de DELTA_DIAS_ESTAMPA para inteiro. Usando valor padrão (5).")
    except Exception as e:
        print(f"Erro inesperado ao carregar o arquivo de configuração: {e}. Usando valor padrão (5).")
    return 5

def preencher_producao(
        ws: Worksheet,
        df_priorizado: pd.DataFrame,
        quantidade: int, 
        setor: str, 
        linha: int, 
        corte, 
        data_inicio=None, 
        calendario_path=DEFAULT_CALENDARIO_PATH, 
        planilha_path=None, 
        workbook=None, 
        salvar: bool = True,
        priorizar_estampa_terca_quinta: bool = None
    ):
    """
    Preenche a produção a partir do setor especificado, propagando para os demais setores na ordem.
    
    Args:
        ws: Worksheet do openpyxl
        df_priorizado: DataFrame do pandas com os dados de produção
        quantidade: Quantidade a ser produzida
        setor: Setor inicial
        linha: Linha na planilha
        corte: Tipo de corte ('Corte manual' ou 'Corte laser')
        data_inicio: Data de início (opcional)
        calendario_path: Caminho para o arquivo de calendário
        planilha_path: Caminho para o arquivo da planilha (para salvar nova versão)
        workbook: Objeto workbook do openpyxl (necessário para salvar)
        salvar: Se True, salva uma nova versão da planilha
        priorizar_estampa_terca_quinta: Se True, prioriza terças e quintas para o setor Estampa
        
    Returns:
        tuple: (primeiro_dia_usado, ultimo_dia_usado, delay)
    """
    # Verifica se o setor é válido
    if setor not in SETOR_ORDEM:
        raise ValueError(f"Setor '{setor}' não reconhecido. Deve ser um dos: {SETOR_ORDEM}")
    
    # Determina o índice do setor inicial na ordem de processamento
    setor_idx = SETOR_ORDEM.index(setor)
    setores_processar = SETOR_ORDEM[setor_idx:]
    
    # Se data_inicio não for fornecida, usa a data atual
    if data_inicio is None:
        data_inicio = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    elif isinstance(data_inicio, str):
        data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
    
    # Para rastrear a data atual de processamento para cada setor
    data_atual = data_inicio
    ultimo_dia_usado = None
    primeiro_dia_usado = None
    delay = 0
    
    # Por padrão, sempre priorizar terças e quintas para o setor Estampa
    if priorizar_estampa_terca_quinta is None:
        priorizar_estampa_terca_quinta = True
    
    # Obtendo informações do pedido atual para referência
    pedido = ws.cell(row=linha, column=1).value
    entrega = ws.cell(row=linha, column=2).value
    cliente = ws.cell(row=linha, column=3).value
    produto = ws.cell(row=linha, column=4).value

    
    print(f"\n--- Informações do Pedido ---")
    print(f"Pedido: {pedido}")
    print(f"Cliente: {cliente}")
    print(f"Produto: {produto}")
    print(f"Quantidade: {quantidade}")
    print(f"Tipo de Corte: {corte}")
    print(f"------------------------\n")
        
    # Processa cada setor na ordem
    for i, setor_nome in enumerate(setores_processar):
        # Verifica se o tipo de corte deve ser ignorado
        if (setor_nome == 'Corte manual' and corte == 'Corte laser') or \
           (setor_nome == 'Corte laser' and corte == 'Corte manual'):
            continue
        
        # Se não for o primeiro setor, avança para o próximo dia útil
        if i > 0:
            # Avança um dia para garantir que estamos no dia seguinte
            if ultimo_dia_usado:
                # Use o último dia que foi realmente usado pelo setor anterior
                data_proxima = ultimo_dia_usado + timedelta(days=1)
            else:
                # Fallback se não tivermos registro do último dia usado
                data_proxima = data_atual + timedelta(days=1)
                
            # Agora buscamos o próximo dia útil a partir desse dia seguinte
            proximos_dias = obter_proximos_dias_uteis(data_proxima, 1, calendario_path)
            if proximos_dias:
                data_atual = proximos_dias[0]
            else:
                # Avança um dia como fallback
                data_atual = data_proxima
        
        # Resetamos o último dia usado para este novo setor
        setor_ultimo_dia_usado = None
        
        # Caso especial para o setor Estampa
        if setor_nome == 'Estampa':
            
            # Primeiro processa com priorização (terças e quintas)
            resultados = _processar_setor(
                ws, setor_nome, quantidade, linha, data_atual, 
                calendario_path, primeiro_dia_usado, setor_ultimo_dia_usado, delay,
                priorizar_estampa_terca_quinta=True
            )
            
            primeiro_dia_usado = resultados[0] if resultados[0] is not None else primeiro_dia_usado
            setor_ultimo_dia_usado = resultados[1]
            delay = resultados[2]
            

            if isinstance(entrega, str):
                try:
                    data_entrega = datetime.strptime(entrega, "%d/%m/%Y")

                except ValueError:
                    data_entrega = None

            if data_entrega and setor_ultimo_dia_usado:
                
                # Verifica se terminou pelo menos 5 dias antes da data_entrega
                diferenca_dias = (data_entrega - setor_ultimo_dia_usado).days

                print(f"Diferença em dias: {diferenca_dias}")
                
                delta_dias_estampa = obter_delta_dias_estampa()

                if diferenca_dias is not None and diferenca_dias <= delta_dias_estampa:
                    print(f"\nAlerta: O setor Estampa termina em {setor_ultimo_dia_usado.strftime('%d/%m/%Y')}, "
                            f"que é {(data_entrega - setor_ultimo_dia_usado).days} dias antes do Prazo de Entrega "
                            f"({data_entrega.strftime('%d/%m/%Y')}).")
                    
                    # Pergunta ao usuário se deseja refazer sem priorização
                    refazer_sem_priorizar = inquirer.select(
                        message="Deseja refazer o planejamento do setor Estampa sem priorizar terças e quintas?",
                        choices=["Sim", "Não"],
                        default="Sim"
                    ).execute()

                    
                    if refazer_sem_priorizar == "Sim":
                        # Limpa as células preenchidas anteriormente
                        linha_setor = linha + SETOR_ORDEM.index(setor_nome) + 1
                        for col in range(1, ws.max_column + 1):
                            cell = ws.cell(row=linha_setor, column=col)
                            if cell.value is not None:
                                cell.value = None
                        
                        # Refaz o processamento sem priorização
                        resultados = _processar_setor(
                            ws, setor_nome, quantidade, linha, data_atual, 
                            calendario_path, primeiro_dia_usado, None, delay,
                            priorizar_estampa_terca_quinta=False
                        )
                        
                        primeiro_dia_usado = resultados[0] if resultados[0] is not None else primeiro_dia_usado
                        setor_ultimo_dia_usado = resultados[1]
                        delay = resultados[2]
            else:
                print("Erro: Não foi possível calcular a diferença de dias. Verifique os valores de 'entrega' e 'setor_ultimo_dia_usado'.")
        else:
            # Processa o setor atual normalmente
            resultados = _processar_setor(
                ws, setor_nome, quantidade, linha, data_atual, 
                calendario_path, primeiro_dia_usado, setor_ultimo_dia_usado, delay,
                priorizar_estampa_terca_quinta
            )
            
            primeiro_dia_usado = resultados[0] if resultados[0] is not None else primeiro_dia_usado
            setor_ultimo_dia_usado = resultados[1]
            delay = resultados[2]
        
        # Atualiza o último dia usado para o próximo setor
        ultimo_dia_usado = setor_ultimo_dia_usado
    
    # Salvar a planilha em uma nova versão se um caminho e o workbook foram fornecidos
    if planilha_path and workbook and salvar:
        salvar_nova_versao(planilha_path, workbook)

    return primeiro_dia_usado, ultimo_dia_usado, delay

def _processar_setor(
        ws: Worksheet, setor_nome: str, quantidade: int, linha: int, 
        data_atual: datetime, calendario_path: str, 
        primeiro_dia_usado=None, ultimo_dia_usado=None, delay=0,
        priorizar_estampa_terca_quinta=False
    ):
    """
    Processa um setor específico, distribuindo a produção pelos dias úteis.
    
    Args:
        ws: Worksheet do openpyxl
        setor_nome: Nome do setor a ser processado
        quantidade: Quantidade a ser produzida
        linha: Linha base na planilha
        data_atual: Data atual para iniciar o processamento
        calendario_path: Caminho para o arquivo de calendário
        primeiro_dia_usado: Primeiro dia usado em todo o processamento
        ultimo_dia_usado: Último dia usado por este setor
        delay: Atraso acumulado
        priorizar_estampa_terca_quinta: Se True, prioriza terças e quintas para o setor Estampa
        
    Returns:
        tuple: (primeiro_dia_usado, ultimo_dia_usado, delay)
    """
    # Calcula a linha do setor atual (linha do pedido + offset do setor)
    linha_setor = linha + SETOR_ORDEM.index(setor_nome) + 1
    
    # Obtém o limite de produção diário para este setor
    linha_limite = SETOR_ORDEM.index(setor_nome) + 3  # +3 para corresponder à posição na planilha
    valor_limite_max = obter_limite_producao(ws, linha_limite)
    
    qtd_restante = quantidade
    
    # Enquanto houver quantidade restante, continue processando o setor atual
    while qtd_restante > 0:
        # Calcula quantos dias serão necessários para processar toda a quantidade restante
        if valor_limite_max > 0:
            dias_necessarios = (qtd_restante + valor_limite_max - 1) // valor_limite_max  # Arredonda para cima
        else:
            dias_necessarios = 1  # Se não há limite, assume que tudo pode ser feito em um dia
        
        # Limita o número de dias buscados por vez para evitar carregar toda a tabela
        dias_por_iteracao = min(dias_necessarios, 30)  # Busca no máximo 30 dias por vez
        
        # Obtém os próximos dias úteis necessários
        dias_uteis = obter_proximos_dias_uteis(data_atual, dias_por_iteracao, calendario_path)
        
        if not dias_uteis:
            # Se não encontrar dias úteis, saímos do loop deste setor
            break
        
        # Distribui a produção pelos dias úteis
        for dia_util in dias_uteis:
            # Verifica se deve pular este dia para o setor Estampa
            if setor_nome == 'Estampa' and priorizar_estampa_terca_quinta:
                # Verifica se o dia atual é terça-feira (1) ou quinta-feira (3)
                if dia_util.weekday() != 1 and dia_util.weekday() != 3:
                    # Se não for terça ou quinta, incrementa o delay e continua
                    delay += 1
                    continue
            
            # Encontra a coluna correspondente à data
            col = encontrar_coluna_por_data(ws, dia_util)
            
            if col is None:
                continue
            
            # Calcula o valor já planejado para este setor nesta data
            valor_planejado = calcular_producao_planejada(ws, setor_nome, col, linha_setor)
            print(f"\n🔍 Depuração: Dia {dia_util.strftime('%d/%m/%Y')}, Setor: {setor_nome}")
            print(f"  Valor já planejado: {valor_planejado}")

            # Calcula o limite disponível para o dia atual
            valor_limite = max(0, valor_limite_max - valor_planejado)
            print(f"  Limite máximo diário: {valor_limite_max}")
            print(f"  Limite disponível para o dia: {valor_limite}")
            
            # Determina quanto produzir neste dia
            producao_dia = min(valor_limite, qtd_restante)
            print(f"  Quantidade restante: {qtd_restante}")
            print(f"  Quantidade a ser produzida neste dia: {producao_dia}")

            try:
                ws.cell(row=linha_setor, column=col, value=producao_dia)
                print(f"  ✔ Produção registrada: {producao_dia} unidades")

                
                # Registra este dia como o último usado pelo setor atual
                if producao_dia > 0:
                    ultimo_dia_usado = dia_util
                
                if primeiro_dia_usado is None and producao_dia > 0:
                    primeiro_dia_usado = dia_util

                if producao_dia == 0:
                    delay += 1
                    
            except Exception as e:
                print(f"DEBUG - ERRO ao registrar produção: {e}")
            
            qtd_restante -= producao_dia
            print(f"  Quantidade restante após registro: {qtd_restante}")

            
            # Se toda a quantidade foi distribuída, para
            if qtd_restante <= 0:
                break
        
        # Se ainda houver quantidade restante, avança a data atual para continuar
        if qtd_restante > 0:
            if ultimo_dia_usado:
                data_atual = ultimo_dia_usado + timedelta(days=1)
            else:
                data_atual = dias_uteis[-1] + timedelta(days=1)
        
        # Se foram processados todos os dias disponíveis e ainda há qtd_restante,
        # mas não conseguimos alocar nada, devemos sair do loop para evitar um loop infinito
        if all(encontrar_coluna_por_data(ws, dia) is None for dia in dias_uteis) and len(dias_uteis) > 0:
            break
    
    return primeiro_dia_usado, ultimo_dia_usado, delay