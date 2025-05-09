"""
Funções de utilidade para lidar com o calendário de dias úteis.
"""

import pandas as pd
from datetime import datetime
from automation.core.constants import DEFAULT_CALENDARIO_PATH

def carregar_calendario(calendario_path=DEFAULT_CALENDARIO_PATH):
    """
    Carrega o calendário de dias úteis a partir do arquivo CSV.
    
    Args:
        calendario_path (str): Caminho para o arquivo do calendário.
        
    Returns:
        pandas.DataFrame: DataFrame contendo o calendário com datas convertidas.
    """
    try:
        df_cal = pd.read_csv(calendario_path)
        df_cal['DATA'] = pd.to_datetime(df_cal['DATA'], format="%d/%m/%Y")
        return df_cal
    except Exception as e:
        raise IOError(f"Erro ao carregar calendário: {e}")

def obter_proximos_dias_uteis(data_inicio, dias_necessarios, calendario_path=DEFAULT_CALENDARIO_PATH):
    """
    Obtém uma lista de dias úteis a partir da data fornecida (inclusive).
    
    Args:
        data_inicio (datetime ou str): Data inicial para busca.
        dias_necessarios (int): Quantidade de dias úteis a serem obtidos.
        calendario_path (str): Caminho para o arquivo de calendário.
        
    Returns:
        list: Lista de objetos datetime representando os dias úteis.
    """
    df_cal = carregar_calendario(calendario_path)
    
    # Filtra apenas dias úteis
    dias_uteis = df_cal[df_cal['VALOR'] == 'UTIL']['DATA'].tolist()
    
    # Converte data_inicio para datetime se for string
    if isinstance(data_inicio, str):
        data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
    
    # Usa a data exata fornecida
    data_proxima = data_inicio
    
    # Encontra a posição do próximo dia útil a partir da data fornecida (inclusive)
    encontrou = False
    proxima_data_util_idx = 0
    
    for i, data in enumerate(dias_uteis):
        if data >= data_proxima:
            proxima_data_util_idx = i
            encontrou = True
            break
    
    if not encontrou:
        return []
    
    # Retorna os próximos dias úteis necessários
    dias_selecionados = dias_uteis[proxima_data_util_idx:proxima_data_util_idx + dias_necessarios]
    return dias_selecionados