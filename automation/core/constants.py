"""
Constantes utilizadas no sistema de planejamento de produção.
"""

# Ordem dos setores no processo de produção
SETOR_ORDEM = [
    'PCP', 'Separação MP', 'Corte manual', 'Impressão',
    'Estampa', 'Corte laser', 'Distribuição', 'Costura', 'Arremate', 'Embalagem'
]

# Caminhos padrão
DEFAULT_CALENDARIO_PATH = 'data/_CALENDARIO.csv'
DEFAULT_CONFIG_PATH = 'data/_CONFIG.csv'
DEFAULT_EXP_PATH = 'exp/'

import pandas as pd
from automation.core.constants import DEFAULT_CONFIG_PATH

def obter_valor_parametro(parametro: str):
    """
    Consulta o arquivo DEFAULT_CONFIG_PATH e retorna o VALOR do PARAMETRO fornecido.

    Args:
        parametro (str): O nome do parâmetro a ser consultado.

    Returns:
        str: O valor do parâmetro, se encontrado.
        None: Se o parâmetro não for encontrado.
    """
    try:
        # Lê o arquivo de configuração
        df = pd.read_csv(DEFAULT_CONFIG_PATH, encoding='utf-16')
        
        # Verifica se as colunas esperadas estão presentes
        if 'PARAMETRO' not in df.columns or 'VALOR' not in df.columns:
            raise ValueError("O arquivo de configuração não contém as colunas 'PARAMETRO' e 'VALOR'.")
        
        # Busca o parâmetro no DataFrame
        resultado = df.loc[df['PARAMETRO'] == parametro, 'VALOR']
        
        # Retorna o valor se encontrado, caso contrário retorna None
        if not resultado.empty:
            return resultado.iloc[0]
        else:
            print(f"⚠ Parâmetro '{parametro}' não encontrado no arquivo de configuração.")
            return None
    except FileNotFoundError:
        print(f"⚠ Arquivo de configuração não encontrado: {DEFAULT_CONFIG_PATH}")
        return None
    except Exception as e:
        print(f"⚠ Erro ao consultar o parâmetro '{parametro}': {e}")
        return None