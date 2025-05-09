"""
Constantes utilizadas no sistema de planejamento de produção.
"""

# Ordem dos setores no processo de produção
SETOR_ORDEM = [
    'PCP', 'Separação MP', 'Corte manual', 'Impressão',
    'Estampa', 'Corte laser', 'Costura', 'Arremate', 'Embalagem'
]

# Caminhos padrão
DEFAULT_CALENDARIO_PATH = 'data/_CALENDARIO.csv'
DEFAULT_CONFIG_PATH = 'data/_CONFIG.csv'
DEFAULT_EXP_PATH = 'exp/'