"""
Funções de utilidade para manipulação de arquivos.
"""

from datetime import datetime
from pathlib import Path
from automation.core.constants import DEFAULT_CONFIG_PATH, obter_valor_parametro
from automation.core.excel_utils import obter_carga_producao


def salvar_nova_versao(caminho_original, workbook):
    """
    Salva a planilha em uma nova versão na pasta exp/
    
    Args:
        caminho_original (str): Caminho do arquivo original.
        workbook: Objeto workbook do openpyxl.
        
    Returns:
        str or None: Caminho da nova versão salva ou None em caso de erro.
    """
    try:
        # Garantir que a pasta exp/ existe
        pasta_exp = Path("exp")
        pasta_exp.mkdir(exist_ok=True)
        
        # Obter o nome do arquivo original
        base_nome = Path(caminho_original).stem
        extensao = Path(caminho_original).suffix
        
        carga_prod = obter_carga_producao(config_path=DEFAULT_CONFIG_PATH)
        priorizar_estampa_value = obter_valor_parametro('PRIORIDADE_ESTAMPA')
        if priorizar_estampa_value:
            priorizado = "P"
        else:
            priorizado = "N"



        # Gerar um nome de arquivo com timestamp
        timestamp = datetime.now().strftime("%Y_%m_%d__%H_%M_%S")
        novo_nome = f"{base_nome}_c{carga_prod}_{priorizado}__{timestamp}_{extensao}"
        novo_caminho = pasta_exp / novo_nome
        
        # Salvar a planilha no novo caminho
        workbook.save(str(novo_caminho))
        return str(novo_caminho)
    except Exception as e:
        return None